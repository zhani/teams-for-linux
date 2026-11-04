// app/terminalApi/chatService.js

/**
 * Client for the Teams chat service (IC3 "chatsvc"), the internal API the
 * Teams web app itself uses for chats.
 *
 * Why not Graph: the token borrowed from Teams lacks Chat.Read/Chat.ReadBasic/
 * Chat.ReadWrite, so Graph refuses listing chats, reading chat messages,
 * marking chats read and editing chat messages. New messages are still sent
 * through Graph (ChatMessage.Send), because ADR-015 found chatsvc sends from
 * outside the Teams client unreliable.
 *
 * Results are mapped to the Graph chat/chatMessage shape (graphShape.js) so
 * callers can treat both sources alike. This is an undocumented API and may
 * change.
 */

const {
  isChatConversation,
  isControlMessage,
  mriToUserId,
  toGraphChat,
  toGraphMessage,
} = require("./graphShape");

const IC3_RESOURCE = "https://ic3.teams.office.com";
const SKYPE_RESOURCE = "https://api.spaces.skype.com";
const AUTHZ_URL = "https://teams.microsoft.com/api/authsvc/v1.0/authz";
const REQUEST_TIMEOUT_MS = 30000;
const MEMBERS_TTL_MS = 10 * 60 * 1000;
const MEMBER_FETCH_CONCURRENCY = 8;
const MESSAGES_VIEW = "msnp24Equivalent|supportsMessageProperties";

// Runs in the Teams page. Finds the chat service base URL from requests Teams
// has already made, e.g. https://teams.microsoft.com/api/chatsvc/emea.
const DISCOVER_BASE_SCRIPT = String.raw`(() => {
  const names = performance.getEntriesByType('resource').map((e) => e.name);
  for (let i = names.length - 1; i >= 0; i--) {
    const m = names[i].match(/^(https:\/\/[^/]+(?:\/api\/chatsvc\/[a-z0-9-]+)?)\/v1\/users\/ME\//i);
    if (m) return m[1];
  }
  return null;
})()`;

function tokenScript(resource) {
  return `(async () => {
    const h = window.teamsForLinuxReactHandler;
    if (!h || typeof h.acquireToken !== 'function') return { success: false, error: 'ReactHandler not available' };
    return await h.acquireToken(${JSON.stringify(resource)});
  })()`;
}

const activityTime = (conv) =>
  String(conv.lastMessage?.originalarrivaltime || conv.lastMessage?.composetime || "");

async function mapLimited(items, limit, fn) {
  const results = new Array(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const i = next++;
      results[i] = await fn(items[i]);
    }
  });
  await Promise.all(workers);
  return results;
}

class ChatServiceClient {
  #graph;
  #base = null;
  #auth = null; // { headers, mode }
  #membersCache = new Map(); // chatId -> { at, members }
  #myUserId = null;

  constructor(graphApiClient) {
    this.#graph = graphApiClient;
  }

  async #run(script) {
    const wc = this.#graph.mainWindow?.webContents;
    if (!wc) throw new Error("Main window not initialized");
    return await wc.executeJavaScript(script);
  }

  async #token(resource) {
    const result = await this.#run(tokenScript(resource));
    if (!result?.success || !result.token) {
      throw new Error(`Could not get chat service token: ${result?.error || "unknown error"}`);
    }
    return result.token;
  }

  async #fetchJson(url, options = {}) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
    try {
      const response = await fetch(url, { ...options, signal: controller.signal });
      const text = await response.text();
      let data = null;
      try {
        data = text ? JSON.parse(text) : null;
      } catch {
        // Non-JSON error bodies are reported by status below.
      }
      return { ok: response.ok, status: response.status, data };
    } finally {
      clearTimeout(timeout);
    }
  }

  // Preferred: the IC3 bearer token against the base URL Teams itself uses.
  // Fallback: exchange a Skype-audience token at authsvc for a skypetoken and
  // the regional chat service URL, which older chatsvc deployments require.
  async #connect(forceFallback = false) {
    if (!forceFallback) {
      const base = await this.#run(DISCOVER_BASE_SCRIPT).catch(() => null);
      if (base) {
        this.#base = base;
        this.#auth = { headers: { Authorization: `Bearer ${await this.#token(IC3_RESOURCE)}` }, mode: "bearer" };
        return;
      }
    }
    const skype = await this.#token(SKYPE_RESOURCE);
    const res = await this.#fetchJson(AUTHZ_URL, {
      method: "POST",
      headers: { Authorization: `Bearer ${skype}`, "Content-Type": "application/json" },
    });
    const skypeToken = res.data?.tokens?.skypeToken;
    const chatService = res.data?.regionGtms?.chatServiceAfd || res.data?.regionGtms?.chatService;
    if (!res.ok || !skypeToken || !chatService) {
      throw new Error(`Chat service discovery failed (HTTP ${res.status})`);
    }
    this.#base = chatService.replace(/\/$/, "");
    this.#auth = {
      headers: { Authentication: `skypetoken=${skypeToken}`, "X-Skypetoken": skypeToken },
      mode: "skypetoken",
    };
  }

  /**
   * Request against the chat service. `pathOrUrl` is a path under the base,
   * or an absolute URL that must itself be under the base (paging links).
   */
  async #request(method, pathOrUrl, body) {
    if (!this.#base) await this.#connect();
    const send = () => {
      const url = pathOrUrl.startsWith("https://") ? pathOrUrl : `${this.#base}${pathOrUrl}`;
      if (!url.startsWith(`${this.#base}/`)) throw new Error("Refusing chat service URL outside the service base");
      return this.#fetchJson(url, {
        method,
        headers: { ...this.#auth.headers, ...(body ? { "Content-Type": "application/json" } : {}) },
        ...(body ? { body: JSON.stringify(body) } : {}),
      });
    };
    let res = await send();
    if (res.status === 401 || res.status === 403) {
      // Token expired, or bearer auth not accepted here: reconnect once, using
      // the skypetoken route if bearer was just rejected.
      await this.#connect(this.#auth.mode === "bearer");
      res = await send();
    }
    if (!res.ok) {
      const detail = res.data?.message || res.data?.errorCode || "";
      const error = new Error(`Chat service request failed (HTTP ${res.status})${detail ? `: ${detail}` : ""}`);
      error.status = res.status;
      throw error;
    }
    return res.data;
  }

  async myUserId() {
    if (this.#myUserId) return this.#myUserId;
    const me = await this.#graph.makeRequest("/me?$select=id");
    this.#myUserId = me?.success ? String(me.data?.id || "").toLowerCase() : null;
    return this.#myUserId;
  }

  /** Graph member list of a chat (ChatMember.Read is in the token), cached. */
  async members(chatId) {
    const hit = this.#membersCache.get(chatId);
    if (hit && Date.now() - hit.at < MEMBERS_TTL_MS) return hit.members;
    const result = await this.#graph.makeRequest(`/chats/${encodeURIComponent(chatId)}/members`).catch(() => null);
    const members = result?.success ? result.data?.value || [] : [];
    this.#membersCache.set(chatId, { at: Date.now(), members });
    return members;
  }

  async #conversations(limit) {
    // Over-fetch: the conversation list also carries channels and system feeds.
    const data = await this.#request("GET", `/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=${Math.min(limit * 3, 300)}`);
    return (data?.conversations || [])
      .filter(isChatConversation)
      .filter((c) => c.lastMessage?.id)
      .sort((a, b) => activityTime(b).localeCompare(activityTime(a)))
      .slice(0, limit);
  }

  // --- Graph-shaped operations, used by the Graph proxy (graphProxy.js) ---

  /** GET /me/chats: most recently active chats, optionally with members. */
  async listChats({ top = 50, expandMembers = false } = {}) {
    const convs = await this.#conversations(Math.min(Number(top) || 50, 100));
    const members = expandMembers
      ? await mapLimited(convs, MEMBER_FETCH_CONCURRENCY, (c) => this.members(c.id))
      : [];
    return { value: convs.map((c, i) => toGraphChat(c, expandMembers ? members[i] : undefined)) };
  }

  /** GET /chats/{id} */
  async getChat(chatId, { expandMembers = false } = {}) {
    const conv = await this.#request("GET", `/v1/users/ME/conversations/${encodeURIComponent(chatId)}?view=msnp24Equivalent`);
    return toGraphChat(conv, expandMembers ? await this.members(chatId) : undefined);
  }

  /**
   * GET /chats/{id}/messages, newest first. Paging follows chatsvc's
   * backwardLink, carried opaquely in Graph's $skiptoken.
   */
  async listMessages(chatId, { top = 50, skipToken = null } = {}) {
    const size = Math.min(Number(top) || 50, 200);
    const path = skipToken
      ? Buffer.from(skipToken, "base64url").toString("utf8")
      : `/v1/users/ME/conversations/${encodeURIComponent(chatId)}/messages?view=${MESSAGES_VIEW}&pageSize=${size}&startTime=1`;
    const data = await this.#request("GET", path);
    const messages = data?.messages || [];
    const value = messages.filter((m) => !isControlMessage(m)).map(toGraphMessage);
    const back = data?._metadata?.backwardLink;
    const result = { value };
    if (back && messages.length) {
      result["@odata.nextLink"] =
        `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(chatId)}/messages?$skiptoken=${Buffer.from(back).toString("base64url")}`;
    }
    return result;
  }

  /** POST /chats/{id}/markChatReadForUser: move the read horizon to the latest message. */
  async markRead(chatId) {
    const conv = await this.#request("GET", `/v1/users/ME/conversations/${encodeURIComponent(chatId)}?view=msnp24Equivalent`);
    const last = conv?.lastMessage;
    if (!last?.id) return;
    const horizon = `${last.id};${Date.now()};${last.clientmessageid || last.id}`;
    await this.#request(
      "PUT",
      `/v1/users/ME/conversations/${encodeURIComponent(chatId)}/properties?name=consumptionhorizon`,
      { consumptionhorizon: horizon }
    );
  }

  /** PATCH /chats/{id}/messages/{mid}: replace a message's HTML body. */
  async editMessage(chatId, messageId, html) {
    await this.#request(
      "PUT",
      `/v1/users/ME/conversations/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`,
      { content: html, messagetype: "RichText/Html", contenttype: "text" }
    );
  }

  // --- CLI helpers (cli.js `chats` / `chat`): titles computed server-side ---

  /** Recent chats with a human title in `topic`, most recently active first. */
  async getChats(top = 30) {
    const convs = await this.#conversations(Math.min(Number(top) || 30, 100));
    const me = await this.myUserId();
    const value = [];
    for (const conv of convs) {
      const chat = toGraphChat(conv);
      if (!chat.topic) {
        const others = (await this.members(conv.id))
          .filter((m) => (m.userId || "").toLowerCase() !== me)
          .map((m) => m.displayName)
          .filter(Boolean);
        chat.topic = others.join(", ") || conv.lastMessage?.imdisplayname || "chat";
      }
      value.push(chat);
    }
    return { success: true, data: { value } };
  }

  /** Recent messages of a chat, newest first. */
  async getChatMessages(chatId, top = 20) {
    return { success: true, data: await this.listMessages(chatId, { top }) };
  }
}

module.exports = { ChatServiceClient, toGraphMessage, isChatConversation, mriToUserId, DISCOVER_BASE_SCRIPT };
