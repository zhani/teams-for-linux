// app/terminalApi/chatService.js

/**
 * Read-only client for the Teams chat service (IC3 "chatsvc"), the internal
 * API the Teams web app itself uses for chats.
 *
 * Why not Graph: the token borrowed from Teams lacks Chat.Read/Chat.ReadBasic,
 * so Graph refuses /me/chats and /chats/{id}/messages. Sending still goes
 * through Graph (ChatMessage.Send), because ADR-015 found chatsvc sends from
 * outside the Teams client unreliable.
 *
 * Results are normalised to the Graph chat/chatMessage shape so callers can
 * treat both sources alike. This is an undocumented API and may change.
 */

const IC3_RESOURCE = "https://ic3.teams.office.com";
const SKYPE_RESOURCE = "https://api.spaces.skype.com";
const AUTHZ_URL = "https://teams.microsoft.com/api/authsvc/v1.0/authz";
const REQUEST_TIMEOUT_MS = 30000;

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

/** Extract the AAD object id from an MRI such as ".../contacts/8:orgid:<guid>". Pure. */
function mriToUserId(mri) {
  const m = String(mri || "").match(/8:orgid:([0-9a-f-]{36})/i);
  return m ? m[1].toLowerCase() : null;
}

/** Whether a chatsvc conversation is a chat (not a channel, system feed, etc.). Pure. */
function isChatConversation(conv) {
  const id = conv?.id || "";
  if (!id.startsWith("19:")) return false; // 48:notifications, 48:calllogs, ...
  const type = (conv.threadProperties?.threadType || "").toLowerCase();
  return !["space", "topic"].includes(type); // channels
}

/** Map a chatsvc message to the Graph chatMessage shape. Pure. */
function toGraphMessage(msg) {
  const type = String(msg?.messagetype || "");
  // RichText/Media_* are meeting recordings and call cards, Event/* and
  // ThreadActivity/* are membership and call events: none are user text.
  const isUser = /^(Text|RichText|RichText\/Html)$/i.test(type);
  return {
    id: msg?.id,
    createdDateTime: msg?.originalarrivaltime || msg?.composetime || null,
    messageType: isUser ? "message" : "systemEventMessage",
    deletedDateTime: msg?.properties?.deletetime ? new Date(Number(msg.properties.deletetime)).toISOString() : null,
    from: isUser ? { user: { displayName: msg?.imdisplayname || "unknown", id: mriToUserId(msg?.from) } } : null,
    body: { contentType: /^RichText/i.test(type) ? "html" : "text", content: msg?.content || "" },
  };
}

class ChatServiceClient {
  #graph;
  #base = null;
  #auth = null; // { headers, expiresAt }
  #titleCache = new Map();

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

  async #get(pathAndQuery) {
    if (!this.#base) await this.#connect();
    let res = await this.#fetchJson(`${this.#base}${pathAndQuery}`, { headers: this.#auth.headers });
    if (res.status === 401 || res.status === 403) {
      // Token expired, or bearer auth not accepted here: reconnect once, using
      // the skypetoken route if bearer was just rejected.
      await this.#connect(this.#auth.mode === "bearer");
      res = await this.#fetchJson(`${this.#base}${pathAndQuery}`, { headers: this.#auth.headers });
    }
    if (!res.ok) {
      const detail = res.data?.message || res.data?.errorCode || "";
      throw new Error(`Chat service request failed (HTTP ${res.status})${detail ? `: ${detail}` : ""}`);
    }
    return res.data;
  }

  // 1:1 and unnamed group chats have no topic; name them after the other
  // members, which Graph exposes via ChatMember.Read.
  async #title(conv, myUserId) {
    const topic = conv.threadProperties?.topic;
    if (topic) return topic;
    if (this.#titleCache.has(conv.id)) return this.#titleCache.get(conv.id);
    let title = null;
    const members = await this.#graph.makeRequest(`/chats/${encodeURIComponent(conv.id)}/members`).catch(() => null);
    if (members?.success) {
      const others = (members.data?.value || [])
        .filter((m) => (m.userId || "").toLowerCase() !== myUserId)
        .map((m) => m.displayName)
        .filter(Boolean);
      title = others.join(", ") || null;
    }
    title = title || conv.lastMessage?.imdisplayname || "chat";
    this.#titleCache.set(conv.id, title);
    return title;
  }

  async #myUserId() {
    if (this.myUserId) return this.myUserId;
    const me = await this.#graph.makeRequest("/me?$select=id");
    this.myUserId = me?.success ? String(me.data?.id || "").toLowerCase() : null;
    return this.myUserId;
  }

  /** Recent chats, most recently active first, in the Graph chat shape. */
  async getChats(top = 30) {
    const limit = Math.min(Number(top) || 30, 100);
    // Over-fetch: the conversation list also carries channels and system feeds.
    const data = await this.#get(`/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=${Math.min(limit * 3, 200)}`);
    const myUserId = await this.#myUserId();
    const chats = (data?.conversations || [])
      .filter(isChatConversation)
      .filter((c) => c.lastMessage?.id)
      .sort((a, b) => String(b.lastMessage.originalarrivaltime || b.lastMessage.composetime || "")
        .localeCompare(String(a.lastMessage.originalarrivaltime || a.lastMessage.composetime || "")))
      .slice(0, limit);
    const value = [];
    for (const conv of chats) {
      value.push({
        id: conv.id,
        topic: await this.#title(conv, myUserId),
        chatType: conv.threadProperties?.threadType || "oneOnOne",
        lastMessagePreview: toGraphMessage(conv.lastMessage),
      });
    }
    return { success: true, data: { value } };
  }

  /** Recent messages of a chat, newest first, in the Graph chatMessage shape. */
  async getChatMessages(chatId, top = 20) {
    const size = Math.min(Number(top) || 20, 200);
    const data = await this.#get(
      `/v1/users/ME/conversations/${encodeURIComponent(chatId)}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=${size}&startTime=1`
    );
    return { success: true, data: { value: (data?.messages || []).map(toGraphMessage) } };
  }
}

module.exports = { ChatServiceClient, toGraphMessage, isChatConversation, mriToUserId, DISCOVER_BASE_SCRIPT };
