// app/terminalApi/graphProxy.js

/**
 * Executes Microsoft Graph requests on behalf of a local client (the Go TUI
 * in tui/) using the token borrowed from the Teams session, so the client
 * needs no sign-in of its own.
 *
 * Requests the token cannot serve (no Chat.Read / Chat.ReadWrite) are answered
 * from the Teams chat service in Graph shape instead:
 *   GET   /me/chats                          list chats
 *   GET   /chats/{id}                        one chat
 *   GET   /chats/{id}/messages               messages, paged via $skiptoken
 *   POST  /chats/{id}/markChatReadForUser    move the read horizon
 *   PATCH /chats/{id}/messages/{mid}         edit (also the client's "delete")
 * Everything else is forwarded to Graph unchanged.
 *
 * Only https://graph.microsoft.com URLs are accepted, so the token is never
 * sent anywhere else. Bodies travel base64-encoded so binary content
 * (images, uploads) survives the JSON socket protocol.
 */

const GRAPH_ORIGIN = "https://graph.microsoft.com";
const REQUEST_TIMEOUT_MS = 60000;

const CHAT_ROUTES = [
  { method: "GET", re: /^\/(?:v1\.0|beta)\/me\/chats$/, op: "listChats" },
  { method: "GET", re: /^\/(?:v1\.0|beta)\/chats\/([^/]+)$/, op: "getChat" },
  { method: "GET", re: /^\/(?:v1\.0|beta)\/chats\/([^/]+)\/messages$/, op: "listMessages" },
  { method: "POST", re: /^\/(?:v1\.0|beta)\/chats\/([^/]+)\/markChatReadForUser$/, op: "markRead" },
  { method: "PATCH", re: /^\/(?:v1\.0|beta)\/chats\/([^/]+)\/messages\/([^/]+)$/, op: "editMessage" },
];

/** Match a Graph request to a chat-service route. Pure; exported for testing. */
function matchChatRoute(method, url) {
  const parsed = new URL(url);
  const path = decodeURIComponent(parsed.pathname);
  for (const route of CHAT_ROUTES) {
    if (route.method !== method.toUpperCase()) continue;
    const m = path.match(route.re);
    if (m) return { op: route.op, args: m.slice(1), query: parsed.searchParams };
  }
  return null;
}

// tmux does not join skin-tone modifiers onto the preceding emoji, so it
// draws e.g. 👍🏽 four columns wide while TUI layout libraries count two, and
// every such line wraps and corrupts the screen. Drop the modifiers; the
// emoji then shows in its default tone.
const SKIN_TONES = /[\u{1F3FB}-\u{1F3FF}]/gu;

/** Remove skin-tone modifiers from JSON text. Pure; exported for testing. */
function stripSkinTones(text) {
  return text.replace(SKIN_TONES, "");
}

const json = (status, data) => ({
  status,
  contentType: "application/json",
  body: data === undefined ? "" : Buffer.from(stripSkinTones(JSON.stringify(data))).toString("base64"),
});

const graphError = (status, message) => json(status, { error: { code: status === 403 ? "Forbidden" : "BadRequest", message } });

class GraphProxy {
  #graph;
  #chats;

  constructor(graphApiClient, chatService) {
    this.#graph = graphApiClient;
    this.#chats = chatService;
  }

  /**
   * @param {{method: string, url: string, contentType?: string, body?: string}} req
   *   body is base64
   * @returns {Promise<{status: number, contentType: string, body: string}>} body is base64
   */
  async handle(req) {
    const method = String(req.method || "GET").toUpperCase();
    let url;
    try {
      url = new URL(req.url);
    } catch {
      return graphError(400, "Invalid URL");
    }
    if (url.origin !== GRAPH_ORIGIN) return graphError(400, "Only graph.microsoft.com requests are proxied");

    const route = matchChatRoute(method, url.href);
    if (route) {
      try {
        return await this.#chatRoute(route, req);
      } catch (error) {
        return graphError(error.status || 502, error.message);
      }
    }
    return await this.#forward(method, url.href, req);
  }

  async #chatRoute({ op, args, query }, req) {
    const expandMembers = /\bmembers\b/.test(query.get("$expand") || "");
    switch (op) {
      case "listChats":
        return json(200, await this.#chats.listChats({ top: query.get("$top") || 50, expandMembers }));
      case "getChat":
        return json(200, await this.#chats.getChat(args[0], { expandMembers }));
      case "listMessages":
        return json(
          200,
          await this.#chats.listMessages(args[0], { top: query.get("$top") || 50, skipToken: query.get("$skiptoken") })
        );
      case "markRead":
        await this.#chats.markRead(args[0]);
        return json(204);
      case "editMessage": {
        const payload = req.body ? JSON.parse(Buffer.from(req.body, "base64").toString("utf8")) : {};
        const content = payload?.body?.content;
        if (typeof content !== "string") return graphError(400, "Only body edits are supported");
        // Graph clients send plain text or HTML; chatsvc stores HTML.
        const html = payload.body.contentType === "html" ? content : content.replaceAll("&", "&amp;").replaceAll("<", "&lt;");
        await this.#chats.editMessage(args[0], args[1], html);
        return json(204);
      }
      default:
        return graphError(400, `Unsupported chat route ${op}`);
    }
  }

  async #forward(method, url, req) {
    const token = await this.#graph.acquireToken();
    if (!token?.success || !token.token) return graphError(401, "Teams session has no Graph token (signed out?)");
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
    try {
      const res = await fetch(url, {
        method,
        headers: {
          Authorization: `Bearer ${token.token}`,
          ...(req.contentType ? { "Content-Type": req.contentType } : {}),
          ...(req.range ? { Range: req.range } : {}),
        },
        ...(req.body ? { body: Buffer.from(req.body, "base64") } : {}),
        signal: controller.signal,
      });
      let body = Buffer.from(await res.arrayBuffer());
      const contentType = res.headers.get("content-type") || "";
      if (contentType.includes("json")) body = Buffer.from(stripSkinTones(body.toString("utf8")));
      return {
        status: res.status,
        contentType,
        body: body.toString("base64"),
      };
    } catch (error) {
      return graphError(502, error.name === "AbortError" ? "Graph request timed out" : error.message);
    } finally {
      clearTimeout(timeout);
    }
  }
}

module.exports = { GraphProxy, matchChatRoute, stripSkinTones };
