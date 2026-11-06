// app/terminalApi/index.js

/**
 * Local terminal API: lets a CLI on the same machine read and post Teams
 * channel and chat messages through the borrowed Graph token, so the app can
 * run headless (e.g. under Xvfb) and be used over SSH.
 *
 * Security:
 * - Listens only on a Unix socket under $XDG_RUNTIME_DIR (a 0700 per-user
 *   directory), chmod'ed to 0600, so only the same user can connect.
 * - Never logs message bodies, names, team/channel names or URLs.
 * - Surfaces the MFA number-matching code so the user can approve the sign-in
 *   on their phone; it does not bypass MFA.
 */

const net = require("node:net");
const fs = require("node:fs");
const path = require("node:path");
const { app, webContents } = require("electron");
const { isLoginUrl } = require("../ssoPasswordPrefill");
const { ChatServiceClient } = require("./chatService");
const { buildLoginStepScript } = require("./loginDriver");
const { GraphProxy } = require("./graphProxy");

const SOCKET_NAME = "teams-for-linux-terminal.sock";
const MFA_WAIT_MS = 120000;
// Large enough for proxied file uploads and pasted images (base64).
const MAX_LINE_BYTES = 64 * 1024 * 1024;
const LOGIN_TIMEOUT_MS = 5 * 60 * 1000;
const LOGIN_STEP_MS = 1500;
// A fresh start may still be loading Teams; wait this long for a sign-in
// page before concluding the session is already signed in.
const LOGIN_APPEAR_MS = 20000;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Runs in the login page's frame. Resolves with the number-matching code once
// Microsoft renders it, or null after the timeout.
const MFA_CODE_SCRIPT = `(() => new Promise((resolve) => {
  const started = Date.now();
  const read = () => {
    const el = document.querySelector('#idRichContext_DisplaySign, [data-testid*="DisplaySign" i]');
    const text = el && (el.textContent || '').trim();
    return text && /^\\d{1,3}$/.test(text) ? text : null;
  };
  const tick = () => {
    const code = read();
    if (code) return resolve(code);
    if (Date.now() - started > ${MFA_WAIT_MS}) return resolve(null);
    setTimeout(tick, 1000);
  };
  tick();
}))()`;

function defaultSocketPath() {
  // Lets a second instance (e.g. a throwaway test profile) use its own socket.
  if (process.env.TEAMS_CLI_SOCKET) return process.env.TEAMS_CLI_SOCKET;
  const dir = process.env.XDG_RUNTIME_DIR || app.getPath("userData");
  return path.join(dir, SOCKET_NAME);
}

/**
 * Resolve a channel spec ("Team/Channel" by name, or "teamId/channelId") to
 * ids. Exact case-insensitive name match wins over prefix match. Pure;
 * exported for testing.
 */
function matchByName(items, query, field = "displayName") {
  const q = String(query).toLowerCase();
  const byId = items.find((i) => i.id === query);
  if (byId) return byId;
  const exact = items.filter((i) => (i[field] || "").toLowerCase() === q);
  if (exact.length === 1) return exact[0];
  const prefix = items.filter((i) => (i[field] || "").toLowerCase().startsWith(q));
  if (prefix.length === 1) return prefix[0];
  if (exact.length + prefix.length > 1) {
    throw new Error(`Ambiguous name "${query}": ${[...new Set([...exact, ...prefix])].map((i) => i[field]).join(", ")}`);
  }
  throw new Error(`No match for "${query}"`);
}

class TerminalApiServer {
  #config;
  #graph;
  #chats;
  #proxy;
  #server = null;
  #socketPath;
  #subscribers = new Set();
  #teamsCache = null;
  #channelsCache = new Map();
  #loginContents = new Map(); // webContents id -> webContents showing sign-in
  #loginRunning = false;
  #lastMfa = { code: null, at: 0 };
  #pendingMfaFrames = new WeakSet();

  constructor(config, graphApiClient, { socketPath, chatService } = {}) {
    this.#config = config;
    this.#graph = graphApiClient;
    // Graph can send chat messages but not read them (no Chat.Read scope),
    // so reads go through the Teams chat service.
    this.#chats = chatService || new ChatServiceClient(graphApiClient);
    this.#proxy = new GraphProxy(graphApiClient, this.#chats);
    this.#socketPath = socketPath || defaultSocketPath();
  }

  get socketPath() {
    return this.#socketPath;
  }

  async start() {
    if (await this.#socketInUse()) {
      throw new Error("Terminal API socket is in use by another instance; set TEAMS_CLI_SOCKET to use a different one");
    }
    this.#removeStaleSocket();
    this.#server = net.createServer((socket) => this.#onConnection(socket));
    await new Promise((resolve, reject) => {
      this.#server.once("error", reject);
      this.#server.listen(this.#socketPath, () => {
        this.#server.off("error", reject);
        resolve();
      });
    });
    fs.chmodSync(this.#socketPath, 0o600);
    console.info("[TERMINAL_API] Listening on local socket");
  }

  /** Watch login pages for the MFA number-matching code. Main-process only. */
  watchLogins() {
    const extraHosts = this.#config.auth?.webLogin?.extraHosts || [];
    const attach = (wc) => {
      const wcId = wc.id;
      const check = () => {
        let frames;
        try {
          frames = wc.mainFrame.framesInSubtree;
        } catch {
          return;
        }
        const loginFrames = frames.filter((f) => {
          try {
            return isLoginUrl(f.url, extraHosts);
          } catch {
            return false;
          }
        });
        // Teams keeps hidden MSAL silent-renew iframes on the login host, so
        // only a top-level login page counts as "sign-in needed".
        const topIsLogin = loginFrames.includes(wc.mainFrame);
        if (topIsLogin && !this.#loginContents.has(wcId)) {
          this.#loginContents.set(wcId, wc);
          console.info("[TERMINAL_API] Sign-in page shown");
          this.#broadcast({ event: "login-required" });
        } else if (!topIsLogin) {
          this.#leaveLogin(wcId);
        }
        for (const frame of loginFrames) this.#watchMfaFrame(frame);
      };
      wc.on("dom-ready", check);
      wc.on("did-navigate", check);
      wc.on("did-frame-navigate", check);
      wc.once("destroyed", () => this.#leaveLogin(wcId));
      check();
    };
    // The main window already exists when this module starts, so cover it too.
    for (const wc of webContents.getAllWebContents()) attach(wc);
    app.on("web-contents-created", (_event, wc) => attach(wc));
  }

  // Both the page watcher and the login loop see the same code repeatedly
  // while it waits for approval; report each code once.
  #reportMfaCode(code) {
    const now = Date.now();
    if (code === this.#lastMfa.code && now - this.#lastMfa.at < MFA_WAIT_MS) return;
    this.#lastMfa = { code, at: now };
    console.info(`[TERMINAL_API] Approve sign-in on your phone: ${code}`);
    this.#broadcast({ event: "mfa-code", code });
  }

  #leaveLogin(wcId) {
    if (!this.#loginContents.delete(wcId)) return;
    if (this.#loginContents.size === 0) {
      console.info("[TERMINAL_API] Sign-in page closed");
      this.#broadcast({ event: "login-complete" });
    }
  }

  async #watchMfaFrame(frame) {
    if (this.#pendingMfaFrames.has(frame)) return;
    this.#pendingMfaFrames.add(frame);
    try {
      const code = await frame.executeJavaScript(MFA_CODE_SCRIPT, true);
      if (code) this.#reportMfaCode(code);
    } catch {
      // Frame navigated away or was destroyed; the next navigation re-checks.
    } finally {
      this.#pendingMfaFrames.delete(frame);
    }
  }

  dispose() {
    for (const socket of this.#subscribers) socket.destroy();
    this.#subscribers.clear();
    this.#server?.close();
    this.#server = null;
    this.#removeStaleSocket();
  }

  #socketInUse() {
    return new Promise((resolve) => {
      const probe = net.createConnection(this.#socketPath);
      probe.once("connect", () => {
        probe.destroy();
        resolve(true);
      });
      probe.once("error", () => resolve(false));
    });
  }

  #removeStaleSocket() {
    try {
      if (fs.lstatSync(this.#socketPath).isSocket()) fs.unlinkSync(this.#socketPath);
    } catch {
      // Not there: nothing to clean.
    }
  }

  #broadcast(message) {
    const line = JSON.stringify(message) + "\n";
    for (const socket of this.#subscribers) socket.write(line);
  }

  #onConnection(socket) {
    let buffer = "";
    socket.setEncoding("utf8");
    socket.on("data", (chunk) => {
      buffer += chunk;
      if (buffer.length > MAX_LINE_BYTES) {
        socket.destroy();
        return;
      }
      let newline;
      while ((newline = buffer.indexOf("\n")) >= 0) {
        const line = buffer.slice(0, newline).trim();
        buffer = buffer.slice(newline + 1);
        if (line) this.#handleLine(socket, line);
      }
    });
    socket.on("close", () => this.#subscribers.delete(socket));
    socket.on("error", () => this.#subscribers.delete(socket));
  }

  async #handleLine(socket, line) {
    let request;
    try {
      request = JSON.parse(line);
    } catch {
      socket.write(JSON.stringify({ id: null, error: "Invalid JSON" }) + "\n");
      return;
    }
    const { id = null, method, params = {} } = request;
    try {
      const result = await this.dispatch(method, params, socket);
      if (!socket.destroyed) socket.write(JSON.stringify({ id, result }) + "\n");
    } catch (error) {
      console.warn("[TERMINAL_API] Request failed", { method });
      if (!socket.destroyed) socket.write(JSON.stringify({ id, error: error.message }) + "\n");
    }
  }

  /** Run one API method. Public for testing. */
  async dispatch(method, params, socket) {
    switch (method) {
      case "status": {
        const info = await this.#graph.getTokenInfo();
        return {
          loginVisible: this.#loginContents.size > 0,
          signedIn: !!info.success,
          scopes: info.data?.scopes || [],
          expiry: info.data?.expiry || null,
          error: info.success ? undefined : info.error,
        };
      }
      case "subscribe":
        this.#subscribers.add(socket);
        return { subscribed: true, loginVisible: this.#loginContents.size > 0 };
      case "teams":
        return await this.#teams(true);
      case "channels": {
        const team = await this.#resolveTeam(params.team);
        return { team, channels: await this.#channels(team.id, true) };
      }
      case "messages": {
        const { team, channel } = await this.#resolveChannel(params.channel);
        return this.#unwrap(await this.#graph.getChannelMessages(team.id, channel.id, params.top));
      }
      case "replies": {
        const { team, channel } = await this.#resolveChannel(params.channel);
        return this.#unwrap(
          await this.#graph.getChannelMessageReplies(team.id, channel.id, this.#required(params, "messageId"), params.top)
        );
      }
      case "post": {
        const { team, channel } = await this.#resolveChannel(params.channel);
        return this.#unwrap(
          await this.#graph.sendChannelMessage(team.id, channel.id, this.#required(params, "text"), params.replyTo || null)
        );
      }
      case "chats":
        return this.#unwrap(await this.#chats.getChats(params.top));
      case "chatMessages":
        return this.#unwrap(await this.#chats.getChatMessages(this.#required(params, "chatId"), params.top));
      case "graph":
        // Raw Graph passthrough for the Go TUI; see graphProxy.js.
        return await this.#proxy.handle(params);
      case "login":
        return await this.#login(params, socket);
      case "screenshot":
        return await this.#screenshot();
      case "chatSend":
        return this.#unwrap(
          await this.#graph.sendChatMessage(this.#required(params, "chatId"), this.#required(params, "text"))
        );
      default:
        throw new Error(`Unknown method: ${method}`);
    }
  }

  /**
   * Step through the web sign-in on behalf of a terminal user. Progress goes
   * to the requesting socket as {event:'login-progress'}; the password is
   * only typed into the sign-in page and never logged or broadcast.
   */
  async #login(params, socket) {
    if (this.#loginRunning) throw new Error("A login is already in progress");
    this.#loginRunning = true;
    const user = params.user || this.#config.auth?.webLogin?.user || null;
    const password = params.password || null;
    const prefillPassword = !!this.#config.auth?.webLogin?.passwordCommand;
    const started = Date.now();
    let waitStarted = started;
    let sawLogin = false;
    let resetDone = false;
    let last = null;
    const progress = (step) => {
      const key = JSON.stringify(step);
      if (key === last || socket.destroyed) return;
      last = key;
      socket.write(JSON.stringify({ event: "login-progress", ...step }) + "\n");
    };
    try {
      while (Date.now() - started < LOGIN_TIMEOUT_MS) {
        const wc = [...this.#loginContents.values()].find((w) => !w.isDestroyed());
        if (!wc) {
          if (sawLogin) return { signedIn: true };
          if (Date.now() - waitStarted > LOGIN_APPEAR_MS) {
            if ((await this.#graph.getTokenInfo())?.success) return { signedIn: true, alreadySignedIn: true };
            // An expired cached account makes Teams fail silent SSO and stay put
            // instead of redirecting, so drop its auth cache to force the sign-in page.
            if (resetDone) throw new Error("Session expired, but no sign-in page appeared after clearing the cache");
            resetDone = true;
            progress({ state: "resetting-session" });
            await this.#resetTeamsSession();
            waitStarted = Date.now();
            continue;
          }
          progress({ state: "waiting-for-sign-in-page" });
          await sleep(LOGIN_STEP_MS);
          continue;
        }
        sawLogin = true;
        const step = await wc.mainFrame
          .executeJavaScript(buildLoginStepScript({ user, password }), true)
          .catch(() => ({ state: "loading" }));
        switch (step?.state) {
          case "email":
            throw new Error("The sign-in page asks for an email: run `teams-cli login <email>`");
          case "password":
            if (!prefillPassword) throw new Error("The sign-in page asks for a password but none was given");
            break; // auth.webLogin.passwordCommand fills it
          case "error":
            throw new Error(`Sign-in page reports: ${step.error}`);
          case "otc-unsupported":
            throw new Error("The sign-in asks for a one-time code; switch MFA to Authenticator approval or sign in via VNC");
          case "mfa":
            this.#reportMfaCode(step.code);
            break;
          default:
            break;
        }
        progress(step || { state: "loading" });
        await sleep(LOGIN_STEP_MS);
      }
      throw new Error("Sign-in did not finish within 5 minutes");
    } finally {
      this.#loginRunning = false;
    }
  }

  async #resetTeamsSession() {
    const wc = this.#graph.mainWindow?.webContents;
    if (!wc || wc.isDestroyed()) throw new Error("Teams window is not available");
    const origin = new URL(wc.getURL()).origin;
    console.info("[TERMINAL_API] Clearing Teams auth cache to force sign-in");
    await wc.session.clearStorageData({
      origin,
      storages: ["localstorage", "indexdb", "serviceworkers", "cachestorage"],
    });
    wc.reload();
  }

  /** PNG of the sign-in window if one is open, else the main window. */
  async #screenshot() {
    const wc =
      [...this.#loginContents.values()].find((w) => !w.isDestroyed()) ||
      this.#graph.mainWindow?.webContents;
    if (!wc || wc.isDestroyed()) throw new Error("No window to capture");
    const image = await wc.capturePage();
    return { png: image.toPNG().toString("base64"), url: wc.getURL().split("?")[0] };
  }

  #required(params, key) {
    if (params[key] === undefined || params[key] === null || params[key] === "") {
      throw new Error(`Missing parameter: ${key}`);
    }
    return params[key];
  }

  #unwrap(result) {
    if (!result?.success) {
      const status = result?.status ? ` (HTTP ${result.status})` : "";
      throw new Error(`${result?.error || "Graph request failed"}${status}`);
    }
    return result.data ?? { ok: true };
  }

  async #teams(refresh = false) {
    if (!this.#teamsCache || refresh) {
      this.#teamsCache = this.#unwrap(await this.#graph.getJoinedTeams()).value || [];
    }
    return this.#teamsCache;
  }

  async #channels(teamId, refresh = false) {
    if (!this.#channelsCache.has(teamId) || refresh) {
      this.#channelsCache.set(teamId, this.#unwrap(await this.#graph.getTeamChannels(teamId)).value || []);
    }
    return this.#channelsCache.get(teamId);
  }

  async #resolveTeam(spec) {
    if (!spec) throw new Error("Missing parameter: team");
    return matchByName(await this.#teams(), spec);
  }

  async #resolveChannel(spec) {
    if (!spec || !String(spec).includes("/")) {
      throw new Error('Channel must be given as "Team/Channel"');
    }
    const slash = String(spec).lastIndexOf("/");
    const team = await this.#resolveTeam(spec.slice(0, slash));
    const channel = matchByName(await this.#channels(team.id), spec.slice(slash + 1));
    return { team, channel };
  }
}

module.exports = { TerminalApiServer, matchByName, MFA_CODE_SCRIPT, SOCKET_NAME };
