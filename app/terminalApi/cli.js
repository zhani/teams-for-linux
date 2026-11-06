#!/usr/bin/env node
// Terminal client for the teams-for-linux terminal API (app/terminalApi).
// Plain Node, no dependencies: run with `node cli.js ...` or, with an installed
// app, `ELECTRON_RUN_AS_NODE=1 teams-for-linux /path/to/cli.js ...`.

const net = require("node:net");
const path = require("node:path");
const os = require("node:os");
const fs = require("node:fs");
const readline = require("node:readline");

const SOCKET_NAME = "teams-for-linux-terminal.sock";
const WATCH_POLL_MS = 30000;

const USAGE = `Usage: teams-cli <command> [args]

  status                              Sign-in state and token scopes
  teams                               List joined teams
  channels <team>                     List channels of a team
  read <Team/Channel> [-n N] [--all] Show recent channel messages (--all: include system events)
  replies <Team/Channel> <msgId>      Show replies to a message
  post <Team/Channel> <text|->        Post a message ('-' reads stdin)
  reply <Team/Channel> <msgId> <text|->  Reply to a message
  chats [-n N]                        List chats (# index usable below)
  chat <chatId|#> [-n N]              Show recent chat messages
  send <chatId|#> <text|->            Send a chat message
  watch                               Print sign-in/MFA events and new chat messages
  login [email]                       Sign in from the terminal (prompts for the password)
  screenshot [file.png]               Save an image of the (headless) window

Socket: $TEAMS_CLI_SOCKET, else $XDG_RUNTIME_DIR/${SOCKET_NAME}`;

function socketPath() {
  if (process.env.TEAMS_CLI_SOCKET) return process.env.TEAMS_CLI_SOCKET;
  const dir = process.env.XDG_RUNTIME_DIR || path.join(os.homedir(), ".config", "teams-for-linux");
  return path.join(dir, SOCKET_NAME);
}

/** Convert a Teams HTML message body to readable plain text. Pure; exported for testing. */
function htmlToText(html) {
  if (!html) return "";
  return String(html)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li)>/gi, "\n")
    .replace(/<attachment[^>]*>.*?<\/attachment>/gi, "[attachment]")
    .replace(/<img[^>]*>/gi, "[image]")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#(\d+);/g, (_m, n) => String.fromCodePoint(Number(n)))
    .replace(/&amp;/g, "&")
    .split("\n")
    // Keep leading indentation (code blocks); drop trailing &nbsp; padding.
    .map((line) => line.trimEnd())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function connect() {
  const socket = net.createConnection(socketPath());
  let nextId = 1;
  const pending = new Map();
  const listeners = [];
  let buffer = "";
  socket.setEncoding("utf8");
  socket.on("data", (chunk) => {
    buffer += chunk;
    let nl;
    while ((nl = buffer.indexOf("\n")) >= 0) {
      const msg = JSON.parse(buffer.slice(0, nl));
      buffer = buffer.slice(nl + 1);
      if (msg.event) listeners.forEach((fn) => fn(msg));
      else if (pending.has(msg.id)) {
        const { resolve, reject } = pending.get(msg.id);
        pending.delete(msg.id);
        msg.error ? reject(new Error(msg.error)) : resolve(msg.result);
      }
    }
  });
  const ready = new Promise((resolve, reject) => {
    socket.once("connect", resolve);
    socket.once("error", (err) =>
      reject(new Error(`Cannot reach teams-for-linux at ${socketPath()} (${err.code}). Is it running with terminalApi.enabled?`))
    );
  });
  return {
    async call(method, params = {}) {
      await ready;
      const id = nextId++;
      return new Promise((resolve, reject) => {
        pending.set(id, { resolve, reject });
        socket.write(JSON.stringify({ id, method, params }) + "\n");
      });
    },
    onEvent: (fn) => listeners.push(fn),
    close: () => socket.end(),
  };
}

function takeFlag(args, flag) {
  const i = args.indexOf(flag);
  if (i < 0) return undefined;
  const value = args[i + 1];
  args.splice(i, 2);
  return value;
}

async function readText(arg) {
  if (arg !== "-") return arg;
  let data = "";
  for await (const chunk of process.stdin) data += chunk;
  return data.trimEnd();
}

function ask(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => rl.question(question, (answer) => {
    rl.close();
    resolve(answer.trim());
  }));
}

// Reads a line without echoing it, so the password never shows on screen.
function askHidden(question) {
  if (!process.stdin.isTTY) return ask(question);
  process.stdout.write(question);
  return new Promise((resolve) => {
    let value = "";
    const finish = () => {
      process.stdin.setRawMode(false);
      process.stdin.pause();
      process.stdin.off("data", onData);
      process.stdout.write("\n");
      resolve(value);
    };
    const onData = (chunk) => {
      for (const ch of chunk) {
        if (ch === "\r" || ch === "\n") return finish();
        if (ch === "\u0003") {
          process.stdin.setRawMode(false);
          process.stdout.write("\n");
          process.exit(130);
        }
        if (ch === "\u007f" || ch === "\b") value = value.slice(0, -1);
        else value += ch;
      }
    };
    process.stdin.setRawMode(true);
    process.stdin.setEncoding("utf8");
    process.stdin.resume();
    process.stdin.on("data", onData);
  });
}

const LOGIN_MESSAGES = {
  "waiting-for-sign-in-page": "Waiting for the sign-in page...",
  "resetting-session": "Session expired: clearing the cached Teams sign-in to show the sign-in page...",
  "email-submitted": "Email entered.",
  "password-submitted": "Password entered.",
  "account-picked": "Account selected.",
  "account-other": "Choosing 'Use another account'.",
  "mfa-wait": "Waiting for you to approve the sign-in on your phone...",
  "kmsi-accepted": "Answered 'Stay signed in?' with Yes.",
  password: "Waiting for auth.webLogin.passwordCommand to fill the password...",
};

async function login(client, email) {
  const user = email || (await ask("Email: "));
  const password = await askHidden("Password: ");
  client.onEvent((e) => {
    if (e.event !== "login-progress") return;
    if (e.state === "mfa") console.log(`*** Approve the sign-in in Microsoft Authenticator: ${e.code}`);
    else if (e.state === "unknown") console.log(`(sign-in page: ${e.title || "unrecognised step"}; try 'screenshot')`);
    else if (LOGIN_MESSAGES[e.state]) console.log(LOGIN_MESSAGES[e.state]);
  });
  const result = await client.call("login", { user, password });
  console.log(result.alreadySignedIn ? "No sign-in page appeared: already signed in." : "Signed in.");
}

const fmtTime = (iso) => (iso ? new Date(iso).toLocaleString() : "");
const sender = (m) => m.from?.user?.displayName || m.from?.application?.displayName || "system";

// Graph orders channel threads by latest activity (including replies), so
// sort by post time for a readable transcript; system events have no body.
function printMessages(messages, { all = false, indent = "" } = {}) {
  messages
    .filter((m) => all || (m.messageType || "message") === "message")
    .sort((a, b) => String(a.createdDateTime).localeCompare(String(b.createdDateTime)))
    .forEach((m) => printMessage(m, indent));
}

function printMessage(m, indent = "") {
  if (m.deletedDateTime) return;
  const text = htmlToText(m.body?.content).replace(/\n/g, `\n${indent}    `);
  console.log(`${indent}[${fmtTime(m.createdDateTime)}] ${sender(m)}: ${text}`);
  console.log(`${indent}    id=${m.id}`);
}

function chatTitle(chat, myName) {
  if (chat.topic) return chat.topic;
  const others = (chat.members || []).map((m) => m.displayName).filter((n) => n && n !== myName);
  return others.join(", ") || chat.chatType || "chat";
}

async function resolveChat(client, ref) {
  if (/^#?\d+$/.test(ref)) {
    const chats = (await client.call("chats", { top: 50 })).value || [];
    const chat = chats[Number(ref.replace("#", "")) - 1];
    if (!chat) throw new Error(`No chat #${ref}`);
    return chat.id;
  }
  return ref;
}

async function watch(client) {
  client.onEvent((e) => {
    if (e.event === "login-required") console.log("*** Sign-in page shown in teams-for-linux");
    if (e.event === "mfa-code") console.log(`*** Approve the sign-in in Microsoft Authenticator: ${e.code}`);
    if (e.event === "login-complete") console.log("*** Sign-in page closed");
  });
  await client.call("subscribe");
  console.log("Watching (Ctrl-C to stop)...");
  const seen = new Map();
  let first = true;
  const poll = async () => {
    try {
      const chats = (await client.call("chats", { top: 30 })).value || [];
      for (const chat of chats) {
        const last = chat.lastMessagePreview;
        if (!last || seen.get(chat.id) === last.id) continue;
        seen.set(chat.id, last.id);
        if (!first && (last.messageType || "message") === "message") {
          console.log(`[${fmtTime(last.createdDateTime)}] (${chatTitle(chat)}) ${sender(last)}: ${htmlToText(last.body?.content)}`);
        }
      }
      first = false;
    } catch (error) {
      console.error(`watch: ${error.message}`);
    }
  };
  await poll();
  setInterval(poll, WATCH_POLL_MS);
}

async function main(argv) {
  const args = [...argv];
  const command = args.shift();
  const top = takeFlag(args, "-n");
  const allIndex = args.indexOf("--all");
  const all = allIndex >= 0;
  if (all) args.splice(allIndex, 1);
  if (!command || command === "-h" || command === "--help") {
    console.log(USAGE);
    return;
  }
  const client = connect();
  try {
    switch (command) {
      case "status": {
        const s = await client.call("status");
        console.log(`Signed in: ${s.signedIn ? "yes" : "no"}${s.loginVisible ? " (sign-in page is open)" : ""}`);
        if (s.error) console.log(`Error: ${s.error}`);
        if (s.expiry) console.log(`Token expires: ${fmtTime(s.expiry)}`);
        if (s.scopes.length) console.log(`Scopes: ${s.scopes.join(" ")}`);
        break;
      }
      case "teams":
        for (const t of await client.call("teams")) console.log(`${t.displayName}\n    id=${t.id}`);
        break;
      case "channels": {
        const { team, channels } = await client.call("channels", { team: args.join(" ") });
        for (const c of channels) console.log(`${team.displayName}/${c.displayName}${c.membershipType !== "standard" ? ` (${c.membershipType})` : ""}`);
        break;
      }
      case "read": {
        const data = await client.call("messages", { channel: args[0], top });
        printMessages(data.value || [], { all });
        break;
      }
      case "replies": {
        const data = await client.call("replies", { channel: args[0], messageId: args[1], top });
        printMessages(data.value || [], { all, indent: "  " });
        break;
      }
      case "post":
        await client.call("post", { channel: args[0], text: await readText(args.slice(1).join(" ")) });
        console.log("Posted.");
        break;
      case "reply":
        await client.call("post", { channel: args[0], replyTo: args[1], text: await readText(args.slice(2).join(" ")) });
        console.log("Replied.");
        break;
      case "chats": {
        const chats = (await client.call("chats", { top: top || 30 })).value || [];
        chats.forEach((c, i) => {
          const last = c.lastMessagePreview;
          const preview = !last ? "" : last.messageType && last.messageType !== "message"
            ? "[event]"
            : `${sender(last)}: ${htmlToText(last.body?.content).replace(/\s+/g, " ").slice(0, 60)}`;
          console.log(`#${i + 1} ${chatTitle(c)}  [${fmtTime(last?.createdDateTime)}]\n    ${preview}`);
        });
        break;
      }
      case "chat": {
        const chatId = await resolveChat(client, args[0]);
        const data = await client.call("chatMessages", { chatId, top });
        printMessages(data.value || [], { all });
        break;
      }
      case "send": {
        const chatId = await resolveChat(client, args[0]);
        await client.call("chatSend", { chatId, text: await readText(args.slice(1).join(" ")) });
        console.log("Sent.");
        break;
      }
      case "login":
        await login(client, args[0]);
        break;
      case "screenshot": {
        const file = args[0] || "teams-screenshot.png";
        const { png, url } = await client.call("screenshot");
        // The window can show private messages; keep the file owner-only.
        fs.writeFileSync(file, Buffer.from(png, "base64"), { mode: 0o600 });
        console.log(`Saved ${file} (${url})`);
        break;
      }
      case "watch":
        await watch(client);
        return; // keep the connection open
      default:
        console.error(USAGE);
        process.exitCode = 2;
    }
  } catch (error) {
    console.error(`Error: ${error.message}`);
    process.exitCode = 1;
  }
  client.close();
}

if (require.main === module) main(process.argv.slice(2));

module.exports = { htmlToText, socketPath };
