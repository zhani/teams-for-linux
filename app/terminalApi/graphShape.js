// app/terminalApi/graphShape.js

/**
 * Pure mappers from Teams chat service (chatsvc) objects to the Microsoft
 * Graph chat / chatMessage JSON shape, so Graph clients (cli.js, the Go TUI
 * in tui/) can read chats although the borrowed token lacks Chat.Read.
 *
 * Only the fields those clients read are produced. chatsvc message ids are
 * the same arrival-time ids Graph uses, so Graph write calls that the token
 * does allow (send, setReaction) work on mapped messages.
 */

const MENTION_SPAN =
  /<span[^>]*itemtype="http:\/\/schema\.skype\.com\/Mention"[^>]*itemid="(\d+)"[^>]*>(.*?)<\/span>/gi;

/** AAD object id from an MRI such as ".../contacts/8:orgid:<guid>". */
function mriToUserId(mri) {
  const m = String(mri || "").match(/8:orgid:([0-9a-f-]{36})/i);
  return m ? m[1].toLowerCase() : null;
}

function toIso(ms) {
  const n = Number(ms);
  return Number.isFinite(n) && n > 0 ? new Date(n).toISOString() : null;
}

function parseJsonProp(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value;
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

// chatsvc quotes a replied-to message inline; Graph uses an <attachment>
// placeholder plus a messageReference attachment, which clients render.
const REPLY_QUOTE =
  /<blockquote[^>]*itemtype="http:\/\/schema\.skype\.com\/Reply"[^>]*>([\s\S]*?)<\/blockquote>/gi;

function htmlToPlain(html) {
  return String(html || "")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_m, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_m, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, "&")
    .replace(/\s+/g, " ")
    .trim();
}

/** Replace chatsvc reply quotes with Graph messageReference attachments. */
function extractReplyQuotes(content) {
  const references = [];
  const html = content.replace(REPLY_QUOTE, (block, inner) => {
    const id = (block.match(/itemid="(\d+)"/) || [])[1];
    if (!id) return block;
    const sender = inner.match(/<strong[^>]*itemprop="mri"[^>]*itemid="([^"]*)"[^>]*>([\s\S]*?)<\/strong>/i);
    const preview = (inner.match(/<p[^>]*itemprop="preview"[^>]*>([\s\S]*?)<\/p>/i) || [])[1];
    references.push({
      id,
      contentType: "messageReference",
      content: JSON.stringify({
        messageId: id,
        messagePreview: htmlToPlain(preview),
        messageSender: {
          user: { userIdentityType: "aadUser", id: mriToUserId(sender?.[1]), displayName: htmlToPlain(sender?.[2]) || null },
        },
      }),
    });
    return `<attachment id="${id}"></attachment>`;
  });
  return { html, references };
}

// chatsvc reaction keys -> Graph reactionType. The six classic names pass
// through; others (and skin-tone variants) become the emoji itself, as Graph
// reports them.
const REACTION_EMOJI = {
  like: "👍", yes: "👍", heart: "❤️", laugh: "😆", surprised: "😮", sad: "😢", angry: "😡",
};
const CLASSIC_REACTIONS = new Set(["like", "heart", "laugh", "surprised", "sad", "angry"]);

function reactionType(key) {
  const m = String(key || "").match(/^(.*?)(?:-tone([1-5]))?$/);
  const base = m[1];
  const tone = m[2] ? String.fromCodePoint(0x1f3fa + Number(m[2])) : "";
  if (!tone && CLASSIC_REACTIONS.has(base)) return base;
  // Unicode-named keys such as "1f44f_clappinghands".
  const hex = base.match(/^([0-9a-f]{4,6})_/i);
  const emoji = REACTION_EMOJI[base] || (hex ? String.fromCodePoint(parseInt(hex[1], 16)) : null);
  return emoji ? emoji.replace(/\ufe0f$/, "") + tone : key;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\n", "<br>");
}

function xmlValues(xml, tag) {
  const re = new RegExp(`<${tag}>([^<]*)</${tag}>`, "gi");
  return [...String(xml || "").matchAll(re)].map((m) => m[1]);
}

/** Whether a chatsvc conversation is a chat (not a channel, system feed, ...). */
function isChatConversation(conv) {
  const id = conv?.id || "";
  if (!id.startsWith("19:")) return false; // 48:notifications, 48:calllogs, ...
  const type = (conv.threadProperties?.threadType || "").toLowerCase();
  return !["space", "topic"].includes(type); // channels
}

/** Graph chatType for a chatsvc conversation. */
function chatType(conv) {
  const id = conv?.id || "";
  const props = conv?.threadProperties || {};
  if (id.endsWith("@unq.gbl.spaces") || props.productThreadType === "OneToOneChat") return "oneOnOne";
  if ((props.threadType || "").toLowerCase() === "meeting" || id.includes("meeting_")) return "meeting";
  return "group";
}

/**
 * Graph eventDetail for chatsvc system messages. Unknown types become
 * "<name>EventMessageDetail" so clients can still humanise the name.
 */
function eventDetail(type, content) {
  const name = String(type).split("/").pop() || "unknown";
  const friendly = xmlValues(content, "friendlyname").map((n) => ({ displayName: n }));
  switch (type) {
    case "ThreadActivity/AddMember":
      return { "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail", members: friendly };
    case "ThreadActivity/DeleteMember":
      return { "@odata.type": "#microsoft.graph.membersDeletedEventMessageDetail", members: friendly };
    case "ThreadActivity/TopicUpdate":
      return {
        "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail",
        chatDisplayName: xmlValues(content, "value")[0] || "",
      };
    case "Event/Call": {
      const ended = /<partlist[^>]*type="ended"/i.test(content) || /<ended\/>/i.test(content);
      const seconds = Number(xmlValues(content, "duration")[0]);
      return {
        "@odata.type": ended
          ? "#microsoft.graph.callEndedEventMessageDetail"
          : "#microsoft.graph.callStartedEventMessageDetail",
        ...(ended && Number.isFinite(seconds) ? { callDuration: `PT${Math.round(seconds)}S` } : {}),
      };
    }
    case "RichText/Media_CallRecording":
      return { "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail" };
    case "RichText/Media_CallTranscript":
      return { "@odata.type": "#microsoft.graph.callTranscriptEventMessageDetail" };
    default:
      return { "@odata.type": `#microsoft.graph.${name.charAt(0).toLowerCase()}${name.slice(1)}EventMessageDetail` };
  }
}

const USER_TYPES = /^(Text|RichText|RichText\/Html)$/i;

/** Whether a chatsvc message is only a client signal (typing etc.), never shown. */
function isControlMessage(msg) {
  return /^Control\//i.test(String(msg?.messagetype || ""));
}

/** Map a chatsvc message to the Graph chatMessage shape. */
function toGraphMessage(msg) {
  const type = String(msg?.messagetype || "");
  const props = msg?.properties || {};
  const isUser = USER_TYPES.test(type);
  const userId = mriToUserId(msg?.from);

  // chatsvc keeps raw CRLF (Graph never does); a bare \r makes terminals
  // jump to column 0 and corrupts TUI layouts.
  let content = String(msg?.content || "").replace(/\r\n?/g, "\n");
  if (isUser && !/^RichText/i.test(type)) content = escapeHtml(content);
  const { html: withoutQuotes, references } = extractReplyQuotes(content);
  content = withoutQuotes;
  // chatsvc marks mentions with schema.skype.com spans; Graph uses <at id>.
  content = content.replace(MENTION_SPAN, '<at id="$1">$2</at>');

  const mentions = parseJsonProp(props.mentions).map((m) => ({
    id: Number(m.itemid ?? m.itemId ?? 0),
    mentionText: m.displayName || "",
    mentioned: { user: { id: mriToUserId(m.mri), displayName: m.displayName || null } },
  }));

  const reactions = parseJsonProp(props.emotions).flatMap((emotion) =>
    (emotion.users || []).map((u) => ({
      reactionType: reactionType(emotion.key),
      createdDateTime: toIso(u.time),
      user: { user: { id: mriToUserId(u.mri), displayName: null } },
    }))
  );

  const attachments = references.concat(parseJsonProp(props.files).map((f, i) => ({
    id: f.itemid || f.id || `file-${i}`,
    name: f.fileName || f.title || null,
    contentType: "reference",
    contentUrl: f.fileInfo?.fileUrl || f.objectUrl || null,
  })));

  const deleted = toIso(props.deletetime);
  return {
    id: msg?.id,
    createdDateTime: msg?.originalarrivaltime || msg?.composetime || null,
    lastEditedDateTime: toIso(props.edittime),
    deletedDateTime: deleted,
    messageType: isUser ? "message" : "systemEventMessage",
    subject: props.subject || null,
    from: isUser && !deleted ? { user: { id: userId, displayName: msg?.imdisplayname || null } } : null,
    // Graph gives system events this placeholder body, not chatsvc's raw XML.
    body: { contentType: "html", content: deleted ? "" : isUser ? content : "<systemEventMessage/>" },
    eventDetail: isUser ? null : eventDetail(type, msg?.content),
    attachments,
    mentions,
    reactions,
  };
}

/**
 * Map a chatsvc conversation to the Graph chat shape. `members` is the Graph
 * member list (fetched separately) or undefined when not expanded.
 */
function toGraphChat(conv, members) {
  const last = conv?.lastMessage;
  const horizon = String(conv?.properties?.consumptionhorizon || "").split(";")[0];
  const chat = {
    id: conv.id,
    topic: conv.threadProperties?.topic || null,
    chatType: chatType(conv),
    lastUpdatedDateTime: last?.originalarrivaltime || last?.composetime || null,
    viewpoint: { lastMessageReadDateTime: toIso(horizon) || "0001-01-01T00:00:00Z" },
    lastMessagePreview: last?.id && !isControlMessage(last) ? toGraphMessage(last) : null,
  };
  if (members) chat.members = members;
  return chat;
}

module.exports = {
  mriToUserId,
  isChatConversation,
  isControlMessage,
  chatType,
  eventDetail,
  toGraphMessage,
  toGraphChat,
  reactionType,
};
