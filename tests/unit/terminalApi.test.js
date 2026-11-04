'use strict';

const { describe, it, before, after } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const net = require('node:net');
const os = require('node:os');
const path = require('node:path');
const options = require('../../app/config/options');

const electronPath = require.resolve('electron');
const modulePath = require.resolve('../../app/terminalApi');

function loadModule() {
	require.cache[electronPath] = {
		id: electronPath,
		filename: electronPath,
		loaded: true,
		exports: {
			app: { getPath: () => os.tmpdir(), on() {} },
			webContents: { getAllWebContents: () => [] },
		},
	};
	delete require.cache[modulePath];
	return require(modulePath);
}

function fakeGraph() {
	const calls = [];
	return {
		calls,
		getTokenInfo: async () => ({ success: true, data: { scopes: ['ChannelMessage.Send'], expiry: 'x' } }),
		getJoinedTeams: async () => ({ success: true, data: { value: [{ id: 't1', displayName: 'Engineering' }, { id: 't2', displayName: 'Engagement' }] } }),
		getTeamChannels: async (teamId) => ({ success: true, data: { value: [{ id: `${teamId}-c1`, displayName: 'General' }] } }),
		getChannelMessages: async (...args) => { calls.push(['getChannelMessages', ...args]); return { success: true, data: { value: [] } }; },
		sendChannelMessage: async (...args) => { calls.push(['sendChannelMessage', ...args]); return { success: true, data: { id: 'm1' } }; },
	};
}

function request(socketPath, payload) {
	return new Promise((resolve, reject) => {
		const socket = net.createConnection(socketPath, () => socket.write(JSON.stringify(payload) + '\n'));
		let data = '';
		socket.setEncoding('utf8');
		socket.on('data', (chunk) => {
			data += chunk;
			if (data.includes('\n')) {
				socket.end();
				resolve(JSON.parse(data.split('\n')[0]));
			}
		});
		socket.on('error', reject);
	});
}

it('defines terminalApi as disabled by default and requiring restart', () => {
	assert.deepEqual(options.terminalApi.default, { enabled: false });
	assert.equal(options.terminalApi.fields.enabled.type, 'boolean');
	assert.equal(options.terminalApi.applyMode, 'restart');
});

describe('matchByName', () => {
	const { matchByName } = loadModule();
	const items = [{ id: 'a', displayName: 'Engineering' }, { id: 'b', displayName: 'Engagement' }, { id: 'c', displayName: 'Eng' }];

	it('prefers exact case-insensitive matches, then ids, then unique prefixes', () => {
		assert.equal(matchByName(items, 'eng').id, 'c');
		assert.equal(matchByName(items, 'b').id, 'b');
		assert.equal(matchByName(items, 'engi').id, 'a');
	});

	it('rejects ambiguous and unknown names', () => {
		assert.throws(() => matchByName(items.slice(0, 2), 'eng'), /Ambiguous/);
		assert.throws(() => matchByName(items, 'sales'), /No match/);
	});
});

describe('TerminalApiServer socket', () => {
	let dir;
	let server;
	let graph;

	before(async () => {
		const { TerminalApiServer } = loadModule();
		dir = fs.mkdtempSync(path.join(os.tmpdir(), 'tfl-term-'));
		graph = fakeGraph();
		const chatService = { getChats: async () => ({ success: false, error: 'Forbidden', status: 403 }) };
		server = new TerminalApiServer({}, graph, { socketPath: path.join(dir, 'api.sock'), chatService });
		await server.start();
	});

	after(() => {
		server.dispose();
		fs.rmSync(dir, { recursive: true, force: true });
		delete require.cache[electronPath];
	});

	it('restricts the socket to the owner', () => {
		assert.equal(fs.statSync(server.socketPath).mode & 0o777, 0o600);
	});

	it('answers status with the token scopes', async () => {
		const reply = await request(server.socketPath, { id: 1, method: 'status' });
		assert.equal(reply.id, 1);
		assert.equal(reply.result.signedIn, true);
		assert.deepEqual(reply.result.scopes, ['ChannelMessage.Send']);
	});

	it('resolves "Team/Channel" names to ids when posting', async () => {
		const reply = await request(server.socketPath, { id: 2, method: 'post', params: { channel: 'engineering/gen', text: 'hi' } });
		assert.deepEqual(reply.result, { id: 'm1' });
		assert.deepEqual(graph.calls.at(-1), ['sendChannelMessage', 't1', 't1-c1', 'hi', null]);
	});

	it('returns Graph failures and bad requests as errors', async () => {
		assert.match((await request(server.socketPath, { id: 3, method: 'chats' })).error, /Forbidden \(HTTP 403\)/);
		assert.match((await request(server.socketPath, { id: 4, method: 'post', params: { channel: 'Engineering/General' } })).error, /Missing parameter: text/);
		assert.match((await request(server.socketPath, { id: 5, method: 'nope' })).error, /Unknown method/);
	});

	it('removes the socket on dispose', () => {
		const socketPath = server.socketPath;
		server.dispose();
		assert.equal(fs.existsSync(socketPath), false);
	});
});

describe('GraphApiClient.getTokenInfo', () => {
	it('returns sorted scopes without exposing the token', async () => {
		const GraphApiClient = require('../../app/graphApi');
		const client = new GraphApiClient({ graphApi: { enabled: true } });
		const payload = Buffer.from(JSON.stringify({ scp: 'User.Read Chat.Read' })).toString('base64url');
		client.acquireToken = async () => ({ success: true, token: `h.${payload}.s`, expiry: 'e' });
		const info = await client.getTokenInfo();
		assert.deepEqual(info, { success: true, data: { scopes: ['Chat.Read', 'User.Read'], expiry: 'e' } });
	});
});

describe('cli htmlToText', () => {
	const { htmlToText } = require('../../app/terminalApi/cli');

	it('strips tags and decodes entities', () => {
		assert.equal(htmlToText('<p>Hi &amp; <b>bye</b></p><div>a&lt;b&#39;</div>'), "Hi & bye\na<b'");
		assert.equal(htmlToText(null), '');
	});

	it('collapses whitespace-only lines from Teams paragraph markup', () => {
		assert.equal(htmlToText('<p>Hello</p>\n<p>&nbsp;</p>\n<p>&nbsp;</p>\n<p>  World</p><img src="x">'), 'Hello\n\n  World\n[image]');
	});
});

describe('chatService normalisation', () => {
	const { toGraphMessage, isChatConversation, mriToUserId } = require('../../app/terminalApi/chatService');

	it('maps chatsvc messages to the Graph chatMessage shape', () => {
		const msg = toGraphMessage({
			id: '1', messagetype: 'RichText/Html', content: '<p>hi</p>', imdisplayname: 'Alice',
			originalarrivaltime: '2026-11-04T10:00:00Z', from: 'https://x/v1/users/ME/contacts/8:orgid:0F0E0D0C-0000-0000-0000-000000000001',
		});
		assert.equal(msg.messageType, 'message');
		assert.equal(msg.from.user.displayName, 'Alice');
		assert.equal(msg.from.user.id, '0f0e0d0c-0000-0000-0000-000000000001');
		assert.equal(msg.body.content, '<p>hi</p>');
		assert.equal(toGraphMessage({ id: '2', messagetype: 'ThreadActivity/AddMember' }).messageType, 'systemEventMessage');
		assert.equal(toGraphMessage({ id: '3', messagetype: 'RichText/Media_CallRecording' }).messageType, 'systemEventMessage');
		assert.equal(mriToUserId('8:live:someone'), null);
	});

	it('keeps chats and drops channels and system feeds', () => {
		assert.equal(isChatConversation({ id: '19:abc@unq.gbl.spaces' }), true);
		assert.equal(isChatConversation({ id: '19:abc@thread.v2', threadProperties: { threadType: 'chat' } }), true);
		assert.equal(isChatConversation({ id: '19:abc@thread.tacv2', threadProperties: { threadType: 'space' } }), false);
		assert.equal(isChatConversation({ id: '48:notifications' }), false);
	});
});

describe('TerminalApiServer socket ownership', () => {
	it('refuses to take over a socket another instance is serving, and honours TEAMS_CLI_SOCKET', async () => {
		const { TerminalApiServer } = loadModule();
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'tfl-own-'));
		const socketPath = path.join(dir, 'api.sock');
		const previous = process.env.TEAMS_CLI_SOCKET;
		process.env.TEAMS_CLI_SOCKET = socketPath;
		const first = new TerminalApiServer({}, fakeGraph(), { chatService: {} });
		try {
			assert.equal(first.socketPath, socketPath);
			await first.start();
			const second = new TerminalApiServer({}, fakeGraph(), { chatService: {} });
			await assert.rejects(second.start(), /in use by another instance/);
			assert.equal(fs.existsSync(socketPath), true);
		} finally {
			first.dispose();
			if (previous === undefined) delete process.env.TEAMS_CLI_SOCKET;
			else process.env.TEAMS_CLI_SOCKET = previous;
			fs.rmSync(dir, { recursive: true, force: true });
			delete require.cache[electronPath];
		}
	});
});

describe('loginDriver', () => {
	const { buildLoginStepScript } = require('../../app/terminalApi/loginDriver');

	it('embeds credentials as JSON literals so quotes cannot break the script', () => {
		const script = buildLoginStepScript({ user: 'a@b.c', password: `p'"\`\${x}` });
		assert.ok(script.includes(JSON.stringify(`p'"\`\${x}`)));
		assert.doesNotThrow(() => new Function(`return ${script}`));
	});
});

describe('graphShape', () => {
	const { toGraphMessage, toGraphChat, chatType, eventDetail } = require('../../app/terminalApi/graphShape');

	it('maps reactions, mentions, files and edits from chatsvc properties', () => {
		const msg = toGraphMessage({
			id: '10', messagetype: 'RichText/Html', originalarrivaltime: '2026-11-04T10:00:00Z',
			from: '8:orgid:00000000-0000-0000-0000-00000000000a', imdisplayname: 'A',
			content: '<p>hi <span itemscope="" itemtype="http://schema.skype.com/Mention" itemid="0">Bo</span></p>',
			properties: {
				emotions: [{ key: 'like', users: [{ mri: '8:orgid:00000000-0000-0000-0000-00000000000b', time: 1793700000000 }] }],
				mentions: JSON.stringify([{ itemid: '0', mri: '8:orgid:00000000-0000-0000-0000-00000000000b', displayName: 'Bo' }]),
				files: JSON.stringify([{ itemid: 'f1', fileName: 'a.txt', fileInfo: { fileUrl: 'https://x/a.txt' } }]),
				edittime: '1793700001000',
			},
		});
		assert.equal(msg.body.content, '<p>hi <at id="0">Bo</at></p>');
		assert.deepEqual(msg.reactions.map((r) => [r.reactionType, r.user.user.id]), [['like', '00000000-0000-0000-0000-00000000000b']]);
		assert.equal(msg.mentions[0].mentioned.user.displayName, 'Bo');
		assert.equal(msg.attachments[0].name, 'a.txt');
		assert.ok(msg.lastEditedDateTime);
	});

	it('escapes plain-text bodies and hides deleted content', () => {
		assert.equal(toGraphMessage({ id: '1', messagetype: 'Text', content: 'a<b\nc' }).body.content, 'a&lt;b<br>c');
		assert.equal(toGraphMessage({ id: '3', messagetype: 'Text', content: 'a\r\nb\rc' }).body.content, 'a<br>b<br>c');
		assert.equal(toGraphMessage({ id: '4', messagetype: 'RichText/Media_CallRecording', content: '<URIObject/>' }).body.content, '<systemEventMessage/>');
		const del = toGraphMessage({ id: '2', messagetype: 'RichText/Html', content: 'x', properties: { deletetime: '1793700000000' } });
		assert.equal(del.body.content, '');
		assert.equal(del.from, null);
	});

	it('turns chatsvc reply quotes into Graph messageReference attachments', () => {
		const msg = toGraphMessage({
			id: '20', messagetype: 'RichText/Html',
			content: '<div>\n<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="17"><strong itemprop="mri" itemid="8:orgid:00000000-0000-0000-0000-00000000000c">Cy D</strong><span itemprop="time" itemid="17"></span><p itemprop="preview">It&#39;s &amp; ok</p></blockquote>\nreply</div>',
		});
		assert.equal(msg.body.content, '<div>\n<attachment id="17"></attachment>\nreply</div>');
		const ref = JSON.parse(msg.attachments[0].content);
		assert.equal(msg.attachments[0].contentType, 'messageReference');
		assert.deepEqual([ref.messageId, ref.messagePreview, ref.messageSender.user.displayName], ['17', "It's & ok", 'Cy D']);
	});

	it('maps chatsvc reaction keys to Graph reaction types', () => {
		const { reactionType } = require('../../app/terminalApi/graphShape');
		assert.equal(reactionType('heart'), 'heart');
		assert.equal(reactionType('like'), 'like');
		assert.equal(reactionType('yes-tone3'), '👍🏽');
		assert.equal(reactionType('1f44f_clappinghands'), '👏');
		assert.equal(reactionType('weirdcustom'), 'weirdcustom');
	});

	it('maps system events to Graph eventDetail types', () => {
		assert.equal(eventDetail('Event/Call', '<partlist type="ended"><duration>90</duration></partlist>').callDuration, 'PT90S');
		assert.match(eventDetail('ThreadActivity/AddMember', '<friendlyname>Ann</friendlyname>')['@odata.type'], /membersAdded/);
		assert.equal(eventDetail('ThreadActivity/MemberJoined', '')['@odata.type'], '#microsoft.graph.memberJoinedEventMessageDetail');
	});

	it('derives chat type and unread viewpoint', () => {
		assert.equal(chatType({ id: '19:a_b@unq.gbl.spaces' }), 'oneOnOne');
		assert.equal(chatType({ id: '19:meeting_x@thread.v2' }), 'meeting');
		assert.equal(chatType({ id: '19:x@thread.v2', threadProperties: { threadType: 'chat' } }), 'group');
		const chat = toGraphChat({ id: '19:x@thread.v2', properties: { consumptionhorizon: '1793700000000;1;1' }, lastMessage: { id: '5', messagetype: 'Text', content: 'x', originalarrivaltime: '2026-11-04T10:00:00Z' } });
		assert.equal(chat.viewpoint.lastMessageReadDateTime, new Date(1793700000000).toISOString());
		assert.equal(chat.members, undefined);
	});
});

describe('graphProxy', () => {
	const { GraphProxy, matchChatRoute } = require('../../app/terminalApi/graphProxy');
	const G = 'https://graph.microsoft.com/v1.0';

	it('routes only the chat operations the token cannot do', () => {
		assert.equal(matchChatRoute('GET', `${G}/me/chats?$top=50`).op, 'listChats');
		assert.deepEqual(matchChatRoute('GET', `${G}/chats/19%3Aabc%40thread.v2/messages?$top=5`).args, ['19:abc@thread.v2']);
		assert.equal(matchChatRoute('PATCH', `${G}/chats/c/messages/m`).op, 'editMessage');
		assert.equal(matchChatRoute('POST', `${G}/chats/c/messages`), null); // send stays on Graph
		assert.equal(matchChatRoute('POST', `${G}/chats/c/messages/m/setReaction`), null);
		assert.equal(matchChatRoute('GET', `${G}/teams/t/channels`), null);
	});

	it('strips skin-tone modifiers that tmux draws as separate wide cells', () => {
		const { stripSkinTones } = require('../../app/terminalApi/graphProxy');
		assert.equal(stripSkinTones('{"r":"👍🏽","t":"hi 👋🏿 👨‍💻"}'), '{"r":"👍","t":"hi 👋 👨‍💻"}');
	});

	it('refuses non-Graph hosts so the token is never sent elsewhere', async () => {
		let tokenAsked = false;
		const proxy = new GraphProxy({ acquireToken: async () => { tokenAsked = true; return { success: true, token: 't' }; } }, {});
		const res = await proxy.handle({ method: 'GET', url: 'https://graph.microsoft.com.evil.example/v1.0/me' });
		assert.equal(res.status, 400);
		assert.equal(tokenAsked, false);
	});

	it('answers chat routes from the chat service in Graph shape', async () => {
		const chats = { listMessages: async (id, opts) => ({ value: [{ id: 'm1' }], seen: [id, opts.top] }) };
		const proxy = new GraphProxy({}, chats);
		const res = await proxy.handle({ method: 'GET', url: `${G}/chats/c1/messages?$top=7` });
		assert.equal(res.status, 200);
		assert.deepEqual(JSON.parse(Buffer.from(res.body, 'base64')).seen, ['c1', '7']);
	});
});
