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
