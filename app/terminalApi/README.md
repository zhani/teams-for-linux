# Terminal API

Read and post Teams channel and chat messages from a terminal. This is meant for running
teams-for-linux headless on a Linux box and using it over SSH.

## How it works

- The app keeps the signed-in Teams web session. All requests use the Graph token borrowed from it
  through `GraphApiClient` (see [Graph API](../graphApi/README.md)), so no Azure app registration
  or tenant consent is needed.
- `index.js` (`TerminalApiServer`) listens on a Unix socket at
  `$XDG_RUNTIME_DIR/teams-for-linux-terminal.sock`, falling back to the app's userData directory.
  - The socket is set to `0600`, so only your user can connect.
  - The protocol is newline-delimited JSON: you send `{id, method, params}` and get back
    `{id, result}` or `{id, error}`.
- When a Microsoft sign-in page opens, the server sends a `login-required` event to subscribers.
- When the number-matching MFA page shows its code (`#idRichContext_DisplaySign`), the server sends
  an `mfa-code` event and logs `Approve sign-in on your phone: NN`. You still approve the sign-in in
  Microsoft Authenticator yourself; this only shows you the number.
- `chatService.js` reads chats through the Teams chat service (`chatsvc`), the internal API the
  Teams web app uses. The borrowed Graph token has no `Chat.Read` scope, so Graph cannot read chats.
  It uses a token for `https://ic3.teams.office.com` and the regional base URL taken from requests
  Teams has already made. If no such request is found, it falls back to a skypetoken from authsvc.
  Sending chat messages still goes through Graph (`ChatMessage.Send`).
- `cli.js` is a client with no dependencies.

## Configuration

```json
{
  "graphApi": { "enabled": true },
  "terminalApi": { "enabled": true }
}
```

To have the sign-in form filled in without typing, use the existing `auth.webLogin.*` options:

```json
{
  "auth": {
    "webLogin": {
      "user": "you@example.com",
      "passwordCommand": "pass show teams",
      "autoSubmit": true
    }
  }
}
```

## Running headless

Electron needs a display, so give it a virtual one with Xvfb. Install it with
`sudo apt install xvfb`, then create `~/.config/systemd/user/teams-for-linux.service`:

```ini
[Unit]
Description=Teams for Linux (headless)

[Service]
ExecStart=/usr/bin/xvfb-run -a /usr/bin/teams-for-linux
Restart=on-failure

[Install]
WantedBy=default.target
```

On a Wayland desktop, Electron still opens a visible window under `xvfb-run`
because it prefers Wayland. Hide it from Wayland and force X11:

```sh
env -u WAYLAND_DISPLAY XDG_SESSION_TYPE=x11 xvfb-run -a teams-for-linux --ozone-platform=x11
# from a source checkout:
npm run start:headless
```

In the systemd unit, use
`ExecStart=/usr/bin/env -u WAYLAND_DISPLAY XDG_SESSION_TYPE=x11 /usr/bin/xvfb-run -a /usr/bin/teams-for-linux --ozone-platform=x11`.
A user service started over SSH normally has no `WAYLAND_DISPLAY`, but the
flags make it safe either way.

Then run:

```sh
systemctl --user enable --now teams-for-linux
loginctl enable-linger "$USER"   # keep it running after you log out of SSH
```

## CLI

```sh
alias teams-cli='node /path/to/app/terminalApi/cli.js'
# or, using the installed app's bundled Node:
# alias teams-cli='ELECTRON_RUN_AS_NODE=1 /opt/teams-for-linux/teams-for-linux /opt/teams-for-linux/resources/app.asar/app/terminalApi/cli.js'

teams-cli status                         # signed in? which Graph scopes does the token carry?
teams-cli teams
teams-cli channels Engineering
teams-cli read "Engineering/General" -n 10
teams-cli replies "Engineering/General" <messageId>
teams-cli post "Engineering/General" "Hello from the terminal"
teams-cli reply "Engineering/General" <messageId> "Thanks!"
echo "multi-line text" | teams-cli post "Engineering/General" -
teams-cli chats
teams-cli chat 1 -n 20                   # '#'-index from `chats`, or a chat id
teams-cli send 1 "On my way"
teams-cli watch                          # sign-in/MFA events and new chat messages
teams-cli login you@example.com          # sign in headless; prompts for the password
teams-cli screenshot /tmp/teams.png      # image of the invisible window, for debugging
```

Names in `Team/Channel` are matched case-insensitively. An exact name wins; otherwise a unique
prefix is accepted.

## Signing in headless

When the session expires, the app shows the Microsoft sign-in page on the invisible display.
`teams-cli login <email>` asks for your password without echoing it. It then works through the
pages with `loginDriver.js`: email, password, account picker, number-matching MFA and "Stay signed
in?". Each page is acted on at most once. It prints the two-digit number for you to enter in
Microsoft Authenticator. The password is only typed into the sign-in page; it is never logged or
saved.

- A password the page rejects is never typed again, because repeated failures can lock the
  account.
- One-time-code MFA (SMS or an authenticator code) is not supported. Use Authenticator approval,
  or sign in over VNC.
- If it stops on an unrecognised page, run `teams-cli screenshot` to see what is showing.
  Federated company sign-in pages are not handled.
- The signed-in session lives in the profile directory and survives restarts.

To test the flow without touching your real session, run a second instance with its own profile
and socket:

```sh
E2E_USER_DATA_DIR=/tmp/tfl-test TEAMS_CLI_SOCKET=/tmp/tfl-test.sock npm run start:headless
TEAMS_CLI_SOCKET=/tmp/tfl-test.sock teams-cli login you@example.com
```

`/tmp/tfl-test/config.json` needs `graphApi.enabled` and `terminalApi.enabled` set to `true`. The
server will not start on a socket that another instance is still serving.

## Limitations

- **Scopes are fixed by Microsoft.** The token belongs to the Teams first-party client, and some
  endpoints may return 403 for your tenant. Reading channel messages, for example, needs
  `ChannelMessage.Read.All`. Run `teams-cli status` to see what the token carries.
- `watch` polls the chat list every 30 seconds. It does not watch channels.
- Messages are sent as plain text.
- Reading chats relies on the undocumented `chatsvc` API, which Microsoft can change without notice.
- Nothing about message content, people, teams, channels or URLs is logged.
