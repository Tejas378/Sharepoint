# Local NVIDIA chat

A standalone chatbot that runs on your machine. No SharePoint, no deployment,
no Azure/Cloudflare account.

```
local-chat/
  server.js            local HTTP server + /api/chat proxy
  public/              chat UI (HTML/CSS/JS)
  .env                 your key (ignored by git)
  .env.example         template
```

## Run it

Requires Node 18 or newer.

**Windows, no terminal:** double-click `local-chat\start.cmd`. It starts the
server and opens <http://127.0.0.1:4173> in your browser.

**Terminal:**

```powershell
cd local-chat
npm start
```

Then open <http://127.0.0.1:4173>.

`npm start` only runs `node server.js`; there are no dependencies to install.
Stop the server with `Ctrl+C` in its window.

## How it works

- The browser calls `POST /api/chat` on the same local server, so there is no
  CORS problem and no key in the browser.
- `server.js` forwards the request to
  `https://integrate.api.nvidia.com/v1/chat/completions` with
  `Authorization: Bearer <NVIDIA_API_KEY>` and streams the SSE response back.
- The server binds to `127.0.0.1` by default, so it is not reachable from other
  machines.

## Where to see the request

The browser never contacts NVIDIA, so DevTools → **Network** only shows
`POST http://127.0.0.1:4173/api/chat`. The outbound call appears in the server
window:

```
[11:03:15] POST /api/chat
[11:03:15] -> POST https://integrate.api.nvidia.com/v1/chat/completions model=nvidia/nemotron-3.5-super-vl-120b-a12b stream=true
```

If the upstream answers, the next line is `<- HTTP 200 text/event-stream …`.
If nothing arrives, the request is sent but never answered; the server gives up
after `CHAT_TIMEOUT_MS` and the UI shows a `504` banner. While waiting, the
composer shows `Request sent to NVIDIA through this server · waiting Ns`, and
pressing **Stop** (or reloading) cancels the upstream call as well.

## The key

`.env` is listed in `local-chat/.gitignore` and is not part of the repository.
Edit it to change the key:

```
NVIDIA_API_KEY=nvapi-...
```

Optional settings in the same file:

| Variable | Default | Purpose |
| --- | --- | --- |
| `PORT` | `4173` | Port for the local server. |
| `HOST` | `127.0.0.1` | Bind address. Keep it on localhost. |
| `CHAT_TIMEOUT_MS` | `180000` | How long to wait for NVIDIA before answering `504`. |

The key that was pasted into the earlier chat should be rotated in the NVIDIA
console; replace the value here after rotating.

## Features

- Streaming replies with a Stop button.
- Multi-turn conversation, Enter to send, Shift+Enter for a new line.
- Image attachment (up to 5 MB) sent as an `image_url` content part.
- Settings for model, system prompt, temperature, top-p and max tokens, saved
  in this browser.
- Default model: `nvidia/nemotron-3.5-super-vl-120b-a12b` (vision + text).
- Reasoning models stream their chain of thought separately; it appears in a
  collapsible **Reasoning** block and collapses automatically when the answer
  starts. `deepseek-ai/deepseek-v4.1-flash` streams `reasoning_content` too, but
  that model does not answer on the current account.
- Plain JSON responses also work if you point the server at a non-streaming
  backend.

## Troubleshooting

| Symptom | Cause |
| --- | --- |
| Banner says the key is missing | `.env` not found or `NVIDIA_API_KEY` empty; restart `npm start`. |
| `HTTP 401` | Key is invalid or revoked. |
| `HTTP 402/403` | The key has no access to the model; change **Model** in settings. |
| `HTTP 429` | Rate limited; wait and retry. |
| `Could not reach the local server` | The `npm start` process stopped. |
| `HTTP 504: NVIDIA did not respond within …` | NVIDIA accepted the request but sent nothing. The key appears in `/v1/models` but inference hangs, which usually means the key is flagged, out of credit, or the model is unavailable for that account. Rotate the key, or change **Model** in settings. |
| Port already in use | Set `PORT=4174` in `.env` and restart. |

## Scope

This folder is independent of the SPFx project. It does not touch `src/` and is
not included in the SharePoint bundle.
