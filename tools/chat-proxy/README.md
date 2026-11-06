# NVIDIA chat proxy

The SharePoint chatbot (`components/research/ChatbotPanel.tsx`) calls an
OpenAI-compatible endpoint. It cannot call NVIDIA directly:

- `https://integrate.api.nvidia.com/v1/chat/completions` returns no
  `Access-Control-Allow-Origin` header, so the browser blocks the request.
- An SPFx web part is client-side code. Any API key embedded in it is readable
  by every visitor, and the key would be committed to git.

This folder contains two small proxies that add the key server-side. Pick one.

> **Rotate the key first.** The key shared in chat should be treated as
> compromised. Create a new key in the NVIDIA console and put the new value in
> the proxy secret. Do not commit it.

## Option A: Cloudflare Worker (recommended, free, browser setup)

1. Open <https://dash.cloudflare.com> → **Workers & Pages** → **Create** →
   **Worker**.
2. Paste `cloudflare-worker.js` as the worker code and deploy.
3. Add an **encrypted secret** named `NVIDIA_API_KEY` with your key.
4. (Recommended) Add a plain variable `ALLOWED_ORIGINS` set to your SharePoint
   origins, separated by commas, for example:
   `https://contoso.sharepoint.com,https://contoso-my.sharepoint.com`
5. Copy the worker URL, e.g. `https://nvidia-chat.example.workers.dev`.
6. In the SharePoint chatbot, open **Settings** and paste that URL into
   **Proxy URL**.

The client accepts the worker base URL, a `/v1` URL, or the full
`/v1/chat/completions` URL.

## Option B: Node 18+ server

Deploy `server.js` and `package.json` to any HTTPS host (Azure App Service,
Container Apps, Render, Fly.io, a VM). Then set:

| Variable | Purpose |
| --- | --- |
| `NVIDIA_API_KEY` | Required. Server-side key. |
| `ALLOWED_ORIGINS` | Comma-separated SharePoint origins. Empty allows any origin. |
| `PORT` | Listening port, default `8787`. |

```bash
NVIDIA_API_KEY=nvapi-... \
ALLOWED_ORIGINS=https://contoso.sharepoint.com \
npm start
```

The proxy must be HTTPS when the SharePoint page is HTTPS. An
`http://localhost` proxy is blocked as mixed content by the browser.

## Option C: Power Automate (no code hosting)

The client also accepts a JSON response, so a flow works if you prefer not to
host anything:

1. Create an **instant cloud flow** with the **When an HTTP request is
   received** trigger.
2. Add an **HTTP** action: `POST` to
   `https://integrate.api.nvidia.com/v1/chat/completions`, header
   `Authorization: Bearer <key>`, body `triggerBody()`.
3. Add a **Response** action: status `200`, body `body('HTTP')`, and a header
   `Access-Control-Allow-Origin` set to your SharePoint origin (or `*`).
4. Use the flow's HTTP POST URL as the **Proxy URL** in the chatbot.

Power Automate does not stream; the client detects a JSON response and renders
it when it arrives.

## What the client sends

```json
{
  "model": "deepseek-ai/deepseek-v4.1-flash",
  "messages": [{ "role": "user", "content": "Hello" }],
  "temperature": 1,
  "top_p": 0.95,
  "max_tokens": 4096,
  "stream": true
}
```

Image turns use the OpenAI content-parts shape:

```json
{
  "role": "user",
  "content": [
    { "type": "text", "text": "Describe this image." },
    { "type": "image_url", "image_url": { "url": "data:image/jpeg;base64,..." } }
  ]
}
```

## Hardening notes

- Keep `ALLOWED_ORIGINS` to your tenant, not `*`.
- Keep the key in the proxy's secret store, never in `src/`.
- Add rate limiting at the proxy or provider level before wider use.
- The proxy forwards any JSON body; restrict `model` server-side if you want to
  prevent callers from switching to a more expensive model.
