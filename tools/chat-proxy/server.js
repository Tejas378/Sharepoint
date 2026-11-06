/**
 * NVIDIA chat proxy for the SharePoint chatbot (plain Node, no dependencies).
 *
 * Requires Node 18+ (global fetch and Readable.fromWeb).
 *
 *   NVIDIA_API_KEY=nvapi-... \
 *   ALLOWED_ORIGINS=https://contoso.sharepoint.com \
 *   PORT=8787 \
 *   npm start
 *
 * Serve it over HTTPS when SharePoint runs on HTTPS; an http://localhost proxy
 * is blocked as mixed content by the browser.
 */

const http = require("node:http");
const { Readable } = require("node:stream");

const UPSTREAM = "https://integrate.api.nvidia.com/v1/chat/completions";
const PORT = Number(process.env.PORT || 8787);
const API_KEY = process.env.NVIDIA_API_KEY || "";
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || "")
  .split(",")
  .map((value) => value.trim().replace(/\/+$/, ""))
  .filter(Boolean);
const MAX_BODY_BYTES = 25 * 1024 * 1024;

function corsHeaders(origin) {
  const allowed = ALLOWED_ORIGINS.length === 0 || ALLOWED_ORIGINS.indexOf(origin) >= 0;
  return {
    "Access-Control-Allow-Origin": allowed ? (origin || "*") : "null",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Accept",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin",
  };
}

function sendJson(res, status, body, cors) {
  res.writeHead(status, Object.assign({
    "Content-Type": "application/json",
    "Cache-Control": "no-store",
  }, cors));
  res.end(JSON.stringify(body));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > MAX_BODY_BYTES) {
        reject(new Error("Request body is too large."));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

const server = http.createServer(async (req, res) => {
  const origin = (req.headers.origin || "").replace(/\/+$/, "");
  const cors = corsHeaders(origin);

  if (req.method === "OPTIONS") {
    res.writeHead(204, cors);
    res.end();
    return;
  }
  if (req.method !== "POST") {
    sendJson(res, 405, { error: "Use POST." }, cors);
    return;
  }
  if (ALLOWED_ORIGINS.length > 0 && ALLOWED_ORIGINS.indexOf(origin) < 0) {
    sendJson(res, 403, { error: "Origin not allowed." }, cors);
    return;
  }
  if (!API_KEY) {
    sendJson(res, 500, { error: "NVIDIA_API_KEY is not set on the proxy." }, cors);
    return;
  }

  let payload;
  try {
    const raw = await readBody(req);
    payload = JSON.parse(raw || "{}");
  } catch (error) {
    sendJson(res, 400, { error: error instanceof Error ? error.message : "Invalid request body." }, cors);
    return;
  }

  try {
    const upstream = await fetch(UPSTREAM, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Accept": req.headers.accept || "text/event-stream",
        "Authorization": "Bearer " + API_KEY,
      },
      body: JSON.stringify(payload),
    });

    const headers = Object.assign({}, cors, {
      "Content-Type": upstream.headers.get("content-type") || "application/json",
      "Cache-Control": "no-store",
    });
    res.writeHead(upstream.status, headers);
    if (upstream.body) {
      Readable.fromWeb(upstream.body).pipe(res);
    } else {
      res.end();
    }
  } catch (error) {
    sendJson(res, 502, { error: error instanceof Error ? error.message : "Upstream request failed." }, cors);
  }
});

server.listen(PORT, () => {
  console.log("NVIDIA chat proxy listening on port " + PORT);
});
