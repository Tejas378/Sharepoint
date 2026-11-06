/**
 * Local-only chat server for an OpenAI-compatible NVIDIA endpoint.
 *
 * - Serves the static UI from ./public
 * - Proxies POST /api/chat to NVIDIA, adding the key server-side
 * - Streams the upstream SSE response straight back to the browser
 *
 * No dependencies. Node 18+.
 */

const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const { Readable } = require("node:stream");
const { pipeline } = require("node:stream/promises");

const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, "public");
const UPSTREAM = process.env.NVIDIA_CHAT_URL || "https://integrate.api.nvidia.com/v1/chat/completions";
const MAX_BODY_BYTES = 30 * 1024 * 1024;

loadEnv(path.join(ROOT, ".env"));

const PORT = Number(process.env.PORT || 4173);
const HOST = process.env.HOST || "127.0.0.1";
const API_KEY = process.env.NVIDIA_API_KEY || "";
const UPSTREAM_TIMEOUT_MS = Number(process.env.CHAT_TIMEOUT_MS || 180000);

const UPSTREAM_HOST = (() => {
  try {
    return new URL(UPSTREAM).host;
  } catch (_error) {
    return UPSTREAM;
  }
})();

function log(message) {
  const now = new Date();
  const stamp = now.toTimeString().slice(0, 8);
  console.log("[" + stamp + "] " + message);
}

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".ico": "image/x-icon",
};

function loadEnv(file) {
  try {
    const text = fs.readFileSync(file, "utf8");
    text.split(/\r?\n/).forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed || trimmed.indexOf("#") === 0) return;
      const split = trimmed.indexOf("=");
      if (split < 1) return;
      const key = trimmed.slice(0, split).trim();
      let value = trimmed.slice(split + 1).trim();
      if ((value[0] === '"' && value[value.length - 1] === '"') ||
          (value[0] === "'" && value[value.length - 1] === "'")) {
        value = value.slice(1, -1);
      }
      if (key && process.env[key] === undefined) process.env[key] = value;
    });
  } catch (_error) {
    /* .env is optional */
  }
}

function sendJson(res, status, body) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  res.end(JSON.stringify(body));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > MAX_BODY_BYTES) {
        reject(new Error("Request body is too large (limit 30 MB)."));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function serveStatic(req, res, pathname) {
  const relative = pathname === "/" ? "index.html" : pathname.replace(/^\/+/, "");
  const resolved = path.resolve(PUBLIC_DIR, relative);
  const prefix = PUBLIC_DIR + path.sep;
  if (resolved !== path.join(PUBLIC_DIR, "index.html") && resolved.indexOf(prefix) !== 0) {
    sendJson(res, 403, { error: "Forbidden." });
    return;
  }
  fs.readFile(resolved, (error, data) => {
    if (error) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": MIME[path.extname(resolved).toLowerCase()] || "application/octet-stream",
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
}

async function handleChat(req, res) {
  if (!API_KEY) {
    sendJson(res, 500, { error: "NVIDIA_API_KEY is missing. Add it to local-chat/.env and restart the server." });
    return;
  }

  let payload;
  try {
    const raw = await readBody(req);
    payload = JSON.parse(raw || "{}");
  } catch (error) {
    sendJson(res, 400, { error: error instanceof Error ? error.message : "Invalid JSON body." });
    return;
  }

  const controller = new AbortController();
  let timedOut = false;
  let finished = false;
  const timer = setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, UPSTREAM_TIMEOUT_MS);
  const started = Date.now();

  // Stop paying for a reply nobody will read: cancel upstream if the browser
  // closes the tab, reloads, or presses Stop. `finished` keeps keep-alive
  // socket cleanup from being reported as a client abort.
  res.on("close", () => {
    if (!finished && !res.writableEnded) {
      log("!! client closed the connection, cancelling upstream");
      controller.abort();
    }
  });
  const model = payload && payload.model ? String(payload.model) : "(no model)";
  log("-> POST " + UPSTREAM + " model=" + model + " stream=" + !!(payload && payload.stream));

  try {
    const upstream = await fetch(UPSTREAM, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Accept": req.headers.accept || "text/event-stream",
        "Authorization": "Bearer " + API_KEY,
      },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });
    clearTimeout(timer);
    log("<- HTTP " + upstream.status + " " + (upstream.headers.get("content-type") || "") +
      " after " + (Date.now() - started) + " ms");

    res.writeHead(upstream.status, {
      "Content-Type": upstream.headers.get("content-type") || "application/json",
      "Cache-Control": "no-store",
      "X-Accel-Buffering": "no",
    });

    if (upstream.body) {
      await pipeline(Readable.fromWeb(upstream.body), res);
      finished = true;
      log("<> stream finished after " + (Date.now() - started) + " ms");
    } else {
      res.end();
      finished = true;
    }
  } catch (error) {
    finished = true;
    clearTimeout(timer);
    const clientGone = controller.signal.aborted && !timedOut;
    const message = timedOut
      ? "NVIDIA did not respond within " + Math.round(UPSTREAM_TIMEOUT_MS / 1000) + "s. " +
        "Try again, lower Max tokens, or pick another model in Settings."
      : error instanceof Error
        ? error.message
        : "Upstream request failed.";
    if (!clientGone) {
      log("!! " + (timedOut ? "timeout" : "error") + " after " + (Date.now() - started) + " ms: " + message);
    }
    if (!res.headersSent) {
      if (!clientGone) sendJson(res, timedOut ? 504 : 502, { error: message });
    } else if (!res.writableEnded) {
      res.end();
    }
  }
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url || "/", "http://" + (req.headers.host || "localhost"));

  if (url.pathname.indexOf("/api/") === 0) {
    log(req.method + " " + url.pathname);
  }

  if (url.pathname === "/api/health") {
    sendJson(res, 200, {
      ok: true,
      keyConfigured: !!API_KEY,
      upstream: UPSTREAM,
      upstreamHost: UPSTREAM_HOST,
      timeoutMs: UPSTREAM_TIMEOUT_MS,
    });
    return;
  }

  if (url.pathname === "/api/chat") {
    if (req.method !== "POST") {
      sendJson(res, 405, { error: "Use POST." });
      return;
    }
    handleChat(req, res);
    return;
  }

  if (req.method !== "GET" && req.method !== "HEAD") {
    sendJson(res, 405, { error: "Use GET." });
    return;
  }
  serveStatic(req, res, url.pathname);
});

server.listen(PORT, HOST, () => {
  const shown = HOST === "0.0.0.0" ? "localhost" : HOST;
  console.log("");
  console.log("  Local NVIDIA chat");
  console.log("  URL:      http://" + shown + ":" + PORT);
  console.log("  Upstream: " + UPSTREAM);
  console.log("  Key:      " + (API_KEY ? "loaded from .env" : "MISSING - add NVIDIA_API_KEY to local-chat/.env"));
  console.log("  Timeout:  " + Math.round(UPSTREAM_TIMEOUT_MS / 1000) + "s");
  console.log("  Stop:     Ctrl+C");
  console.log("");
  log("ready");
});
