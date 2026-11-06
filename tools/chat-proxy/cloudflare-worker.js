/**
 * NVIDIA chat proxy for the SharePoint chatbot (Cloudflare Worker).
 *
 * The browser must never hold the NVIDIA API key, and integrate.api.nvidia.com
 * does not send CORS headers, so a direct call from SharePoint is blocked.
 * This worker adds the key server-side and returns CORS headers.
 *
 * Setup:
 *   1. Create a Worker and paste this file.
 *   2. Add an encrypted secret named NVIDIA_API_KEY.
 *   3. Optionally add ALLOWED_ORIGINS, e.g.
 *      "https://contoso.sharepoint.com,https://contoso-my.sharepoint.com".
 *   4. Deploy and paste the worker URL into the chatbot settings.
 */

const UPSTREAM = "https://integrate.api.nvidia.com/v1/chat/completions";

function corsHeaders(origin, allowed) {
  const originAllowed = allowed.length === 0 || allowed.indexOf(origin) >= 0;
  return {
    "Access-Control-Allow-Origin": originAllowed ? (origin || "*") : "null",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Accept",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin",
  };
}

function json(body, status, cors) {
  return new Response(JSON.stringify(body), {
    status: status,
    headers: Object.assign({ "Content-Type": "application/json", "Cache-Control": "no-store" }, cors),
  });
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get("Origin") || "";
    const allowed = (env.ALLOWED_ORIGINS || "")
      .split(",")
      .map((value) => value.trim().replace(/\/+$/, ""))
      .filter(Boolean);
    const cors = corsHeaders(origin.replace(/\/+$/, ""), allowed);

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: cors });
    }
    if (request.method !== "POST") {
      return json({ error: "Use POST." }, 405, cors);
    }
    if (allowed.length > 0 && allowed.indexOf(origin.replace(/\/+$/, "")) < 0) {
      return json({ error: "Origin not allowed." }, 403, cors);
    }
    if (!env.NVIDIA_API_KEY) {
      return json({ error: "NVIDIA_API_KEY is not configured on the worker." }, 500, cors);
    }

    let body;
    try {
      body = await request.json();
    } catch (_error) {
      return json({ error: "Request body must be JSON." }, 400, cors);
    }

    const upstream = await fetch(UPSTREAM, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Accept": request.headers.get("Accept") || "text/event-stream",
        "Authorization": "Bearer " + env.NVIDIA_API_KEY,
      },
      body: JSON.stringify(body),
    });

    const headers = Object.assign({}, cors);
    headers["Content-Type"] = upstream.headers.get("Content-Type") || "application/json";
    headers["Cache-Control"] = "no-store";

    // Pass the body through untouched so SSE streaming keeps working.
    return new Response(upstream.body, { status: upstream.status, headers: headers });
  },
};
