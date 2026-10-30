import { HttpRequestSpec, Transport, TransportResponse } from "./types";

/** Supplies a short-lived Microsoft Graph access token. Keep token acquisition
 * outside the request engine so callers can use MSAL or another approved flow. */
export type GraphTokenProvider = () => Promise<string>;

/** Create a Graph-only transport. It never sends SharePoint cookies or digests. */
export function createGraphTransport(
  getAccessToken: GraphTokenProvider,
  fetchImpl?: typeof fetch
): Transport {
  if (typeof getAccessToken !== "function") {
    throw new Error("Graph transport requires an access-token provider.");
  }
  const doFetch = fetchImpl || fetch;
  return async function (spec: HttpRequestSpec): Promise<TransportResponse> {
    if (spec.family !== "microsoft-graph") {
      throw new Error("Graph transport only accepts microsoft-graph requests.");
    }
    const url = new URL(spec.url);
    if (url.protocol !== "https:" || url.hostname.toLowerCase() !== "graph.microsoft.com") {
      throw new Error("Graph transport only permits HTTPS requests to graph.microsoft.com.");
    }
    const token = await getAccessToken();
    if (!token || /\s/.test(token)) {
      throw new Error("Graph token provider returned an empty or invalid token.");
    }
    const headers: Record<string, string> = {};
    Object.keys(spec.headers).forEach(function (key) {
      if (key.toLowerCase() !== "authorization" && key.toLowerCase() !== "x-requestdigest") {
        headers[key] = spec.headers[key];
      }
    });
    headers.Authorization = "Bearer " + token;
    const response = await doFetch(url.toString(), {
      method: spec.method,
      headers: headers,
      body: spec.body,
    });
    const responseHeaders: Record<string, string> = {};
    response.headers.forEach(function (value, key) {
      responseHeaders[key.toLowerCase()] = value;
    });
    return {
      status: response.status,
      statusText: response.statusText,
      headers: responseHeaders,
      bodyText: await response.text(),
    };
  };
}
