import { getEndpoint } from "./registry";
import {
  BuildRequestInput,
  EndpointDefinition,
  HttpMethod,
  HttpRequestSpec,
  ODataMode,
  QueryPair,
  QueryValue,
} from "./types";
import { joinUrl, resolveTemplate, serializeQuery, toQueryPairs } from "./url";

export const DIGEST_PLACEHOLDER = "{{REQUEST_DIGEST}}";

const READ_METHODS: HttpMethod[] = ["GET", "HEAD", "OPTIONS"];

function acceptHeader(family: string, mode: ODataMode): string {
  if (family === "microsoft-graph") {
    return "application/json";
  }
  if (mode === "nometadata") {
    return "application/json;odata=nometadata";
  }
  if (mode === "minimal") {
    return "application/json;odata=minimalmetadata";
  }
  return "application/json;odata=verbose";
}

function contentTypeFor(
  endpoint: EndpointDefinition,
  mode: ODataMode,
  hasBody: boolean
): string | undefined {
  if (!hasBody) {
    return undefined;
  }
  const declared = endpoint.requestBody ? endpoint.requestBody.contentType : "";
  if (declared && declared.indexOf("multipart/") === 0) {
    return declared;
  }
  if (declared && declared.indexOf("application/json") !== 0) {
    return declared; // octet-stream, text/plain...
  }
  if (endpoint.family === "microsoft-graph") {
    return "application/json";
  }
  if (mode === "nometadata") {
    return "application/json;odata=nometadata";
  }
  if (mode === "minimal") {
    return "application/json;odata=minimalmetadata";
  }
  return "application/json;odata=verbose";
}

function isJsonContentType(contentType: string): boolean {
  return contentType.indexOf("application/json") === 0;
}

function serializeBody(body: unknown, contentType: string): string {
  if (typeof body === "string") {
    return body;
  }
  if (body === undefined) {
    return "";
  }
  if (isJsonContentType(contentType) || contentType === "") {
    return JSON.stringify(body);
  }
  if (typeof body === "number" || typeof body === "boolean") {
    return String(body);
  }
  // Non-JSON endpoints (file upload, batch) should be given a string body.
  return JSON.stringify(body);
}

/**
 * Builds a complete, inspectable HTTP request for a registry endpoint.
 *
 * The builder never sends anything. It is deterministic so the same input
 * always produces the same raw request, which makes baseline/mutation
 * comparison and Burp exports reliable.
 */
export function buildRequest(endpointId: string, input?: BuildRequestInput): HttpRequestSpec {
  const endpoint = getEndpoint(endpointId);
  const options: BuildRequestInput = input || {};
  const notes: string[] = [];

  const intendedMethod: HttpMethod =
    options.method ||
    (endpoint.methods.length > 0 ? endpoint.methods[0] : "GET");

  if (endpoint.methods.indexOf(intendedMethod) < 0) {
    notes.push(
      "Method " + intendedMethod + " is not declared for this endpoint (declared: " +
        endpoint.methods.join(", ") + "). Treat this request as a controlled probe."
    );
  }

  const odataMode: ODataMode = options.odataMode || "verbose";
  const resolved = resolveTemplate(endpoint.urlTemplate, options.path, endpoint.id);
  const pathKeys = Object.keys(options.path || {});
  for (let i = 0; i < pathKeys.length; i++) {
    if (!resolved.used[pathKeys[i]]) {
      notes.push("Path value '" + pathKeys[i] + "' was supplied but is not used by the template.");
    }
  }

  const queryPairs: QueryPair[] = [];
  appendAll(queryPairs, toQueryPairs(options.query));
  appendAll(queryPairs, options.queryPairs || []);

  const queryString = serializeQuery(
    queryPairs,
    options.encodeQueryValues === undefined ? true : options.encodeQueryValues
  );
  const rawQuery = options.rawQuery || "";
  const combinedQuery = [queryString, rawQuery].filter(function (part) {
    return !!part;
  }).join("&");

  const url = joinUrl(options.siteUrl || "", resolved.path) + (combinedQuery ? "?" + combinedQuery : "");

  const headers: Record<string, string> = {};
  headers["Accept"] = acceptHeader(endpoint.family, odataMode);

  const bodyText =
    options.body === undefined ? undefined : serializeBody(options.body, endpoint.requestBody ? endpoint.requestBody.contentType : "");
  if (bodyText !== undefined) {
    const contentType = contentTypeFor(endpoint, odataMode, true);
    if (contentType) {
      headers["Content-Type"] = contentType;
    }
    if (READ_METHODS.indexOf(intendedMethod) >= 0) {
      notes.push("A request body was supplied for a method that normally has no body.");
    }
  }

  const needsDigest =
    (endpoint.requiresDigest || endpoint.auth === "cookie+digest") &&
    READ_METHODS.indexOf(intendedMethod) < 0;
  if (needsDigest) {
    const placeholderEnabled = options.digestPlaceholder !== false;
    if (options.digest) {
      headers["X-RequestDigest"] = options.digest;
    } else if (placeholderEnabled) {
      headers["X-RequestDigest"] = DIGEST_PLACEHOLDER;
      notes.push(
        "X-RequestDigest is a placeholder. The executor replaces it; exported requests stay redacted unless you choose otherwise."
      );
    } else {
      headers["X-RequestDigest"] = "*";
      notes.push("No digest supplied; '*' is sent. SharePoint will normally reject this write.");
    }
  }

  if (endpoint.conditional) {
    const defaultIfMatch = options.defaultIfMatch === undefined ? "*" : options.defaultIfMatch;
    const shouldDefaultIfMatch =
      !!defaultIfMatch &&
      options.ifMatch === undefined &&
      endpoint.destructive &&
      (intendedMethod === "PATCH" || intendedMethod === "PUT" || intendedMethod === "DELETE" || intendedMethod === "MERGE");
    if (shouldDefaultIfMatch) {
      headers["If-Match"] = defaultIfMatch;
      notes.push("If-Match defaulted to '" + defaultIfMatch + "'.");
    }
    if (options.ifMatch !== undefined) {
      headers["If-Match"] = options.ifMatch;
    }
    if (options.ifNoneMatch !== undefined) {
      headers["If-None-Match"] = options.ifNoneMatch;
    }
  }

  let sentMethod: HttpMethod = intendedMethod;
  const useOverride = options.useMethodOverride !== false;
  if (intendedMethod === "MERGE") {
    if (useOverride) {
      sentMethod = "POST";
      headers["X-HTTP-Method"] = "MERGE";
      notes.push("MERGE sent as POST with X-HTTP-Method: MERGE (SharePoint convention).");
    } else {
      notes.push("MERGE sent as the literal request method.");
    }
  }

  const extraHeaders = options.headers || {};
  const extraNames = Object.keys(extraHeaders);
  for (let i = 0; i < extraNames.length; i++) {
    headers[extraNames[i]] = extraHeaders[extraNames[i]];
  }

  if (endpoint.requiredHeaders) {
    for (let i = 0; i < endpoint.requiredHeaders.length; i++) {
      const name = endpoint.requiredHeaders[i];
      if (!headers[name]) {
        notes.push("Endpoint declares required header '" + name + "' but no value was supplied.");
      }
    }
  }

  return {
    method: sentMethod,
    intendedMethod: intendedMethod,
    url: url,
    headers: headers,
    body: bodyText,
    endpointId: endpoint.id,
    family: endpoint.family,
    requiresDigest: needsDigest,
    destructive: endpoint.destructive,
    buildNotes: notes,
  };
}

function appendAll(target: QueryPair[], source: QueryPair[]): void {
  for (let i = 0; i < source.length; i++) {
    target.push(source[i]);
  }
}

/** Convenience: build a request for a specific list/item context. */
export function buildListItemRequest(
  endpointId: string,
  siteUrl: string,
  listId: string,
  itemId: number,
  extra?: BuildRequestInput
): HttpRequestSpec {
  const base: BuildRequestInput = {
    siteUrl: siteUrl,
    path: { listId: listId, itemId: itemId },
  };
  return buildRequest(endpointId, mergeInput(base, extra));
}

function mergeInput(base: BuildRequestInput, extra?: BuildRequestInput): BuildRequestInput {
  if (!extra) {
    return base;
  }
  const merged: BuildRequestInput = {
    siteUrl: extra.siteUrl !== undefined ? extra.siteUrl : base.siteUrl,
    path: mergeRecords(base.path, extra.path),
    query: mergeQuery(base.query, extra.query),
    queryPairs: (base.queryPairs || []).concat(extra.queryPairs || []),
    rawQuery: extra.rawQuery !== undefined ? extra.rawQuery : base.rawQuery,
    body: extra.body !== undefined ? extra.body : base.body,
    method: extra.method,
    headers: mergeRecords(base.headers, extra.headers),
    digest: extra.digest,
    odataMode: extra.odataMode,
    ifMatch: extra.ifMatch,
    ifNoneMatch: extra.ifNoneMatch,
    encodeQueryValues: extra.encodeQueryValues,
    digestPlaceholder: extra.digestPlaceholder,
    useMethodOverride: extra.useMethodOverride,
    defaultIfMatch: extra.defaultIfMatch,
  };
  return merged;
}

function mergeRecords<T>(
  base: Record<string, T> | undefined,
  extra: Record<string, T> | undefined
): Record<string, T> | undefined {
  if (!base && !extra) {
    return undefined;
  }
  const result: Record<string, T> = {};
  const baseKeys = Object.keys(base || {});
  for (let i = 0; i < baseKeys.length; i++) {
    result[baseKeys[i]] = (base as Record<string, T>)[baseKeys[i]];
  }
  const extraKeys = Object.keys(extra || {});
  for (let i = 0; i < extraKeys.length; i++) {
    result[extraKeys[i]] = (extra as Record<string, T>)[extraKeys[i]];
  }
  return result;
}

function mergeQuery(
  base: Record<string, QueryValue> | undefined,
  extra: Record<string, QueryValue> | undefined
): Record<string, QueryValue> | undefined {
  return mergeRecords<QueryValue>(base, extra);
}
