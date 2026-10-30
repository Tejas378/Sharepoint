import { QueryPair, QueryValue } from "./types";

/** Removes a trailing slash so joining is predictable. */
export function trimTrailingSlash(value: string): string {
  if (!value) {
    return "";
  }
  return value.charAt(value.length - 1) === "/" ? value.substring(0, value.length - 1) : value;
}

export function joinUrl(base: string, path: string): string {
  if (!base) {
    return path;
  }
  if (!path) {
    return base;
  }
  if (/^https?:\/\//i.test(path)) {
    return path;
  }
  const left = trimTrailingSlash(base);
  const right = path.charAt(0) === "/" ? path : "/" + path;
  return left + right;
}

/** Escapes a raw value for an OData single-quoted string literal. */
export function odataString(value: string): string {
  return "'" + String(value).replace(/'/g, "''") + "'";
}

/** Builds guid'xxxxxxxx-...' with a validated GUID. */
export function odataGuid(value: string): string {
  const normalized = String(value).replace(/[{}]/g, "").toLowerCase();
  if (!isGuid(normalized)) {
    throw new Error("Expected a GUID, received: " + value);
  }
  return "guid'" + normalized + "'";
}

export function isGuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

/** True when the value is a 32 digit GUID without dashes. */
export function isCompactGuid(value: string): boolean {
  return /^[0-9a-f]{32}$/i.test(value);
}

/**
 * Encodes a query value while keeping common OData punctuation readable.
 * The encoded output is still valid for SharePoint and easier to review in Burp.
 */
export function encodeQueryValue(value: string): string {
  return encodeURIComponent(value)
    .replace(/%2C/gi, ",")
    .replace(/%28/gi, "(")
    .replace(/%29/gi, ")")
    .replace(/%27/gi, "'")
    .replace(/%3A/gi, ":")
    .replace(/%20/g, "%20");
}

/** Encodes a path segment (file name, folder name, list path). */
export function encodePathSegment(value: string): string {
  return encodeURIComponent(value);
}

/** Keeps OData option names ($select) and aliases (@a1) readable. */
export function encodeQueryName(value: string): string {
  return encodeURIComponent(value).replace(/%24/gi, "$").replace(/%40/gi, "@");
}

const LIST_QUERY_OPTIONS = ["$select", "$expand", "$orderby"];

/** Normalises the convenience query map into ordered pairs. */
export function toQueryPairs(query?: Record<string, QueryValue>): QueryPair[] {
  const pairs: QueryPair[] = [];
  if (!query) {
    return pairs;
  }
  const names = Object.keys(query);
  for (let i = 0; i < names.length; i++) {
    pairs.push({ name: names[i], value: query[names[i]] });
  }
  return pairs;
}

/**
 * Serialises query pairs. Array values are joined with "," for $select /
 * $expand / $orderby, and repeated as separate parameters otherwise.
 */
export function serializeQuery(
  pairs: QueryPair[],
  encodeValues: boolean = true
): string {
  const parts: string[] = [];
  for (let i = 0; i < pairs.length; i++) {
    const pair = pairs[i];
    if (pair.value === undefined || pair.value === null) {
      continue;
    }
    const values = pair.value instanceof Array ? pair.value : [pair.value];
    if (values.length === 0) {
      continue;
    }
    const joinWithComma = LIST_QUERY_OPTIONS.indexOf(pair.name) >= 0;
    const serialized: string[] = [];
    for (let j = 0; j < values.length; j++) {
      const raw = String(values[j]);
      serialized.push(!encodeValues || pair.raw ? raw : encodeQueryValue(raw));
    }
    parts.push(
      encodeQueryName(pair.name) + "=" + serialized.join(joinWithComma ? "," : "&" + encodeQueryName(pair.name) + "=")
    );
  }
  return parts.join("&");
}

export interface PlaceholderResolution {
  path: string;
  used: Record<string, boolean>;
}

/**
 * Resolves {name} / {name:modifier} placeholders in a URL template.
 * Throws a descriptive error for missing values so failed generations are
 * obvious instead of producing silently broken URLs.
 */
export function resolveTemplate(
  template: string,
  values: Record<string, string | number> | undefined,
  endpointId: string
): PlaceholderResolution {
  const used: Record<string, boolean> = {};
  const source = values === undefined ? {} : values;
  const resolved = template.replace(/\{([A-Za-z0-9_]+)(?::([A-Za-z]+))?\}/g, function (
    match: string,
    name: string,
    modifier?: string
  ): string {
    const hasValue = Object.prototype.hasOwnProperty.call(source, name);
    if (!hasValue) {
      throw new Error(
        "Missing path placeholder {" + name + "} while building endpoint '" + endpointId + "'"
      );
    }
    used[name] = true;
    const value = source[name];
    return applyModifier(String(value), modifier, endpointId, name);
  });
  return { path: resolved, used: used };
}

function applyModifier(
  value: string,
  modifier: string | undefined,
  endpointId: string,
  name: string
): string {
  switch (modifier) {
    case undefined:
    case "":
    case "raw":
      return value;
    case "odata":
      return odataString(value);
    case "odataEncoded":
      return encodeURIComponent(odataString(value));
    case "guid":
      return odataGuid(value);
    case "path":
      return encodePathSegment(value);
    case "int":
      if (!/^-?\d+$/.test(value)) {
        throw new Error(
          "Placeholder {" + name + ":int} for '" + endpointId + "' expected an integer, received: " + value
        );
      }
      return value;
    default:
      throw new Error(
        "Unknown placeholder modifier ':" + modifier + "' in endpoint '" + endpointId + "'"
      );
  }
}

/** Returns the origin (scheme + host + port) of a URL. */
export function getOrigin(url: string): string {
  const match = /^(https?:\/\/[^/?#]+)/i.exec(url);
  return match ? match[1] : "";
}

/** Returns path + query without the origin, which is what Burp request lines use. */
export function getPathAndQuery(url: string): string {
  const match = /^https?:\/\/[^/?#]+(\/[^#]*)?/i.exec(url);
  if (match) {
    return match[1] || "/";
  }
  return url.charAt(0) === "/" ? url : "/" + url;
}

/**
 * Rewrites a legacy absolute URL onto the configured tenant.
 *
 * The historical services embed the original tenant host in hundreds of string
 * literals. Instead of editing every literal (and risking a typo that points a
 * probe at the wrong host), the request path normalises known legacy hosts to
 * the configured site URL. Only hosts the caller explicitly lists are rewritten,
 * so an unrelated absolute URL is never silently redirected.
 */
export function rewriteTenantUrl(
  url: string,
  siteUrl: string,
  legacyHosts: string[]
): string {
  if (!url || !/^https?:\/\//i.test(url)) {
    return url;
  }
  const origin = getOrigin(url);
  const targetOrigin = getOrigin(siteUrl);
  if (!origin || !targetOrigin || origin.toLowerCase() === targetOrigin.toLowerCase()) {
    return url;
  }
  const host = origin.replace(/^https?:\/\//i, "").toLowerCase();
  for (let i = 0; i < legacyHosts.length; i++) {
    if (legacyHosts[i].toLowerCase() === host) {
      return targetOrigin + url.substring(origin.length);
    }
  }
  return url;
}
