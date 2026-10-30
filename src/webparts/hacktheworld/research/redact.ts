/**
 * Secret redaction helpers. Everything that can be exported, logged or shown
 * in the UI goes through these functions first.
 */

export const REDACTED = "[REDACTED]";

const SECRET_HEADER_NAMES = [
  "authorization",
  "cookie",
  "set-cookie",
  "x-requestdigest",
  "x-request-digest",
  "x-api-key",
  "x-amz-security-token",
  "proxy-authorization",
];

/** Query / JSON keys that carry credentials. `$skiptoken` is deliberately allowed. */
const SECRET_KEY_PATTERN =
  /(^|[^a-z])(password|passwd|pwd|secret|clientsecret|client_secret|accesstoken|access_token|refreshtoken|refresh_token|idtoken|id_token|bearer|authorization|cookie|digest|requestdigest|saml|assertion|credential|apikey|api_key|accesskey|access_key|sig|signature)($|[^a-z])/i;

const SAFE_KEY_PATTERN = /^\$|^(skiptoken|changetoken|viewtoken|formdigesttimeoutseconds)$/i;

export function isSecretKey(key: string): boolean {
  if (!key || SAFE_KEY_PATTERN.test(key)) {
    return false;
  }
  return SECRET_KEY_PATTERN.test(key);
}

export function redactHeaders(headers: Record<string, string>): Record<string, string> {
  const result: Record<string, string> = {};
  const names = Object.keys(headers);
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    result[name] = SECRET_HEADER_NAMES.indexOf(name.toLowerCase()) >= 0
      ? REDACTED
      : headers[name];
  }
  return result;
}

/**
 * Redacts credential-looking query parameters while keeping the rest of the
 * URL readable for manual comparison in Burp.
 */
export function redactUrl(url: string): string {
  const hashIndex = url.indexOf("#");
  const withoutHash = hashIndex >= 0 ? url.substring(0, hashIndex) : url;
  const qIndex = withoutHash.indexOf("?");
  if (qIndex < 0) {
    return url;
  }
  const base = withoutHash.substring(0, qIndex);
  const query = withoutHash.substring(qIndex + 1);
  const parts = query.split("&");
  for (let i = 0; i < parts.length; i++) {
    const eq = parts[i].indexOf("=");
    if (eq < 0) {
      continue;
    }
    const rawName = parts[i].substring(0, eq);
    const name = decodeURIComponentQuietly(rawName);
    if (isSecretKey(name)) {
      parts[i] = rawName + "=" + REDACTED;
    }
  }
  const suffix = hashIndex >= 0 ? url.substring(hashIndex) : "";
  return base + "?" + parts.join("&") + suffix;
}

/** Recursively redacts JSON values. Handles cyclic-free JSON data only. */
export function redactJson(value: unknown, depth: number = 0): unknown {
  if (depth > 12 || value === null || typeof value !== "object") {
    return value;
  }
  if (value instanceof Array) {
    const out: unknown[] = [];
    for (let i = 0; i < value.length; i++) {
      out.push(redactJson(value[i], depth + 1));
    }
    return out;
  }
  const source = value as Record<string, unknown>;
  const out: Record<string, unknown> = {};
  const keys = Object.keys(source);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    out[key] = isSecretKey(key) ? REDACTED : redactJson(source[key], depth + 1);
  }
  return out;
}

/** Best-effort text redaction for logs and error messages. */
export function redactText(text: string): string {
  if (!text) {
    return text;
  }
  return text
    .replace(/Bearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer " + REDACTED)
    .replace(/("(?:access_token|refresh_token|id_token|password|client_secret)"\s*:\s*")[^"]*(")/gi, "$1" + REDACTED + "$2")
    .replace(/(X-RequestDigest\s*:\s*)[^\s,]+/gi, "$1" + REDACTED);
}

/** Truncates a response body for evidence storage. */
export function preview(text: string, maxLength: number = 2000): string {
  if (!text) {
    return "";
  }
  if (text.length <= maxLength) {
    return text;
  }
  return text.substring(0, maxLength) + "\n… [" + (text.length - maxLength) + " chars truncated]";
}

function decodeURIComponentQuietly(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch (error) {
    return value;
  }
}
