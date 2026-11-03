/**
 * Secret redaction helpers. Everything that can be exported, logged or shown
 * in the UI goes through these functions first.
 */

export const REDACTED = "[REDACTED]";

const SECRET_HEADER_NAMES = [
  "authorization",
  "proxy-authorization",
  "cookie",
  "set-cookie",
  "x-requestdigest",
  "x-request-digest",
  "x-api-key",
  "x-amz-security-token",
  "x-auth-token",
  "x-csrf-token",
  "x-xsrf-token",
];

/** Query / JSON / header keys that carry credentials. `$skiptoken` is deliberately allowed. */
const SECRET_KEY_PATTERN =
  /(^|[^a-z])(password|passwd|pwd|passphrase|secret|clientsecret|client_secret|accesstoken|access_token|refreshtoken|refresh_token|idtoken|id_token|bearer|authorization|cookie|digest|requestdigest|saml|samlresponse|assertion|credential|credentials|apikey|api_key|accesskey|access_key|sig|signature|authtoken|auth_token|xsrf|csrf|sessiontoken|session_token)($|[^a-z])/i;

const SAFE_KEY_PATTERN =
  /^\$|^(skiptoken|changetoken|viewtoken|formdigesttimeoutseconds|samlrequest)$/i;

/**
 * Keys in free text look like `name=value` or `name: value`. Keep the key and
 * the separator so the log stays readable, replace only the value.
 */
const SECRET_KEY_SOURCE =
  "password|passwd|pwd|passphrase|secret|client[_-]?secret|access[_-]?token|refresh[_-]?token|id[_-]?token|bearer|authorization|cookie|digest|requestdigest|saml|samlresponse|assertion|credential|credentials|api[_-]?key|access[_-]?key|sig|signature|auth[_-]?token|xsrf|csrf|session[_-]?token";

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
 *
 * Also removes embedded userinfo (`https://user:pass@host/...`) because that is
 * a credential even though it is not a query parameter.
 */
export function redactUrl(url: string): string {
  if (!url) {
    return url;
  }
  const hashIndex = url.indexOf("#");
  const withoutHash = hashIndex >= 0 ? url.substring(0, hashIndex) : url;
  const qIndex = withoutHash.indexOf("?");
  const base = qIndex >= 0 ? withoutHash.substring(0, qIndex) : withoutHash;
  const query = qIndex >= 0 ? withoutHash.substring(qIndex + 1) : "";
  const sanitizedBase = base.replace(/^(https?:\/\/)([^/@\s]+)@/i, "$1" + REDACTED + "@");

  if (qIndex < 0) {
    return sanitizedBase + (hashIndex >= 0 ? url.substring(hashIndex) : "");
  }

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
  return sanitizedBase + "?" + parts.join("&") + suffix;
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

/**
 * Best-effort text redaction for logs and error messages.
 *
 * Covers header lines, cookie values, bearer/basic credentials, JSON members
 * and `key=value` / `key: value` pairs. The function is intentionally
 * idempotent: redacting an already-redacted string is a no-op.
 */
export function redactText(text: string): string {
  if (!text) {
    return text;
  }

  let output = text;

  // Header lines whose whole value is sensitive (Cookie: a=b; c=d).
  output = output.replace(
    /((?:proxy-)?authorization\s*:\s*)[^\r\n]+/gi,
    "$1" + REDACTED
  );
  output = output.replace(
    /((?:set-)?cookie\s*:\s*)[^\r\n]+/gi,
    "$1" + REDACTED
  );
  output = output.replace(
    /((?:x-requestdigest|x-request-digest)\s*:\s*)[^\r\n]+/gi,
    "$1" + REDACTED
  );

  // Authorization schemes can also appear without a header name.
  output = output.replace(/\b(Bearer|Basic)\s+[A-Za-z0-9._~+/=-]+/gi, "$1 " + REDACTED);

  // JSON members: {"password":"hunter2"}. Keep the key so triage stays readable.
  output = output.replace(
    new RegExp('("(?:' + SECRET_KEY_SOURCE + ')"\\s*:\\s*")[^"]*(")', "gi"),
    "$1" + REDACTED + "$2"
  );

  // Query-string / form / log pairs: password=hunter2, digest=abc, token: xyz.
  output = output.replace(
    new RegExp("((?:^|[?&;,\\s])(?:" + SECRET_KEY_SOURCE + ")\\s*[=:]\\s*)(\"?)([^\"&;,\\s]+)(\\2)", "gi"),
    function (_match: string, prefix: string, quote: string, _value: string, closing: string): string {
      return prefix + quote + REDACTED + closing;
    }
  );

  return output;
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
