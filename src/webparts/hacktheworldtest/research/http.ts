import { HttpRequestSpec } from "./types";
import { getOrigin, getPathAndQuery } from "./url";

/**
 * Renders a request as raw HTTP for Burp Suite / other manual tools.
 *
 * This is shared by the request printer and the case-plan exporter so both
 * produce byte-compatible request lines and header ordering.
 */
export function toRawHttp(spec: HttpRequestSpec): string {
  const origin = getOrigin(spec.url);
  const lines: string[] = [];
  lines.push(spec.method + " " + getPathAndQuery(spec.url) + " HTTP/1.1");
  lines.push("Host: " + origin.replace(/^https?:\/\//, ""));
  const names = Object.keys(spec.headers);
  for (let i = 0; i < names.length; i++) {
    lines.push(names[i] + ": " + spec.headers[names[i]]);
  }
  lines.push("");
  if (spec.body !== undefined) {
    lines.push(spec.body);
  }
  return lines.join("\r\n");
}

/** Human-readable comment block used in exported plans. */
export function describeRequest(spec: HttpRequestSpec): string[] {
  const lines: string[] = [];
  lines.push("# " + spec.endpointId + " (" + spec.family + ") intended=" + spec.intendedMethod);
  for (let i = 0; i < spec.buildNotes.length; i++) {
    lines.push("# note: " + spec.buildNotes[i]);
  }
  lines.push("# This request was NOT sent. Paste it into Burp against an authorized tenant.");
  return lines;
}
