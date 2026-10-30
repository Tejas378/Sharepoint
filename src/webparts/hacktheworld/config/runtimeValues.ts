import { loadTestConfig } from "./loader";

export type ConfiguredTimestamp =
  | "futureExpiration"
  | "pastExpiration"
  | "createdDate"
  | "modifiedDate"
  | "lastModified"
  | "lastUserModified";

/** Read a timestamp resolved from a relative marker in test-config.json. */
export function configuredTimestamp(name: ConfiguredTimestamp): string {
  const config = loadTestConfig();
  const constants = config.testData.constants as { dates?: Record<string, unknown> } | undefined;
  const value = constants && constants.dates && constants.dates[name];
  if (typeof value !== "string") {
    throw new Error("Missing testData.constants.dates." + name + " in test-config.json.");
  }
  return value;
}

export function configuredSiteUrl(): string {
  return loadTestConfig().environment.siteUrl;
}

export function configuredSiteHost(): string {
  return new URL(configuredSiteUrl()).host;
}
