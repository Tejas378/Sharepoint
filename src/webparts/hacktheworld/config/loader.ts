import rawConfig from "../../../config/test-config.json";

/**
 * Centralised test configuration.
 *
 * This module is the single source of truth for SharePoint identifiers used by
 * the web part and the research toolkit. It deliberately lives outside
 * research/ because it is not SPFx-specific, but it has no Node dependencies so
 * the same code runs in the browser, in tests and in tools/validate-config.mjs
 * (which re-implements only the CLI shell, not the rules).
 *
 * Rules enforced here:
 *  - no credentials or tokens in the configuration file
 *  - identifiers are validated before anything is used
 *  - destructive capabilities are opt-in and host-scoped
 */

export type Severity = "error" | "warning";

export interface ConfigIssue {
  path: string;
  message: string;
  severity: Severity;
}

export interface ValidationResult {
  ok: boolean;
  issues: ConfigIssue[];
}

export interface ListConfig {
  id: string;
  title: string;
  template: number;
  baseType: number;
  itemEntityType?: string;
  serverRelativeUrl: string;
}

export interface UserConfig {
  upn: string;
  loginName: string;
  id: number;
}

export interface SafetyConfig {
  allowDestructive: boolean;
  allowPermissionChanges: boolean;
  allowOutboundWebhooks: boolean;
  allowLargePayloads: boolean;
  allowedHosts: string[];
  maxRequestsPerRun: number;
  maxPayloadBytes: number;
  requestTimeoutMs: number;
  minimumDelayBetweenRequestsMs: number;
  notes?: string;
}

export interface TestConfig {
  version: number;
  description?: string;
  environment: {
    name: string;
    tenantUrl: string;
    tenantId: string;
    siteUrl: string;
    siteCollectionId: string;
    siteId: string;
    webId: string;
    webUrl: string;
  };
  resources: {
    lists: Record<string, ListConfig>;
    items: {
      primaryItemId: number;
      secondaryItemId: number;
      nonExistentItemId: number;
      boundaryItemIds: number[];
      stringIds: string[];
    };
    folders: Record<string, string | number>;
    files: Record<string, string>;
    fields: Record<string, string>;
    contentTypes: Record<string, string>;
    views: Record<string, string>;
    attachments: Record<string, string>;
    pages: Record<string, string>;
  };
  principals: {
    users: { primary: UserConfig; secondary: UserConfig };
    groups: {
      primaryGroupId: number;
      primaryGroupName: string;
      ownerGroupId: number;
      memberGroupId: number;
      visitorGroupId: number;
      shareWith: string[];
      groupsToEnumerate: number[];
    };
    roleDefinitions: {
      read: number;
      contribute: number;
      edit: number;
      design: number;
      fullControl: number;
      restrictedRead: number;
      all: number[];
    };
    principalIds: number[];
    shareLinkKinds: number[];
    shareLinkRoleKinds: number[];
    scope: number[];
    permissionKinds: number[];
  };
  taxonomy: Record<string, string>;
  search: { queryText: string; rowLimit: number; sortList: string[] };
  testData: Record<string, unknown>;
  safety: SafetyConfig;
}

/** Deep-partial overrides supplied at runtime (property pane, tests, tools). */
export interface ConfigOverrides {
  environment?: Partial<TestConfig["environment"]>;
  resources?: {
    lists?: Record<string, Partial<ListConfig>>;
    items?: Partial<TestConfig["resources"]["items"]>;
    folders?: Record<string, string | number>;
    files?: Record<string, string>;
    fields?: Record<string, string>;
    contentTypes?: Record<string, string>;
    views?: Record<string, string>;
    attachments?: Record<string, string>;
    pages?: Record<string, string>;
  };
  principals?: {
    users?: { primary?: Partial<UserConfig>; secondary?: Partial<UserConfig> };
    groups?: Partial<TestConfig["principals"]["groups"]>;
    roleDefinitions?: Record<string, number | number[]>;
    principalIds?: number[];
    shareLinkKinds?: number[];
    shareLinkRoleKinds?: number[];
    scope?: number[];
    permissionKinds?: number[];
  };
  testData?: Record<string, unknown>;
  safety?: Partial<SafetyConfig>;
}

export class ConfigError extends Error {
  public readonly issues: ConfigIssue[];
  constructor(issues: ConfigIssue[]) {
    super(formatIssues(issues));
    this.name = "ConfigError";
    this.issues = issues;
  }
}

const SECRET_KEY_PATTERN = /^(password|passwd|pwd|secret|clientsecret|client_secret|accesstoken|access_token|refreshtoken|refresh_token|bearer|apikey|api_key|apiKey|authorization|cookie|connectionstring)$/i;
const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const HOST_PATTERN = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$/i;
const KNOWN_SECTIONS = [
  "version",
  "description",
  "$schema",
  "environment",
  "resources",
  "principals",
  "taxonomy",
  "search",
  "testData",
  "safety",
];

export function isGuid(value: unknown): boolean {
  return typeof value === "string" && GUID_PATTERN.test(value);
}

/** SharePoint accepts GUIDs with braces; the config normalises them away. */
export function normaliseGuid(value: string): string {
  return String(value).replace(/[{}]/g, "").toLowerCase();
}

export function isHttpsUrl(value: unknown): boolean {
  if (typeof value !== "string" || value.length === 0) {
    return false;
  }
  const match = /^https:\/\/([^/?#]+)(\/[^?#]*)?$/i.exec(value);
  return !!match;
}

export function urlHost(value: string): string {
  const match = /^https?:\/\/([^/?#]+)/i.exec(value);
  return match ? match[1].toLowerCase() : "";
}

export function isUpn(value: unknown): boolean {
  return typeof value === "string" && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(value);
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isPositiveInt(value: unknown): boolean {
  return typeof value === "number" && isFinite(value) && Math.floor(value) === value && value > 0;
}

function isNonNegativeInt(value: unknown): boolean {
  return typeof value === "number" && isFinite(value) && Math.floor(value) === value && value >= 0;
}

/** Rejects a URL that is outside the configured host allow-list. */
export function assertAllowedUrl(url: string, config: TestConfig): void {
  const host = urlHost(url).split(":")[0];
  if (url.toLowerCase().indexOf("https://") !== 0) {
    throw new Error("Refusing non-HTTPS target: " + url);
  }
  if (config.safety.allowedHosts.length === 0) {
    throw new Error("No hosts are allowed by safety.allowedHosts; refusing to send requests.");
  }
  if (config.safety.allowedHosts.indexOf(host) < 0) {
    throw new Error(
      "Refusing target host '" + host + "'. Allowed hosts: " + config.safety.allowedHosts.join(", ")
    );
  }
}

/** Guards destructive operations behind an explicit, host-scoped switch. */
export function assertDestructiveAllowed(config: TestConfig, operation: string): void {
  if (!config.safety.allowDestructive) {
    throw new Error(
      "Operation '" + operation + "' is destructive and safety.allowDestructive is false."
    );
  }
}

export function assertPermissionChangeAllowed(config: TestConfig, operation: string): void {
  assertDestructiveAllowed(config, operation);
  if (!config.safety.allowPermissionChanges) {
    throw new Error(
      "Operation '" + operation + "' changes permissions and safety.allowPermissionChanges is false."
    );
  }
}

export function assertWebhookAllowed(config: TestConfig): void {
  assertDestructiveAllowed(config, "create webhook subscription");
  if (!config.safety.allowOutboundWebhooks) {
    throw new Error("Creating webhooks is disabled (safety.allowOutboundWebhooks).");
  }
}

export function getPrimaryList(config: TestConfig): ListConfig {
  const list = config.resources.lists.primary;
  if (!list) {
    throw new ConfigError([
      { path: "resources.lists.primary", message: "Primary list is not configured.", severity: "error" },
    ]);
  }
  return list;
}

export function itemIdOrDefault(config: TestConfig, explicit?: number): number {
  return explicit !== undefined ? explicit : config.resources.items.primaryItemId;
}

/**
 * Validates an arbitrary object against the configuration contract.
 * Returns every issue instead of throwing, so tooling can print a full report.
 */
export function validateTestConfig(raw: unknown): ValidationResult {
  const issues: ConfigIssue[] = [];
  const add = (path: string, message: string, severity: Severity = "error"): void => {
    issues.push({ path, message, severity });
  };

  if (!isPlainObject(raw)) {
    add("", "Configuration must be a JSON object.");
    return { ok: false, issues };
  }

  const config = raw as Record<string, unknown>;

  // Unknown sections are usually typos; reject them so a misspelt key cannot be
  // silently ignored while the real setting keeps its default.
  Object.keys(config).forEach(function (key) {
    if (KNOWN_SECTIONS.indexOf(key) < 0) {
      add(key, "Unknown top-level section. Fix the spelling or remove it.");
    }
  });

  if (config.version !== 1) {
    add("version", "Unsupported config version. Expected 1.");
  }

  scanForSecrets(config, "", add);

  const environment = config.environment;
  if (!isPlainObject(environment)) {
    add("environment", "Missing environment section.");
  } else {
    requireString(environment, "name", "environment.name", add);
    if (!isHttpsUrl(environment.tenantUrl)) {
      add("environment.tenantUrl", "Must be an https:// URL.");
    }
    if (!isHttpsUrl(environment.siteUrl)) {
      add("environment.siteUrl", "Must be an https:// URL.");
    }
    if (
      isHttpsUrl(environment.tenantUrl) &&
      isHttpsUrl(environment.siteUrl) &&
      urlHost(environment.tenantUrl as string) !== urlHost(environment.siteUrl as string)
    ) {
      add(
        "environment.siteUrl",
        "Site host must match the tenant host (" + urlHost(environment.tenantUrl as string) + ")."
      );
    }
    ["tenantId", "siteCollectionId", "siteId", "webId"].forEach(function (key) {
      const value = environment[key];
      if (typeof value === "string" && value.length > 0 && !isGuid(value)) {
        add("environment." + key, "Must be a GUID or an empty string.");
      }
    });
  }

  const resources = config.resources;
  if (!isPlainObject(resources)) {
    add("resources", "Missing resources section.");
  } else {
    validateLists(resources.lists, add);
    validateItems(resources.items, add);
    ["folders", "files", "fields", "contentTypes", "views", "attachments", "pages"].forEach(
      function (section) {
        if (!isPlainObject(resources[section])) {
          add("resources." + section, "Missing section.");
        }
      }
    );
    const views = resources.views;
    if (isPlainObject(views) && typeof views.viewId === "string" && views.viewId.length > 0) {
      if (!isGuid(views.viewId)) {
        add("resources.views.viewId", "Must be a GUID or an empty string.");
      }
    }
    const contentTypes = resources.contentTypes;
    if (isPlainObject(contentTypes)) {
      Object.keys(contentTypes).forEach(function (key) {
        const value = contentTypes[key];
        if (typeof value !== "string" || value.length === 0) {
          add("resources.contentTypes." + key, "Must be a non-empty string.");
        } else if (key !== "nonexistentContentTypeId" && !/^0x[0-9A-F]+$/i.test(value)) {
          add("resources.contentTypes." + key, "Content type id must start with 0x.");
        }
      });
    }
  }

  validatePrincipals(config.principals, add);
  validateSearch(config.search, add);
  validateSafety(config.safety, environment, add);

  if (!isPlainObject(config.testData)) {
    add("testData", "Missing testData section.");
  }

  const errors = issues.filter(function (issue) {
    return issue.severity === "error";
  });
  return { ok: errors.length === 0, issues };
}

function validateLists(lists: unknown, add: AddIssue): void {
  if (!isPlainObject(lists)) {
    add("resources.lists", "Missing lists section.");
    return;
  }
  const keys = Object.keys(lists);
  if (keys.length === 0) {
    add("resources.lists", "At least one list must be configured.");
    return;
  }
  keys.forEach(function (key) {
    const path = "resources.lists." + key;
    const list = lists[key];
    if (!isPlainObject(list)) {
      add(path, "Must be an object.");
      return;
    }
    if (typeof list.title !== "string" || list.title.length === 0) {
      add(path + ".title", "Required, non-empty list title.");
    }
    if (typeof list.id !== "string") {
      add(path + ".id", "Must be a GUID or an empty string.");
    } else if (list.id.length > 0 && !isGuid(list.id)) {
      add(path + ".id", "Must be a GUID or an empty string.");
    }
    if (!isNonNegativeInt(list.template)) {
      add(path + ".template", "Must be a non-negative integer (100 list, 101 document library, 119 site pages).");
    }
    if (typeof list.serverRelativeUrl !== "string") {
      add(path + ".serverRelativeUrl", "Must be a string.");
    }
  });
  const primary = lists.primary;
  if (!isPlainObject(primary)) {
    add("resources.lists.primary", "A primary list entry is required.");
  } else if (primary.id === "" && primary.title === "") {
    add("resources.lists.primary", "Configure at least the title or the GUID of the primary list.");
  } else if (primary.id === "") {
    add(
      "resources.lists.primary.id",
      "Primary list has no GUID. Endpoints that require a list GUID cannot be exercised.",
      "warning"
    );
  }
}

function validateItems(items: unknown, add: AddIssue): void {
  if (!isPlainObject(items)) {
    add("resources.items", "Missing items section.");
    return;
  }
  ["primaryItemId", "secondaryItemId", "nonExistentItemId"].forEach(function (key) {
    if (!isPositiveInt(items[key])) {
      add("resources.items." + key, "Must be a positive integer.");
    }
  });
  if (!Array.isArray(items.boundaryItemIds)) {
    add("resources.items.boundaryItemIds", "Must be an array of integers.");
  }
  if (!Array.isArray(items.stringIds)) {
    add("resources.items.stringIds", "Must be an array of strings.");
  }
}

function validatePrincipals(principals: unknown, add: AddIssue): void {
  if (!isPlainObject(principals)) {
    add("principals", "Missing principals section.");
    return;
  }
  const users = principals.users;
  if (!isPlainObject(users)) {
    add("principals.users", "Missing users section.");
  } else {
    ["primary", "secondary"].forEach(function (key) {
      const path = "principals.users." + key;
      const user = users[key];
      if (!isPlainObject(user)) {
        add(path, "Missing user entry.");
        return;
      }
      if (!isUpn(user.upn)) {
        add(path + ".upn", "Must be a user principal name.");
      }
      if (typeof user.loginName !== "string" || user.loginName.indexOf("|") < 0) {
        add(path + ".loginName", "Must be a claims login name (i:0#.f|membership|...).");
      }
      if (!isNonNegativeInt(user.id)) {
        add(path + ".id", "Must be a non-negative integer (0 when not yet resolved).");
      }
    });
  }
  const groups = principals.groups;
  if (!isPlainObject(groups)) {
    add("principals.groups", "Missing groups section.");
  } else {
    ["primaryGroupId", "ownerGroupId", "memberGroupId", "visitorGroupId"].forEach(function (key) {
      if (!isPositiveInt(groups[key])) {
        add("principals.groups." + key, "Must be a positive integer.");
      }
    });
    if (!Array.isArray(groups.groupsToEnumerate)) {
      add("principals.groups.groupsToEnumerate", "Must be an array of group ids.");
    }
    if (!Array.isArray(groups.shareWith) || (groups.shareWith as unknown[]).length === 0) {
      add("principals.groups.shareWith", "Must be a non-empty array of identities to share with.");
    }
  }
  const definitions = principals.roleDefinitions;
  if (!isPlainObject(definitions)) {
    add("principals.roleDefinitions", "Missing roleDefinitions section.");
  } else {
    Object.keys(definitions).forEach(function (key) {
      const value = definitions[key];
      const ok = Array.isArray(value)
        ? value.every(isPositiveInt)
        : isPositiveInt(value);
      if (!ok) {
        add("principals.roleDefinitions." + key, "Must be a positive integer or an array of them.");
      }
    });
  }
  ["principalIds", "shareLinkKinds", "shareLinkRoleKinds", "scope", "permissionKinds"].forEach(
    function (key) {
      if (!Array.isArray(principals[key])) {
        add("principals." + key, "Must be an array.");
      }
    }
  );
}

function validateSearch(search: unknown, add: AddIssue): void {
  if (!isPlainObject(search)) {
    add("search", "Missing search section.");
    return;
  }
  if (typeof search.queryText !== "string" || search.queryText.length === 0) {
    add("search.queryText", "Required.");
  }
  if (!isPositiveInt(search.rowLimit)) {
    add("search.rowLimit", "Must be a positive integer.");
  }
  if (!Array.isArray(search.sortList)) {
    add("search.sortList", "Must be an array.");
  }
}

function validateSafety(safety: unknown, environment: unknown, add: AddIssue): void {
  if (!isPlainObject(safety)) {
    add("safety", "Missing safety section. Request execution is disabled without it.");
    return;
  }
  [
    "allowDestructive",
    "allowPermissionChanges",
    "allowOutboundWebhooks",
    "allowLargePayloads",
  ].forEach(function (key) {
    if (typeof safety[key] !== "boolean") {
      add("safety." + key, "Must be a boolean.");
    }
  });
  const hosts = safety.allowedHosts;
  if (!Array.isArray(hosts) || hosts.length === 0) {
    add("safety.allowedHosts", "Must contain at least one host.");
  } else {
    hosts.forEach(function (host, index) {
      if (typeof host !== "string" || !HOST_PATTERN.test(host)) {
        add("safety.allowedHosts[" + index + "]", "Must be a bare hostname (no scheme or path).");
      }
    });
    if (isPlainObject(environment) && isHttpsUrl(environment.siteUrl)) {
      const siteHost = urlHost(environment.siteUrl as string).split(":")[0];
      if ((hosts as string[]).indexOf(siteHost) < 0) {
        add(
          "safety.allowedHosts",
          "The configured site host '" + siteHost + "' is not in the allow-list."
        );
      }
    }
  }
  [
    "maxRequestsPerRun",
    "maxPayloadBytes",
    "requestTimeoutMs",
    "minimumDelayBetweenRequestsMs",
  ].forEach(function (key) {
    if (!isPositiveInt(safety[key])) {
      add("safety." + key, "Must be a positive integer.");
    }
  });
  if (typeof safety.allowDestructive === "boolean" && safety.allowDestructive) {
    add(
      "safety.allowDestructive",
      "Destructive testing is enabled. Only use a throwaway site collection.",
      "warning"
    );
  }
}

type AddIssue = (path: string, message: string, severity?: Severity) => void;

/**
 * Walks the whole object and rejects secret-looking keys. This is the guard that
 * keeps credentials out of a file that is committed to source control.
 */
function scanForSecrets(value: unknown, path: string, add: AddIssue): void {
  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) {
      scanForSecrets(value[i], path + "[" + i + "]", add);
    }
    return;
  }
  if (!isPlainObject(value)) {
    return;
  }
  Object.keys(value).forEach(function (key) {
    const childPath = path ? path + "." + key : key;
    if (SECRET_KEY_PATTERN.test(key)) {
      add(
        childPath,
        "Secret-looking key in configuration. Credentials must come from the runtime/environment, never this file."
      );
    }
    scanForSecrets(value[key], childPath, add);
  });
}

function requireString(
  source: Record<string, unknown>,
  key: string,
  path: string,
  add: AddIssue
): void {
  if (typeof source[key] !== "string" || (source[key] as string).length === 0) {
    add(path, "Required, non-empty string.");
  }
}

/** Recursive merge used for overrides. Arrays replace, objects merge. */
export function mergeOverrides<T>(base: T, overrides: unknown): T {
  if (!isPlainObject(base) || !isPlainObject(overrides)) {
    return (overrides === undefined ? base : (overrides as T));
  }
  const result: Record<string, unknown> = {};
  Object.keys(base).forEach(function (key) {
    result[key] = (base as Record<string, unknown>)[key];
  });
  Object.keys(overrides).forEach(function (key) {
    const baseValue = (base as Record<string, unknown>)[key];
    const overrideValue = overrides[key];
    result[key] =
      isPlainObject(baseValue) && isPlainObject(overrideValue)
        ? mergeOverrides(baseValue, overrideValue)
        : overrideValue;
  });
  return result as T;
}

export function formatIssues(issues: ConfigIssue[]): string {
  if (issues.length === 0) {
    return "Configuration is valid.";
  }
  const lines = issues.map(function (issue) {
    return (
      "  [" + issue.severity + "] " +
      (issue.path || "<root>") + ": " + issue.message
    );
  });
  const errors = issues.filter(function (issue) {
    return issue.severity === "error";
  }).length;
  return (
    "Invalid test configuration (" + errors + " error(s), " +
    (issues.length - errors) + " warning(s)):\n" + lines.join("\n")
  );
}

/**
 * Loads, merges and validates the configuration.
 * Throws ConfigError with every error instead of failing on the first one.
 */
export function loadTestConfig(overrides?: ConfigOverrides, options?: { allowWarnings?: boolean }): TestConfig {
  const merged = overrides ? mergeOverrides(rawConfig, overrides) : rawConfig;
  const result = validateTestConfig(merged);
  const errors = result.issues.filter(function (issue) {
    return issue.severity === "error";
  });
  if (errors.length > 0 || (result.issues.length > 0 && options && options.allowWarnings === false)) {
    throw new ConfigError(result.issues);
  }
  return normalise(merged as unknown as TestConfig);
}

/** Trimmed, normalised copy so callers never compare raw whitespace. */
function normalise(config: TestConfig): TestConfig {
  const copy = JSON.parse(JSON.stringify(config)) as TestConfig;
  copy.environment.tenantUrl = copy.environment.tenantUrl.replace(/\/+$/, "");
  copy.environment.siteUrl = copy.environment.siteUrl.replace(/\/+$/, "");
  Object.keys(copy.resources.lists).forEach(function (key) {
    const list = copy.resources.lists[key];
    if (list && list.id) {
      list.id = normaliseGuid(list.id);
    }
  });
  if (copy.resources.views.viewId) {
    copy.resources.views.viewId = normaliseGuid(copy.resources.views.viewId);
  }
  copy.safety.allowedHosts = copy.safety.allowedHosts.map(function (host) {
    return host.toLowerCase();
  });
  return copy;
}

/** Convenience accessor used by tools: all GUID-like values that must exist. */
export function collectIdentifiers(config: TestConfig): Array<{ path: string; value: string }> {
  const out: Array<{ path: string; value: string }> = [];
  Object.keys(config.resources.lists).forEach(function (key) {
    const id = config.resources.lists[key].id;
    if (id) {
      out.push({ path: "resources.lists." + key + ".id", value: id });
    }
  });
  if (config.resources.views.viewId) {
    out.push({ path: "resources.views.viewId", value: config.resources.views.viewId });
  }
  return out;
}

export default rawConfig;
