import rawConfig from "../../src/config/test-config.json";
import {
  assertAllowedUrl,
  assertDestructiveAllowed,
  assertPermissionChangeAllowed,
  assertWebhookAllowed,
  collectIdentifiers,
  ConfigError,
  formatIssues,
  getPrimaryList,
  isGuid,
  isHttpsUrl,
  isUpn,
  itemIdOrDefault,
  loadTestConfig,
  mergeOverrides,
  normaliseGuid,
  urlHost,
  validateTestConfig,
} from "../../src/webparts/hacktheworld/config/loader";
import { assertEquals, assertThrows, assertTrue, group, test } from "./harness";

type Json = Record<string, any>;

function clone(): Json {
  return JSON.parse(JSON.stringify(rawConfig)) as Json;
}

function withMutation(mutate: (config: Json) => void): Json {
  const config = clone();
  mutate(config);
  return config;
}

function errorPaths(config: Json): string[] {
  return validateTestConfig(config).issues
    .filter(function (issue) {
      return issue.severity === "error";
    })
    .map(function (issue) {
      return issue.path;
    });
}

function hasIssue(config: Json, path: string, fragment?: string): boolean {
  return validateTestConfig(config).issues.some(function (issue) {
    return issue.path === path && (!fragment || issue.message.indexOf(fragment) >= 0);
  });
}

const BASE = loadTestConfig();

export function registerConfigTests(): void {
  group("config: shipped configuration");

  test("the committed configuration validates", function () {
    const result = validateTestConfig(rawConfig);
    const errors = result.issues.filter(function (issue) {
      return issue.severity === "error";
    });
    assertEquals(errors, [], "committed config must have no errors");
  });

  test("loadTestConfig returns a normalised copy", function () {
    assertEquals(BASE.environment.siteUrl, "https://tmsobb2.sharepoint.com");
    assertEquals(BASE.resources.lists.primary.id, "4a81ba4f-84ed-4e91-a731-9a82c0ede37e");
    assertEquals(BASE.safety.allowedHosts[0], "tmsobb2.sharepoint.com");
  });

  test("the derived CONFIG consumed by Detaservice still has its role list", function () {
    assertTrue(Array.isArray(BASE.principals.roleDefinitions.all), "roleDefinitions.all must stay an array");
    assertTrue(BASE.principals.roleDefinitions.all.length > 0, "role list must not be empty");
  });

  test("normalisation lower-cases list GUIDs", function () {
    const config = loadTestConfig({
      resources: { lists: { primary: { id: "4A81BA4F-84ED-4E91-A731-9A82C0EDE37E" } } },
    });
    assertEquals(config.resources.lists.primary.id, "4a81ba4f-84ed-4e91-a731-9a82c0ede37e");
  });

  test("all expected host identifiers are collected for tooling", function () {
    const identifiers = collectIdentifiers(BASE);
    const primary = identifiers.filter(function (entry) {
      return entry.path === "resources.lists.primary.id";
    })[0];
    assertTrue(!!primary, "primary list id must be collected");
    assertEquals(primary.value, BASE.resources.lists.primary.id);
  });

  group("config: structural validation");

  test("a non-object is rejected", function () {
    assertTrue(hasIssue("not-an-object" as never, "", "JSON object"), "scalars must be rejected");
  });

  test("an unsupported version is rejected", function () {
    assertTrue(hasIssue(withMutation(function (c) { c.version = 2; }), "version", "Expected 1"));
  });

  test("unknown top-level sections are rejected as typos", function () {
    const config = withMutation(function (c) {
      c.typoSection = true;
    });
    const result = validateTestConfig(config);
    assertTrue(!result.ok, "an unknown section must fail validation");
    assertTrue(
      result.issues.some(function (issue) {
        return issue.path === "typoSection" && issue.severity === "error";
      }),
      "expected an error for the unknown section"
    );
  });

  test("secret-looking keys are rejected anywhere in the tree", function () {
    const config = withMutation(function (c) {
      c.testData.credentials = { access_token: "abc" };
    });
    assertTrue(hasIssue(config, "testData.credentials.access_token", "Secret-looking key"));
  });

  test("missing environment, resources, principals, search, testData are reported", function () {
    const config = withMutation(function (c) {
      delete c.environment;
      delete c.resources;
      delete c.principals;
      delete c.search;
      delete c.testData;
    });
    const paths = errorPaths(config);
    ["environment", "resources", "principals", "search", "testData"].forEach(function (section) {
      assertTrue(paths.indexOf(section) >= 0, "expected an error for " + section);
    });
  });

  group("config: environment");

  test("a tenant/site host mismatch is rejected", function () {
    const config = withMutation(function (c) {
      c.environment.siteUrl = "https://other.sharepoint.com/sites/x";
    });
    assertTrue(hasIssue(config, "environment.siteUrl", "must match"));
  });

  test("plain http is rejected", function () {
    const config = withMutation(function (c) {
      c.environment.tenantUrl = "http://tmsobb2.sharepoint.com";
    });
    assertTrue(hasIssue(config, "environment.tenantUrl", "https://"));
  });

  test("malformed GUIDs are rejected but empty strings are allowed", function () {
    const malformed = withMutation(function (c) {
      c.environment.webId = "not-a-guid";
    });
    assertTrue(hasIssue(malformed, "environment.webId", "GUID"));

    const empty = withMutation(function (c) {
      c.environment.webId = "";
    });
    assertEquals(errorPaths(empty), []);
  });

  group("config: resources");

  test("a list with a non-GUID id is rejected", function () {
    const config = withMutation(function (c) {
      c.resources.lists.primary.id = "not-a-guid";
    });
    assertTrue(hasIssue(config, "resources.lists.primary.id", "GUID"));
  });

  test("a list template must be a non-negative integer", function () {
    const config = withMutation(function (c) {
      c.resources.lists.primary.template = -1;
    });
    assertTrue(hasIssue(config, "resources.lists.primary.template", "non-negative"));
  });

  test("the primary list must exist", function () {
    const config = withMutation(function (c) {
      delete c.resources.lists.primary;
    });
    assertTrue(hasIssue(config, "resources.lists.primary", "required"));
  });

  test("an empty primary list is a warning when the title is configured", function () {
    const config = withMutation(function (c) {
      c.resources.lists.primary.id = "";
    });
    const result = validateTestConfig(config);
    assertTrue(
      result.issues.some(function (issue) {
        return issue.path === "resources.lists.primary.id" && issue.severity === "warning";
      }),
      "expected a warning about endpoints that need a GUID"
    );
  });

  test("item ids must be positive integers", function () {
    const config = withMutation(function (c) {
      c.resources.items.primaryItemId = 0;
      c.resources.items.secondaryItemId = "3";
    });
    const paths = errorPaths(config);
    assertTrue(paths.indexOf("resources.items.primaryItemId") >= 0, "zero must be rejected");
    assertTrue(paths.indexOf("resources.items.secondaryItemId") >= 0, "strings must be rejected");
  });

  test("content type ids must start with 0x", function () {
    const config = withMutation(function (c) {
      c.resources.contentTypes.itemContentTypeId = "0100";
    });
    assertTrue(hasIssue(config, "resources.contentTypes.itemContentTypeId", "0x"));
  });

  test("a view id must be a GUID", function () {
    const config = withMutation(function (c) {
      c.resources.views.viewId = "view-42";
    });
    assertTrue(hasIssue(config, "resources.views.viewId", "GUID"));
  });

  group("config: principals");

  test("a user upn and claims login name are required", function () {
    const config = withMutation(function (c) {
      c.principals.users.primary.upn = "not-a-upn";
      c.principals.users.primary.loginName = "plain-name";
    });
    const paths = errorPaths(config);
    assertTrue(paths.indexOf("principals.users.primary.upn") >= 0, "upn must be validated");
    assertTrue(paths.indexOf("principals.users.primary.loginName") >= 0, "login name must be validated");
  });

  test("group ids must be positive", function () {
    const config = withMutation(function (c) {
      c.principals.groups.primaryGroupId = 0;
    });
    assertTrue(hasIssue(config, "principals.groups.primaryGroupId", "positive"));
  });

  test("role definitions accept both scalars and arrays", function () {
    const config = withMutation(function (c) {
      c.principals.roleDefinitions.bad = 0;
    });
    assertTrue(hasIssue(config, "principals.roleDefinitions.bad", "positive"));
  });

  test("shareWith must not be empty", function () {
    const config = withMutation(function (c) {
      c.principals.groups.shareWith = [];
    });
    assertTrue(hasIssue(config, "principals.groups.shareWith", "non-empty"));
  });

  group("config: safety");

  test("a missing safety section is fatal", function () {
    const config = withMutation(function (c) {
      delete c.safety;
    });
    assertTrue(hasIssue(config, "safety", "disabled"));
  });

  test("allowedHosts must contain the site host", function () {
    const config = withMutation(function (c) {
      c.safety.allowedHosts = ["contoso.sharepoint.com"];
    });
    assertTrue(hasIssue(config, "safety.allowedHosts", "not in the allow-list"));
  });

  test("allowedHosts entries cannot include a scheme or path", function () {
    const config = withMutation(function (c) {
      c.safety.allowedHosts = ["https://tmsobb2.sharepoint.com"];
    });
    assertTrue(hasIssue(config, "safety.allowedHosts[0]", "bare hostname"));
  });

  test("safety limits must be positive integers", function () {
    const config = withMutation(function (c) {
      c.safety.maxRequestsPerRun = 0;
      c.safety.requestTimeoutMs = "1000";
    });
    const paths = errorPaths(config);
    assertTrue(paths.indexOf("safety.maxRequestsPerRun") >= 0, "zero limit must be rejected");
    assertTrue(paths.indexOf("safety.requestTimeoutMs") >= 0, "string timeout must be rejected");
  });

  test("enabling destructive testing is surfaced as a warning", function () {
    const config = withMutation(function (c) {
      c.safety.allowDestructive = true;
    });
    const result = validateTestConfig(config);
    assertTrue(result.ok, "enabling the switch must not be an error");
    assertTrue(
      result.issues.some(function (issue) {
        return issue.path === "safety.allowDestructive" && issue.severity === "warning";
      }),
      "expected a warning"
    );
  });

  group("config: runtime guards");

  test("assertAllowedUrl accepts an allow-listed https target", function () {
    assertAllowedUrl("https://tmsobb2.sharepoint.com/_api/web", BASE);
  });

  test("assertAllowedUrl rejects another host", function () {
    assertThrows(function () {
      assertAllowedUrl("https://evil.example/_api/web", BASE);
    }, "Refusing target host");
  });

  test("assertAllowedUrl rejects plain http", function () {
    assertThrows(function () {
      assertAllowedUrl("http://tmsobb2.sharepoint.com/_api/web", BASE);
    }, "non-HTTPS");
  });

  test("assertAllowedUrl rejects everything when the allow-list is empty", function () {
    const noHosts = mergeOverrides(BASE, { safety: { allowedHosts: [] } });
    assertThrows(function () {
      assertAllowedUrl("https://tmsobb2.sharepoint.com/_api/web", noHosts as never);
    }, "No hosts are allowed");
  });

  test("destructive, permission and webhook operations are off by default", function () {
    assertThrows(function () {
      assertDestructiveAllowed(BASE, "delete list item");
    }, "allowDestructive is false");
    assertThrows(function () {
      assertPermissionChangeAllowed(BASE, "break inheritance");
    }, "allowDestructive is false");
    assertThrows(function () {
      assertWebhookAllowed(BASE);
    }, "allowDestructive is false");
  });

  test("each guard can be enabled deliberately", function () {
    const permissive = mergeOverrides(BASE, {
      safety: {
        allowDestructive: true,
        allowPermissionChanges: true,
        allowOutboundWebhooks: true,
      },
    });
    assertDestructiveAllowed(permissive as never, "delete list item");
    assertPermissionChangeAllowed(permissive as never, "break inheritance");
    assertWebhookAllowed(permissive as never);

    const onlyDestructive = mergeOverrides(BASE, { safety: { allowDestructive: true } });
    assertThrows(function () {
      assertPermissionChangeAllowed(onlyDestructive as never, "break inheritance");
    }, "allowPermissionChanges is false");
    assertThrows(function () {
      assertWebhookAllowed(onlyDestructive as never);
    }, "allowOutboundWebhooks");
  });

  test("getPrimaryList returns the configured list or throws", function () {
    assertEquals(getPrimaryList(BASE).title, "emp");
    const noLists = mergeOverrides(BASE, { resources: { lists: { primary: null } } });
    assertThrows(function () {
      getPrimaryList(noLists as never);
    }, "not configured");
  });

  test("itemIdOrDefault prefers the explicit id", function () {
    assertEquals(itemIdOrDefault(BASE), BASE.resources.items.primaryItemId);
    assertEquals(itemIdOrDefault(BASE, 99), 99);
  });

  group("config: overrides and errors");

  test("overrides deep-merge and validate", function () {
    const config = loadTestConfig({
      environment: { siteUrl: "https://tmsobb2.sharepoint.com/sites/other" },
      resources: { items: { primaryItemId: 7 } },
      safety: { maxRequestsPerRun: 10 },
    });
    assertEquals(config.environment.siteUrl, "https://tmsobb2.sharepoint.com/sites/other");
    assertEquals(config.resources.items.primaryItemId, 7);
    assertEquals(config.resources.items.secondaryItemId, BASE.resources.items.secondaryItemId);
    assertEquals(config.safety.maxRequestsPerRun, 10);
  });

  test("arrays replace rather than concatenate", function () {
    const merged = mergeOverrides(
      { list: [1, 2, 3], nested: { a: 1, b: 2 } },
      { list: [9], nested: { b: 3 } }
    ) as { list: number[]; nested: Record<string, number> };
    assertEquals(merged.list, [9]);
    assertEquals(merged.nested, { a: 1, b: 3 });
  });

  test("invalid overrides throw ConfigError with every issue", function () {
    let caught: ConfigError | null = null;
    try {
      loadTestConfig({ resources: { items: { primaryItemId: -1 } } });
    } catch (error) {
      caught = error as ConfigError;
    }
    if (!caught) {
      throw new Error("expected ConfigError");
    }
    assertTrue(caught instanceof ConfigError, "error must be a ConfigError");
    assertTrue(caught.issues.length > 0, "ConfigError must report the issues");
    assertTrue(
      caught.issues.some(function (issue) {
        return issue.path === "resources.items.primaryItemId";
      }),
      "ConfigError must name the offending path"
    );
  });

  test("formatIssues summarises errors and warnings", function () {
    const result = validateTestConfig(
      withMutation(function (c) {
        c.version = 2;
        c.typoSection = true;
      })
    );
    const text = formatIssues(result.issues);
    assertTrue(text.indexOf("error") >= 0, "expected an error count");
    assertTrue(text.indexOf("version") >= 0, "expected the failing path");
  });

  test("formatIssues reports a valid configuration", function () {
    assertEquals(formatIssues([]), "Configuration is valid.");
  });

  group("config: primitives");

  test("isGuid accepts dashed GUIDs in either case", function () {
    assertTrue(isGuid("4A81BA4F-84ED-4E91-A731-9A82C0EDE37E"), "upper case GUID should pass");
    assertTrue(!isGuid("4a81ba4f84ed4e91a7319a82c0ede37e"), "compact GUID is not accepted here");
  });

  test("normaliseGuid strips braces and lower-cases", function () {
    assertEquals(normaliseGuid("{4A81BA4F-84ED-4E91-A731-9A82C0EDE37E}"), "4a81ba4f-84ed-4e91-a731-9a82c0ede37e");
  });

  test("isHttpsUrl only accepts https URLs", function () {
    assertTrue(isHttpsUrl("https://contoso.sharepoint.com"), "https should pass");
    assertTrue(!isHttpsUrl("http://contoso.sharepoint.com"), "http must fail");
    assertTrue(!isHttpsUrl("/sites/team"), "relative URL must fail");
  });

  test("urlHost extracts a lower-case host", function () {
    assertEquals(urlHost("https://Contoso.SharePoint.com/x"), "contoso.sharepoint.com");
    assertEquals(urlHost("/relative"), "");
  });

  test("isUpn requires an at-sign and a dotted domain", function () {
    assertTrue(isUpn("user@contoso.onmicrosoft.com"), "valid UPN should pass");
    assertTrue(!isUpn("user@contoso"), "domain without a dot must fail");
    assertTrue(!isUpn("user"), "missing at-sign must fail");
  });
}
