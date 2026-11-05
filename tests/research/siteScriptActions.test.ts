import { DigestProvider } from "../../src/webparts/hacktheworld/research/digest";
import { RequestExecutor } from "../../src/webparts/hacktheworld/research/executor";
import { buildRequest } from "../../src/webparts/hacktheworld/research/request";
import {
  createSiteScriptActionDefinition,
  createSiteScriptContentTypeSubaction,
  createSiteScriptListSubaction,
  parseSiteScriptActionDefinition,
  serializeSiteScriptActionDefinition,
  SITE_SCRIPT_LIST_SUBACTION_VERBS,
  SITE_SCRIPT_VERBS,
  validateSiteScriptActionDefinition,
} from "../../src/webparts/hacktheworld/research/siteScriptActions";
import { TransportResponse } from "../../src/webparts/hacktheworld/research/types";
import { assertContains, assertEquals, assertLength, assertThrows, group, test } from "./harness";

const SITE = "https://contoso.sharepoint.com/sites/team";

export function registerSiteScriptActionTests(): void {
  group("site script action schema and execution");

  test("provides valid example definitions for every supported top-level verb", function () {
    assertLength(SITE_SCRIPT_VERBS.slice(), 21, "top-level verbs");
    SITE_SCRIPT_VERBS.forEach((verb) => validateSiteScriptActionDefinition(createSiteScriptActionDefinition(verb)));
  });

  test("supports every createSPList subaction with a valid independent example", function () {
    assertLength(SITE_SCRIPT_LIST_SUBACTION_VERBS.slice(), 18, "list subactions");
    SITE_SCRIPT_LIST_SUBACTION_VERBS.forEach((verb) => {
      const definition = createSiteScriptActionDefinition("createSPList") as { verb: string; subactions?: unknown[] };
      definition.subactions = [createSiteScriptListSubaction(verb)];
      validateSiteScriptActionDefinition(definition);
    });
  });

  test("supports both createContentType subactions", function () {
    ["addSiteColumn", "removeSiteColumn"].forEach((verb) => {
      const definition = createSiteScriptActionDefinition("createContentType") as { verb: string; subactions?: unknown[] };
      definition.subactions = [createSiteScriptContentTypeSubaction(verb as "addSiteColumn" | "removeSiteColumn")];
      validateSiteScriptActionDefinition(definition);
    });
  });

  test("rejects unsupported verbs, missing required properties, and invalid types", function () {
    assertThrows(() => validateSiteScriptActionDefinition({ verb: "inventedVerb" }), "Unsupported site script verb");
    assertThrows(() => validateSiteScriptActionDefinition({ verb: "installSolution", id: "id" }), "actionDefinition.name is required");
    assertThrows(() => validateSiteScriptActionDefinition({ verb: "createSPList", listName: "List", templateType: "100" }), "templateType must be number");
  });

  test("checks enum values, nested subactions, and schema-declared extra properties", function () {
    assertThrows(() => validateSiteScriptActionDefinition({ verb: "addPrincipalToSPGroup", principal: "user@example.invalid", group: "Administrators" }), "required pattern");
    const list = createSiteScriptActionDefinition("createSPList") as { verb: string; subactions?: unknown[]; [key: string]: unknown };
    list.subactions = [{ verb: "addFolder", path: 5 }];
    assertThrows(() => validateSiteScriptActionDefinition(list), "subactions\\[0\\].path must be string");
    list.subactions = [{ verb: "notASubaction", path: "Research" }];
    assertThrows(() => validateSiteScriptActionDefinition(list), "unsupported or missing verb");
    list.subactions = [{ name: "Research View" }];
    validateSiteScriptActionDefinition(list);
    assertThrows(() => validateSiteScriptActionDefinition({ verb: "createSPList", listName: "List", templateType: 100, unknownProperty: true }), "unknownProperty is not supported");
  });

  test("parses and serializes nested action definitions as JSON", function () {
    const action = createSiteScriptActionDefinition("createSPList") as { verb: string; subactions?: unknown[] };
    action.subactions = [createSiteScriptListSubaction("addSPView")];
    const serialized = serializeSiteScriptActionDefinition(action);
    assertEquals(parseSiteScriptActionDefinition(serialized), action);
    assertThrows(() => parseSiteScriptActionDefinition("{"), "valid JSON");
  });

  test("passes stringified actionDefinition through the existing endpoint and handles SharePoint success", async function () {
    const action = createSiteScriptActionDefinition("createSPList");
    const actionJson = serializeSiteScriptActionDefinition(action);
    const spec = buildRequest("siteScripts.executeAction", {
      siteUrl: SITE,
      body: { actionDefinition: actionJson },
    });
    const calls: string[] = [];
    const transport = async (request: typeof spec): Promise<TransportResponse> => {
      calls.push(request.body || "");
      return { status: 200, bodyText: "{\"d\":{\"ExecuteSiteScriptAction\":\"ok\"}}", headers: {} };
    };
    const digest = new DigestProvider({ fetchDigest: async () => ({ digest: "0x-test-digest", expiresInSeconds: 1800 }) });
    const executor = new RequestExecutor(transport, digest, { maxTransientRetries: 0, maxDigestRetries: 0 });
    const record = await executor.execute(spec);
    assertEquals(spec.url, SITE + "/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.ExecuteSiteScriptAction()");
    assertEquals(spec.method, "POST");
    assertEquals(JSON.parse(calls[0]).actionDefinition, actionJson);
    assertEquals(record.status, 200);
    assertContains(record.responseBody || "", "ExecuteSiteScriptAction", "successful SharePoint response body");
  });

  test("preserves SharePoint API validation errors in the execution record", async function () {
    const actionJson = serializeSiteScriptActionDefinition(createSiteScriptActionDefinition("setRegionalSettings"));
    const spec = buildRequest("siteScripts.executeAction", { siteUrl: SITE, body: { actionDefinition: actionJson } });
    const digest = new DigestProvider({ fetchDigest: async () => ({ digest: "0x-test-digest", expiresInSeconds: 1800 }) });
    const executor = new RequestExecutor(async () => ({ status: 400, bodyText: "Invalid site script action", headers: {} }), digest, { maxTransientRetries: 0, maxDigestRetries: 0 });
    const record = await executor.execute(spec);
    assertEquals(record.status, 400);
    assertEquals(record.responseClass, "validation");
    assertContains(record.responseBody || "", "Invalid site script action", "SharePoint validation response body");
  });
}
