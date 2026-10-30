import {
  ENDPOINTS,
  METHODS_OF_INTEREST,
  getEndpoint,
  hasEndpoint,
  listEndpointIdsByFamily,
  listEndpoints,
  listModules,
} from "../../src/webparts/hacktheworld/research/registry";
import { assertEquals, assertThrows, assertTrue, group, test } from "./harness";

const PLACEHOLDER = /\{([A-Za-z0-9_]+)(?::([A-Za-z]+))?\}/g;

export function registerRegistryTests(): void {
  group("registry: integrity");

  test("endpoint ids are unique", function () {
    const seen: Record<string, boolean> = {};
    for (let i = 0; i < ENDPOINTS.length; i++) {
      const id = ENDPOINTS[i].id;
      assertTrue(!seen[id], "duplicate endpoint id: " + id);
      seen[id] = true;
    }
  });

  test("every endpoint declares a module, operation, method and evidence source", function () {
    for (let i = 0; i < ENDPOINTS.length; i++) {
      const endpoint = ENDPOINTS[i];
      assertTrue(endpoint.module.length > 0, endpoint.id + " has no module");
      assertTrue(endpoint.operation.length > 0, endpoint.id + " has no operation");
      assertTrue(endpoint.methods.length > 0, endpoint.id + " has no methods");
      assertTrue(endpoint.sources.length > 0, endpoint.id + " has no evidence source");
      assertTrue(endpoint.payloadFamilies.length > 0, endpoint.id + " has no payload families");
      assertTrue(endpoint.description.length > 0, endpoint.id + " has no description");
    }
  });

  test("SharePoint templates are relative and Graph templates are absolute", function () {
    for (let i = 0; i < ENDPOINTS.length; i++) {
      const endpoint = ENDPOINTS[i];
      if (endpoint.family === "microsoft-graph") {
        assertTrue(
          /^https:\/\/graph\.microsoft\.com\//.test(endpoint.urlTemplate),
          endpoint.id + " must be an absolute graph.microsoft.com URL"
        );
      } else {
        assertTrue(
          endpoint.urlTemplate.charAt(0) === "/",
          endpoint.id + " must be a site-relative template"
        );
      }
    }
  });

  test("no template hardcodes a tenant host", function () {
    for (let i = 0; i < ENDPOINTS.length; i++) {
      const template = ENDPOINTS[i].urlTemplate;
      assertTrue(
        template.indexOf("sharepoint.com") < 0 || ENDPOINTS[i].family === "microsoft-graph",
        ENDPOINTS[i].id + " hardcodes a tenant host: " + template
      );
    }
  });

  test("every URL placeholder is declared as a parameter", function () {
    for (let i = 0; i < ENDPOINTS.length; i++) {
      const endpoint = ENDPOINTS[i];
      const names: string[] = [];
      let match: RegExpExecArray | null;
      PLACEHOLDER.lastIndex = 0;
      while ((match = PLACEHOLDER.exec(endpoint.urlTemplate)) !== null) {
        names.push(match[1]);
      }
      for (let j = 0; j < names.length; j++) {
        const declared = endpoint.parameters.some(function (parameter) {
          return parameter.name === names[j];
        });
        assertTrue(
          declared,
          endpoint.id + " uses {" + names[j] + "} but does not declare it"
        );
      }
    }
  });

  test("required parameters are marked required and have a source", function () {
    for (let i = 0; i < ENDPOINTS.length; i++) {
      const endpoint = ENDPOINTS[i];
      for (let j = 0; j < endpoint.parameters.length; j++) {
        const parameter = endpoint.parameters[j];
        assertTrue(parameter.description.length > 0, endpoint.id + "." + parameter.name + " has no description");
        assertTrue(parameter.source.length > 0, endpoint.id + "." + parameter.name + " has no source");
      }
    }
  });

  test("destructive endpoints either need a digest or document why not", function () {
    for (let i = 0; i < ENDPOINTS.length; i++) {
      const endpoint = ENDPOINTS[i];
      if (endpoint.destructive) {
        assertTrue(
          endpoint.requiresDigest || endpoint.auth === "bearer-graph",
          endpoint.id + " is destructive but does not require a digest"
        );
      }
    }
  });

  group("registry: lookups");

  test("getEndpoint returns the definition and hasEndpoint agrees", function () {
    assertEquals(getEndpoint("lists.collection.get").module, "lists");
    assertTrue(hasEndpoint("lists.collection.get"), "known id should be found");
    assertTrue(!hasEndpoint("nope.nope"), "unknown id should not be found");
  });

  test("getEndpoint throws a helpful error for unknown ids", function () {
    assertThrows(function () {
      getEndpoint("nope.nope");
    }, "Unknown endpoint");
  });

  test("listEndpoints returns a copy, not the internal array", function () {
    const copy = listEndpoints();
    copy.push(copy[0]);
    assertTrue(copy.length !== ENDPOINTS.length, "listEndpoints must not expose the internal array");
  });

  test("modules are sorted and contain lists and items", function () {
    const modules = listModules();
    const sorted = modules.slice(0).sort();
    assertEquals(modules, sorted);
    assertTrue(modules.indexOf("lists") >= 0, "lists module must exist");
    assertTrue(modules.indexOf("items") >= 0, "items module must exist");
  });

  test("family filter returns only that family", function () {
    const rest = listEndpointIdsByFamily("sharepoint-rest");
    assertTrue(rest.length > 0, "expected sharepoint-rest endpoints");
    for (let i = 0; i < rest.length; i++) {
      assertEquals(getEndpoint(rest[i]).family, "sharepoint-rest");
    }
  });

  test("all eight HTTP methods are represented in METHODS_OF_INTEREST", function () {
    assertEquals(METHODS_OF_INTEREST.length, 8);
    ["GET", "POST", "PATCH", "PUT", "DELETE", "HEAD", "OPTIONS", "MERGE"].forEach(function (method) {
      assertTrue(METHODS_OF_INTEREST.indexOf(method as never) >= 0, method + " missing");
    });
  });

  group("registry: lists coverage (added from PnPjs 4.9.0 source)");

  [
    "lists.effectiveBasePermissions",
    "lists.eventReceivers",
    "lists.relatedFields",
    "lists.irmSettings",
    "lists.renderListFormData",
    "lists.reserveListItemId",
    "lists.forms",
    "lists.subscriptions",
    "lists.subscription.add",
    "lists.userCustomActions",
    "lists.defaultView",
    "lists.roleAssignments",
    "lists.breakRoleInheritance",
    "lists.resetRoleInheritance",
    "lists.ensureSiteAssetsLibrary",
    "lists.ensureSitePagesLibrary",
  ].forEach(function (id) {
    test("registers " + id, function () {
      const endpoint = getEndpoint(id);
      assertEquals(endpoint.module, "lists");
      assertTrue(endpoint.urlTemplate.indexOf("/_api/") === 0, "must target _api");
      assertTrue(endpoint.sources.indexOf("pnpjs-source") >= 0 || endpoint.sources.indexOf("msdocs") >= 0, "needs a source");
    });
  });

  test("registry covers the documented list surface", function () {
    const onlyListEndpoints = listEndpoints().filter(function (endpoint) {
      return endpoint.module === "lists";
    });
    assertTrue(
      onlyListEndpoints.length >= 29,
      "expected at least 29 list endpoints, got " + onlyListEndpoints.length
    );
  });
}
