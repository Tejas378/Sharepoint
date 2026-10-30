/**
 * Backward-compatible facade over the centralised test configuration.
 *
 * The legacy `CONFIG` shape consumed by DataService.ts is preserved, but every
 * value is now derived from src/config/test-config.json. Do not add literals
 * here: change the JSON instead.
 *
 * `sendRequest` is the shared, corrected low-level helper. The previous
 * implementation lived in config.js and had four defects, fixed below:
 *   1. every HTTP response resolved (even 4xx/5xx), so callers parsed errors;
 *   2. bodies were copied char-code-by-char into a Blob, corrupting non-ASCII;
 *   3. the digest was read from localStorage with a "*" fallback and no
 *      acquisition, so the first write always failed;
 *   4. If-Match: * was sent on GET and no timeout was applied.
 */
import { DigestProvider, parseContextInfoResponse } from "../research/digest";
import { rewriteTenantUrl } from "../research/url";
import {
  assertAllowedUrl,
  loadTestConfig,
  TestConfig,
} from "../config/loader";

const TEST_CONFIG: TestConfig = loadTestConfig();

/**
 * Hosts that earlier versions of the data services embedded in URL literals.
 * Requests to these hosts are rewritten onto environment.siteUrl so changing
 * tenant no longer requires editing hundreds of strings. New code must build
 * URLs from CONFIG / loadTestConfig(), not from this list.
 */
const LEGACY_TENANT_HOSTS = ["tmsobb2.sharepoint.com"];

export const HTTP_METHODS: string[] = ["GET", "POST", "PUT", "PATCH", "MERGE", "DELETE"];

function siteUrl(): string {
  return TEST_CONFIG.environment.siteUrl;
}

function listRoute(): string {
  return "_api/web/lists('" + TEST_CONFIG.resources.lists.primary.id + "')";
}

/** First title from the centralised valid item payload, used by legacy helpers. */
function itemTitle(): string {
  const item = TEST_CONFIG.testData.item as
    | { valid?: { Title?: string } }
    | undefined;
  const title = item && item.valid && item.valid.Title;
  return typeof title === "string" && title.length > 0 ? title : "Title probe";
}

/**
 * Legacy configuration object. Field names are kept exactly as they were so
 * existing code keeps compiling and behaving the same way.
 *
 * @deprecated New code should import { loadTestConfig } from "../config/loader"
 * and read the typed sections instead.
 */
export const CONFIG = {
  lid: TEST_CONFIG.resources.lists.primary.id,
  itemId: TEST_CONFIG.resources.items.primaryItemId,
  folderId: Number(TEST_CONFIG.resources.folders.primaryFolderId),
  listName: TEST_CONFIG.resources.lists.primary.title,
  LibraryName: TEST_CONFIG.resources.lists.documents.title,
  subFolderName: String(TEST_CONFIG.resources.folders.testFolderName),
  viewName: String(TEST_CONFIG.resources.views.viewName),
  fileName: TEST_CONFIG.resources.files.testFileName,
  SITE_URL: TEST_CONFIG.environment.siteUrl,
  personGroup: TEST_CONFIG.principals.groups.shareWith,
  userLoginName: TEST_CONFIG.principals.users.primary.loginName,
  kinds: TEST_CONFIG.principals.shareLinkKinds.filter(function (kind) {
    return kind >= 2;
  }),
  scope: TEST_CONFIG.principals.scope,
  groupIds: TEST_CONFIG.principals.groups.groupsToEnumerate,
  groupId: TEST_CONFIG.principals.groups.primaryGroupId,
  groupName: TEST_CONFIG.principals.groups.primaryGroupName,
  userid: TEST_CONFIG.principals.users.primary.id,
  roles: TEST_CONFIG.principals.roleDefinitions.all,
  roles2: TEST_CONFIG.principals.shareLinkRoleKinds,
  principleIds: TEST_CONFIG.principals.principalIds,
  HTTPMethods: HTTP_METHODS,

  FOLDER_OPERATIONS: {
    TEST_FOLDER_NAME: String(TEST_CONFIG.resources.folders.testFolderName),
    TEST_FOLDER_PATH: String(TEST_CONFIG.resources.folders.serverRelativePath),
    UPDATED_FOLDER_NAME: String(TEST_CONFIG.resources.folders.createdFolderName),
    TEST_ITEM_TITLE: itemTitle(),
    TEST_ITEM_TITLE_2: itemTitle() + " 2",
    TEST_ITEM_TITLE_3: itemTitle() + " 3",
    FOLDER_PAYLOAD_PROG_ID: "11111111111111111111",
    FOLDER_PAYLOAD_SERVER_URL: "/Lists/" + TEST_CONFIG.resources.lists.primary.title + "/",
    FOLDER_PAYLOAD_NAME: "folder-probe",
    FOLDER_PAYLOAD_WELCOME_PAGE: "welcome-probe.aspx",
    SHARE_LINK_KIND: TEST_CONFIG.principals.shareLinkKinds.length > 2 ? TEST_CONFIG.principals.shareLinkKinds[2] : 2,
    DEFAULT_COLUMN_VALUE: 1,
  },

  API_ENDPOINTS: {
    VALIDATE_UPDATE_FETCH_LIST_ITEM:
      siteUrl() + "/_api/web/GetList(@a1)/items(@a2)/ValidateUpdateFetchListItem()",
    GET_LIST_DATA_AS_STREAM: siteUrl() + "/_api/SP.List.GetListDataAsStream",
    FOLDER_OPERATIONS_BASE:
      siteUrl() + "/" + listRoute() + "/items(" + TEST_CONFIG.resources.items.primaryItemId + ")/folder",
    MOVE_COPY_UTIL: siteUrl() + "/_api/SP.MoveCopyUtil.MoveFileByPath()",
  },

  VIEW_CONFIG: {
    VIEW_ID: TEST_CONFIG.resources.views.viewId,
    RENDER_OPTIONS: 1496583,
    RENDER_OPTIONS_2: 1069063,
    VIEW_XML: TEST_CONFIG.resources.views.viewXml,
  },

  SHARING: {
    EMAILS: {
      TMSOBB: TEST_CONFIG.principals.users.primary.upn,
      ATTACKER: TEST_CONFIG.principals.users.secondary.upn,
    },
    SHARING_PERMISSIONS: [
      { email: TEST_CONFIG.principals.users.primary.upn },
      { email: TEST_CONFIG.principals.users.secondary.upn },
    ],
  },

  MOVE_COPY: {
    DEST_PATH:
      siteUrl() + "/Lists/" + TEST_CONFIG.resources.lists.primary.title + "/" +
      TEST_CONFIG.resources.lists.primary.title + "/1_.000",
    SRC_PATH:
      siteUrl() + "/Lists/" + TEST_CONFIG.resources.lists.primary.title + "/" +
      TEST_CONFIG.resources.lists.primary.title + "/1_.000",
    RETAIN_EDITOR_ON_MOVE: true,
    BYPASS_SHARED_LOCKS: true,
  },

  DATA_SERVICE: {
    FOLDER_NAME: String(TEST_CONFIG.resources.folders.createdFolderName),
  },
};

/** Exposed for callers that need the typed configuration. */
export const config: TestConfig = TEST_CONFIG;

const READ_METHODS = ["GET", "HEAD", "OPTIONS"];

let digestProvider: DigestProvider | null = null;

function getDigestProvider(): DigestProvider {
  if (!digestProvider) {
    digestProvider = new DigestProvider({
      fetchDigest: async function () {
        const response = await fetch(TEST_CONFIG.environment.siteUrl + "/_api/contextinfo", {
          method: "POST",
          headers: { Accept: "application/json;odata=verbose" },
          credentials: "include",
        });
        if (!response.ok) {
          throw new Error("contextinfo failed with HTTP " + response.status);
        }
        return parseContextInfoResponse(await response.json());
      },
    });
  }
  return digestProvider;
}

/** Drops the cached request digest (call after a digest rejection). */
export function invalidateDigest(): void {
  getDigestProvider().invalidate();
}

function truncate(value: string, max: number): string {
  return value.length > max ? value.substring(0, max) + "...[truncated]" : value;
}

/**
 * Single low-level request helper for absolute SharePoint URLs.
 *
 * Throws on network errors, timeouts and non-2xx responses. JSON parsing only
 * happens when requested and tolerates an empty body.
 *
 * `options.observeRejection` attaches a logging handler to the returned promise
 * without changing it. Legacy fire-and-forget probes then cannot surface an
 * unhandled rejection, while callers that await still receive the error.
 */
export interface SendRequestOptions {
  accept?: string;
  contentType?: string;
  /** null omits If-Match; a string overrides it; undefined uses the defaults. */
  ifMatch?: string | null;
  timeoutMs?: number;
  observeRejection?: boolean;
}

export function normaliseLegacyUrl(url: string): string {
  return rewriteTenantUrl(url, siteUrl(), LEGACY_TENANT_HOSTS);
}

export function sendRequest(
  method: string,
  url: string,
  body: string | Record<string, unknown> = "",
  parseJson: boolean = true,
  options: SendRequestOptions = {}
): Promise<unknown> {
  const pending = performRequest(method, url, body, parseJson, options);
  if (options.observeRejection) {
    pending.catch(function (error: unknown) {
      console.warn(
        "[sendRequest] observed failure: " +
          (error instanceof Error ? error.message : String(error))
      );
    });
  }
  return pending;
}

async function performRequest(
  method: string,
  url: string,
  body: string | Record<string, unknown>,
  parseJson: boolean,
  options: SendRequestOptions
): Promise<unknown> {
  url = normaliseLegacyUrl(url);
  assertAllowedUrl(url, TEST_CONFIG);

  const intended = String(method || "GET").toUpperCase();
  const isRead = READ_METHODS.indexOf(intended) >= 0;
  const headers: Record<string, string> = {
    Accept: options.accept || "application/json;odata=verbose",
  };

  let sentMethod = intended;
  if (intended === "MERGE") {
    sentMethod = "POST";
    headers["X-HTTP-Method"] = "MERGE";
  }

  const payload =
    body === "" || body === undefined || body === null
      ? undefined
      : typeof body === "string"
      ? body
      : JSON.stringify(body);

  if (!isRead) {
    headers["Content-Type"] = options.contentType || "application/json;odata=verbose";
    headers["X-RequestDigest"] = await getDigestProvider().get(false);
    if (options.ifMatch === null) {
      // Deliberately no conditional header (e.g. probing unconditional writes).
    } else if (options.ifMatch !== undefined) {
      headers["If-Match"] = options.ifMatch;
    } else if (
      intended === "DELETE" ||
      intended === "PATCH" ||
      intended === "PUT" ||
      intended === "MERGE"
    ) {
      headers["If-Match"] = "*";
    }
  }

  const timeoutMs =
    options.timeoutMs !== undefined ? options.timeoutMs : TEST_CONFIG.safety.requestTimeoutMs;
  const controller = typeof AbortController !== "undefined" ? new AbortController() : undefined;
  const timer =
    controller !== undefined
      ? setTimeout(function () {
          controller.abort();
        }, timeoutMs)
      : undefined;

  let response: Response;
  try {
    response = await fetch(url, {
      method: sentMethod,
      headers: headers,
      credentials: "include",
      body: payload,
      signal: controller ? controller.signal : undefined,
    });
  } catch (error) {
    throw new Error(
      "Network error calling " + url + ": " + (error instanceof Error ? error.message : String(error))
    );
  } finally {
    if (timer !== undefined) {
      clearTimeout(timer);
    }
  }

  const text = await response.text();
  if (!response.ok) {
    const detail = truncate(text, 500);
    if (response.status === 403) {
      // Let the next write fetch a fresh digest instead of reusing a stale one.
      invalidateDigest();
    }
    throw new Error(
      "HTTP " + response.status + " " + response.statusText + " from " + url + (detail ? ": " + detail : "")
    );
  }

  if (!parseJson || text === "") {
    return text;
  }
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error("Response was not valid JSON from " + url + ": " + truncate(text, 200));
  }
}
