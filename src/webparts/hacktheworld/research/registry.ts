import {
  EndpointDefinition,
  EndpointParameter,
  EvidenceKind,
  HttpMethod,
  ParamType,
  PayloadFamily,
  RequestBodyDefinition,
  VerificationStatus,
} from "./types";

/**
 * Endpoint registry.
 *
 * Rules for adding entries:
 *  - URL templates must be traceable to Microsoft documentation, the installed
 *    PnPjs 4.9.0 source, or observed traffic (mark the source accordingly).
 *  - Never list a parameter as "documented" unless it is in one of those
 *    sources. Guesses use verification: "unverified".
 *  - Registering an endpoint means "the request can be generated". It does NOT
 *    mean the request was sent or that the behaviour was verified.
 */

const SP_SELECT = "$select";
const SP_FILTER = "$filter";
const SP_EXPAND = "$expand";
const SP_ORDERBY = "$orderby";
const SP_TOP = "$top";
const SP_SKIPTOKEN = "$skiptoken";

function queryParam(
  name: string,
  description: string,
  example?: string | number | boolean,
  source: EvidenceKind = "msdocs",
  verification: VerificationStatus = "documented"
): EndpointParameter {
  return {
    name: name,
    location: "query",
    type: "string",
    required: false,
    description: description,
    example: example,
    source: source,
    verification: verification,
  };
}

function pathParam(
  name: string,
  type: ParamType,
  description: string,
  example?: string | number | boolean,
  source: EvidenceKind = "msdocs",
  verification: VerificationStatus = "documented"
): EndpointParameter {
  return {
    name: name,
    location: "path",
    type: type,
    required: true,
    description: description,
    example: example,
    source: source,
    verification: verification,
  };
}

function bodyParam(
  name: string,
  description: string,
  type: ParamType = "json",
  source: EvidenceKind = "msdocs",
  verification: VerificationStatus = "documented"
): EndpointParameter {
  return {
    name: name,
    location: "body",
    type: type,
    required: false,
    description: description,
    source: source,
    verification: verification,
  };
}

function jsonBody(
  contentType: string,
  description: string,
  example: unknown,
  options?: {
    odataType?: string;
    requiredProperties?: string[];
    protectedProperties?: string[];
  }
): RequestBodyDefinition {
  return {
    contentType: contentType,
    description: description,
    example: example,
    odataType: options && options.odataType,
    requiredProperties: options && options.requiredProperties,
    protectedProperties: options && options.protectedProperties,
  };
}

function endpoint(def: EndpointDefinition): EndpointDefinition {
  return def;
}

const LIST_QUERY_PARAMS: EndpointParameter[] = [
  queryParam(SP_SELECT, "Comma separated fields to return.", "Id,Title"),
  queryParam(SP_FILTER, "OData filter expression.", "Id gt 0"),
  queryParam(SP_EXPAND, "Comma separated navigation properties to expand.", "Author"),
  queryParam(SP_ORDERBY, "OData ordering expression.", "Id desc"),
  queryParam(SP_TOP, "Maximum number of items to return.", 100),
  queryParam(SP_SKIPTOKEN, "Server pagination token (Paged=TRUE&p_ID=n).", "Paged=TRUE&p_ID=10"),
];

const ITEM_PAYLOAD_FAMILIES: PayloadFamily[] = [
  "idor",
  "mass-assignment",
  "odata-injection",
  "xss",
  "sqli",
  "method-confusion",
  "input-validation",
  "business-logic",
  "authz",
];

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

export const ENDPOINTS: EndpointDefinition[] = [
  // ---------------------------------------------------------------------------
  // Context / digest
  // ---------------------------------------------------------------------------
  endpoint({
    id: "rest.contextinfo",
    family: "sharepoint-rest",
    module: "context",
    operation: "Get request digest and web context",
    methods: ["POST"],
    urlTemplate: "/_api/contextinfo",
    description:
      "Returns FormDigestValue / FormDigestTimeoutSeconds used for write operations. This is the only endpoint that must not send X-RequestDigest.",
    parameters: [],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["input-validation"],
    sources: ["msdocs", "project-code"],
  }),

  // ---------------------------------------------------------------------------
  // Lists
  // ---------------------------------------------------------------------------
  endpoint({
    id: "lists.collection.get",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Enumerate lists",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists",
    description: "Returns every list in the current web.",
    parameters: LIST_QUERY_PARAMS,
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "input-validation", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "lists.collection.add",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Create list",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists",
    description: "Creates a list. Documented body is an SP.List entity.",
    parameters: [
      bodyParam("Title", "List title."),
      bodyParam("BaseTemplate", "List template id (100 generic, 101 document library).", "int"),
      bodyParam("Description", "List description."),
      bodyParam("ContentTypesEnabled", "Enable content types.", "bool"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "SP.List creation payload.",
      { __metadata: { type: "SP.List" }, Title: "Research list", BaseTemplate: 100, Description: "" },
      { odataType: "SP.List", requiredProperties: ["Title", "BaseTemplate"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["mass-assignment", "xss", "input-validation", "business-logic"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "lists.getByTitle",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Get list by title",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists/getByTitle({listTitle:odata})",
    description: "Resolves a list by display title.",
    parameters: [
      pathParam("listTitle", "string", "List display title.", "Documents"),
      queryParam(SP_SELECT, "Fields to return.", "Id,Title"),
      queryParam(SP_EXPAND, "Navigation properties to expand.", "Fields"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "input-validation", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "lists.getById",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Get list by id",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})",
    description: "Resolves a list by GUID. Alternative identifier form for the same resource.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      queryParam(SP_SELECT, "Fields to return.", "Id,Title"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "odata-injection", "input-validation"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "lists.update",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Update list properties",
    methods: ["PATCH", "MERGE"],
    urlTemplate: "/_api/web/lists({listId:guid})",
    description:
      "Updates list metadata. SharePoint accepts PATCH with application/json;odata=verbose or POST + X-HTTP-Method: MERGE.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("Title", "New title."),
      bodyParam("Description", "New description."),
      bodyParam("Hidden", "Hide the list.", "bool"),
      bodyParam("EnableVersioning", "Enable versioning.", "bool"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "SP.List update payload.",
      { __metadata: { type: "SP.List" }, Title: "Updated title" },
      {
        odataType: "SP.List",
        protectedProperties: ["Id", "Created", "Modified", "AuthorId", "EditorId", "ItemCount"],
      }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: true,
    payloadFamilies: ["mass-assignment", "xss", "input-validation", "business-logic", "method-confusion"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "lists.delete",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Delete list",
    methods: ["DELETE"],
    urlTemplate: "/_api/web/lists({listId:guid})",
    description: "Deletes a list. Destructive: only run against disposable research lists.",
    parameters: [pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000")],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: true,
    payloadFamilies: ["idor", "method-confusion", "authz", "business-logic"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "lists.recycle",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Recycle list",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/recycle",
    description: "Moves a list to the recycle bin.",
    parameters: [pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000")],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["idor", "method-confusion", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "lists.getChanges",
    family: "sharepoint-rest",
    module: "changes",
    operation: "Get list changes",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/getchanges",
    description: "Returns changes for a list using an SP.ChangeQuery payload.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("query", "SP.ChangeQuery object.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Change query payload.",
      {
        __metadata: { type: "SP.ChangeQuery" },
        Add: true,
        Update: true,
        DeleteObject: true,
        Web: true,
        List: true,
        Item: true,
        RowLimit: 100,
      },
      { odataType: "SP.ChangeQuery" }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["input-validation", "business-logic", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "lists.getItemChangesSinceToken",
    family: "sharepoint-rest",
    module: "changes",
    operation: "Get item changes since token",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/getlistitemchangessincetoken",
    description:
      "Returns item changes as XML since ChangeToken. PnPjs posts { query: { ChangeToken, RowLimit } }.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("query", "Change log query object.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Change log query.",
      { query: { ChangeToken: "*", RowLimit: 100 } },
      { requiredProperties: ["query"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["input-validation", "business-logic"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "lists.renderListDataAsStream",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Render list data as stream",
    methods: ["POST"],
    urlTemplate:
      "/_api/web/GetListUsingPath(DecodedUrl=@listPath)/RenderListDataAsStream?@listPath={listPath:odataEncoded}",
    description:
      "Modern list rendering endpoint. Accepts RenderListDataParameters including ViewXml and RenderOptions; returned metadata often includes permissions and field schema.",
    parameters: [
      pathParam("listPath", "path", "Server relative list path, e.g. /Lists/Tasks.", "/Lists/Tasks"),
      queryParam("RootFolder", "Folder to render.", "/Lists/Tasks/Folder"),
      queryParam("View", "View GUID to apply.", "00000000-0000-0000-0000-000000000000"),
      queryParam("TryNewExperienceSingle", "Modern experience flag.", "TRUE"),
      bodyParam("parameters", "SP.RenderListDataParameters object.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Render parameters. ViewXml is the interesting input.",
      {
        parameters: {
          __metadata: { type: "SP.RenderListDataParameters" },
          RenderOptions: 1496583,
          AddRequiredFields: true,
          RequireFolderColoringFields: true,
        },
      }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "xss", "input-validation", "business-logic"],
    sources: ["project-code", "observed-traffic"],
  }),
  endpoint({
    id: "lists.renderListData",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Render list data (ViewXml)",
    methods: ["POST"],
    urlTemplate:
      "/_api/web/lists({listId:guid})/renderlistdata(@viewXml)?@viewXml={viewXml:odataEncoded}",
    description: "Server-side CAML rendering. ViewXml is passed as an @alias value.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("viewXml", "string", "CAML ViewXml.", "<View><RowLimit>10</RowLimit></View>"),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "xss", "input-validation"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "lists.getItemsByCamlQuery",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Get items by CAML query",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/getitems",
    description: "CAML query endpoint. Supports row limits up to the list view threshold.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("query", "CAML query object (ViewXml, RowLimit, DatesInUtc, FolderServerRelativeUrl).", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "CAML query payload.",
      {
        query: {
          ViewXml: "<View><Query><Where><IsNotNull><FieldRef Name='Title'/></IsNotNull></Where></Query></View>",
          RowLimit: 100,
        },
      },
      { requiredProperties: ["query"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "input-validation", "xss", "business-logic"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "lists.addValidateUpdateItemUsingPath",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Add / validate item using path",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/AddValidateUpdateItemUsingPath()",
    description:
      "Creates or validates an item in a folder path. PnPjs posts { formValues, bNewDocumentUpdate, checkInComment, FolderPath }.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("formValues", "Array of {FieldName, FieldValue}.", "json"),
      bodyParam("FolderPath", "Target folder resource path.", "json"),
      bodyParam("bNewDocumentUpdate", "Document library update flag.", "bool"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Validate/update payload.",
      {
        formValues: [{ FieldName: "Title", FieldValue: "Controlled probe" }],
        bNewDocumentUpdate: false,
        checkInComment: "",
        FolderPath: { DecodedUrl: "/Lists/Tasks" },
      },
      { requiredProperties: ["formValues", "FolderPath"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["mass-assignment", "xss", "input-validation", "business-logic"],
    sources: ["pnpjs-source", "project-code"],
  }),

  // ---------------------------------------------------------------------------
  // List configuration, permissions and sub-resources
  //
  // URLs below were read from the installed @pnp/sp 4.9.0 sources:
  //   node_modules/@pnp/sp/lists/types.js        (EffectiveBasePermissions,
  //     EventReceivers, getRelatedFields, InformationRightsManagementSettings,
  //     renderlistformdata, reservelistitemid, ensuresiteassetslibrary,
  //     ensuresitepageslibrary, forms)
  //   node_modules/@pnp/sp/security/funcs.js     (breakroleinheritance,
  //     resetroleinheritance)
  //   node_modules/@pnp/sp/security/types.js     (roleassignments)
  //   node_modules/@pnp/sp/user-custom-actions/* (usercustomactions)
  //   node_modules/@pnp/sp/subscriptions/types.js(subscriptions)
  // These were missing from the registry before this change; nothing that was
  // already registered was duplicated.
  // ---------------------------------------------------------------------------
  endpoint({
    id: "lists.effectiveBasePermissions",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Get list effective base permissions",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/EffectiveBasePermissions",
    description:
      "Returns the permission mask the current user has on this list. Useful to establish an authorization baseline before probing item-level access.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["authz", "idor"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "lists.eventReceivers",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Enumerate list event receivers",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/EventReceivers",
    description:
      "Lists the server-side event receivers attached to the list. Reveals assembly names, URLs and (for remote receivers) endpoints.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      queryParam(SP_SELECT, "Fields to return.", "Id,ReceiverName,ReceiverClass,ReceiverAssembly"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["authz", "odata-injection", "input-validation"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "lists.relatedFields",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Get related fields of a lookup",
    methods: ["GET", "POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/getRelatedFields",
    description:
      "Returns fields projected from a lookup relationship. PnPjs exposes it as a queryable (GET); the REST method is also callable as a POST with fieldId/groupId.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      queryParam("fieldId", "Lookup field id.", "00000000-0000-0000-0000-000000000000", "pnpjs-source", "unverified"),
      queryParam("groupId", "Field group id.", "00000000-0000-0000-0000-000000000000", "pnpjs-source", "unverified"),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "input-validation", "authz"],
    sources: ["pnpjs-source"],
    notes: "Query parameter names come from the REST method signature and are marked unverified.",
  }),
  endpoint({
    id: "lists.irmSettings",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Get / update IRM settings",
    methods: ["GET", "PATCH"],
    urlTemplate: "/_api/web/lists({listId:guid})/InformationRightsManagementSettings",
    description:
      "Information Rights Management configuration for the list. Writable properties include IrmEnabled, IrmExpire, IrmReject and IrmAppliedTo.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("IrmEnabled", "Enable IRM on the list.", "bool"),
      bodyParam("IrmReject", "Reject documents that do not support IRM.", "bool"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "SP.InformationRightsManagementSettings update payload.",
      { __metadata: { type: "SP.InformationRightsManagementSettings" }, IrmEnabled: false },
      { odataType: "SP.InformationRightsManagementSettings" }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["mass-assignment", "input-validation", "authz"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "lists.renderListFormData",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Render list form data",
    methods: ["POST"],
    urlTemplate:
      "/_api/web/lists({listId:guid})/renderlistformdata(itemid={itemId:int},formid='{formId:raw}',mode='{mode:raw}')",
    description:
      "Returns the field values and form schema used by the classic edit/display forms (ControlMode 1=Display, 2=Edit, 3=New).",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Target item id.", 1),
      pathParam("formId", "string", "Form id.", "00000000-0000-0000-0000-000000000000"),
      pathParam("mode", "enum", "1 Display, 2 Edit, 3 New.", 2),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "input-validation", "authz"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "lists.reserveListItemId",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Reserve a list item id",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/reservelistitemid",
    description:
      "Reserves the next item id so an item can be created idempotently. No body. Returns the reserved id as an integer.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["business-logic", "input-validation"],
    sources: ["pnpjs-source"],
    notes: "Consumes an id from the list sequence; harmless but not free of side effects.",
  }),
  endpoint({
    id: "lists.forms",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Enumerate list forms",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/Forms",
    description: "Lists the Display/Edit/New form definitions attached to the list.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      queryParam(SP_SELECT, "Fields to return.", "Id,FormType,ServerRelativeUrl"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["authz", "odata-injection"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "lists.subscriptions",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Enumerate webhook subscriptions",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/subscriptions",
    description:
      "Lists webhook subscriptions for the list. Exposes notification endpoints and clientState values.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["authz", "input-validation", "ssrf"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "lists.subscription.add",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Create webhook subscription",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/subscriptions",
    description:
      "Registers a webhook. The service immediately calls notificationUrl, so only point it at an endpoint you control. SSRF-sensitive.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("notificationUrl", "Callback URL.", "string"),
      bodyParam("expirationDateTime", "Expiry (max 180 days).", "date"),
      bodyParam("clientState", "Opaque value echoed to the callback.", "string"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "SP.Subscription creation payload.",
      {
        "resource": "https://contoso.sharepoint.com/_api/web/lists('00000000-0000-0000-0000-000000000000')",
        "notificationUrl": "https://example.invalid/webhook",
        "expirationDateTime": "2026-01-01T00:00:00Z",
        "clientState": "research",
      },
      { requiredProperties: ["resource", "notificationUrl", "expirationDateTime"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["ssrf", "input-validation", "business-logic", "xss"],
    sources: ["pnpjs-source"],
    notes:
      "Creates tenant state and triggers an outbound request. Requires explicit out-of-band approval; never use a third-party URL.",
  }),
  endpoint({
    id: "lists.userCustomActions",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Enumerate list user custom actions",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/UserCustomActions",
    description:
      "Custom actions (scripts, ribbon buttons) scoped to the list. ScriptSrc/ScriptBlock are a stored-XSS sink and an information disclosure source.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      queryParam(SP_SELECT, "Fields to return.", "Id,Name,Location,ScriptSrc"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["xss", "authz", "odata-injection"],
    sources: ["pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "lists.defaultView",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Get default view",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/DefaultView",
    description: "Returns the list's default view definition.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      queryParam(SP_SELECT, "Fields to return.", "Id,Title,ViewFields"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "input-validation", "authz"],
    sources: ["msdocs"],
    notes: "DefaultView is a documented SP.List navigation property; not wrapped by this PnPjs version.",
  }),
  endpoint({
    id: "lists.roleAssignments",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Enumerate list role assignments",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/roleassignments",
    description:
      "ACL of the list itself. Compare with item and web role assignments to find inheritance gaps.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      queryParam(SP_EXPAND, "Expand member and bindings.", "Member,RoleDefinitionBindings"),
      queryParam(SP_SELECT, "Fields to return.", "PrincipalId,Member/Title"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["authz", "idor", "odata-injection"],
    sources: ["pnpjs-source", "msdocs"],
  }),
  endpoint({
    id: "lists.breakRoleInheritance",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Break list role inheritance",
    methods: ["POST"],
    urlTemplate:
      "/_api/web/lists({listId:guid})/breakroleinheritance(copyroleassignments={copyRoleAssignments:raw},clearsubscopes={clearSubscopes:raw})",
    description:
      "Creates unique permissions on the list. Reversible only with resetroleinheritance; the original inherited ACL is not restored.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("copyRoleAssignments", "bool", "Copy current assignments before breaking.", false),
      pathParam("clearSubscopes", "bool", "Clear unique permissions on children.", false),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["authz", "business-logic", "input-validation"],
    sources: ["pnpjs-source"],
    notes: "Destructive permission change. Disabled by default in the test configuration.",
  }),
  endpoint({
    id: "lists.resetRoleInheritance",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Reset list role inheritance",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/resetroleinheritance",
    description: "Restores inherited permissions on the list.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["authz", "business-logic"],
    sources: ["pnpjs-source"],
    notes: "Destructive permission change. Disabled by default in the test configuration.",
  }),
  endpoint({
    id: "lists.ensureSiteAssetsLibrary",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Ensure Site Assets library",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists/ensuresiteassetslibrary",
    description:
      "Returns (creating if needed) the Site Assets library. Creates a library on a web that does not have one.",
    parameters: [],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["business-logic", "authz"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "lists.ensureSitePagesLibrary",
    family: "sharepoint-rest",
    module: "lists",
    operation: "Ensure Site Pages library",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists/ensuresitepageslibrary",
    description:
      "Returns (creating if needed) the Site Pages library. Creates a library on a web that does not have one.",
    parameters: [],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["business-logic", "authz"],
    sources: ["pnpjs-source"],
  }),

  // ---------------------------------------------------------------------------
  // List items
  // ---------------------------------------------------------------------------
  endpoint({
    id: "items.collection.get",
    family: "sharepoint-rest",
    module: "items",
    operation: "Enumerate list items",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/items",
    description:
      "Primary item collection. Supports OData query options and server-driven paging via $skiptoken / __next.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      queryParam(SP_SELECT, "Fields to return.", "Id,Title,Created,Modified"),
      queryParam(SP_FILTER, "OData filter. All list items are security-trimmed by the service.", "Id gt 0"),
      queryParam(SP_EXPAND, "Navigation properties to expand.", "Author,Editor"),
      queryParam(SP_ORDERBY, "Ordering expression.", "Id desc"),
      queryParam(SP_TOP, "Page size.", 100),
      queryParam(SP_SKIPTOKEN, "Paging token returned in __next.", "Paged=TRUE&p_ID=10"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "odata-injection", "input-validation", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "items.collection.add",
    family: "sharepoint-rest",
    module: "items",
    operation: "Create list item",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/items",
    description:
      "Creates a list item. The odata type name must match the list entity type (SP.Data.<List>ListItem) for verbose payloads; nometadata accepts plain JSON.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("body", "Field values keyed by internal name.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Item creation payload.",
      { __metadata: { type: "SP.Data.TasksListItem" }, Title: "Controlled probe" },
      {
        odataType: "SP.Data.<ListName>ListItem",
        requiredProperties: ["Title"],
        protectedProperties: ["Id", "Created", "Modified", "AuthorId", "EditorId", "GUID", "ContentTypeId"],
      }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ITEM_PAYLOAD_FAMILIES,
    sources: ["msdocs", "pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "items.getById",
    family: "sharepoint-rest",
    module: "items",
    operation: "Get list item",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})",
    description:
      "Returns a single item. The same item can also be addressed via getItemByStringId, parent list navigation and the v2.0 API.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      queryParam(SP_SELECT, "Fields to return.", "Id,Title"),
      queryParam(SP_EXPAND, "Navigation properties to expand.", "Author"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "odata-injection", "input-validation", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "items.getByStringId",
    family: "sharepoint-rest",
    module: "items",
    operation: "Get item by string id (BCS)",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/getItemByStringId({stringId:odata})",
    description:
      "Alternate item identifier used for external/BCS lists. Useful for identifier-parsing tests.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("stringId", "string", "String identifier.", "1"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "input-validation"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "items.update",
    family: "sharepoint-rest",
    module: "items",
    operation: "Update list item",
    methods: ["PATCH", "MERGE"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})",
    description:
      "Updates fields on an item. Use If-Match with the item ETag to explore conditional update behaviour; PATCH is the modern verb, MERGE the legacy override.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      bodyParam("fields", "Field values keyed by internal name.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Item update payload.",
      { __metadata: { type: "SP.Data.TasksListItem" }, Title: "Updated title" },
      {
        odataType: "SP.Data.<ListName>ListItem",
        protectedProperties: ["Id", "Created", "Modified", "AuthorId", "EditorId", "GUID", "ContentTypeId"],
      }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: true,
    payloadFamilies: ITEM_PAYLOAD_FAMILIES,
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "items.delete",
    family: "sharepoint-rest",
    module: "items",
    operation: "Delete list item",
    methods: ["DELETE"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})",
    description:
      "Deletes an item. If-Match controls optimistic concurrency. deleteObject is the legacy POST equivalent.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: true,
    payloadFamilies: ["idor", "method-confusion", "authz", "business-logic"],
    sources: ["msdocs", "pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "items.recycle",
    family: "sharepoint-rest",
    module: "items",
    operation: "Recycle list item",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/recycle",
    description: "Moves an item to the recycle bin.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["idor", "method-confusion", "authz"],
    sources: ["pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "items.validateUpdateListItem",
    family: "sharepoint-rest",
    module: "items",
    operation: "Validate / update item fields",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/validateupdatelistitem",
    description:
      "Validates field values without a raw property update. Used by modern forms; a good place to observe field-level validation and protected property handling.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      bodyParam("formValues", "Array of {FieldName, FieldValue, HasException, ErrorMessage}.", "json"),
      bodyParam("bNewDocumentUpdate", "Document library update flag.", "bool"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Form values payload. FieldValue is a string in the editor protocol.",
      {
        formValues: [{ FieldName: "Title", FieldValue: "Validated title", HasException: false, ErrorMessage: "" }],
        bNewDocumentUpdate: false,
      },
      { requiredProperties: ["formValues"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["mass-assignment", "xss", "input-validation", "business-logic"],
    sources: ["pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "items.fieldValuesAsText",
    family: "sharepoint-rest",
    module: "items",
    operation: "Get field values as text",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/FieldValuesAsText",
    description: "Returns display-formatted field values, including computed people/lookup text.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "xss", "odata-injection", "authz"],
    sources: ["pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "items.fieldValuesAsHTML",
    family: "sharepoint-rest",
    module: "items",
    operation: "Get field values as HTML",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/FieldValuesAsHTML",
    description:
      "Returns server-rendered HTML for field values. Relevant when checking how stored markup is encoded before it reaches a browser.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["xss", "idor", "authz"],
    sources: ["pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "items.versions",
    family: "sharepoint-rest",
    module: "items",
    operation: "Enumerate item versions",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/versions",
    description: "Lists historical versions. Version data can bypass current-field permissions logic.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "input-validation", "authz"],
    sources: ["pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "items.version.get",
    family: "sharepoint-rest",
    module: "items",
    operation: "Get item version",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/versions({versionId:int})",
    description: "Returns one historical version by VersionId.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      pathParam("versionId", "int", "Integer version id.", 512),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "input-validation", "authz"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "items.attachments",
    family: "sharepoint-rest",
    module: "attachments",
    operation: "Enumerate attachments",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/attachmentFiles",
    description: "Lists attachment metadata for an item.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "authz", "input-validation"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "items.attachment.add",
    family: "sharepoint-rest",
    module: "attachments",
    operation: "Add attachment",
    methods: ["POST"],
    urlTemplate:
      "/_api/web/lists({listId:guid})/items({itemId:int})/attachmentFiles/add(FileName={fileName:odata})",
    description:
      "Uploads an attachment. Body is the base64 content (not JSON). PnPjs sends raw content with no JSON wrapper.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      pathParam("fileName", "string", "Attachment file name.", "research.txt"),
      bodyParam("content", "Base64 file content or raw payload.", "binary"),
    ],
    requestBody: {
      contentType: "application/octet-stream",
      description: "Base64 string for JSON clients; PnPjs posts the value as the request body.",
      example: "dGVzdA==",
    },
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["path-traversal", "xss", "input-validation", "ssrf", "authz"],
    sources: ["msdocs", "pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "items.attachment.get",
    family: "sharepoint-rest",
    module: "attachments",
    operation: "Get attachment",
    methods: ["GET"],
    urlTemplate:
      "/_api/web/lists({listId:guid})/items({itemId:int})/attachmentFiles({fileName:odata})",
    description: "Returns attachment metadata. Append /$value for the raw content.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      pathParam("fileName", "string", "Attachment file name.", "research.txt"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "path-traversal", "xss", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "items.attachment.delete",
    family: "sharepoint-rest",
    module: "attachments",
    operation: "Delete attachment",
    methods: ["DELETE"],
    urlTemplate:
      "/_api/web/lists({listId:guid})/items({itemId:int})/attachmentFiles({fileName:odata})",
    description: "Deletes one attachment.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      pathParam("fileName", "string", "Attachment file name.", "research.txt"),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: true,
    payloadFamilies: ["idor", "path-traversal", "method-confusion", "authz"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "items.comments",
    family: "sharepoint-rest",
    module: "comments",
    operation: "Enumerate item comments",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/comments",
    description: "Lists modern comments for an item.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "xss", "authz"],
    sources: ["project-code", "observed-traffic"],
  }),
  endpoint({
    id: "items.comment.add",
    family: "sharepoint-rest",
    module: "comments",
    operation: "Add item comment",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/comments",
    description:
      "Adds a comment. Text is rendered in the item pane; mention payloads use a separate mentions array.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      bodyParam("text", "Comment text.", "string"),
      bodyParam("mentions", "Mention objects.", "json"),
    ],
    requestBody: jsonBody("application/json", "Comment payload.", { text: "Controlled probe" }),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["xss", "sqli", "mass-assignment", "input-validation", "business-logic"],
    sources: ["project-code", "observed-traffic"],
  }),
  endpoint({
    id: "items.sharingInformation",
    family: "sharepoint-rest",
    module: "permissions",
    operation: "Get item sharing information",
    methods: ["POST"],
    urlTemplate:
      "/_api/web/lists({listId:guid})/items({itemId:int})/GetSharingInformation?$Expand=permissionsInformation",
    description:
      "Returns effective sharing state and sharing links for an item. Useful for checking link permissions and audience scope.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      bodyParam("request", "maxPrincipalsToReturn / populateInheritedLinks.", "json"),
    ],
    requestBody: jsonBody(
      "application/json",
      "Sharing info request.",
      { request: { maxPrincipalsToReturn: 100, populateInheritedLinks: true } }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["authz", "idor", "business-logic"],
    sources: ["pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "items.shareLink",
    family: "sharepoint-rest",
    module: "sharing",
    operation: "Create sharing link on item",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/shareLink",
    description:
      "Creates a sharing link. The request.settings object controls linkKind, role, expiration, scope and password-related fields.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      bodyParam("request", "Sharing link request.", "json"),
    ],
    requestBody: jsonBody(
      "application/json",
      "Sharing link payload.",
      {
        request: {
          createLink: true,
          emailData: null,
          settings: { expiration: null, linkKind: 5, role: 2, allowAnonymousAccess: false },
        },
      },
      { requiredProperties: ["request"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["business-logic", "mass-assignment", "authz", "input-validation"],
    sources: ["project-code", "observed-traffic"],
  }),
  endpoint({
    id: "items.roleAssignments",
    family: "sharepoint-rest",
    module: "permissions",
    operation: "Get item role assignments",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/roleAssignments",
    description:
      "Returns explicit role assignments for the item. An empty collection can simply mean inheritance, not a permission bypass.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
      queryParam(SP_EXPAND, "Expand Member / RoleDefinitionBindings.", "Member"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["authz", "idor", "odata-injection"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "items.effectiveBasePermissions",
    family: "sharepoint-rest",
    module: "permissions",
    operation: "Get effective base permissions",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/items({itemId:int})/EffectiveBasePermissions",
    description:
      "Returns the effective permission mask for the current user. Compare masks between users rather than treating a numeric mask as a finding.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Integer item id.", 1),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["authz", "idor"],
    sources: ["pnpjs-source", "project-code"],
  }),

  // ---------------------------------------------------------------------------
  // Fields / content types / views
  // ---------------------------------------------------------------------------
  endpoint({
    id: "fields.collection",
    family: "sharepoint-rest",
    module: "fields",
    operation: "Enumerate list fields",
    methods: ["GET", "POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/fields",
    description:
      "GET lists fields; POST (with SP.Field body) creates a field. Field schema is a useful source of field internal names for later payloads.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      queryParam(SP_SELECT, "Fields to return.", "Id,InternalName,Title,TypeAsString"),
      queryParam(SP_FILTER, "Filter expression.", "Hidden eq false"),
      bodyParam("field", "SP.Field properties for creation.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Field creation payload.",
      { __metadata: { type: "SP.Field" }, Title: "ResearchField", FieldTypeKind: 2, Required: false },
      { odataType: "SP.Field" }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["mass-assignment", "xss", "input-validation", "business-logic"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "fields.getByTitle",
    family: "sharepoint-rest",
    module: "fields",
    operation: "Get / update / delete field",
    methods: ["GET", "PATCH", "DELETE"],
    urlTemplate: "/_api/web/lists({listId:guid})/fields/getByTitle({fieldTitle:odata})",
    description: "Field metadata by display title. Update uses SP.Field; delete removes the column.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("fieldTitle", "string", "Field display title.", "Title"),
      bodyParam("field", "SP.Field update properties.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Field update payload.",
      { __metadata: { type: "SP.Field" }, Description: "Updated description" },
      { odataType: "SP.Field", protectedProperties: ["Id", "InternalName", "SchemaXml"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: true,
    payloadFamilies: ["mass-assignment", "xss", "input-validation", "method-confusion", "business-logic"],
    sources: ["msdocs", "pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "fields.createFieldAsXml",
    family: "sharepoint-rest",
    module: "fields",
    operation: "Create field from CAML schema",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/fields/createfieldasxml",
    description:
      "Creates a field from a SchemaXml string. The schema text is a high-value input validation surface.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("parameters", "XML schema and AddFieldOptions.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Field schema payload.",
      {
        parameters: {
          __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
          SchemaXml: "<Field Type='Text' Name='Research' DisplayName='Research' />",
          Options: 2,
        },
      },
      { requiredProperties: ["parameters"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["xss", "input-validation", "mass-assignment", "business-logic"],
    sources: ["pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "contentTypes.collection",
    family: "sharepoint-rest",
    module: "content-types",
    operation: "Enumerate list content types",
    methods: ["GET"],
    urlTemplate: "/_api/web/lists({listId:guid})/contentTypes",
    description: "Lists content types available on the list.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      queryParam(SP_SELECT, "Fields to return.", "StringId,Name"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "contentTypes.byId",
    family: "sharepoint-rest",
    module: "content-types",
    operation: "Get content type",
    methods: ["GET", "PATCH", "DELETE"],
    urlTemplate: "/_api/web/lists({listId:guid})/contentTypes({contentTypeId})",
    description: "Content type metadata by string id (0x...).",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("contentTypeId", "string", "Content type id.", "0x0100"),
      bodyParam("contentType", "SP.ContentType update payload.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Content type update payload.",
      { __metadata: { type: "SP.ContentType" }, Description: "Updated description" },
      { odataType: "SP.ContentType", protectedProperties: ["StringId", "Id"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: true,
    payloadFamilies: ["idor", "mass-assignment", "xss", "method-confusion"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "contentTypes.addAvailable",
    family: "sharepoint-rest",
    module: "content-types",
    operation: "Add available content type",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/contentTypes/addAvailableContentType",
    description: "Adds a site content type to the list.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("contentTypeId", "Content type id to add.", "string"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Content type id payload.",
      { contentTypeId: "0x0100" },
      { requiredProperties: ["contentTypeId"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["mass-assignment", "input-validation", "business-logic"],
    sources: ["pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "views.collection",
    family: "sharepoint-rest",
    module: "views",
    operation: "Enumerate / create views",
    methods: ["GET", "POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/views",
    description:
      "GET lists views. POST creates a view; ViewCreationInformation controls ViewFields, Query and form/view type.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("parameters", "SP.ViewCreationInformation.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "View creation payload.",
      {
        parameters: {
          __metadata: { type: "SP.ViewCreationInformation" },
          Title: "ResearchView",
          RowLimit: 30,
          Paged: true,
          PersonalView: false,
          ViewTypeKind: 0,
          ViewFields: { __metadata: { type: "Collection(Edm.String)" }, results: ["ID", "Title"] },
        },
      },
      { requiredProperties: ["parameters"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["odata-injection", "xss", "mass-assignment", "input-validation"],
    sources: ["msdocs", "pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "views.byId",
    family: "sharepoint-rest",
    module: "views",
    operation: "Get / update / delete view",
    methods: ["GET", "PATCH", "DELETE"],
    urlTemplate: "/_api/web/lists({listId:guid})/views({viewId:guid})",
    description:
      "View metadata. ViewXml/CustomFormatter are stored markup that is later rendered, so they are XSS-relevant inputs.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("viewId", "guid", "View GUID.", "00000000-0000-0000-0000-000000000000"),
      bodyParam("view", "SP.View update payload.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "View update payload.",
      { __metadata: { type: "SP.View" }, Title: "Updated view", Hidden: false },
      { odataType: "SP.View", protectedProperties: ["Id", "ServerRelativeUrl"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: true,
    payloadFamilies: ["xss", "odata-injection", "mass-assignment", "input-validation", "method-confusion"],
    sources: ["msdocs", "pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "views.renderAsHtml",
    family: "sharepoint-rest",
    module: "views",
    operation: "Render view as HTML",
    methods: ["POST"],
    urlTemplate: "/_api/web/lists({listId:guid})/views({viewId:guid})/renderAsHtml",
    description: "Server-side view rendering. Output is HTML that includes list data.",
    parameters: [
      pathParam("listId", "guid", "List GUID.", "00000000-0000-0000-0000-000000000000"),
      pathParam("viewId", "guid", "View GUID.", "00000000-0000-0000-0000-000000000000"),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["xss", "odata-injection", "authz"],
    sources: ["pnpjs-source", "project-code"],
  }),

  // ---------------------------------------------------------------------------
  // Files and folders
  // ---------------------------------------------------------------------------
  endpoint({
    id: "files.byPath",
    family: "sharepoint-rest",
    module: "files",
    operation: "Get file by server relative path",
    methods: ["GET", "DELETE"],
    urlTemplate:
      "/_api/web/GetFileByServerRelativePath(DecodedUrl=@filePath)?@filePath={filePath:odataEncoded}",
    description:
      "Addresses a file by decoded server-relative path. Path handling is a key traversal / authz surface.",
    parameters: [
      pathParam("filePath", "path", "Server relative file path.", "/sites/team/Shared Documents/file.txt"),
      queryParam(SP_SELECT, "Fields to return.", "Length,Name,ServerRelativeUrl"),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: true,
    payloadFamilies: ["idor", "path-traversal", "authz", "method-confusion", "input-validation"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "files.addUsingPath",
    family: "sharepoint-rest",
    module: "files",
    operation: "Upload file using path",
    methods: ["POST"],
    urlTemplate:
      "/_api/web/GetFolderByServerRelativePath(DecodedUrl=@folderPath)/Files/AddUsingPath(DecodedUrl=@filePath,Overwrite={overwrite})?@folderPath={folderPath:odataEncoded}&@filePath={filePath:odataEncoded}",
    description:
      "Modern upload endpoint used by the SharePoint UI. Supports overwrite and AutoCheckoutOnInvalidData flags.",
    parameters: [
      pathParam("folderPath", "path", "Target folder.", "/sites/team/Shared Documents"),
      pathParam("filePath", "path", "File name relative to the folder.", "research.txt"),
      pathParam("overwrite", "bool", "Overwrite existing file.", true),
      queryParam("AutoCheckoutOnInvalidData", "Auto checkout flag.", "true"),
      bodyParam("content", "File content.", "binary"),
    ],
    requestBody: {
      contentType: "text/plain",
      description: "Raw file bytes (or text for manual probes).",
      example: "research",
    },
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["path-traversal", "ssrf", "xss", "input-validation", "authz", "mass-assignment"],
    sources: ["project-code", "observed-traffic"],
  }),
  endpoint({
    id: "files.versions",
    family: "sharepoint-rest",
    module: "files",
    operation: "Enumerate file versions",
    methods: ["GET"],
    urlTemplate:
      "/_api/web/GetFileByServerRelativePath(DecodedUrl=@filePath)/versions?@filePath={filePath:odataEncoded}",
    description: "Lists file versions; version contents may expose pre-redaction values.",
    parameters: [
      pathParam("filePath", "path", "Server relative file path.", "/sites/team/Shared Documents/file.txt"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "authz", "input-validation"],
    sources: ["pnpjs-source"],
  }),
  endpoint({
    id: "folders.byPath",
    family: "sharepoint-rest",
    module: "folders",
    operation: "Get folder by path",
    methods: ["GET", "DELETE"],
    urlTemplate:
      "/_api/web/GetFolderByServerRelativePath(DecodedUrl=@folderPath)?@folderPath={folderPath:odataEncoded}",
    description: "Resolves a folder by decoded path.",
    parameters: [
      pathParam("folderPath", "path", "Server relative folder path.", "/sites/team/Shared Documents"),
      queryParam(SP_SELECT, "Fields to return.", "Name,ServerRelativeUrl,ItemCount"),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: true,
    payloadFamilies: ["idor", "path-traversal", "authz", "method-confusion"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "folders.add",
    family: "sharepoint-rest",
    module: "folders",
    operation: "Create folder",
    methods: ["POST"],
    urlTemplate: "/_api/web/folders/add({folderPath:odata})",
    description: "Creates a folder under a server-relative path.",
    parameters: [
      pathParam("folderPath", "path", "Folder to create.", "/sites/team/Shared Documents/Research"),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["path-traversal", "xss", "input-validation", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "moveCopy.moveFileByPath",
    family: "sharepoint-rest",
    module: "files",
    operation: "Move file by path",
    methods: ["POST"],
    urlTemplate: "/_api/SP.MoveCopyUtil.MoveFileByPath()",
    description:
      "Server-side move. srcPath/destPath are SP.ResourcePath objects; options include ShouldBypassSharedLocks.",
    parameters: [
      bodyParam("srcPath", "SP.ResourcePath source.", "json"),
      bodyParam("destPath", "SP.ResourcePath destination.", "json"),
      bodyParam("options", "SP.MoveCopyOptions.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Move payload.",
      {
        srcPath: {
          __metadata: { type: "SP.ResourcePath" },
          DecodedUrl: "https://contoso.sharepoint.com/sites/team/Shared Documents/a.txt",
        },
        destPath: {
          __metadata: { type: "SP.ResourcePath" },
          DecodedUrl: "https://contoso.sharepoint.com/sites/team/Shared Documents/b.txt",
        },
        options: {
          __metadata: { type: "SP.MoveCopyOptions" },
          RetainEditorAndModifiedOnMove: true,
          ShouldBypassSharedLocks: true,
        },
      },
      { requiredProperties: ["srcPath", "destPath"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["path-traversal", "ssrf", "idor", "authz", "mass-assignment"],
    sources: ["project-code", "observed-traffic"],
  }),

  // ---------------------------------------------------------------------------
  // Users, groups, permissions
  // ---------------------------------------------------------------------------
  endpoint({
    id: "siteUsers.collection",
    family: "sharepoint-rest",
    module: "users",
    operation: "Enumerate site users",
    methods: ["GET"],
    urlTemplate: "/_api/web/siteUsers",
    description: "Lists users known to the web. Useful for mapping principals before permission tests.",
    parameters: [
      queryParam(SP_SELECT, "Fields to return.", "Id,Title,LoginName,Email"),
      queryParam(SP_FILTER, "Filter expression.", "IsSiteAdmin eq true"),
      queryParam(SP_TOP, "Page size.", 100),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "idor", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "siteUsers.ensureUser",
    family: "sharepoint-rest",
    module: "users",
    operation: "Ensure user",
    methods: ["POST"],
    urlTemplate: "/_api/web/ensureuser",
    description: "Resolves a login name to a site user, creating the user info record if needed.",
    parameters: [bodyParam("logonName", "Login name or email.", "string")],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Ensure user payload.",
      { logonName: "i:0#.f|membership|user@contoso.onmicrosoft.com" },
      { requiredProperties: ["logonName"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["input-validation", "xss", "business-logic", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "siteGroups.collection",
    family: "sharepoint-rest",
    module: "groups",
    operation: "Enumerate site groups",
    methods: ["GET", "POST"],
    urlTemplate: "/_api/web/siteGroups",
    description: "Lists groups or creates one (SP.Group body).",
    parameters: [
      queryParam(SP_SELECT, "Fields to return.", "Id,Title,Description"),
      bodyParam("group", "SP.Group creation payload.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Group creation payload.",
      { __metadata: { type: "SP.Group" }, Title: "Research group", Description: "" },
      { odataType: "SP.Group", requiredProperties: ["Title"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["mass-assignment", "xss", "input-validation", "business-logic"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "siteGroups.users",
    family: "sharepoint-rest",
    module: "groups",
    operation: "List / add group members",
    methods: ["GET", "POST"],
    urlTemplate: "/_api/web/siteGroups({groupId:int})/users",
    description: "Lists members or adds a user ({ LoginName }).",
    parameters: [
      pathParam("groupId", "int", "Site group id.", 5),
      bodyParam("user", "User creation payload.", "json"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Add member payload.",
      { __metadata: { type: "SP.User" }, LoginName: "i:0#.f|membership|user@contoso.onmicrosoft.com" },
      { odataType: "SP.User", requiredProperties: ["LoginName"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["idor", "mass-assignment", "input-validation", "business-logic", "authz"],
    sources: ["msdocs", "pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "roleAssignments.add",
    family: "sharepoint-rest",
    module: "permissions",
    operation: "Add role assignment",
    methods: ["POST"],
    urlTemplate:
      "/_api/web/roleAssignments/addroleassignment(principalid={principalId:int},roledefid={roleDefId:int})",
    description:
      "Grants a role definition to a principal. principalid is a site user/group id; roledefid is a permission level id.",
    parameters: [
      pathParam("principalId", "int", "Principal id.", 1),
      pathParam("roleDefId", "int", "Role definition id.", 1073741826),
    ],
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["idor", "authz", "business-logic", "method-confusion"],
    sources: ["msdocs", "pnpjs-source", "project-code"],
  }),
  endpoint({
    id: "roleDefinitions.collection",
    family: "sharepoint-rest",
    module: "permissions",
    operation: "Enumerate role definitions",
    methods: ["GET"],
    urlTemplate: "/_api/web/roleDefinitions",
    description: "Lists permission levels and their base permission masks.",
    parameters: [queryParam(SP_SELECT, "Fields to return.", "Id,Name,BasePermissions")],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "web.shareObject",
    family: "sharepoint-rest",
    module: "sharing",
    operation: "Share web object",
    methods: ["POST"],
    urlTemplate: "/_api/SP.Web.ShareObject",
    description:
      "Shares a URL with principals or a link. Payload uses shareObject parameters (url, peoplePickerInput, role, propagateAcl, sendEmail, includeAnonymousLinkInEmail).",
    parameters: [
      bodyParam("url", "Absolute URL to share.", "string"),
      bodyParam("peoplePickerInput", "Serialised people picker JSON.", "string"),
      bodyParam("role", "Role value.", "int"),
      bodyParam("group", "Group id.", "int"),
    ],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "ShareObject payload.",
      {
        url: "https://contoso.sharepoint.com/sites/team/Shared Documents",
        peoplePickerInput: '[{"Key":"user@contoso.onmicrosoft.com"}]',
        role: 1,
        propagateAcl: true,
        sendEmail: false,
        includeAnonymousLinkInEmail: false,
      },
      { requiredProperties: ["url", "role"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["ssrf", "idor", "mass-assignment", "business-logic", "authz", "xss"],
    sources: ["pnpjs-source", "project-code", "observed-traffic"],
  }),

  // ---------------------------------------------------------------------------
  // Search
  // ---------------------------------------------------------------------------
  endpoint({
    id: "search.query",
    family: "sharepoint-rest",
    module: "search",
    operation: "Run search query",
    methods: ["GET", "POST"],
    urlTemplate: "/_api/search/query?querytext={queryText:odataEncoded}",
    description:
      "Search REST endpoint. Query text, refiners, select properties and sorting are all inputs; results are security trimmed.",
    parameters: [
      pathParam("queryText", "string", "KQL query text.", "*"),
      queryParam("RowLimit", "Rows to return.", 50),
      queryParam("SelectProperties", "Managed properties to project.", "Title,Path,Author"),
      queryParam("RefinementFilters", "Refiner filters.", "[]"),
      queryParam("SortList", "Sort definition JSON.", "[]"),
      queryParam("ClientType", "Client type.", "Modern"),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "xss", "sqli", "input-validation", "authz"],
    sources: ["msdocs", "project-code"],
  }),
  endpoint({
    id: "search.postquery",
    family: "sharepoint-rest",
    module: "search",
    operation: "Run search postquery",
    methods: ["POST"],
    urlTemplate: "/_api/search/postquery",
    description: "POST search API used by the modern UI; supports larger query payloads.",
    parameters: [bodyParam("request", "Search request object.", "json")],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Search request.",
      {
        request: {
          Querytext: "*",
          RowLimit: 50,
          SelectProperties: { results: ["Title", "Path"] },
          ClientType: "Modern",
        },
      },
      { requiredProperties: ["request"] }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "xss", "sqli", "input-validation", "authz"],
    sources: ["msdocs", "project-code"],
  }),

  // ---------------------------------------------------------------------------
  // Web / changes / batch
  // ---------------------------------------------------------------------------
  endpoint({
    id: "web.getChanges",
    family: "sharepoint-rest",
    module: "changes",
    operation: "Get web changes",
    methods: ["POST"],
    urlTemplate: "/_api/web/getchanges",
    description: "Returns changes for the web using SP.ChangeQuery.",
    parameters: [bodyParam("query", "SP.ChangeQuery object.", "json")],
    requestBody: jsonBody(
      "application/json;odata=verbose",
      "Change query.",
      {
        __metadata: { type: "SP.ChangeQuery" },
        Add: true,
        Update: true,
        DeleteObject: true,
        Web: true,
        List: true,
        Item: true,
        RowLimit: 100,
      },
      { odataType: "SP.ChangeQuery" }
    ),
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: false,
    conditional: false,
    payloadFamilies: ["input-validation", "business-logic", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),
  endpoint({
    id: "batch.post",
    family: "sharepoint-rest",
    module: "batch",
    operation: "Submit OData batch",
    methods: ["POST"],
    urlTemplate: "/_api/$batch",
    description:
      "Batches multiple operations in one request. Each part has its own method, URL and body; useful for atomicity and authorization edge cases.",
    parameters: [
      bodyParam("batch", "multipart/mixed body with boundary.", "binary"),
    ],
    requestBody: {
      contentType: "multipart/mixed; boundary=batch_research",
      description:
        "Each part starts with --batch_research and Content-Type: application/http; msgtype=request.",
      example:
        "--batch_research\r\nContent-Type: application/http; msgtype=request\r\n\r\nGET /_api/web HTTP/1.1\r\nAccept: application/json;odata=nometadata\r\n\r\n--batch_research--\r\n",
    },
    auth: "cookie+digest",
    requiresDigest: true,
    destructive: true,
    conditional: false,
    payloadFamilies: ["method-confusion", "authz", "mass-assignment", "input-validation", "business-logic"],
    sources: ["msdocs", "project-code"],
  }),
  endpoint({
    id: "web.recycleBin",
    family: "sharepoint-rest",
    module: "recycle-bin",
    operation: "Enumerate recycle bin",
    methods: ["GET"],
    urlTemplate: "/_api/web/RecycleBin",
    description: "Lists deleted items. Deleted content often retains metadata worth reviewing.",
    parameters: [
      queryParam(SP_SELECT, "Fields to return.", "Id,Title,ItemState,DeletedDate"),
      queryParam(SP_TOP, "Page size.", 100),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "idor", "authz"],
    sources: ["msdocs", "pnpjs-source"],
  }),

  // ---------------------------------------------------------------------------
  // SharePoint v2.0 API (different serialisation from _api)
  // ---------------------------------------------------------------------------
  endpoint({
    id: "v2.site.lists",
    family: "sharepoint-v2",
    module: "lists",
    operation: "v2.0 enumerate lists",
    methods: ["GET"],
    urlTemplate: "/_api/v2.0/sites/{siteId}/lists",
    description:
      "SharePoint v2.0 API. Response shape differs from _api (no d/results wrapper). Do not reuse _api assumptions.",
    parameters: [
      pathParam("siteId", "string", "Site id (hostname,siteCollectionId,webId).", "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000"),
      queryParam(SP_TOP, "Page size.", 100),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "odata-injection", "authz"],
    sources: ["project-code", "observed-traffic"],
  }),
  endpoint({
    id: "v2.site.list.items",
    family: "sharepoint-v2",
    module: "items",
    operation: "v2.0 get list item",
    methods: ["GET"],
    urlTemplate: "/_api/v2.0/sites/{siteId}/lists/{listId}/items/{itemId:int}",
    description: "Alternative item representation used by some modern components.",
    parameters: [
      pathParam("siteId", "string", "Site id.", "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000"),
      pathParam("listId", "string", "List id.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Item id.", 1),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "authz", "input-validation"],
    sources: ["project-code", "observed-traffic"],
  }),
  endpoint({
    id: "v2.site.list.items.activities",
    family: "sharepoint-v2",
    module: "items",
    operation: "v2.0 item activities",
    methods: ["GET"],
    urlTemplate: "/_api/v2.0/sites/{siteId}/lists/{listId}/items/{itemId:int}/activities",
    description: "Activity feed for an item. Supports $expand=listItem($select=...).",
    parameters: [
      pathParam("siteId", "string", "Site id.", "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000"),
      pathParam("listId", "string", "List id.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Item id.", 1),
      queryParam(SP_EXPAND, "Expansion expression.", "listItem($select=id,name,webUrl)"),
      queryParam(SP_TOP, "Page size.", 25),
    ],
    auth: "cookie",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "odata-injection", "authz"],
    sources: ["project-code", "observed-traffic"],
  }),

  // ---------------------------------------------------------------------------
  // Microsoft Graph (different auth, URL and envelope conventions)
  // ---------------------------------------------------------------------------
  endpoint({
    id: "graph.sites.get",
    family: "microsoft-graph",
    module: "sites",
    operation: "Get site by hostname and path",
    methods: ["GET"],
    urlTemplate: GRAPH_BASE + "/sites/{hostname}:/{sitePath}",
    description: "Graph site lookup. Requires a Graph access token with Sites.Read.All or equivalent.",
    parameters: [
      pathParam("hostname", "string", "Tenant hostname.", "contoso.sharepoint.com"),
      pathParam("sitePath", "path", "Site server-relative path.", "/sites/team"),
    ],
    auth: "bearer-graph",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["idor", "input-validation"],
    sources: ["msdocs"],
  }),
  endpoint({
    id: "graph.sites.lists",
    family: "microsoft-graph",
    module: "lists",
    operation: "List site lists (Graph)",
    methods: ["GET"],
    urlTemplate: GRAPH_BASE + "/sites/{siteId}/lists",
    description: "Graph list enumeration. Graph uses bearer tokens, not SharePoint cookies/digests.",
    parameters: [
      pathParam("siteId", "string", "Graph site id.", "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000"),
      queryParam(SP_SELECT, "Fields to return.", "id,name,webUrl"),
      queryParam(SP_EXPAND, "Expansion.", "columns"),
    ],
    auth: "bearer-graph",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "idor", "authz"],
    sources: ["msdocs"],
  }),
  endpoint({
    id: "graph.sites.lists.items",
    family: "microsoft-graph",
    module: "items",
    operation: "List items (Graph)",
    methods: ["GET", "POST"],
    urlTemplate: GRAPH_BASE + "/sites/{siteId}/lists/{listId}/items",
    description:
      "Graph list items. Field values live under the fields facet; item create uses { fields: {...} }.",
    parameters: [
      pathParam("siteId", "string", "Graph site id.", "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000"),
      pathParam("listId", "string", "List id.", "00000000-0000-0000-0000-000000000000"),
      queryParam(SP_EXPAND, "Expansion expression.", "fields"),
      queryParam(SP_TOP, "Page size.", 100),
      bodyParam("fields", "Field values.", "json"),
    ],
    requestBody: jsonBody(
      "application/json",
      "Graph item payload.",
      { fields: { Title: "Controlled probe" } },
      { requiredProperties: ["fields"], protectedProperties: ["createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy"] }
    ),
    auth: "bearer-graph",
    requiresDigest: false,
    destructive: true,
    conditional: false,
    payloadFamilies: ["idor", "mass-assignment", "odata-injection", "xss", "input-validation", "authz"],
    sources: ["msdocs"],
  }),
  endpoint({
    id: "graph.sites.lists.items.get",
    family: "microsoft-graph",
    module: "items",
    operation: "Get list item (Graph)",
    methods: ["PATCH", "DELETE"],
    urlTemplate: GRAPH_BASE + "/sites/{siteId}/lists/{listId}/items/{itemId}",
    description: "Graph update/delete. Update uses { fields: {...} }; ETag via If-Match is supported.",
    parameters: [
      pathParam("siteId", "string", "Graph site id.", "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000"),
      pathParam("listId", "string", "List id.", "00000000-0000-0000-0000-000000000000"),
      pathParam("itemId", "int", "Item id.", 1),
      bodyParam("fields", "Field values.", "json"),
    ],
    requestBody: jsonBody(
      "application/json",
      "Graph update payload.",
      { fields: { Title: "Updated title" } },
      { requiredProperties: ["fields"] }
    ),
    auth: "bearer-graph",
    requiresDigest: false,
    destructive: true,
    conditional: true,
    payloadFamilies: ["idor", "mass-assignment", "method-confusion", "authz", "business-logic"],
    sources: ["msdocs"],
  }),
  endpoint({
    id: "graph.search.query",
    family: "microsoft-graph",
    module: "search",
    operation: "Graph search query",
    methods: ["POST"],
    urlTemplate: GRAPH_BASE + "/search/query",
    description: "Cross-workload search. Graph search entities differ from SharePoint search managed properties.",
    parameters: [bodyParam("requests", "Search request array.", "json")],
    requestBody: jsonBody(
      "application/json",
      "Graph search payload.",
      {
        requests: [
          {
            entityTypes: ["listItem"],
            query: { queryString: "*" },
            from: 0,
            size: 25,
          },
        ],
      },
      { requiredProperties: ["requests"] }
    ),
    auth: "bearer-graph",
    requiresDigest: false,
    destructive: false,
    conditional: false,
    payloadFamilies: ["odata-injection", "xss", "sqli", "input-validation", "authz"],
    sources: ["msdocs"],
  }),
];

const ENDPOINTS_BY_ID: Record<string, EndpointDefinition> = (function buildIndex() {
  const index: Record<string, EndpointDefinition> = {};
  for (let i = 0; i < ENDPOINTS.length; i++) {
    const item = ENDPOINTS[i];
    if (index[item.id]) {
      throw new Error("Duplicate endpoint id in registry: " + item.id);
    }
    index[item.id] = item;
  }
  return index;
})();

export function getEndpoint(id: string): EndpointDefinition {
  const found = ENDPOINTS_BY_ID[id];
  if (!found) {
    throw new Error(
      "Unknown endpoint '" + id + "'. Known ids: " + Object.keys(ENDPOINTS_BY_ID).join(", ")
    );
  }
  return found;
}

export function hasEndpoint(id: string): boolean {
  return Object.prototype.hasOwnProperty.call(ENDPOINTS_BY_ID, id);
}

export function listEndpoints(): EndpointDefinition[] {
  return ENDPOINTS.slice(0);
}

export function listModules(): string[] {
  const modules: string[] = [];
  for (let i = 0; i < ENDPOINTS.length; i++) {
    if (modules.indexOf(ENDPOINTS[i].module) < 0) {
      modules.push(ENDPOINTS[i].module);
    }
  }
  return modules.sort();
}

export function listEndpointIdsByFamily(family: string): string[] {
  const ids: string[] = [];
  for (let i = 0; i < ENDPOINTS.length; i++) {
    if (ENDPOINTS[i].family === family) {
      ids.push(ENDPOINTS[i].id);
    }
  }
  return ids;
}

export const METHODS_OF_INTEREST: HttpMethod[] = ["GET", "POST", "PATCH", "PUT", "DELETE", "HEAD", "OPTIONS", "MERGE"];