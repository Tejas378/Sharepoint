# Security test coverage review

## What the old implementation covers

The legacy UI and services contain hand-written SharePoint operations and
several role, principal and link-kind loops. Those probes are useful as a
source of parameter ideas. Some loops are fire-and-forget `forEach` calls with
async callbacks, catch errors by logging them, and do not inspect the final
permission state. A request being attempted is not reliable evidence that the
whole permutation completed or that authorization was correct.

## What the research runner covers

The research registry currently declares 476 endpoint definitions: 215
explicit `endpoint(...)` definitions and 261 `catalogEndpoint(...)`
definitions. Each carries a route, method, parameter metadata, auth mode,
write/safety classification, sources and applicable payload families. The UI
lets the operator select one case, preview it, and execute it using the
currently signed-in browser session. The request executor captures status,
response class, timing, retry attempts and a redacted response preview.

The generic payload families are hypotheses, not endpoint-specific
authorization policies. In particular, a response code cannot establish
whether a user was authorized. For writes, record the resulting object or
permission state and verify it with the second account.

## Changes made for two-account UI testing

- Every registry endpoint can now derive cases from its own parameter types,
  declared enum values, required body properties, and protected-property list.
  Input checks include parameter omission, typed boundaries, malformed JSON,
  duplicate and conflicting query keys, and request content-type/Accept variants.
- IDOR and authorization cases use only IDs present in the selected config and
  include bounded pairwise combinations where an endpoint has multiple ID
  parameters. Business-logic enum cases get pairwise combinations; method
  confusion includes controlled HTTP method override probes. The per-family
  endpoint-derived case cap is 200.
- Where the endpoint declares a relevant parameter, the runner also generates
  path-normalization cases, reserved/loopback URL cases without callbacks,
  inert markup probes, non-time-based SQL syntax probes, and OData query cases.
- The runner identifies the signed-in account, target site, config file and
  configured primary/secondary identities. Requests still use the current
  browser session; the displayed config identities are comparison references.
- Share-link creation endpoints now offer a Cartesian product of up to 20
  distinct configured `shareLinkKinds` and 20 `shareLinkRoleKinds` values, with
  the generated plan capped at 200 cases. Numeric codes are
  shown as configured; their semantics are not assumed. Each case is a single
  request and continues to use the existing destructive and permission-change
  safety gates.
- Each executed request is retained in the current page session with the
  account, endpoint, case name and ID overrides, method, URL, status,
  response class and redacted preview. Add an assessment and verification note,
  then download the evidence as JSON. Authentication headers and request
  bodies are excluded.
- The same runner changes are included in both web parts; each shows and uses
  its own config file.

Generated write, method-override, mass-assignment and permission cases are
safety-marked even when the base endpoint is classified as read-only. The
runner requires its existing destructive approval switch, applies
permission-change controls to authorization cases, honors request
budgets/pacing, and asks for per-request confirmation.

## Remaining manual analysis and verification

The endpoint registry is an inventory, not proof that every operation has a
verified authorization contract. The 261 catalog entries use a shared metadata
builder and need endpoint-by-endpoint review for documented parameters,
actual PnP request construction, authorization expectations, and safe
post-operation verification. Values in the legacy config are not necessarily
valid for every SharePoint endpoint or tenant policy.

The generated cases are a bounded test design (up to 300 endpoint-derived
cases per endpoint and family), not an exhaustive Cartesian product and not a
claim that every hypothetical payload has been tested. The
UI does not switch identities: run the same case separately in each browser,
then verify resulting access and state from the other account. Anonymous link
access must be checked in a separate unauthenticated browser context.

For each account, use a known disposable fixture and record its baseline
permissions first. In each browser, run the same request and compare the
resulting permission state from the other account. For generated sharing
links, verify both the returned link metadata and whether the second test
account can perform the expected read or write. Treat external or anonymous
access only as a controlled policy probe on a disposable fixture. Do not infer
that a link is safe from the link code or HTTP status alone.

Microsoft Graph `driveItem/createLink` documents `view`/`edit` types and
`anonymous`/`organization`/`users` scopes. Those string values are specific to
Graph and must not be substituted for the numeric SharePoint REST codes in the
project config without verifying the SharePoint REST contract. Delegated Graph
calls also remain limited by the signed-in user's access as well as the app's
 consented permissions.

## Parameter inventory, hidden fields, and OWASP coverage

Each runner offers **Download all declared API parameters (JSON)**. The export
includes endpoint routes, methods, auth mode, parameter location/type/required
flag, examples and enums, request-body required and protected properties, and
enabled test families. This is the complete project-declared inventory; it
cannot enumerate undocumented server-side parameters. The UI generates bounded
hidden-property hypotheses (such as `isAdmin`, `role`, `permissions`,
`tenantId`, `AuthorId`, and `EditorId`) for endpoints with JSON body examples.
Write probes require existing destructive controls. Acceptance is not proof of
a vulnerability: verify persisted state and effective access from the other
browser.

Cross-tenant IDOR cases require a target under `safety.crossTenantTargets`, its
host in `safety.allowedHosts`, and an explicitly supplied identifier matching
a declared endpoint parameter. Both example configs keep this list empty. Add
only owned or authorized test tenants and fixture IDs. Current cross-tenant
cases cover SharePoint REST only; Graph is omitted because the runner does not
establish a separate foreign-tenant Graph identity/token. Compare returned
object identity and account B's effective access in the second browser.

The UI maps cases to OWASP Top 10:2025. A01 and A05 have endpoint probes for
access control and injection; A02, A06, A08, and A10 have configuration,
workflow, integrity, or malformed-input cases as applicable. A03 supply-chain,
A04 cryptography, A07 logged-out authentication, and A09 audit/alerting require
manual deployment or tenant checks; a payload cannot prove them. See the
[OWASP 2025 Top 10](https://top10.owasp.org/2025/0x00_2025-Introduction/).

The expanded input sets cover integer/GUID/date/boolean/JSON/path boundaries;
null, empty, duplicate, conflicting, malformed, Unicode, encoding, and length
variants; stale ETag and content negotiation checks; SQL boolean/parser and
OData expression probes; controlled URL/path and inert endpoint-derived
rendering probes; declared and unexpected fields; configured user/site/item identifiers; and pairwise
parameter interactions. SQL time-delay, stacked/data-extraction probes and
cloud metadata/internal-range SSRF targets are intentionally excluded. Case
volume is capped so the UI remains usable; this is broad coverage, not literally
every server-side parameter or payload permutation.
