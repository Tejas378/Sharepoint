# 06 - Test evidence

This file records what was actually executed, when, and with what result. It is
deliberately narrow: offline construction tests and static analysis pass, but
**no request was sent to a SharePoint tenant**. Claims that require a live
session are marked as not tested.

Last run: **2026-10-30** (local workstation, Windows, PowerShell).

## 1. Environment

| Item | Value |
| --- | --- |
| Working directory | `C:\Users\Tejas\Documents\Sharepoint bug bounty` |
| Node.js | v24.16.0 |
| TypeScript | 4.7.4 (`@microsoft/rush-stack-compiler-4.7`) |
| SPFx | 1.20.0 |
| `@pnp/sp` | 4.9.0 |
| Git base commit | `9586511` (`First commit`) |
| Tenant configured in `src/config/test-config.json` | `https://tmsobb2.sharepoint.com` (not contacted) |

`package.json` declares `"node": ">=18.17.1 <19.0.0"`, so the current Node
version is outside the supported range for the SPFx toolchain. This does not
affect `tsc` or the dependency-free research harness, but it does mean
`gulp build` was **not** run or validated here.

## 2. Commands and results

All commands were run from the repository root. Output is summarised; the exit
codes are the evidence.

### 2.1 Central configuration validation

```powershell
npm run config:validate
```

```text
Configuration: C:\Users\Tejas\Documents\Sharepoint bug bounty\src\config\test-config.json
OK: valid, no issues.
```

Exit code: **0**.

This validates the committed config against the loader rules (GUID/UPN/URL
shapes, unknown sections, secret-looking keys, safety switches, allow-list).

### 2.2 Research type check and unit suite

```powershell
npm run research:test
```

`research:test` first runs `tsc -p tsconfig.research.json`, then executes the
plain-Node harness.

```text
274 passed, 0 failed, 274 total
```

Exit code: **0**. The suite is split across ten files:

| File | Area |
| --- | --- |
| `tests/research/config.test.ts` | Shipped config, structural validation, identifiers, safety guards, overrides, primitives |
| `tests/research/url.test.ts` | URL joining, OData literals, query serialisation, template resolution |
| `tests/research/request.test.ts` | URL/query/header/body construction, method override, probes |
| `tests/research/registry.test.ts` | Endpoint uniqueness, evidence sources, placeholder consistency, Lists coverage |
| `tests/research/payloads.test.ts` | Payload family metadata and generation |
| `tests/research/variations.test.ts` | Baseline, method/query/body/header/identifier probes |
| `tests/research/executor.test.ts` | Response classification, retries, network failures, digest retry, evidence capture |
| `tests/research/digest.test.ts` | `contextinfo` parsing, caching, expiry, rejection detection |
| `tests/research/redact.test.ts` | Header/URL/JSON/text redaction and preview truncation |
| `tests/research/secrets.test.ts` | Runtime secret resolution and precedence |

The harness has no Jest, no network calls and no tenant dependency. It runs the
same TypeScript that the web part uses for the research modules.

### 2.3 Web-part type check

```powershell
npx tsc --noEmit
```

Exit code: **0**. No diagnostics.

This checks the existing SPFx web parts together with the new config facade,
`constants.ts`, and the research modules under the main `tsconfig.json`.

### 2.4 Coverage report

```powershell
npm run research:coverage
```

Report timestamp: `2026-10-30T05:08:10.741Z`.

| Metric | Value |
| --- | --- |
| Registered endpoints | 87 |
| Implemented | 87 |
| Baseline generated | 87 |
| Executed | **0** |
| Verified | **0** |

Modules (19): attachments, batch, changes, comments, content-types, context,
fields, files, folders, groups, items, lists, permissions, recycle-bin, search,
sharing, sites, users, views.

Declared methods: GET (45 endpoints), POST (43), DELETE (9), PATCH (7), MERGE
(2). HEAD, OPTIONS and PUT exist only as generated variations.

The report also lists 39 destructive endpoints blocked by
`safety.allowDestructive=false` and four uncovered areas: taxonomy, workflow,
tenant-level webhooks, and Microsoft Graph users/groups.

### 2.5 First-step request rendering

```powershell
npm run research:request -- items.getById
```

Produces a complete raw HTTP request for inspection or Burp, for example:

```http
GET /_api/web/lists(guid'4a81ba4f-84ed-4e91-a731-9a82c0ede37e')/items(2) HTTP/1.1
Host: tmsobb2.sharepoint.com
Accept: application/json;odata=nometadata
X-RequestDigest: {{REQUEST_DIGEST}}
```

The tool prints the request and **does not send it**. `{{REQUEST_DIGEST}}` is a
placeholder, not a captured value.

## 3. What the passing tests prove

- The committed config is internally consistent and the safety switches are
  enforced by the loader and guards.
- Every registry entry produces a URL from its declared parameters, with no
  tenant host hardcoded and no undeclared placeholder.
- Request construction handles OData quoting, encoding, method override,
  digest placeholders and conditional headers as designed.
- Response classification and retry policy behave as specified for the tested
  status codes.
- Digest parsing, caching and rejection detection work against fixture
  envelopes.
- Redaction removes credential-shaped headers, query values, JSON keys and
  free text.
- The runtime secret resolver throws on a missing secret and follows the
  documented precedence.

## 4. What the passing tests do **not** prove

- That any SharePoint endpoint exists, accepts the generated shape, or returns
  the expected status. Construction is not execution.
- That the configured GUIDs, item ids, users or groups exist in the tenant.
- That any authorization boundary holds, or that any injection payload is
  handled safely. All such items remain hypotheses in
  `docs/05-vulnerability-report.md`.
- That the legacy UI path is safe. The tests exercise `research/*` and the new
  facade, not the class-local XHR methods in `Detaservice.ts` /
  `DataService2.ts`.
- That the SPFx bundle builds with gulp. `gulp build` / `gulp bundle` were not
  run because the Node version is outside the supported range.
- That rate limiting, throttling or `Retry-After` behaviour matches the
  executor model. No live 429 was observed.

## 5. Evidence log template

Use this for every live request that is meant to support a finding. Store it
outside the repository if it contains tenant data, and redact before sharing.

```text
ID:            EV-0001
UTC timestamp: 2026-10-30T00:00:00Z
Endpoint:      items.getById
Method:        GET
URL:           /_api/web/lists(guid'...')/items(2)
Account:       <test user, no real identity>
Site:          <test site>
Control:       <same request as an authorized user / expected-deny case>
Status:        200
ETag:          <if present>
Duration:      <ms>
Response:      <shape or redacted body>
Conclusion:    confirmed | rejected | inconclusive
Hypothesis:    H-01
Notes:         <what changed, what was ruled out>
```

Promotion rules:

1. A hypothesis becomes **confirmed** only with a control request and either a
   reproducible response or a reproducible denial.
2. A denial is evidence only if the same request succeeds under an account that
   should have access.
3. Anything that touches real user data, creates persistence, or affects another
   user stops the run; see `docs/01-safe-testing-boundaries.md`, section 8.

## 6. How to reproduce

```powershell
npm run config:validate
npm run research:test
npx tsc --noEmit
npm run research:coverage
npm run research:request -- lists.collection.get
```

Expected: three exit codes of 0, `274 passed`, and `executed=0`, `verified=0`
until a live session is deliberately added. If a count changes, update this file
and `PROJECT_CONTEXT.md` in the same change.

## 7. Current limitations

| Limitation | Consequence |
| --- | --- |
| No authenticated session | No live request, no finding, `executed=0`. |
| Node outside `engines` range | SPFx gulp pipeline unverified. |
| Legacy XHR not migrated | UI path still has V-02 to V-07 from the vulnerability report. |
| `doclib` has a separate config | Identifier drift remains possible. |
| Graph needs a separate token | Graph entries are implemented but untested and unreachable via the SharePoint cookie transport. |
| No evidence store | `executed`/`verified` are hardcoded to 0 rather than computed. |
