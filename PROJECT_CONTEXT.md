# PROJECT_CONTEXT

Last updated: 2026-10-30

This file is the handoff document for the SharePoint API research project. It
records what exists, what was actually verified, and what is still unproven.
Read it before changing anything.

## 1. Purpose

Turn an existing SPFx proof-of-concept into a reliable SharePoint API
exploration and manual security-testing toolkit:

- generate valid, inspectable requests for SharePoint REST, SharePoint v2 and
  Microsoft Graph;
- vary methods, parameters, bodies and headers with an explicit rationale;
- export requests to Burp Suite;
- keep identifiers in one place;
- never claim a vulnerability or a test result that was not actually observed.

The toolkit assists manual research. It does not automatically declare
vulnerabilities.

## 2. Architecture

```
src/
  config/
    test-config.json          Single source for identifiers and static test data
    test-config.schema.json   JSON Schema mirror of the loader rules
    secrets.ts                Runtime-only secret resolution (no committed values)
    README.md                 How to change identifiers and supply secrets
  webparts/
    hacktheworld/
      config/loader.ts        Typed loader, validation, safety guards
      research/               SPFx-free request engine (unit tested in Node)
        registry.ts           87 endpoint definitions, 19 modules
        request.ts            URL/header/body construction
        url.ts                OData/URL encoding and template resolution
        variations.ts         Baseline + method/query/body/header/ID probes
        payloads.ts           11 payload families with controls
        digest.ts             Digest acquisition, caching, rejection handling
        executor.ts           Bounded execution, classification, retry policy
        redact.ts             Header/URL/JSON/text redaction
        types.ts              Shared contracts and coverage types
      components/
        config.ts             Legacy CONFIG facade + corrected sendRequest helper
        constants.ts          Legacy constants, now derived from the central config
        Detaservice.ts        Legacy PnPjs + XHR data service (not fully migrated)
        DataService2.ts       Older duplicate XHR data service
        Hacktheworld.tsx      UI (1,700 lines)
        Hacktheworldcopy.tsx  Near-duplicate UI (not wired as the active UI)
    doclib/                   Near-duplicate web part; separately built
    helloWorld/               Unrelated starter web part
tests/research/               274 Node tests, no Jest, no network
tools/
  validate-config.ts          `npm run config:validate`
  coverage-report.ts          `npm run research:coverage`
  print-request.ts            `npm run research:request -- <endpointId>`
docs/                         First-step workflow, boundaries, analysis, report
```

### Runtime flow

1. `src/config/test-config.json` is imported by `config/loader.ts`.
2. `loadTestConfig()` deep-merges optional overrides, validates the whole object
   and returns a normalised `TestConfig`.
3. `components/config.ts` maps the typed config onto the legacy `CONFIG` object
   so `Detaservice.ts` keeps working.
4. `research/*` uses the registry to build requests; it never imports SPFx or
   PnPjs, so it runs in plain Node and in the browser.
5. `tools/print-request.ts` renders a raw HTTP request for Burp. It does not send.
6. `research/executor.ts` is the only component that sends, and only through a
   caller-supplied `Transport`.

## 3. Versions

| Component | Installed version | Evidence |
| --- | --- | --- |
| SharePoint Framework | 1.20.0 | `node_modules/@microsoft/sp-core-library/package.json` |
| PnPjs `@pnp/sp` | 4.9.0 | `node_modules/@pnp/sp/package.json` |
| TypeScript | 4.7.4 (rush-stack) | `npx tsc --version` |
| Node used for research | v24.16.0 | `node --version` |
| Node required by package.json | `>=18.17.1 <19.0.0` | `package.json` |

The Node version mismatch means `gulp build` is not validated in this
environment. The research and web-part type checks both pass with `tsc`.

## 4. API coverage

Registry totals (from `npm run research:coverage`):

| Metric | Value |
| --- | --- |
| Endpoints registered | 87 |
| Modules | 19 |
| Lists endpoints | 29 |
| Items endpoints | 16 |
| Baselines generated from examples | 87 |
| Executed against a live tenant | 0 |
| Verified manually | 0 |

Modules: attachments, batch, changes, comments, content-types, context, fields,
files, folders, groups, items, lists, permissions, recycle-bin, search, sharing,
sites, users, views.

Declared methods: GET (45 endpoints), POST (43), DELETE (9), PATCH (7), MERGE
(2). HEAD, OPTIONS and PUT are generated as variations, not declared methods.

See `docs/04-endpoint-inventory.md` for the full list and
`docs/03-pnpjs-and-endpoint-analysis.md` for how each entry was sourced.

### Status definitions

- **Implemented** - an `EndpointDefinition` exists.
- **Generated** - `buildRequest()` produced a complete URL from documented
  examples. This proves construction, not server behaviour.
- **Executed** - a request was actually sent. **Zero in this repository.**
- **Verified** - manual evidence confirmed the behaviour. **Zero in this
  repository.**

Do not promote an endpoint to executed/verified without a recorded request,
response and timestamp.

## 5. Configuration and secrets

`src/config/test-config.json` is the only place where list/item/user/group/
principal/role/view/content-type identifiers and static test data live. Nothing
in `research/`, tests or the new `constants.ts` hardcodes them.

```powershell
npm run config:validate
```

Validation rejects unknown top-level sections, malformed GUIDs/UPNs/URLs,
missing safety switches, a site host outside `safety.allowedHosts`, and any key
that looks like a credential (`password`, `token`, `cookie`, ...).

Secrets are **not** in the JSON and **not** in source. `src/config/secrets.ts`
resolves them at runtime, in order:

1. `setSecret("TEST_PASSWORD", value)` from code/tests;
2. `sessionStorage["sp-research.secret.TEST_PASSWORD"]`;
3. environment variable `SP_RESEARCH_TEST_PASSWORD` (Node tooling).

A missing secret throws an explanatory error. There is no default.

The three legacy password uses in the hacktheworld data service and the five
literal occurrences in the doclib service were replaced with `getSecret(...)`.

## 6. Request reliability

`components/config.ts` exports a corrected `sendRequest`:

- rejects on 4xx/5xx instead of resolving;
- JSON-encodes bodies without the char-code/Blob corruption;
- acquires a digest from `/_api/contextinfo`, caches it and invalidates on 403;
- sends `If-Match` only for write verbs;
- applies `safety.requestTimeoutMs`;
- restricts targets with `assertAllowedUrl`.

**Known gap:** `Detaservice.ts` still defines its own class-local `sendRequest`
(line 1877) and `sendRequest2` (line 1939), and `DataService2.ts` has more copies
(line 917/958 and others). These retain the historical defects (any non-zero
status resolves, `localStorage["reqd"] ?? "*"`, Blob char-code bodies,
`If-Match: *` on reads, no timeout, 61 hardcoded tenant URLs). Migrating those 26
call sites is a deliberate follow-up because changing rejection behaviour can
turn fire-and-forget probes into unhandled promise rejections. The corrected
helper is available for new code and for the research executor.

## 7. Tests actually run

```powershell
npx tsc -p tsconfig.research.json   # exit 0
node .tmp/research/tests/research/run.js
# 274 passed, 0 failed, 274 total

npx tsc --noEmit                   # exit 0 (web part)
npm run config:validate            # OK: valid, no issues
npm run research:coverage          # 87 endpoints, 0 executed, 0 verified
```

The suite covers config validation and guards, URL/OData encoding, request
construction, registry integrity, payload generation, variation generation,
executor retry/classification, digest parsing/caching, redaction and the runtime
secret resolver. It sends no traffic.

Details and limitations: `docs/06-test-evidence.md`.

## 8. Known defects (static, confirmed by reading the code)

1. **Error handling** - legacy `xhr.onload` resolves whenever `xhr.status` is
   truthy, including 401/403/404/500. Callers parse error pages as success.
2. **Unicode corruption** - `body.charCodeAt(i)` into `Uint8Array` truncates
   every code point above U+00FF, so non-ASCII titles are altered.
3. **Digest** - reads `localStorage["reqd"]`, falls back to `"*"`, never calls
   `/_api/contextinfo`. First write in a fresh session fails.
4. **Conditional headers** - `If-Match: *` is attached even to GET requests.
5. **Hardcoded tenant** - 61 absolute `https://tmsobb2.sharepoint.com/...` URLs
   in the hacktheworld service; changing tenant requires editing code.
6. **Duplication** - three large data services (`Detaservice.ts`,
   `DataService2.ts`, and the doclib copies) implement overlapping behaviour.
7. **Error swallowing** - `new Promise` bodies wrapped in `try/catch` that only
   `console.log`, leaving callers waiting or receiving `undefined`.
8. **No timeout** - a stalled request hangs forever.
9. **Secrets in source** - removed in this change; see section 5.

These are code-quality/security defects in the testing tool, not findings
against SharePoint. The vulnerability report keeps that distinction.

## 9. Safety

- `safety.allowDestructive` defaults to `false`.
- `safety.allowPermissionChanges` defaults to `false`.
- `safety.allowOutboundWebhooks` defaults to `false`.
- `safety.allowedHosts` contains only the configured tenant.
- `maxRequestsPerRun`, `maxPayloadBytes`, `requestTimeoutMs` and
  `minimumDelayBetweenRequestsMs` are enforced by the executor.
- Request exports redact cookies, authorization headers and digests.

Read `docs/01-safe-testing-boundaries.md` before enabling anything.

## 10. Decisions and reasons

| Decision | Reason |
| --- | --- |
| Config at `src/config/` | `tsconfig` sets `rootDir: src`; a root `config/` would fall outside it. |
| Keep `CONFIG` legacy shape | Preserves the existing data service without a rewrite. |
| SPFx-free `research/` | Lets requests be unit tested in plain Node with no tenant. |
| Dependency-free test harness | No Jest install; tests run with `tsc` + `node`. |
| Unknown config sections are errors | A misspelt key must not silently keep a default. |
| Valid-shaped foreign GUID probes | Malformed GUIDs made variation generation throw; IDOR still works with syntactically valid foreign IDs. |
| `executed`/`verified` hardcoded to 0 | No tenant access; claiming otherwise would be false. |
| Legacy XHR not rewired yet | Avoids unhandled rejections in 26 fire-and-forget probes; documented instead. |

## 11. Task status

- [x] Repository assessment and architecture map.
- [x] Installed PnPjs 4.9.0 inspected for Lists.
- [x] Central config, schema, loader, safety guards.
- [x] Lists coverage expanded to 29 endpoints (16 added).
- [x] Request engine: URL, body, headers, digest, ETags, redaction.
- [x] Payload and variation generators with classifications.
- [x] Runtime secret resolution; plaintext passwords removed.
- [x] `constants.ts` derived from central config.
- [x] Config validation, coverage report and raw-request CLI.
- [x] 274 unit tests passing; web-part and research type checks clean.
- [x] Documentation: workflow, boundaries, analysis, inventory, report, evidence.
- [~] Legacy `Detaservice.ts`/`DataService2.ts` migration to the corrected helper.
- [~] Doclib web part still has its own `config.js` and hardcoded URLs.
- [ ] Live-tenant execution and manual verification (blocked: no authenticated session).
- [ ] Taxonomy, workflow and tenant-level webhook coverage.
- [ ] Microsoft Graph authentication/transport separate from SharePoint cookies.
- [ ] Automated rate-limit/backoff evidence from a live tenant.

## 12. Next steps

1. Supply an authenticated browser session for the authorized tenant.
2. Run one GET baseline per module, starting with `lists.collection.get`.
3. Record status, ETag, body shape and duration in an evidence log.
4. Only then enable `allowDestructive` on a throwaway site and test writes.
5. Migrate the legacy XHR call sites to `components/config.ts#sendRequest`.
6. Add taxonomy/workflow modules only with documented or observed endpoints.
7. Promote `executed`/`verified` counts in the coverage report from real evidence.

## 13. How to continue without repeating the investigation

- Read `docs/00-first-step-workflow.md` for the operating sequence.
- Read `docs/01-safe-testing-boundaries.md` before any write test.
- Change identifiers only in `src/config/test-config.json`.
- Add endpoints only with a source (`msdocs`, `pnpjs-source`, `pnpjs-typedef`,
  `observed-traffic`); mark guesses `unverified`.
- Run `npm run config:validate && npm run research:test` after every change.
- Never put credentials, cookies or digests in this file or in the repository.
