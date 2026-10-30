# 02 - Existing code analysis

Scope: the repository as found at commit `9586511` plus the changes recorded in
`PROJECT_CONTEXT.md`. All line numbers refer to the current working tree.

## 1. Repository shape

- SPFx 1.20.0 solution with three web parts: `hacktheworld`, `doclib` and
  `helloWorld`.
- TypeScript/React, bundled by gulp.
- PnPjs 4.9.0 is installed and used for some operations; many others use raw
  `XMLHttpRequest` against hand-written URLs.
- No backend, no database, no test framework wired into `npm test` (that script
  calls `gulp test`, which is the SPFx default and was not run here).
- The active research UI is `hacktheworld/components/Hacktheworld.tsx`.

## 2. Entry points

| Entry point | Role |
| --- | --- |
| `HacktheworldWebPart.ts` | SPFx web part bootstrap, renders `Hacktheworld.tsx`. |
| `Hacktheworld.tsx` (1,700 lines) | UI; one button per operation; constructs `DataService`. |
| `Detaservice.ts` (3,236 lines) | Main data service: PnPjs calls plus raw XHR probes. |
| `DataService2.ts` (1,470 lines) | Older/duplicate raw XHR service. |
| `components/config.ts` | Legacy `CONFIG` facade and corrected `sendRequest`. |
| `components/constants.ts` | Legacy constants, now derived from central config. |
| `doclib/*` | Near-duplicate web part with its own `config.js` and service. |
| `helloWorld/*` | Unrelated starter web part. |

## 3. Duplication

| Pair | Observation |
| --- | --- |
| `hacktheworld/components/Detaservice.ts` vs `doclib/components/Detaservice.ts` | 3,236 vs 3,100 lines; overlapping methods and identical sharing/XHR logic. Not byte-identical. |
| `hacktheworld/components/DataService2.ts` vs `doclib/components/DataService2.ts` | 1,470 vs 2,179 lines; divergent copies of the same idea. |
| `Hacktheworld.tsx` vs `Hacktheworldcopy.tsx` | 1,700 vs 1,491 lines; near-duplicate UI. |
| `Hacktheworld.tsx` vs `doclib/Doclib.tsx` | 1,700 vs 1,576 lines; near-duplicate UI in a second web part. |
| `IHacktheworldProps.ts` | Byte-identical in both web parts (SHA-256 match). |

Consequence: a fix applied to one service does not reach the others. The central
config and the corrected request helper are the first step toward one path; the
services themselves are still duplicated.

## 4. Configuration before this change

- `hacktheworld/components/config.js` contained a large `CONFIG` literal with
  list GUIDs, item ids, user ids, group ids, role ids and tenant URLs, plus an
  XHR `sendRequest`.
- `doclib/components/config.js` had a second, different literal set.
- `components/constants.ts` had a third set: tenant URL, user emails, content
  type id, view id, change tokens, dates and plaintext passwords.
- Tests and research code did not exist.

Changing tenant or list required editing several files, and the values could
drift apart. This is the root cause the central config addresses.

## 5. Confirmed defects

These are defects in the **testing tool**, confirmed by reading the code. They
are not vulnerabilities in SharePoint.

### 5.1 Non-2xx responses resolved as success

`Detaservice.ts` lines 1916-1922:

```ts
xhr.onload = function () {
  if (xhr.status) {
    resolve(xhr.responseText);
  } else {
    reject(xhr.statusText);
  }
};
```

`xhr.status` is truthy for 401, 403, 404 and 500, so the promise resolves with an
error page. `DataService2.ts` has the same pattern at lines 929, 970 and 1011.
Callers then read `res.d...` from an error body and throw unrelated TypeErrors,
or silently treat a denial as success.

### 5.2 Non-ASCII bodies are corrupted

`Detaservice.ts` lines 1924-1930 (and many inline copies):

```ts
var aBody = new Uint8Array(body.length);
for (var i = 0; i < aBody.length; i++) {
  aBody[i] = body.charCodeAt(i);
}
xhr.send(new Blob([aBody]));
```

`charCodeAt` returns UTF-16 code units; assigning to `Uint8Array` keeps only the
low 8 bits. Any title containing `é`, `你` or an emoji is changed before it is
sent. This makes Unicode/XSS payload results unreliable.

### 5.3 Digest is never acquired

`Detaservice.ts` lines 1907-1910:

```ts
xhr.setRequestHeader("X-RequestDigest", localStorage.getItem("reqd") ?? "*");
```

Nothing in the service writes `reqd`. A fresh session sends `*` and SharePoint
rejects the write with a security-validation error. The same pattern appears in
`DataService2.ts` lines 497, 927, 968, 1009 and in the doclib copies.

### 5.4 `If-Match: *` on reads

`Detaservice.ts` line 1911 sets `If-Match: *` unconditionally, including GET.
This is meaningless on reads and hides whether a caller is actually using ETags
for conditional updates.

### 5.5 Hardcoded tenant

`Detaservice.ts` contains 61 occurrences of
`https://tmsobb2.sharepoint.com`. Examples: lines 108, 171, 233, 238, 304. The
tenant cannot be changed without editing code, and a typo can send a request to
the wrong host. `assertAllowedUrl` now protects the corrected path, but the
legacy literals bypass it.

### 5.6 Errors swallowed inside `new Promise`

Several methods wrap the whole promise body in `try/catch` that only
`console.log`s (for example `sendRequest` at lines 1886-1937). A synchronous
error leaves the returned promise pending forever.

### 5.7 No timeout

`XMLHttpRequest` has no `timeout` and `fetch` had no `AbortController`. A stalled
connection hangs the operation and any UI state waiting on it.

### 5.8 Plaintext passwords

Removed in this change:

- `components/constants.ts` had `TEST_PASSWORD`, `TEJAS_PASSWORD` and
  `OKKK_PASSWORD`.
- `doclib/components/Detaservice.ts` had five literal occurrences of the same
  values inside JSON bodies.

They now resolve through `src/config/secrets.ts` at runtime. Nothing is
committed.

### 5.9 Broken response parsing in call sites

Some call sites await the raw-text `sendRequest` and immediately access
`res.d.permissionsInformation` (for example `getShareLinkIds`, line 1997). Since
the helper resolves a string, that access is undefined. This is a consequence of
5.1/5.2 and is not fixed yet.

## 6. What changed

| Area | Change |
| --- | --- |
| Config | `src/config/test-config.json` + schema + loader + README. |
| Legacy facade | `components/config.ts` derives `CONFIG` from the central config. |
| Constants | `components/constants.ts` derives every value from config. |
| Secrets | `src/config/secrets.ts`; plaintext passwords removed. |
| Request engine | New `research/` modules (registry, request, url, variations, payloads, digest, executor, redact). |
| Tests | 274 offline tests in `tests/research/`. |
| Tools | Config validator, coverage report, raw-request printer. |
| Types | `tsconfig.research.json`; `tsconfig.json` gained `resolveJsonModule` and `allowSyntheticDefaultImports`. |

## 7. What did not change

- The UI is unchanged. It still triggers the legacy service methods.
- The legacy class-local `sendRequest`/`sendRequest2` are unchanged. There are
  26 call sites; switching them to the throwing helper can create unhandled
  rejections in fire-and-forget probes.
- `doclib` still uses its own `config.js` and hardcoded URLs. Only its plaintext
  passwords were removed.
- No dependency was upgraded.
- No SharePoint request was sent.

## 8. Risks if the remaining work is skipped

- The UI can still report success on 4xx/5xx.
- Unicode and XSS payloads sent through the legacy path are altered.
- First writes still fail without a digest.
- Two tenants' identifiers can drift between `hacktheworld` and `doclib`.
- The coverage report will keep showing `executed=0`, which is honest but limits
  the value of the toolkit until a live session is used.

## 9. Recommended follow-up order

1. Add a session/transport layer that uses `components/config.ts#sendRequest` or
   `research/executor.ts`.
2. Migrate one method group at a time (reads first, then writes) and add a
   regression test for each group.
3. Delete `DataService2.ts` only after its still-used methods are accounted for.
4. Point `doclib` at the central config or mark it deprecated.
5. Add an evidence log so `executed`/`verified` can be computed from real data.
