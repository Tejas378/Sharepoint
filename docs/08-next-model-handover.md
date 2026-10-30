# SharePoint API coverage handover

Updated: 2026-10-30

## Goal and current state

This project is an authorized SharePoint API security test harness. New registry
operations are exposed through the Research API runner UI; tenant calls must be
performed only against a tenant the tester is authorized to assess.

- Registry: 316 endpoint definitions across 25 modules.
- Coverage generator: 316 implemented; 188 baseline requests can currently be
  generated from examples/configuration.
- Evidence: executed=0, verified=0. No tenant calls were made in this pass.
- Installed PnPjs package: `@pnp/sp` 4.9.0 (package range `^4.9.0`).
- Most recent config validation: 0 errors, 21 warnings. Warnings identify tenant
  values the tester still needs to provide; see `src/config/test-config.json`.

The 316 count describes the current project's endpoint catalog. It is not a
claim that every SharePoint REST, Graph, CSOM, or tenant-specific API exists in
the catalog. Keep auditing official API references, installed PnPjs methods,
and project call sites.

## Authorization/IDOR changes in this pass

`ResearchApiPanel.tsx` now presents configured values for ID-like path/query
parameters and adds **Sweep configured IDs (read only)**. For GET operations it
uses configured list IDs, item IDs (primary, secondary, nonexistent, boundary),
principal IDs, user IDs, group IDs, and role-definition IDs. It combines all
configured ID dimensions present on a route (for example list × item ×
principal) so object checks can exercise those combinations together. This is a sequential read-only sweep and
obeys `safety.maxRequestsPerRun`, `minimumDelayBetweenRequestsMs`, and the
allowed-host check. It refuses destructive operations and non-GET methods.
Results show the selected configured value and HTTP result.

Fill in authorized resource/principal IDs in `src/config/test-config.json`
before sweeping. Empty IDs cannot be inferred safely. Role definition values
are tenant-specific inputs; verify those configured values against the target
tenant. A single signed-in identity tests object-level access decisions for
that identity only. To compare roles, repeat the same configured sweep under
each authorized test account/session; this UI does not switch authentication
identities itself.

## Legacy overlap and duplicate calls

The old components (`src/webparts/hacktheworld/components/Detaservice.ts`,
`DataService2.ts`, `Hacktheworld.tsx`, and `src/webparts/doclib`) contain API
calls that also appear in the new registry. Examples include list-item parent
list/unique-ancestor lookups, sharing APIs (`ShareLink`,
`GetSharingInformation`, `UpdateDocumentSharingInfo`, `UnshareLink`),
`renderListDataAsStream`, and folder/file operations. Some legacy code invokes
the same route repeatedly with different methods or parameter permutations.

An overlap with legacy source does not cause a duplicate runtime call: the new
runner executes only the selected registry request, while legacy controls are a
separate UI path. Do not remove a registry definition solely because a legacy
call exists; that would remove its new-runner coverage. Continue deduplicating
exact method+URL definitions inside the registry. `$batch` has distinct
read-only and change-set request profiles on the same route by design, so keep
their safety/body distinction explicit.

## Current endpoint groups

The generated inventory is `docs/04-endpoint-inventory.md`; source/API audit is
`docs/03-pnpjs-and-endpoint-analysis.md`. Current counts include:

| Area | Endpoints |
| --- | ---: |
| Files / folders / libraries / lists / items / attachments | 117 |
| Fields / content types / views | 58 |
| Users / groups / permissions / sharing | 72 |
| Sites / webs / configuration | 28 |
| Search / changes / batch | 10 |

Counts above are a practical grouping of registry modules, not a statement that
all possible permutations, methods, or API versions are covered.

## Remaining work for the next model

1. Inspect the current diff and registry for exact duplicate method+URL routes;
   preserve intentionally different request-body/safety profiles.
2. Build a machine-readable legacy-to-registry overlap report, including the
   old `doclib` directory, and distinguish same route from duplicate execution.
3. Continue endpoint discovery per module. In particular, verify route/method
   variants and request/response shapes for role assignments, role definitions,
   sharing, file versions/metadata, list and web configuration, search options,
   change queries/tokens, and `$batch` against installed PnPjs and Microsoft
   documentation. Do not claim “all endpoints” until this audit is complete.
4. Review the ID sweep parameter-name mapping against every permission/role
   endpoint. Add tenant-configurable alternate site/web IDs and file/folder
   identifiers if useful for object-level authorization checks. Confirm
   combinations stay within the configured request cap.
5. Fill tenant config warnings using authorized test resources. Change tokens
   must be obtained from the target tenant. Keep disposable IDs/data for write
   tests separate from production resources.
6. Run authorized, controlled tenant validation for baseline GETs first, then
   explicitly approved write cases. Record evidence and expected outcomes;
   inspect `docs/evidence` and update executed/verified reporting.
7. Add two-account authorization comparison guidance or a safe way to select
   distinct signed-in sessions; current UI uses the active SharePoint session.
8. Validate any newly discovered routes against supported SharePoint REST
   behavior; a generated URL is not proof that a route exists or is supported.

## Useful commands

```powershell
npm run config:validate
npm run research:coverage
npm run research:evidence
npx tsc --noEmit
```

These commands compile/report local data; they do not authenticate to or send
requests to the configured SharePoint tenant. The UI button sends a request only
when a user explicitly executes it.
