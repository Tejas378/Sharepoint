# Centralised test configuration

`test-config.json` is the single place where SharePoint identifiers and static
test data live. Nothing else in the project should hardcode a list GUID, item
id, user id, group id, principal id or tenant URL.

## Files

| File                      | Purpose                                                               |
| ------------------------- | --------------------------------------------------------------------- |
| `test-config.json`        | Committed defaults. No secrets.                                        |
| `test-config.schema.json` | Editor/JSON-Schema description of the same contract.                  |
| `secrets.ts`              | Runtime-only secret resolution. No values are stored here.             |
| `../webparts/hacktheworld/config/loader.ts` | Typed loader, validation and safety guards (authoritative). |

## Changing an identifier

Edit the value in `test-config.json`, then:

```powershell
npm run config:validate   # schema + semantic validation, exits non-zero on errors
npm run research:test     # unit tests, including configuration tests
```

Every consumer reads through `loadTestConfig()`, so a change takes effect
everywhere: the web part (`components/config.ts`), the research toolkit and the
test runner.

Typical edits:

```jsonc
"resources": {
  "lists": { "primary": { "id": "<list-guid>", "title": "<list title>" } },
  "items": { "primaryItemId": 2 }
},
"principals": {
  "users": { "primary": { "upn": "user@tenant.onmicrosoft.com", "id": 12 } },
  "groups": { "primaryGroupId": 50 }
}
```

## Environment overrides

The committed file is the default environment. Runtime overrides are merged on
top (arrays replace, objects merge):

```ts
import { loadTestConfig } from "../config/loader";

const config = loadTestConfig({
  environment: { siteUrl: "https://contoso.sharepoint.com/sites/research" },
  resources: { items: { primaryItemId: 7 } },
});
```

For a second tenant, keep a separate branch or a separate config file checked
out locally. `test-config.local.json` is git-ignored for scratch overrides.

## Secrets

Credentials, cookies, bearer tokens and client secrets must **never** appear in
this file. The loader rejects any object key matching
`password|secret|token|apikey|authorization|cookie|...` at any depth, and
`npm run config:validate` fails the build.

Authentication in the SPFx web part comes from the current session
(`SPFx(context)` / `credentials: "include"`). The research transport takes a
caller-supplied cookie or token from the browser or Burp, never from disk.

### Supplying a secret at runtime

`src/config/secrets.ts` resolves known secrets on demand. Nothing is committed
and there is no default; a missing value throws with the expected variable name.

Resolution order, first match wins:

1. **In-memory**, for tests and tools:

   ```ts
   import { setSecret, clearSecret } from "../config/secrets";

   setSecret("TEST_PASSWORD", valueFromCaller);
   try {
     // use it
   } finally {
     clearSecret("TEST_PASSWORD");
   }
   ```

2. **`sessionStorage`**, for the browser research UI. Tab-scoped, but readable by
   any script on the same origin, so never paste production credentials into a
   research tab:

   ```ts
   sessionStorage.setItem("sp-research.secret.TEST_PASSWORD", value);
   ```

3. **Environment variable**, for Node tooling:

   ```powershell
   $env:SP_RESEARCH_TEST_PASSWORD = "..."   # current shell only
   npm run research:test
   ```

   The pattern is `SP_RESEARCH_<NAME>`, where `<NAME>` is the key passed to
   `getSecret`.

Known names are defined in `secrets.ts`. Adding one there is required; an
unknown name is rejected rather than read from an arbitrary environment
variable.

### `constants.ts` migration note

`src/webparts/hacktheworld/components/constants.ts` used to contain tenant URLs,
emails, IDs and plaintext passwords. It now derives every non-secret value from
`test-config.json` and exposes the former passwords through `getSecret(...)`.
The old literal values remain in git history from commit `9586511`: rotate them
and see `docs/05-vulnerability-report.md` (V-01) before sharing the repository.

## Safety switches

`safety` is enforced in code, not advisory:

| Switch                        | Default | Effect when false                                              |
| ----------------------------- | ------- | -------------------------------------------------------------- |
| `allowDestructive`            | `false` | `assertDestructiveAllowed()` throws for DELETE/PATCH/PUT/MERGE. |
| `allowPermissionChanges`      | `false` | Blocks `breakroleinheritance`, `resetroleinheritance`, ACL add. |
| `allowOutboundWebhooks`       | `false` | Blocks webhook creation (server-side SSRF / tenant state).      |
| `allowLargePayloads`          | `false` | Blocks payloads above `maxPayloadBytes`.                        |
| `allowedHosts`                | one host| `assertAllowedUrl()` refuses any other host.                    |
| `maxRequestsPerRun`           | 200     | Executor stops before exceeding the budget.                     |
| `minimumDelayBetweenRequestsMs` | 250   | Pacing floor to avoid accidental load.                          |

Turning a switch on is only acceptable in a throwaway site collection, after the
boundaries in `docs/01-safe-testing-boundaries.md` have been agreed.
