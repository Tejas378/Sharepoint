# 00 - First-step workflow

This is the concrete sequence to follow the first time you run the toolkit, and
every time you start a new research session. It assumes PowerShell on Windows.

If a step fails, stop at that step. Do not continue with a half-valid
configuration.

## Step 0 - Confirm authorization (outside the tool)

Before running anything:

1. You have **written permission** to test the target tenant.
2. The target is a **throwaway site collection** for any destructive test.
3. You know the **out-of-scope** hosts and features.
4. You have a **contact** and a **stop condition** (for example: stop on any
   evidence of real user data, or on the first 5xx wave).

The safety switches in the config enforce some of this mechanically, but they do
not replace authorization. See `01-safe-testing-boundaries.md`.

## Step 1 - Install and type-check

```powershell
npm install
npx tsc -p tsconfig.research.json
npx tsc --noEmit
```

Expected: both commands exit 0. The second one checks the SPFx web part. Do not
run `gulp build` on Node 24; this repository pins Node 18 in `package.json`, and
the SPFx toolchain was not validated here.

## Step 2 - Point the toolkit at your tenant

Edit **only** `src/config/test-config.json`:

```jsonc
"environment": {
  "siteUrl": "https://<tenant>.sharepoint.com/sites/<research-site>"
},
"resources": {
  "lists": {
    "primary": { "id": "<list-guid>", "title": "<list title>" }
  },
  "items": { "primaryItemId": 2 }
},
"safety": {
  "allowedHosts": ["<tenant>.sharepoint.com"]
}
```

Leave `allowDestructive`, `allowPermissionChanges` and `allowOutboundWebhooks`
at `false` until Step 7.

## Step 3 - Validate the configuration

```powershell
npm run config:validate
```

Expected: `OK: valid, no issues.` and exit code 0.

Common failures and what they mean:

| Message | Fix |
| --- | --- |
| `Unknown top-level section` | Fix the spelling; misspelt keys are rejected, not ignored. |
| `Must be an https:// URL` | Use `https://`, not `http://`. |
| `Site host must match the tenant host` | `tenantUrl` and `siteUrl` must share a host. |
| `The configured site host ... is not in the allow-list` | Add the host to `safety.allowedHosts`. |
| `Secret-looking key in configuration` | Remove the credential; supply it at runtime. |

## Step 4 - Run the offline test suite

```powershell
npm run research:test
```

Expected: `274 passed, 0 failed, 274 total`.

This proves the request engine is internally consistent. It sends **no**
traffic, so it does not prove that any endpoint works against your tenant.

## Step 5 - Generate one baseline request

Start with the read-only collection endpoint:

```powershell
npm run research:request -- lists.collection.get
```

Expected output (abridged):

```http
GET /_api/web/lists HTTP/1.1
Host: <tenant>.sharepoint.com
Accept: application/json;odata=verbose
```

Now generate a single item read. `listId` and `itemId` default to the values in
`test-config.json`, so you do not have to retype GUIDs:

```powershell
npm run research:request -- items.getById
```

For a write, use a body file to avoid PowerShell quoting problems:

```powershell
'{"Title":"controlled probe"}' | Set-Content .tmp\body.json
npm run research:request -- items.update --body-file .tmp/body.json --method PATCH
```

The output ends with:

```http
X-RequestDigest: {{REQUEST_DIGEST}}
If-Match: *
```

`{{REQUEST_DIGEST}}` is a placeholder on purpose. The tool never prints a live
digest. Replace it in Burp with a digest fetched from `/_api/contextinfo` in the
same authenticated session, or let `RequestExecutor` fetch it in code.

## Step 6 - Send the baseline and record evidence

1. Paste the raw request into Burp Repeater (right-click -> Paste from file, or
   paste into the raw editor).
2. Send it once.
3. Record, in your own evidence log:
   - timestamp, endpoint id, method, URL path (no query secrets);
   - status code, duration, `Content-Type`, `ETag`;
   - whether the body shape matched the expected envelope;
   - any error text, with tenant-specific values redacted.

Do not mark anything `verified` from one request. A baseline only proves the
request is well-formed and authorized.

## Step 7 - Only if authorized: enable the needed switch

In `src/config/test-config.json`:

```jsonc
"safety": {
  "allowDestructive": true,
  "allowPermissionChanges": false,
  "allowOutboundWebhooks": false
}
```

Then re-run Step 3. The validator will emit a warning that destructive testing
is enabled; that warning is expected and should stay visible in your notes.

Destructive testing order:

1. create a disposable item;
2. update it;
3. add an attachment;
4. delete it;
5. verify the recycle bin state;
6. only then touch permissions, and only on the disposable item.

## Step 8 - Move to variations

Variations are produced by `generateVariations` in
`src/webparts/hacktheworld/research/variations.ts`. They are grouped as:

- `documented` - declared method or documented query option;
- `controlled-probe` - a deliberate mutation with a known expected failure;
- `hypothesis` - a plausible weakness that has **not** been observed.

For every variation, record the baseline response next to the mutated response.
A difference is a lead, not a finding.

## Step 9 - Triage findings

Use this ladder and do not skip a rung:

1. **Request generated** - the URL was constructed.
2. **Request executed** - the server responded.
3. **Behaviour observed** - the response differs from the baseline in a way you
   can reproduce.
4. **Impact demonstrated** - you can show what an attacker gains (data, access,
   state change) without harming real users.
5. **Verified** - steps 1-4 are reproducible and documented.

Only rung 5 goes in the vulnerability report as confirmed.

## Step 10 - Clean up

- Delete disposable items, files, folders, links and permissions you created.
- Restore inheritance if you broke it.
- Remove any webhook/subscription you created.
- Clear runtime secrets: `clearAllSecrets()` or close the tab.
- Set the safety switches back to `false`.

## Quick command reference

| Command | Purpose |
| --- | --- |
| `npm run config:validate` | Validate identifiers, safety switches, no secrets. |
| `npm run research:test` | 274 offline tests. |
| `npm run research:coverage` | Endpoint/module/method coverage; executed=0 by design. |
| `npm run research:request -- <id>` | Print a raw HTTP request; does not send. |
| `npx tsc --noEmit` | Type-check the SPFx web part. |
