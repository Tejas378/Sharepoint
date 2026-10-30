# Evidence folder

`evidence.json` in this folder is the **only** input for the `executed` and
`verified` counts in `npm run research:coverage`.

- It is not committed while it contains real tenant observations. Add it
  locally (or to an ignored private branch) and keep secrets out of it.
- A missing file is not an error: the report shows `executed=0`, `verified=0`
  and says so.
- Every entry is validated and redacted before it can affect a report. A typo
  in `outcome` is rejected, never silently counted.

Start from `evidence.example.json`, which contains one `not-run` placeholder.
See `docs/07-evidence-log.md` for the fields and the manual workflow.

Commands:

```powershell
npm run research:evidence
npm run research:evidence -- --file docs/evidence/evidence.json --json
npm run research:coverage
```
