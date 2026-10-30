# 07 - Evidence log and coverage statuses

This closes the gap between "the toolkit can build a request" and "we actually
tested something". Before this file existed, `executed` and `verified` were
hardcoded to `false` in `tools/coverage-report.ts`, which was honest but made the
coverage report unable to improve from real work.

Now they are computed from `docs/evidence/evidence.json`. If that file is
absent, both stay `0` and the report says why.

## 1. The four statuses stay distinct

| Status | Meaning | Where it comes from |
| --- | --- | --- |
| Implemented | An `EndpointDefinition` exists in the registry. | `research/registry.ts` |
| Generated | `buildRequest()` produced a complete URL. This proves construction only. | `tools/coverage-report.ts` |
| Executed | A request was actually sent and a response (or network failure) was recorded. | `evidence.json` |
| Verified | Manual review confirmed the behaviour. Evidence must exist first. | `evidence.json` |

A generated request is never promoted to executed automatically. Nothing in the
repository sends traffic during a test run.

## 2. Recording evidence

1. Generate the request:
   `npm run research:request -- lists.collection.get`
2. Send it from an authorised browser session or paste it into Burp.
3. Record the observation. Keep the response body short and redacted; never
   store cookies, `Authorization`, `X-RequestDigest` values or passwords.
4. Append an entry to `docs/evidence/evidence.json` (copy
   `docs/evidence/evidence.example.json` to start).
5. Run `npm run research:evidence` and `npm run research:coverage`.

`research/evidence.ts#evidenceFromExecution` builds a redacted entry from a
`RequestExecutor` result. It is available for a future UI or CLI runner; the
legacy UI is **not** wired to it yet, so entries are currently added by hand.

## 3. Schema

```json
{
  "schemaVersion": 1,
  "entries": [
    {
      "endpointId": "lists.collection.get",
      "method": "GET",
      "outcome": "executed",
      "statusCode": 200,
      "executedAt": "2026-10-30T10:15:00.000Z",
      "durationMs": 184,
      "operator": "manual",
      "requestUrl": "https://contoso.sharepoint.com/_api/web/lists?$top=2",
      "responsePreview": "{\"d\":{\"results\":[...]}}",
      "notes": "Baseline returned two lists. No anomaly."
    }
  ]
}
```

| Field | Required | Rules |
| --- | --- | --- |
| `endpointId` | yes | Must match a registry id to affect coverage. Unknown ids are reported. |
| `method` | yes | Non-empty, normalised to uppercase. |
| `outcome` | yes | `not-run`, `executed`, `interesting`, `inconclusive` or `verified`. |
| `statusCode` | yes | Integer `100-599`, or `null` when no response arrived. |
| `executedAt` | yes | ISO-8601 timestamp. |
| `durationMs` | no | Non-negative number. |
| `operator` | no | Free text, redacted. |
| `notes` | no | Free text, redacted. |
| `requestUrl` | no | Unredacted input is accepted, but the stored value is redacted. |
| `responsePreview` | no | Truncated to 2,000 characters after redaction. |

Rules enforced by `parseEvidenceLog()`:

- an unknown `schemaVersion` is an error;
- an unknown `outcome` is an error (a misspelling must not count as execution);
- unknown entry fields are dropped;
- a `not-run` entry does **not** mark the endpoint as executed;
- a `verified` entry also counts as executed.

## 4. How coverage changes

`applyEvidenceToCoverage()` copies the report and:

- marks `status.executed` for an endpoint with any non-`not-run` entry;
- marks `status.verified` for an endpoint with a `verified` entry;
- updates `endpoints.executed`, `endpoints.verified` and
  `modules.<name>.verified`;
- recomputes the `statuses` block from entry outcomes;
- replaces the "no live session" limitation with the actual entry count.

Deduplication is per endpoint for the counts; multiple entries per endpoint are
kept in the evidence report so retries and repeated runs stay visible.

## 5. Commands

```powershell
npm run research:evidence        # human-readable summary
npm run research:evidence -- --json
npm run research:evidence -- --file C:\path\to\evidence.json
npm run research:coverage        # auto-reads docs/evidence/evidence.json
npm run research:test            # offline tests, including the evidence parser
```

`research:coverage` accepts `--no-evidence` to ignore the file and
`--evidence <path>` to use another log.

## 6. What is still not proven

- No live tenant was used while building this repository, so the committed
  evidence file is absent and the counts remain `0`.
- The evidence log records what a human says happened. It is not a substitute
  for reviewing the raw request/response in Burp.
- Microsoft Graph evidence needs its own bearer-token workflow and is not
  produced by the SharePoint cookie transport.
