# 01 - Safe testing boundaries

These boundaries are enforced in code where possible and are mandatory where
they cannot be enforced. "The tool allowed it" is not authorization.

## 1. Scope

### In scope by default

- Only the hosts listed in `safety.allowedHosts`.
- Only the site in `environment.siteUrl`.
- Only accounts and resources created for this research.
- Read-only requests: GET, HEAD, OPTIONS.

### Out of scope unless explicitly authorized in writing

- Production sites, tenant admin centre, OneDrive for real users.
- Other tenants, including `*.sharepoint.com` hosts not in the allow-list.
- Microsoft Graph endpoints other than those you are authorized to call.
- Anything that reads, changes or exfiltrates real user data.
- Denial of service, load testing, brute force, credential stuffing.
- Social engineering, phishing, or contacting real users.
- Persistence: webhooks, subscriptions, workflows, apps, custom actions that
  survive the test window.

`assertAllowedUrl()` refuses any host outside `safety.allowedHosts`, but a
misconfigured allow-list is still your responsibility. Check it in Step 3 of the
first-step workflow.

## 2. Tenant and account setup

Use a dedicated site collection, ideally in a trial or developer tenant.

Minimum setup:

1. One primary test list (`resources.lists.primary`).
2. One secondary item in that list.
3. One test folder.
4. Two test users: a normal member and a second principal for access-control
   comparisons.
5. One test group whose membership you control.

Do **not** point the primary list at a list that contains real business data.
The toolkit generates writes, deletions and permission changes against it.

## 3. Hard limits

| Limit | Default | Where enforced |
| --- | --- | --- |
| Destructive verbs | disabled | `assertDestructiveAllowed`, executor |
| Permission changes | disabled | `assertPermissionChangeAllowed` |
| Outbound webhooks | disabled | `assertWebhookAllowed` |
| Requests per run | 200 | `RequestExecutor` budget |
| Payload size | 65,536 bytes | config / payload generator |
| Request timeout | 30,000 ms | `sendRequest`, executor |
| Minimum spacing | 250 ms | executor pacing |
| Hosts | one tenant | `assertAllowedUrl` |

If a test needs more than 200 requests, split it into separate runs and review
the previous run first. Do not raise the budget to work around a hang.

## 4. Method rules

| Method | Default | Notes |
| --- | --- | --- |
| GET / HEAD / OPTIONS | allowed | Still respect rate limits. |
| POST (read-style actions) | allowed | e.g. `RenderListDataAsStream`, `contextinfo`. |
| POST (create) | needs `allowDestructive` | Creates state. |
| PATCH / MERGE / PUT | needs `allowDestructive` | Modifies state. Always send an explicit `If-Match`. |
| DELETE | needs `allowDestructive` | Prefer recycle over permanent delete. |
| Method-override probes | controlled probe | `POST` + `X-HTTP-Method` must not be used to bypass a disabled switch. |

Never attach every verb to every endpoint. The registry declares genuine
methods; `variations.ts` labels the rest as probes or hypotheses.

## 5. Data handling

### Allowed to store

- Status codes, timings, headers that are not credentials.
- Response shapes with tenant identifiers replaced by placeholders.
- Reproducible request templates with the digest and cookie removed.
- Your own notes and screenshots with secrets redacted.

### Never store or share

- Cookies, bearer tokens, request digests, client secrets, passwords.
- Real user names, emails, phone numbers or document contents.
- Raw responses containing personal data.
- Anything from a tenant you are not authorized to test.

`redact.ts` redacts headers, URL query values, JSON keys and free text. It is a
safety net, not a guarantee. Review exports manually before sharing.

### Secret handling

`src/config/secrets.ts` resolves secrets at runtime only. `sessionStorage` is
tab-scoped but readable by any script on the same origin. Do not paste
production credentials into a research tab.

## 6. Rate limiting and availability

- Keep `minimumDelayBetweenRequestsMs` at 250 ms or higher.
- Do not parallelise destructive requests.
- Stop immediately on:
  - a sustained wave of 429 or 503 responses;
  - any 5xx that is not reproducible with a single request;
  - evidence that a request touched data outside the test resources;
  - any sign that another user is being affected.
- Do not retry non-idempotent requests automatically. `RequestExecutor` only
  retries when the operation is safe and the status is transient.
- Do not use the toolkit to test capacity or exhaust quotas.

## 7. Injection testing boundaries

Payload families exist to test **this tenant's** handling of input. They must
not be used to attack third parties.

| Family | Allowed | Boundary |
| --- | --- | --- |
| IDOR | Yes | Use foreign GUIDs/ids that should 404. Do not enumerate real users. |
| Mass assignment | Yes | Only on disposable items. |
| OData injection | Yes | Only in query options against the test list. |
| XSS | Yes, stored payloads only on disposable items | Do not target real users' browsers or the tenant admin UI. |
| SQL injection | Yes, only where a downstream sink is plausible | Stop at error-based evidence; no data extraction. |
| SSRF | Only with an internal callback you control | No external callbacks; `allowOutboundWebhooks` stays false. |
| Path traversal | Yes, within the test library | Do not read system files or other sites. |
| Method confusion | Yes | Do not use it to bypass the destructive switch. |
| Business logic | Yes, on test state | Do not modify real workflows or approvals. |

If a payload could affect another user, stop and ask for written approval.

## 8. Stop conditions

Stop the run and notify the project owner if any of these occurs:

1. A response contains real user data or document contents.
2. A request changes state outside the test resources.
3. A permission change affects another user's access.
4. A webhook or subscription is created (even unintentionally).
5. A 5xx persists after one retry.
6. You cannot explain what a variation changed.
7. You lose track of which environment you are targeting.

Do not "finish the run first". Stop, record, clean up.

## 9. What the code refuses

The following throw instead of running:

- a URL on a host outside `safety.allowedHosts`;
- a non-HTTPS URL;
- a destructive operation while `allowDestructive` is false;
- a permission change while `allowPermissionChanges` is false;
- a webhook while `allowOutboundWebhooks` is false;
- a payload above `maxPayloadBytes` when `allowLargePayloads` is false;
- a request budget above `maxRequestsPerRun`;
- a missing runtime secret;
- a config with an unknown top-level section or a secret-looking key.

If you find a path that bypasses one of these, that is a bug in the toolkit.
Report it and fix it before continuing.

## 10. Legal and ethical minimum

- Written authorization, scope and a kill switch.
- Minimum necessary access; no real data.
- No persistence beyond the agreed window.
- Responsible disclosure for anything real you find.
- Follow the target's bug bounty policy and Microsoft's terms.
