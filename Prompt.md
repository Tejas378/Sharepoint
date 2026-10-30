# Master Prompt: Expand and Improve My SharePoint API Security Research Project

## Role

Act as a senior application security researcher, bug bounty hunter, SharePoint Online and SharePoint REST API specialist, and experienced TypeScript/Node.js engineer.

You may be running in any AI coding environment. Adapt your workflow to the tools, permissions, and capabilities available to you. Do not assume a specific AI model, IDE, agent, or command-line tool.

## 1. Main Objective

I have an existing project for executing, exploring, and testing SharePoint APIs.

Your task is to **inspect, improve, and expand the existing project** so it covers a much broader SharePoint API attack surface and supports effective manual security research.

The goal is not to automatically declare vulnerabilities. I want to generate valid API requests, explore parameters and methods, inspect responses, identify security-relevant behavior, and manually investigate findings using Burp Suite.

**Important:** Do not rebuild the project from scratch. Understand its existing architecture, preserve working functionality, and make incremental, maintainable improvements.

## 2. Start by Understanding the Existing Project

Before modifying anything:

1. Inspect the repository structure and identify its language, framework, entry points, and execution flow.
2. Review `package.json`, lockfiles, dependencies, configuration, and existing documentation.
3. Identify how API endpoints, HTTP methods, headers, query parameters, request bodies, and authentication are configured.
4. Find how requests are constructed, executed, and displayed.
5. Identify existing endpoint definitions, payload generators, tests, and export functionality.
6. Inspect existing errors, incomplete implementations, duplicated logic, and untested functionality.
7. Determine which SharePoint API families are already supported and which are missing.

Do not claim to have inspected a file, run a command, or verified a result unless you actually did so.

Start with a concise repository assessment and a prioritized implementation plan. Then proceed with the work rather than stopping at recommendations.

## 3. Inspect the Installed PnPjs Implementation

If the project uses `@pnp/sp`, inspect the exact installed version and its available APIs.

- Review installed package source code, TypeScript definitions, exports, and relevant documentation.
- Trace API methods to the actual HTTP requests they generate.
- Identify supported methods, URL patterns, query options, headers, and request-body formats.
- Check which features are implemented by the installed version instead of assuming current documentation matches it.
- Identify useful API methods and modules that the project does not yet expose.
- Avoid unnecessary dependency upgrades.

If PnPjs is not used, inspect the actual SharePoint request implementation and apply the same analysis to the existing code.

## 4. Expand API Coverage

Use the following order of priority.

### Priority 1: SharePoint Lists and List Items

Implement comprehensive coverage for:

- Listing, retrieving, creating, updating, and deleting lists.
- Retrieving list metadata and properties.
- Retrieving, creating, updating, and deleting list items.
- Item identifiers, list identifiers, GUIDs, and alternative identifier formats supported by the API.
- Field selection, expansion, filtering, ordering, pagination, and other supported OData query options.
- Field definitions, field values, required fields, and field types.
- Content types and content-type-specific item behavior.
- Attachments and their associated operations where supported.
- Item-level and list-level permissions where supported.
- ETags and conditional updates.
- Supported batching operations and change-tracking functionality.

Treat Lists and List Items as the initial milestone. Make their coverage reliable before expanding to less important modules.

### Priority 2: Other SharePoint REST API Modules

Expand coverage to relevant modules such as:

- Sites, webs, and subsites.
- Files, folders, libraries, and document metadata.
- Fields and content types.
- Users, groups, role assignments, and permissions.
- Search and search query parameters.
- Change logs and change tokens.
- Site and list configuration.
- Supported batch requests.

Use actual installed APIs, documented endpoints, and observed authorized traffic to determine which operations are valid.

### Priority 3: Microsoft Graph

If Microsoft Graph is part of the project or is a useful extension, support it as a separate API family.

Clearly distinguish Graph endpoints from SharePoint REST endpoints. Do not assume that they use identical authentication scopes, URL structures, headers, payload formats, permissions, or error behavior.

## 5. Expand HTTP Methods and Request Variations

For every applicable endpoint, determine which methods are genuinely supported.

Consider GET, POST, PATCH, PUT, DELETE, HEAD, OPTIONS, and SharePoint-specific update conventions such as MERGE or method-override headers where applicable.

Do not blindly attach every method to every endpoint. Distinguish documented behavior from exploratory probes.

Generate request variations involving:

- Missing, null, empty, malformed, and boundary-value parameters.
- Valid, invalid, nonexistent, and alternate resource identifiers.
- Missing, additional, unexpected, or incorrectly typed JSON properties.
- Different supported content types and serialization formats.
- Query parameter ordering and URL encoding.
- Supported OData query options and combinations.
- Optional headers, conditional headers, and relevant request metadata.
- Valid and invalid field values.
- Normal, empty, malformed, and unexpected response conditions.

Record the rationale for each variation and whether it is a documented request, a controlled probe, or an unverified hypothesis.

## 6. Improve Parameter Discovery

Identify additional parameters using reliable sources:

- Installed library source and TypeScript definitions.
- Official Microsoft documentation.
- API schemas and relevant protocol specifications.
- Existing project code and tests.
- Authorized requests captured in Burp Suite.
- Valid request examples and response metadata.

Where useful, build a parameter registry containing the endpoint, method, parameter name, location, expected type, source, and verification status.

Do not present guessed parameter names as confirmed supported parameters.

## 7. Improve Security Payload Generation

Expand the project's payload generation so it supports manual testing of relevant security classes, including:

- Broken access control and insecure direct object references.
- Mass assignment and unauthorized property modification.
- OData query manipulation and input-validation weaknesses.
- Cross-site scripting where data flows into a browser-rendered context.
- SQL injection where a plausible downstream SQL sink exists.
- Server-side request forgery where a server-side URL-fetching feature exists.
- Path traversal and unsafe file handling where relevant.
- Template injection, deserialization issues, or command execution only where the implementation provides a plausible sink.
- Business-logic errors, workflow bypasses, and state-transition weaknesses.
- Method confusion, incorrect authorization enforcement, and conditional-request edge cases.

For each payload family:

1. Identify the suspected sink or behavior that makes the test relevant.
2. Use endpoint-appropriate input locations and data types.
3. Include normal control requests for comparison.
4. Generate bounded, clearly categorized test cases.
5. Explain what response differences might warrant manual investigation.
6. Keep payload generation separate from automatic vulnerability conclusions.

Do not inject unrelated payloads into every parameter. Do not treat a server error, status code, or unusual response alone as proof of a vulnerability.

Use only systems and tenants for which testing is authorized. Avoid destructive actions, real data exposure, external callbacks, credential theft, persistence, or disruptive exploitation.

## 8. Fix Existing Request-Generation Bugs

Review and correct issues such as:

- Incorrect endpoint URLs and URL encoding.
- Invalid HTTP methods or method-override conventions.
- Incorrect request-body schemas.
- Wrong content types or serialization formats.
- Missing required headers.
- Confusion between SharePoint REST and Microsoft Graph conventions.
- Incorrect parameter placement.
- Invalid identifiers or inconsistent identifier formats.
- Incorrect handling of ETags and conditional requests.
- Poor parsing and presentation of errors.
- Duplicate requests caused by flawed execution logic.
- Broken tests, type errors, and build failures.

Preserve backward compatibility where practical. Explain any behavior-changing fix and add regression tests when possible.

## 9. Implement Correct Request-Digest Handling

If the application performs SharePoint REST operations requiring a request digest, implement a reliable and centralized mechanism.

- Retrieve the digest from `/_api/contextinfo` when appropriate.
- Parse the response correctly.
- Track its expiration.
- Refresh it when expired or rejected.
- Handle concurrent requests without creating unnecessary duplicate refreshes.
- Prevent infinite retry loops.
- Avoid unnecessary digest retrieval for operations that do not require it.
- Keep digest handling separate from Microsoft Graph authentication and request conventions.

Do not hardcode, log, or export live authentication tokens, cookies, request digests, or other secrets.

## 10. Improve ETag and Request Execution Handling

Support conditional requests where applicable, including relevant `If-Match` and `If-None-Match` scenarios.

Distinguish between:

- Expected successful operations.
- Expected validation failures.
- Authentication and authorization failures.
- Conflict and concurrency responses.
- Rate limiting and transient server failures.
- Unexpected responses requiring investigation.

Use bounded retries and sensible timeouts. Do not automatically retry destructive or non-idempotent operations unless the retry is demonstrably safe.

Where practical, record request duration, response status, and relevant response metadata without unnecessarily storing sensitive information.

## 11. Make Burp Suite Manual Testing Easier

I use Burp Suite to analyze requests and responses manually.

Add or improve functionality that allows me to:

- Inspect the complete generated HTTP request.
- Review the URL, method, headers, query parameters, and body.
- Compare baseline requests with mutated requests.
- Export requests in a Burp-friendly format, such as raw HTTP or another practical format supported by the project.
- Identify which parameters and payload variations changed.
- Mark tests as not run, executed, interesting, inconclusive, or verified.
- Record notes and reproduction steps for promising results.

Handle hostnames, authentication, cookies, and other environment-specific values safely. Redact secrets by default when exporting logs or sharing results.

Do not imply that a generated request has been sent or tested unless it actually has.

## 12. Add an Endpoint Registry and Coverage Tracking

Create or improve a structured registry that captures:

- API family.
- Endpoint URL or URL template.
- HTTP method.
- Operation description.
- Parameters and their locations.
- Request-body schema or example.
- Required headers.
- Authentication requirements, where known.
- Payload families applicable to the endpoint.
- Whether the request has been generated, executed, or manually verified.
- Source and evidence supporting the endpoint definition.

Add a coverage report showing:

- Implemented endpoint families.
- Supported methods.
- Implemented request variations.
- Payload families available.
- Tests passing or failing.
- Known gaps and limitations.

Keep these statuses distinct:

- **Implemented:** The code exists.
- **Generated:** The request can be constructed.
- **Executed:** The request was actually sent.
- **Verified:** The behavior was confirmed through appropriate testing.

Do not equate implementation with successful execution or vulnerability confirmation.

## 13. Testing and Quality Requirements

For each meaningful change:

1. Add or update unit tests where practical.
2. Test URL and request construction.
3. Validate parameter serialization and payload generation.
4. Test error handling and malformed responses.
5. Check digest caching and refresh logic if applicable.
6. Test ETag and conditional-request handling where relevant.
7. Run available linting, type checking, and build commands.
8. Run the relevant test suite.
9. Report commands actually executed and their results.

If credentials or a live SharePoint tenant are unavailable, use mocks, fixtures, or safe local tests. Clearly state which behaviors still require live-tenant validation.

Do not silently ignore failures or claim tests passed when they were not run.

## 14. Maintain Persistent Project Documentation

Create or update `PROJECT_CONTEXT.md` at the repository root.

It must record:

- Project purpose and architecture.
- Relevant files and entry points.
- Installed framework and PnPjs versions, where applicable.
- API families and endpoint coverage.
- Request-generation and authentication behavior.
- Known bugs and limitations.
- Implemented changes.
- Tests executed and their actual results.
- Outstanding tasks and next steps.
- Important decisions and reasons for them.

Use these task markers:

- `[x]` Completed and verified.
- `[ ]` Not started.
- `[~]` Partially implemented.
- `[!]` Blocked or requires investigation.

Update this file after meaningful milestones so another AI model or developer can continue without repeating the entire investigation.

Never store credentials, access tokens, cookies, or other secrets in this file.

## 15. Working Rules

- Work within the existing repository and its conventions.
- Inspect before modifying.
- Make small, understandable changes.
- Preserve working functionality.
- Avoid unnecessary rewrites and dependency upgrades.
- Do not invent APIs, parameters, test results, or vulnerabilities.
- Distinguish facts, assumptions, hypotheses, and unverified behavior.
- Explain important changes concisely.
- Ask for clarification only when a missing detail genuinely blocks progress.
- If the environment cannot perform a requested action, explain the limitation and implement the best practical alternative.
- Never expose secrets in terminal output, generated reports, or exported requests.
- Keep all testing within authorized scope.

## 16. Recommended Implementation Order

Follow this sequence unless repository evidence justifies a different order:

**Phase 1 — Repository assessment**
- Map the architecture and existing functionality.
- Identify the current PnPjs version and request-execution mechanism.
- Create or update `PROJECT_CONTEXT.md`.
- Produce a prioritized gap analysis.

**Phase 2 — Lists and List Items**
- Inspect installed implementation details.
- Expand endpoint, method, parameter, and query-option coverage.
- Validate request schemas and serialization.
- Add relevant manual-testing variations.
- Add tests and fix confirmed defects.

**Phase 3 — Request reliability**
- Improve digest handling, ETags, error reporting, timeouts, and safe retry behavior.
- Improve request inspection and Burp-friendly export.

**Phase 4 — Other SharePoint REST modules**
- Add missing modules incrementally.
- Use the same registry, validation, and testing standards.

**Phase 5 — Microsoft Graph**
- Add or improve Graph coverage if it fits the project.
- Maintain separate endpoint, authentication, and request conventions.

**Phase 6 — Final validation**
- Run available builds and tests.
- Generate a coverage report.
- Document remaining gaps.
- Update `PROJECT_CONTEXT.md` with a clear handoff.

## 17. Your First Response and First Actions

Begin by inspecting the actual project. Then report:

1. The repository architecture and main entry points.
2. The framework, dependencies, and installed PnPjs version, if applicable.
3. The existing SharePoint Lists and List Items implementation.
4. Confirmed bugs and concrete coverage gaps.
5. A prioritized plan for incremental improvements.
6. Whether `PROJECT_CONTEXT.md` exists and what needs to be added.

After presenting that assessment, begin Phase 1 and proceed into the first implementation milestone. Do not stop after giving advice or a plan.

**Final objective:** Turn the existing project into a reliable, extensible SharePoint API exploration and manual security-testing toolkit, with broad endpoint coverage, accurate request generation, useful payload variations, clear evidence tracking, and documentation that allows any AI coding model or developer to continue the work.