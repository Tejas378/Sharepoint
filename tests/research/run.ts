import { runAll } from "./harness";
import { registerConfigTests } from "./config.test";
import { registerDigestTests } from "./digest.test";
import { registerExecutorTests } from "./executor.test";
import { registerPayloadTests } from "./payloads.test";
import { registerRedactTests } from "./redact.test";
import { registerRegistryTests } from "./registry.test";
import { registerRequestTests } from "./request.test";
import { registerSecretTests } from "./secrets.test";
import { registerUrlTests } from "./url.test";
import { registerVariationTests } from "./variations.test";

/**
 * Entry point for the dependency-free research test suite.
 *
 * Build and run with:
 *   npm run research:build
 *   npm run research:test
 *
 * These tests exercise request construction, payload generation, digest
 * handling, redaction and configuration validation. They do not send traffic,
 * so a passing run proves the toolkit is internally consistent, not that any
 * SharePoint endpoint behaves a particular way.
 */
async function main(): Promise<void> {
  registerConfigTests();
  registerUrlTests();
  registerRequestTests();
  registerRegistryTests();
  registerPayloadTests();
  registerVariationTests();
  registerExecutorTests();
  registerDigestTests();
  registerRedactTests();
  registerSecretTests();

  const summary = await runAll();
  if (summary.failed > 0) {
    console.log("\nFailures:");
    for (let i = 0; i < summary.failures.length; i++) {
      const failure = summary.failures[i];
      console.log("  [" + failure.group + "] " + failure.name + ": " + failure.error);
    }
    process.exitCode = 1;
    return;
  }
  process.exitCode = 0;
}

main().catch(function (error: unknown) {
  console.error(error instanceof Error ? error.stack || error.message : String(error));
  process.exitCode = 1;
});
