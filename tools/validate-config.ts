import * as fs from "fs";
import * as path from "path";
import {
  formatIssues,
  validateTestConfig,
  ConfigIssue,
} from "../src/webparts/hacktheworld/config/loader";

/**
 * CLI wrapper around validateTestConfig.
 *
 * The validation rules live in src/webparts/hacktheworld/config/loader.ts so the
 * browser, the tests and this tool all enforce exactly the same contract. This
 * file only handles reading the file, printing and exit codes.
 *
 * Usage:
 *   node .tmp/research/tools/validate-config.js [path-to-config.json]
 */

function defaultConfigPath(): string {
  return path.resolve(__dirname, "..", "..", "..", "src", "config", "test-config.json");
}

function main(argv: string[]): number {
  const target = argv.length > 2 ? path.resolve(argv[2]) : defaultConfigPath();
  let text: string;
  try {
    text = fs.readFileSync(target, "utf8");
  } catch (error) {
    console.error("Cannot read configuration file: " + target);
    console.error(error instanceof Error ? error.message : String(error));
    return 2;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch (error) {
    console.error("Configuration file is not valid JSON: " + target);
    console.error(error instanceof Error ? error.message : String(error));
    return 2;
  }

  const result = validateTestConfig(parsed);
  const errors = result.issues.filter(function (issue: ConfigIssue) {
    return issue.severity === "error";
  });
  const warnings = result.issues.length - errors.length;

  console.log("Configuration: " + target);
  if (result.issues.length === 0) {
    console.log("OK: valid, no issues.");
    return 0;
  }
  console.log(formatIssues(result.issues));
  console.log(errors.length + " error(s), " + warnings + " warning(s).");
  return errors.length === 0 ? 0 : 1;
}

process.exitCode = main(process.argv);
