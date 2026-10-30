/**
 * Minimal assertion + reporting harness.
 *
 * Deliberately dependency-free: the research modules are plain TypeScript with
 * no SPFx imports precisely so they can be compiled with tsc and executed in
 * Node without a test framework. See tsconfig.research.json.
 */

export type TestFn = () => void | Promise<void>;

interface Case {
  group: string;
  name: string;
  fn: TestFn;
}

const cases: Case[] = [];
let currentGroup = "general";

export function group(name: string): void {
  currentGroup = name;
}

export function test(name: string, fn: TestFn): void {
  cases.push({ group: currentGroup, name: name, fn: fn });
}

export function assertTrue(value: unknown, message: string = "Expected a truthy value"): void {
  if (!value) {
    throw new Error(message);
  }
}

export function assertEquals(actual: unknown, expected: unknown, message?: string): void {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) {
    throw new Error(
      (message ? message + ": " : "") + "expected " + e + " but received " + a
    );
  }
}

export function assertNotEquals(actual: unknown, forbidden: unknown, message: string): void {
  if (JSON.stringify(actual) === JSON.stringify(forbidden)) {
    throw new Error(message + ": value must not equal " + JSON.stringify(forbidden));
  }
}

export function assertContains(haystack: string, needle: string, message: string): void {
  if (haystack.indexOf(needle) < 0) {
    throw new Error(message + ": expected to find '" + needle + "' in '" + haystack + "'");
  }
}

export function assertThrows(fn: () => unknown, pattern?: string): void {
  let threw = false;
  let message = "";
  try {
    fn();
  } catch (error) {
    threw = true;
    message = error instanceof Error ? error.message : String(error);
  }
  if (!threw) {
    throw new Error("Expected the call to throw" + (pattern ? " matching /" + pattern + "/" : ""));
  }
  if (pattern && message.search(new RegExp(pattern)) < 0) {
    throw new Error("Thrown message '" + message + "' does not match /" + pattern + "/");
  }
}

export function assertLength(value: unknown[], expected: number, message: string): void {
  if (value.length !== expected) {
    throw new Error(message + ": expected length " + expected + " but received " + value.length);
  }
}

export interface RunSummary {
  passed: number;
  failed: number;
  failures: Array<{ group: string; name: string; error: string }>;
}

export async function runAll(): Promise<RunSummary> {
  const failures: Array<{ group: string; name: string; error: string }> = [];
  let passed = 0;
  let lastGroup = "";

  for (let i = 0; i < cases.length; i++) {
    const item = cases[i];
    if (item.group !== lastGroup) {
      console.log("\n" + item.group);
      lastGroup = item.group;
    }
    try {
      await item.fn();
      passed++;
      console.log("  PASS  " + item.name);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      failures.push({ group: item.group, name: item.name, error: message });
      console.log("  FAIL  " + item.name);
      console.log("        " + message);
    }
  }

  console.log(
    "\n" + passed + " passed, " + failures.length + " failed, " + cases.length + " total"
  );
  return { passed: passed, failed: failures.length, failures: failures };
}
