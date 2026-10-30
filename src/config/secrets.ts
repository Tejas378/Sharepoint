/**
 * Runtime secret resolution.
 *
 * No credential is ever committed to this repository. Values are supplied at
 * runtime, in this order of precedence:
 *
 *   1. an in-memory value registered with `setSecret()` (web part / tests)
 *   2. the per-tab `sessionStorage` entry `sp-research.secret.<NAME>`
 *   3. the process environment variable `SP_RESEARCH_<NAME>` (CLI tooling)
 *
 * If a secret is missing, `getSecret()` throws with instructions instead of
 * falling back to a weak or shared default. Callers must handle the error and
 * skip the operation that needs the secret.
 *
 * Storage note: `sessionStorage` is used because it is scoped to the tab and
 * cleared when the tab closes. It is *not* a secure vault: any script running
 * on the same origin can read it. Do not paste production credentials into a
 * research session.
 */

export type SecretName = "TEST_PASSWORD" | "TEJAS_PASSWORD" | "OKKK_PASSWORD";

export const SECRET_NAMES: SecretName[] = ["TEST_PASSWORD", "TEJAS_PASSWORD", "OKKK_PASSWORD"];

const STORAGE_PREFIX = "sp-research.secret.";
const ENV_PREFIX = "SP_RESEARCH_";

/** `process` may not exist in the browser bundle. */
declare const process: { env?: Record<string, string | undefined> } | undefined;

const memory: Record<string, string> = {};

function getSessionStorage(): Storage | null {
  try {
    if (typeof sessionStorage !== "undefined" && sessionStorage !== null) {
      return sessionStorage;
    }
  } catch (error) {
    // Access can throw when storage is blocked by browser privacy settings.
    return null;
  }
  return null;
}

function normaliseName(name: string): SecretName {
  const upper = String(name || "").toUpperCase() as SecretName;
  if (SECRET_NAMES.indexOf(upper) < 0) {
    throw new Error(
      "Unknown secret '" + name + "'. Known secrets: " + SECRET_NAMES.join(", ") + "."
    );
  }
  return upper;
}

function read(name: SecretName): string | undefined {
  if (Object.prototype.hasOwnProperty.call(memory, name)) {
    return memory[name];
  }
  const storage = getSessionStorage();
  if (storage) {
    const stored = storage.getItem(STORAGE_PREFIX + name);
    if (stored !== null && stored.length > 0) {
      return stored;
    }
  }
  if (typeof process !== "undefined" && process && process.env) {
    const fromEnv = process.env[ENV_PREFIX + name];
    if (fromEnv !== undefined && fromEnv.length > 0) {
      return fromEnv;
    }
  }
  return undefined;
}

/** Registers a secret for the current process/tab (not persisted to disk). */
export function setSecret(name: string, value: string): void {
  const key = normaliseName(name);
  if (typeof value !== "string" || value.length === 0) {
    throw new Error("Secret '" + key + "' must be a non-empty string.");
  }
  memory[key] = value;
  const storage = getSessionStorage();
  if (storage) {
    try {
      storage.setItem(STORAGE_PREFIX + key, value);
    } catch (error) {
      // Quota or privacy mode: the in-memory copy still works for this run.
    }
  }
}

/** Removes a secret from memory and sessionStorage. */
export function clearSecret(name: string): void {
  const key = normaliseName(name);
  delete memory[key];
  const storage = getSessionStorage();
  if (storage) {
    try {
      storage.removeItem(STORAGE_PREFIX + key);
    } catch (error) {
      // Ignore: nothing was removed.
    }
  }
}

export function clearAllSecrets(): void {
  for (let i = 0; i < SECRET_NAMES.length; i++) {
    clearSecret(SECRET_NAMES[i]);
  }
}

export function hasSecret(name: string): boolean {
  return read(normaliseName(name)) !== undefined;
}

/**
 * Returns the secret or throws with a message that tells the operator exactly
 * how to supply it. Never returns a default.
 */
export function getSecret(name: string): string {
  const key = normaliseName(name);
  const value = read(key);
  if (value === undefined) {
    throw new Error(
      "Secret '" + key + "' is not available. Supply it at runtime via " +
        "setSecret(\"" + key + "\", value), sessionStorage[\"" + STORAGE_PREFIX + key + "\"], " +
        "or the " + ENV_PREFIX + key + " environment variable. Secrets must never be committed."
    );
  }
  return value;
}

/** Non-throwing lookup for code paths that can degrade gracefully. */
export function tryGetSecret(name: string): string | undefined {
  return read(normaliseName(name));
}
