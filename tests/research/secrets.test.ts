import {
  clearAllSecrets,
  clearSecret,
  getSecret,
  hasSecret,
  SECRET_NAMES,
  setSecret,
  tryGetSecret,
} from "../../src/config/secrets";
import { assertEquals, assertThrows, assertTrue, group, test } from "./harness";

const ENV_NAME = "SP_RESEARCH_TEST_PASSWORD";

function captureError(fn: () => unknown): Error {
  try {
    fn();
  } catch (error) {
    return error instanceof Error ? error : new Error(String(error));
  }
  throw new Error("Expected the function to throw, but it returned normally.");
}

export function registerSecretTests(): void {
  group("secrets: runtime resolution");

  test("every known secret starts unset in a clean test process", function () {
    clearAllSecrets();
    for (let i = 0; i < SECRET_NAMES.length; i++) {
      assertEquals(hasSecret(SECRET_NAMES[i]), false);
    }
  });

  test("reading an unset secret throws instead of returning a default", function () {
    clearAllSecrets();
    const error = captureError(function () {
      getSecret("TEST_PASSWORD");
    });
    assertTrue(error.message.indexOf("TEST_PASSWORD") >= 0, "error must name the missing secret");
    assertTrue(error.message.indexOf(ENV_NAME) >= 0, "error must name the environment variable");
    assertTrue(
      error.message.indexOf("never be committed") >= 0,
      "error must explain the no-commit rule"
    );
  });

  test("setSecret makes a value readable and clears it again", function () {
    setSecret("TEST_PASSWORD", "runtime-only-value");
    assertEquals(getSecret("TEST_PASSWORD"), "runtime-only-value");
    assertEquals(hasSecret("TEST_PASSWORD"), true);
    clearSecret("TEST_PASSWORD");
    assertEquals(hasSecret("TEST_PASSWORD"), false);
    assertEquals(tryGetSecret("TEST_PASSWORD"), undefined);
  });

  test("an empty value is rejected", function () {
    assertThrows(function () {
      setSecret("TEST_PASSWORD", "");
    });
  });

  test("an unknown secret name is rejected", function () {
    assertThrows(function () {
      setSecret("NOT_A_SECRET", "value");
    });
    assertThrows(function () {
      getSecret("NOT_A_SECRET");
    });
  });

  test("the environment variable is used when nothing is set at runtime", function () {
    clearAllSecrets();
    process.env[ENV_NAME] = "from-environment";
    try {
      assertEquals(getSecret("TEST_PASSWORD"), "from-environment");
    } finally {
      delete process.env[ENV_NAME];
      clearAllSecrets();
    }
  });

  test("an in-memory value takes precedence over the environment", function () {
    process.env[ENV_NAME] = "from-environment";
    try {
      setSecret("TEST_PASSWORD", "from-memory");
      assertEquals(getSecret("TEST_PASSWORD"), "from-memory");
    } finally {
      delete process.env[ENV_NAME];
      clearAllSecrets();
    }
  });

  test("clearAllSecrets removes every known secret", function () {
    setSecret("TEST_PASSWORD", "one");
    setSecret("TEJAS_PASSWORD", "two");
    setSecret("OKKK_PASSWORD", "three");
    clearAllSecrets();
    for (let i = 0; i < SECRET_NAMES.length; i++) {
      assertEquals(hasSecret(SECRET_NAMES[i]), false);
    }
  });
}
