import { describe, expect, test } from "bun:test";
import { redactError, redactUrl } from "./redact";
import { WopiHostError } from "./errors";

const TOKEN = "DISTINCTIVESECRETTOKEN12345";

describe("redactUrl", () => {
  test("replaces only the access_token value", () => {
    const redacted = redactUrl(`https://cloud.example.com/wopi/files/a?access_token=${TOKEN}&foo=bar`);
    expect(redacted).toContain("access_token=***");
    expect(redacted).toContain("foo=bar");
    expect(redacted).not.toContain(TOKEN);
  });

  test("accepts a URL object and leaves token free URLs unchanged", () => {
    expect(redactUrl(new URL("https://cloud.example.com/a?x=1"))).toBe("https://cloud.example.com/a?x=1");
  });

  test("redacts uppercase, encoded and fragment token forms", () => {
    expect(redactUrl(`https://cloud.example.com/a?ACCESS_TOKEN=${TOKEN}`)).not.toContain(TOKEN);
    expect(redactUrl(`https://cloud.example.com/a#access_token=${TOKEN}`)).not.toContain(TOKEN);
    expect(redactUrl(`https://cloud.example.com/a?next=%3Faccess_token%3D${TOKEN}`)).not.toContain(TOKEN);
    const redacted = redactError(new Error(`see https://h/a?ACCESS_TOKEN=${TOKEN} and access_token%3D${TOKEN}`));
    expect(redacted.message).not.toContain(TOKEN);
  });

  test("scrubs reason and currentLock on a WopiHostError", () => {
    const err = new WopiHostError("lockMismatch", "mismatch", {
      status: 409,
      currentLock: `lock https://h/a?access_token=${TOKEN}`,
      reason: `held https://h/a?ACCESS_TOKEN=${TOKEN}`,
    });
    const redacted = redactError(err) as WopiHostError;
    expect(redacted.currentLock).not.toContain(TOKEN);
    expect(redacted.reason).not.toContain(TOKEN);
  });
});

describe("redactError", () => {
  test("removes a token embedded in a plain error", () => {
    const redacted = redactError(new Error(`fetch failed for https://h/a?access_token=${TOKEN}`));
    expect(redacted.message).not.toContain(TOKEN);
    expect(String(redacted)).not.toContain(TOKEN);
    expect(redacted.message).toContain("access_token=***");
  });

  test("preserves the kind of a WopiHostError", () => {
    const err = new WopiHostError("unauthorized", `bad https://h/a?access_token=${TOKEN}`, { status: 401 });
    const redacted = redactError(err);
    expect(redacted).toBeInstanceOf(WopiHostError);
    expect((redacted as WopiHostError).kind).toBe("unauthorized");
    expect(String(redacted)).not.toContain(TOKEN);
  });

  test("does not carry an original cause that holds a token", () => {
    const cause = new Error(`fetch failed for https://h/a?access_token=${TOKEN}`);
    const err = new WopiHostError("network", "network error", { cause });
    const redacted = redactError(err);
    expect(redacted.cause).toBeUndefined();
    expect(String(redacted.cause)).not.toContain(TOKEN);
    const plain = redactError(new Error("fetch failed", { cause }));
    expect(plain.cause).toBeUndefined();
    expect(String(plain.cause)).not.toContain(TOKEN);
  });
});
