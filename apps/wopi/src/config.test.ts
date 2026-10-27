import { describe, expect, test } from "bun:test";
import { loadConfig } from "./config";

const hosts = JSON.stringify([
  {
    id: "nextcloud",
    wopiSrcOrigins: ["https://cloud.example.com"],
    frameOrigins: ["https://cloud.example.com"],
  },
]);

function env(overrides: Record<string, string | undefined> = {}): Record<string, string | undefined> {
  return {
    WOPI_PUBLIC_URL: "https://office.example.com",
    WOPI_HOSTS: hosts,
    ...overrides,
  };
}

describe("loadConfig", () => {
  test("parses a valid env and applies defaults", () => {
    const config = loadConfig(env());
    expect(config.publicUrl).toBe("https://office.example.com");
    expect(config.maxFileBytes).toBe(536870912);
    expect(config.hosts).toEqual([
      {
        id: "nextcloud",
        wopiSrcOrigins: ["https://cloud.example.com"],
        frameOrigins: ["https://cloud.example.com"],
      },
    ]);
    expect(config.proofKeyFile).toBe("data/proof-key.pem");
    expect(config.proofOldKeyFile).toBeUndefined();
  });

  test("strips a trailing slash from the public URL", () => {
    expect(loadConfig(env({ WOPI_PUBLIC_URL: "https://office.example.com/" })).publicUrl).toBe(
      "https://office.example.com",
    );
  });

  test("requires WOPI_PUBLIC_URL", () => {
    expect(() => loadConfig(env({ WOPI_PUBLIC_URL: undefined }))).toThrow(/WOPI_PUBLIC_URL/);
    expect(() => loadConfig(env({ WOPI_PUBLIC_URL: "" }))).toThrow(/WOPI_PUBLIC_URL/);
    expect(() => loadConfig(env({ WOPI_PUBLIC_URL: "not a url" }))).toThrow(/WOPI_PUBLIC_URL/);
  });

  test("requires https unless WOPI_ALLOW_HTTP is 1", () => {
    expect(() => loadConfig(env({ WOPI_PUBLIC_URL: "http://office.example.com" }))).toThrow(/https/);
    expect(
      loadConfig(env({ WOPI_PUBLIC_URL: "http://office.example.com", WOPI_ALLOW_HTTP: "1" })).publicUrl,
    ).toBe("http://office.example.com");
    expect(() =>
      loadConfig(env({ WOPI_PUBLIC_URL: "http://office.example.com", WOPI_ALLOW_HTTP: "true" })),
    ).toThrow(/https/);
  });

  test("passes proof key files through and parses a custom size cap", () => {
    const config = loadConfig(
      env({
        WOPI_PROOF_KEY_FILE: "/tmp/current.pem",
        WOPI_PROOF_OLD_KEY_FILE: "/tmp/old.pem",
        WOPI_MAX_FILE_BYTES: "1048576",
      }),
    );
    expect(config.proofKeyFile).toBe("/tmp/current.pem");
    expect(config.proofOldKeyFile).toBe("/tmp/old.pem");
    expect(config.maxFileBytes).toBe(1048576);
  });

  test("rejects an invalid size cap", () => {
    for (const value of ["0", "-1", "1.5", "abc", ""]) {
      expect(() => loadConfig(env({ WOPI_MAX_FILE_BYTES: value }))).toThrow(/WOPI_MAX_FILE_BYTES/);
    }
  });

  test("requires at least one valid host entry", () => {
    expect(() => loadConfig(env({ WOPI_HOSTS: undefined }))).toThrow(/WOPI_HOSTS/);
    expect(() => loadConfig(env({ WOPI_HOSTS: "not json" }))).toThrow(/WOPI_HOSTS/);
    expect(() => loadConfig(env({ WOPI_HOSTS: "{}" }))).toThrow(/WOPI_HOSTS/);
    expect(() => loadConfig(env({ WOPI_HOSTS: "[]" }))).toThrow(/WOPI_HOSTS/);
    expect(() => loadConfig(env({ WOPI_HOSTS: JSON.stringify([{ id: "x" }]) }))).toThrow(/WOPI_HOSTS/);
    expect(
      () =>
        loadConfig(
          env({
            WOPI_HOSTS: JSON.stringify([
              { id: "", wopiSrcOrigins: ["https://a.example.com"], frameOrigins: [] },
            ]),
          }),
        ),
    ).toThrow(/WOPI_HOSTS/);
    expect(
      () =>
        loadConfig(
          env({
            WOPI_HOSTS: JSON.stringify([
              { id: "x", wopiSrcOrigins: [], frameOrigins: ["https://a.example.com"] },
            ]),
          }),
        ),
    ).toThrow(/WOPI_HOSTS/);
    expect(
      () =>
        loadConfig(
          env({
            WOPI_HOSTS: JSON.stringify([
              { id: "x", wopiSrcOrigins: ["https://a.example.com"], frameOrigins: "https://a.example.com" },
            ]),
          }),
        ),
    ).toThrow(/WOPI_HOSTS/);
  });
});
