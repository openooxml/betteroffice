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

  test("rejects a public URL with query, fragment or credentials, keeps a path prefix", () => {
    expect(() => loadConfig(env({ WOPI_PUBLIC_URL: "https://office.example.com/?x=1" }))).toThrow(/WOPI_PUBLIC_URL/);
    expect(() => loadConfig(env({ WOPI_PUBLIC_URL: "https://office.example.com/#frag" }))).toThrow(/WOPI_PUBLIC_URL/);
    expect(() => loadConfig(env({ WOPI_PUBLIC_URL: "https://user:pass@office.example.com" }))).toThrow(/WOPI_PUBLIC_URL/);
    expect(() => loadConfig(env({ WOPI_PUBLIC_URL: "https://user@office.example.com" }))).toThrow(/WOPI_PUBLIC_URL/);
    expect(loadConfig(env({ WOPI_PUBLIC_URL: "https://office.example.com/prefix/" })).publicUrl).toBe(
      "https://office.example.com/prefix",
    );
  });

  test("normalizes host origins and strips a trailing slash", () => {
    const config = loadConfig(
      env({
        WOPI_HOSTS: JSON.stringify([
          {
            id: "a",
            wopiSrcOrigins: ["https://cloud.example.com/"],
            frameOrigins: ["https://cloud.example.com/"],
          },
        ]),
      }),
    );
    expect(config.hosts[0]!.wopiSrcOrigins).toEqual(["https://cloud.example.com"]);
    expect(config.hosts[0]!.frameOrigins).toEqual(["https://cloud.example.com"]);
  });

  test("rejects invalid host origins", () => {
    const bad = [
      "not a url",
      "",
      "   ",
      "https://cloud.example.com/path",
      "https://cloud.example.com/?x=1",
      "https://cloud.example.com/#frag",
      "https://user:pass@cloud.example.com",
      "https://*.example.com",
      "https://cloud.example.com ",
      " https://cloud.example.com",
      "https://cloud.example.com;https://evil.example.com",
      "https://cloud.example.com https://evil.example.com",
      "https://cloud.example.com';alert(1);//",
      "//cloud.example.com",
      "cloud.example.com",
    ];
    for (const origin of bad) {
      expect(() =>
        loadConfig(
          env({
            WOPI_HOSTS: JSON.stringify([
              { id: "a", wopiSrcOrigins: [origin], frameOrigins: ["https://cloud.example.com"] },
            ]),
          }),
        ),
      ).toThrow(/WOPI_HOSTS/);
      expect(() =>
        loadConfig(
          env({
            WOPI_HOSTS: JSON.stringify([
              { id: "a", wopiSrcOrigins: ["https://cloud.example.com"], frameOrigins: [origin] },
            ]),
          }),
        ),
      ).toThrow(/WOPI_HOSTS/);
    }
  });

  test("rejects http host origins unless WOPI_ALLOW_HTTP is 1", () => {
    const hostsWithHttp = JSON.stringify([
      { id: "a", wopiSrcOrigins: ["http://cloud.example.com"], frameOrigins: ["http://cloud.example.com"] },
    ]);
    expect(() => loadConfig(env({ WOPI_HOSTS: hostsWithHttp }))).toThrow(/WOPI_HOSTS/);
    expect(loadConfig(env({ WOPI_HOSTS: hostsWithHttp, WOPI_ALLOW_HTTP: "1" })).hosts[0]!.wopiSrcOrigins).toEqual([
      "http://cloud.example.com",
    ]);
  });

  test("rejects duplicate host ids and a wopiSrc origin shared by two hosts", () => {
    expect(() =>
      loadConfig(
        env({
          WOPI_HOSTS: JSON.stringify([
            { id: "a", wopiSrcOrigins: ["https://one.example.com"], frameOrigins: [] },
            { id: "a", wopiSrcOrigins: ["https://two.example.com"], frameOrigins: [] },
          ]),
        }),
      ),
    ).toThrow(/WOPI_HOSTS/);
    expect(() =>
      loadConfig(
        env({
          WOPI_HOSTS: JSON.stringify([
            { id: "a", wopiSrcOrigins: ["https://cloud.example.com"], frameOrigins: [] },
            { id: "b", wopiSrcOrigins: ["https://cloud.example.com/"], frameOrigins: [] },
          ]),
        }),
      ),
    ).toThrow(/WOPI_HOSTS/);
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
