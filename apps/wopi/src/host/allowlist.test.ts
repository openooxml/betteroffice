import { describe, expect, test } from "bun:test";
import type { WopiConfig } from "../config";
import { resolveHost } from "./allowlist";
import { WopiHostError } from "./errors";

function config(origins: string[]): WopiConfig {
  return {
    publicUrl: "https://office.example.com",
    proofKeyFile: "data/proof-key.pem",
    maxFileBytes: 1024,
    hosts: [{ id: "cloud", wopiSrcOrigins: origins, frameOrigins: origins }],
  };
}

function kindOf(fn: () => unknown): string {
  try {
    fn();
  } catch (err) {
    expect(err).toBeInstanceOf(WopiHostError);
    return (err as WopiHostError).kind;
  }
  throw new Error("expected resolveHost to throw");
}

describe("resolveHost", () => {
  test("returns the matching host for an allowed origin", () => {
    const host = resolveHost(config(["https://cloud.example.com"]), "https://cloud.example.com/wopi/files/1");
    expect(host.id).toBe("cloud");
  });

  test("rejects another origin", () => {
    expect(kindOf(() => resolveHost(config(["https://cloud.example.com"]), "https://evil.example.com/wopi/files/1"))).toBe("notAllowed");
  });

  test("rejects userinfo", () => {
    expect(kindOf(() => resolveHost(config(["https://cloud.example.com"]), "https://user:pass@cloud.example.com/wopi/files/1"))).toBe("notAllowed");
  });

  test("rejects a non default port that is not listed and allows a listed one", () => {
    expect(kindOf(() => resolveHost(config(["https://cloud.example.com"]), "https://cloud.example.com:8443/wopi/files/1"))).toBe("notAllowed");
    expect(resolveHost(config(["https://cloud.example.com:8443"]), "https://cloud.example.com:8443/wopi/files/1").id).toBe("cloud");
    expect(resolveHost(config(["https://cloud.example.com"]), "https://cloud.example.com:443/wopi/files/1").id).toBe("cloud");
  });

  test("requires https unless the entry origin is http", () => {
    expect(kindOf(() => resolveHost(config(["https://cloud.example.com"]), "http://cloud.example.com/wopi/files/1"))).toBe("notAllowed");
    expect(resolveHost(config(["http://cloud.example.com"]), "http://cloud.example.com/wopi/files/1").id).toBe("cloud");
  });

  test("matches IDN hosts after URL parsing", () => {
    expect(resolveHost(config(["https://münchen.example"]), "https://xn--mnchen-3ya.example/wopi/files/1").id).toBe("cloud");
    expect(resolveHost(config(["https://xn--mnchen-3ya.example"]), "https://münchen.example/wopi/files/1").id).toBe("cloud");
  });

  test("rejects an invalid URL", () => {
    expect(kindOf(() => resolveHost(config(["https://cloud.example.com"]), "not a url"))).toBe("notAllowed");
  });
});
