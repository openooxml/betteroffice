import { describe, expect, test } from "bun:test";
import type { WopiConfig } from "./config";
import { generateKeyPair } from "./proof";
import { createHandler, loadServerDeps, resolvePort } from "./server";

const config: WopiConfig = {
  publicUrl: "https://office.example.com",
  proofKeyFile: "data/proof-key.pem",
  maxFileBytes: 536870912,
  hosts: [
    {
      id: "nextcloud",
      wopiSrcOrigins: ["https://cloud.example.com"],
      frameOrigins: ["https://cloud.example.com"],
    },
  ],
};

function handler() {
  return createHandler({ config, currentKey: generateKeyPair(), oldKey: generateKeyPair() });
}

function request(path: string, method = "GET"): Request {
  return new Request(`https://office.example.com${path}`, { method });
}

describe("wopi server routes", () => {
  test("serves discovery as XML with security headers", async () => {
    const response = await handler()(request("/hosting/discovery"));
    expect(response.status).toBe(200);
    expect(response.headers.get("content-type")).toContain("text/xml");
    expect(response.headers.get("x-content-type-options")).toBe("nosniff");
    expect(response.headers.get("referrer-policy")).toBe("no-referrer");
    const body = await response.text();
    expect(body).toContain("<wopi-discovery>");
    expect(body).toContain("<proof-key");
  });

  test("serves capabilities as JSON", async () => {
    const response = await handler()(request("/hosting/capabilities"));
    expect(response.status).toBe(200);
    expect(response.headers.get("content-type")).toContain("application/json");
    expect(response.headers.get("x-content-type-options")).toBe("nosniff");
    expect(response.headers.get("referrer-policy")).toBe("no-referrer");
    const body = (await response.json()) as Record<string, unknown>;
    expect(body.productName).toBe("BetterOffice");
    expect(body.hasTemplateSaveAs).toBe(false);
  });

  test("serves healthz", async () => {
    const response = await handler()(request("/healthz"));
    expect(response.status).toBe(200);
    expect(response.headers.get("x-content-type-options")).toBe("nosniff");
    expect(response.headers.get("referrer-policy")).toBe("no-referrer");
    expect(await response.text()).toContain("ok");
  });

  test("resolves PORT with a 3000 default and a valid range", () => {
    expect(resolvePort({})).toBe(3000);
    expect(resolvePort({ PORT: "" })).toBe(3000);
    expect(resolvePort({ PORT: "8080" })).toBe(8080);
    expect(resolvePort({ PORT: "1" })).toBe(1);
    expect(resolvePort({ PORT: "65535" })).toBe(65535);
    for (const value of ["0", "65536", "-1", "1.5", "abc", " 3000", "3000 "]) {
      expect(() => resolvePort({ PORT: value })).toThrow(/PORT/);
    }
  });

  test("fails at startup when a configured old key file is missing", () => {
    expect(() =>
      loadServerDeps({ ...config, proofOldKeyFile: "/tmp/wopi-missing-old-key.pem" }),
    ).toThrow();
  });

  test("returns 404 for other routes and methods", async () => {
    const handle = handler();
    for (const req of [
      request("/"),
      request("/hosting/discovery", "POST"),
      request("/healthz", "POST"),
    ]) {
      const response = await handle(req);
      expect(response.status).toBe(404);
      expect(response.headers.get("x-content-type-options")).toBe("nosniff");
      expect(response.headers.get("referrer-policy")).toBe("no-referrer");
    }
  });
});
