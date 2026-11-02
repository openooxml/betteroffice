import { describe, expect, test } from "bun:test";
import { UserSession } from "./userSession";

describe("UserSession", () => {
  test("id is 256 bit base64url and token is omitted from JSON", () => {
    const a = new UserSession({ hostId: "h", wopiSrc: "https://cloud.example.com/wopi/files/f1", accessToken: "secret-token-value", accessTokenTtl: 0, userId: "u1", friendlyName: "User One", canWrite: true, joinedAt: 1000 });
    const b = new UserSession({ hostId: "h", wopiSrc: "https://cloud.example.com/wopi/files/f1", accessToken: "other", accessTokenTtl: 0, userId: "u2", friendlyName: "User Two", canWrite: false, joinedAt: 1000 });
    expect(a.id).toMatch(/^[A-Za-z0-9_-]{43}$/);
    expect(a.id).not.toBe(b.id);
    expect(Buffer.from(a.id, "base64url").byteLength).toBe(32);
    expect(a.expiresAt).toBeNull();
    expect(a.connected).toBe(true);
    expect(a.lastSeen).toBe(1000);
    const json = JSON.stringify(a);
    expect(json).not.toContain("secret-token-value");
    expect(json).not.toContain("accessToken");
    expect(JSON.parse(json).accessToken).toBeUndefined();
    expect(String(a)).not.toContain("secret-token-value");
  });

  test("expiresAt comes from the absolute ttl", () => {
    const u = new UserSession({ hostId: "h", wopiSrc: "s", accessToken: "t", accessTokenTtl: 1700003600000, userId: "u", friendlyName: "U", canWrite: true, joinedAt: 5 });
    expect(u.expiresAt).toBe(1700003600000);
  });
});
