import { describe, expect, test } from "bun:test";

import { displayUserName, parseWopiConfig, showCloseButton } from "./config";

const BASE = {
  sessionHandle: "handle-123",
  fileName: "book.xlsx",
  format: "xlsx",
  canWrite: true,
  viewOnlyReason: null,
  postMessageOrigin: "https://cloud.example.com",
  closeUrl: "",
  closePostMessage: false,
  locale: "en-US",
  userFriendlyName: "Alice A",
  tokenExpiresAt: null,
};

describe("parseWopiConfig", () => {
  test("parses a full server config", () => {
    const config = parseWopiConfig(BASE);
    expect(config.sessionHandle).toBe("handle-123");
    expect(config.format).toBe("xlsx");
    expect(config.canWrite).toBe(true);
    expect(config.locale).toBe("en-US");
    expect(config.userFriendlyName).toBe("Alice A");
  });

  test("defaults optional fields", () => {
    const config = parseWopiConfig({ sessionHandle: "h", format: "docx" });
    expect(config.canWrite).toBe(false);
    expect(config.postMessageOrigin).toBe("");
    expect(config.closeUrl).toBe("");
    expect(config.viewOnlyReason).toBeNull();
    expect(config.tokenExpiresAt).toBeNull();
    expect(config.fileName).toBe("document");
  });

  test("rejects a missing handle, bad format and non objects", () => {
    expect(() => parseWopiConfig({ format: "docx" })).toThrow();
    expect(() => parseWopiConfig({ sessionHandle: "h", format: "odt" })).toThrow();
    expect(() => parseWopiConfig(null)).toThrow();
    expect(() => parseWopiConfig("x")).toThrow();
  });

  test("normalizes the postMessage origin and rejects star and bad close urls", () => {
    expect(parseWopiConfig({ ...BASE, postMessageOrigin: "https://cloud.example.com/" }).postMessageOrigin).toBe("https://cloud.example.com");
    expect(parseWopiConfig({ ...BASE, postMessageOrigin: "https://cloud.example.com/some/path" }).postMessageOrigin).toBe("https://cloud.example.com");
    expect(parseWopiConfig({ ...BASE, postMessageOrigin: "*" }).postMessageOrigin).toBe("");
    expect(parseWopiConfig({ ...BASE, postMessageOrigin: "javascript:alert(1)" }).postMessageOrigin).toBe("");
    expect(parseWopiConfig({ ...BASE, closeUrl: "javascript:alert(1)" }).closeUrl).toBe("");
    expect(parseWopiConfig({ ...BASE, closeUrl: "/relative" }).closeUrl).toBe("");
    expect(parseWopiConfig({ ...BASE, closeUrl: "https://cloud.example.com/done" }).closeUrl).toBe("https://cloud.example.com/done");
  });

  test("trims the friendly name", () => {
    expect(parseWopiConfig({ ...BASE, userFriendlyName: "  Alice A  " }).userFriendlyName).toBe("Alice A");
    expect(parseWopiConfig({ ...BASE, userFriendlyName: "   " }).userFriendlyName).toBe("");
  });

  test("display name falls back to Guest for empty or whitespace names", () => {
    expect(displayUserName(parseWopiConfig({ ...BASE, userFriendlyName: "  Alice A  " }).userFriendlyName)).toBe("Alice A");
    expect(displayUserName(parseWopiConfig({ ...BASE, userFriendlyName: "   " }).userFriendlyName)).toBe("Guest");
    expect(displayUserName(parseWopiConfig({ ...BASE, userFriendlyName: "" }).userFriendlyName)).toBe("Guest");
    expect(displayUserName("")).toBe("Guest");
    expect(displayUserName("  Bob B ")).toBe("Bob B");
  });

  test("never keeps an access token field from a hostile config", () => {
    const config = parseWopiConfig({ ...BASE, accessToken: "SECRETTOKEN-config-1", access_token: "SECRETTOKEN-config-2" });
    expect(JSON.stringify(config)).not.toContain("SECRETTOKEN");
  });
});

describe("showCloseButton", () => {
  test("shows when the host listens or a close url exists", () => {
    expect(showCloseButton(parseWopiConfig(BASE))).toBe(true);
    expect(showCloseButton(parseWopiConfig({ ...BASE, postMessageOrigin: "", closePostMessage: true }))).toBe(true);
    expect(showCloseButton(parseWopiConfig({ ...BASE, postMessageOrigin: "", closeUrl: "https://cloud.example.com/close" }))).toBe(true);
  });

  test("hides when the host does not listen", () => {
    expect(showCloseButton(parseWopiConfig({ ...BASE, postMessageOrigin: "", closeUrl: "" }))).toBe(false);
  });
});
