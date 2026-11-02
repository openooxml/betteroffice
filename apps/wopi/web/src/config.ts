import { normalizeHttpOrigin } from "./postMessage";

/** Page configuration embedded by the server in #wopi-config. */
export interface WopiPageConfig {
  sessionHandle: string;
  fileName: string;
  format: "docx" | "xlsx" | "pptx";
  canWrite: boolean;
  viewOnlyReason: string | null;
  postMessageOrigin: string;
  closeUrl: string;
  closePostMessage: boolean;
  locale: string | null;
  userFriendlyName: string;
  tokenExpiresAt: number | null;
}

const FORMATS = new Set(["docx", "xlsx", "pptx"]);

function asString(value: unknown, fallback = ""): string {
  return typeof value === "string" ? value : fallback;
}

function asNullableString(value: unknown): string | null {
  return typeof value === "string" && value !== "" ? value : null;
}

function asHttpUrl(value: unknown): string {
  const raw = asString(value);
  if (raw === "") return "";
  try {
    const url = new URL(raw);
    return url.protocol === "http:" || url.protocol === "https:" ? url.href : "";
  } catch {
    return "";
  }
}

/** Parses and validates the embedded page configuration. */
export function parseWopiConfig(raw: unknown): WopiPageConfig {
  if (typeof raw !== "object" || raw === null) throw new Error("Missing WOPI page configuration.");
  const record = raw as Record<string, unknown>;
  const sessionHandle = asString(record.sessionHandle);
  if (sessionHandle === "") throw new Error("Missing WOPI session handle.");
  const format = asString(record.format);
  if (!FORMATS.has(format)) throw new Error(`Unsupported WOPI format: ${format || "unknown"}.`);
  return {
    sessionHandle,
    fileName: asString(record.fileName, "document"),
    format: format as WopiPageConfig["format"],
    canWrite: record.canWrite === true,
    viewOnlyReason: asNullableString(record.viewOnlyReason),
    postMessageOrigin: normalizeHttpOrigin(asString(record.postMessageOrigin)),
    closeUrl: asHttpUrl(record.closeUrl),
    closePostMessage: record.closePostMessage === true,
    locale: asNullableString(record.locale),
    userFriendlyName: asString(record.userFriendlyName).trim(),
    tokenExpiresAt: typeof record.tokenExpiresAt === "number" && Number.isFinite(record.tokenExpiresAt) ? record.tokenExpiresAt : null,
  };
}

/** Reads the configuration from a #wopi-config script element. */
export function readWopiConfig(doc: { querySelector(selector: string): { textContent: string | null } | null }): WopiPageConfig {
  const element = doc.querySelector("#wopi-config");
  if (!element || !element.textContent) throw new Error("Missing WOPI page configuration.");
  return parseWopiConfig(JSON.parse(element.textContent));
}

/** The name shown for the user, with a Guest fallback. */
export function displayUserName(userFriendlyName: string): string {
  return userFriendlyName.trim() || "Guest";
}

/** Whether the close button should be shown. */
export function showCloseButton(config: WopiPageConfig): boolean {
  return config.closePostMessage || config.postMessageOrigin !== "" || config.closeUrl !== "";
}

/** A random collaboration client id for one browser. */
export function randomClientId(): number {
  return (crypto.getRandomValues(new Uint32Array(1))[0]! & 0x7fffffff) || 1;
}
