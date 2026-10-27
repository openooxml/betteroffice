/** Typed WOPI configuration loaded from the environment. */
export interface WopiHostConfig {
  id: string;
  wopiSrcOrigins: string[];
  frameOrigins: string[];
}

export interface WopiConfig {
  publicUrl: string;
  proofKeyFile: string;
  proofOldKeyFile?: string;
  maxFileBytes: number;
  hosts: WopiHostConfig[];
}

export const DEFAULT_MAX_FILE_BYTES = 536870912;
export const DEFAULT_PROOF_KEY_FILE = "data/proof-key.pem";

function fail(message: string): never {
  throw new Error(message);
}

function parsePublicUrl(env: Record<string, string | undefined>): string {
  const raw = env.WOPI_PUBLIC_URL?.trim();
  if (!raw) fail("WOPI_PUBLIC_URL is required");
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    fail("WOPI_PUBLIC_URL must be a valid URL");
  }
  if (url.protocol !== "https:" && !(url.protocol === "http:" && env.WOPI_ALLOW_HTTP === "1")) {
    fail("WOPI_PUBLIC_URL must use https unless WOPI_ALLOW_HTTP=1");
  }
  return raw.replace(/\/+$/, "");
}

function parseMaxFileBytes(env: Record<string, string | undefined>): number {
  const raw = env.WOPI_MAX_FILE_BYTES?.trim();
  if (raw === undefined || raw === "") {
    if (raw === "") fail("WOPI_MAX_FILE_BYTES must be a positive integer");
    return DEFAULT_MAX_FILE_BYTES;
  }
  if (!/^\d+$/.test(raw)) fail("WOPI_MAX_FILE_BYTES must be a positive integer");
  const value = Number(raw);
  if (!Number.isSafeInteger(value) || value <= 0) fail("WOPI_MAX_FILE_BYTES must be a positive integer");
  return value;
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === "string" && entry.length > 0);
}

function parseHosts(env: Record<string, string | undefined>): WopiHostConfig[] {
  const raw = env.WOPI_HOSTS?.trim();
  if (!raw) fail("WOPI_HOSTS is required");
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    fail("WOPI_HOSTS must be a JSON array of host entries");
  }
  if (!Array.isArray(parsed) || parsed.length === 0) {
    fail("WOPI_HOSTS must be a JSON array with at least one host entry");
  }
  return parsed.map((entry, index) => {
    if (typeof entry !== "object" || entry === null) {
      fail(`WOPI_HOSTS entry ${index} must be an object`);
    }
    const host = entry as Record<string, unknown>;
    if (typeof host.id !== "string" || host.id.length === 0) {
      fail(`WOPI_HOSTS entry ${index} must have a non-empty id`);
    }
    if (!isStringArray(host.wopiSrcOrigins) || host.wopiSrcOrigins.length === 0) {
      fail(`WOPI_HOSTS entry ${index} must have a non-empty wopiSrcOrigins array`);
    }
    if (!isStringArray(host.frameOrigins)) {
      fail(`WOPI_HOSTS entry ${index} must have a frameOrigins array`);
    }
    return {
      id: host.id,
      wopiSrcOrigins: [...host.wopiSrcOrigins],
      frameOrigins: [...host.frameOrigins],
    };
  });
}

function optionalFile(env: Record<string, string | undefined>, name: string): string | undefined {
  const value = env[name]?.trim();
  return value ? value : undefined;
}

export function loadConfig(env: Record<string, string | undefined> = process.env): WopiConfig {
  return {
    publicUrl: parsePublicUrl(env),
    proofKeyFile: optionalFile(env, "WOPI_PROOF_KEY_FILE") ?? DEFAULT_PROOF_KEY_FILE,
    proofOldKeyFile: optionalFile(env, "WOPI_PROOF_OLD_KEY_FILE"),
    maxFileBytes: parseMaxFileBytes(env),
    hosts: parseHosts(env),
  };
}
