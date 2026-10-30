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
  journalDir?: string;
  instanceId?: string;
  maxEditorsPerFile?: number;
}

export const DEFAULT_MAX_FILE_BYTES = 536870912;
export const DEFAULT_PROOF_KEY_FILE = "data/proof-key.pem";

function fail(message: string): never {
  throw new Error(message);
}

function parsePublicUrl(env: Record<string, string | undefined>): string {
  const raw = env.WOPI_PUBLIC_URL?.trim();
  if (!raw) fail("WOPI_PUBLIC_URL is required");
  if (raw.includes("?") || raw.includes("#")) {
    fail("WOPI_PUBLIC_URL must not contain a query, fragment, username or password");
  }
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    fail("WOPI_PUBLIC_URL must be a valid URL");
  }
  if (url.protocol !== "https:" && !(url.protocol === "http:" && env.WOPI_ALLOW_HTTP === "1")) {
    fail("WOPI_PUBLIC_URL must use https unless WOPI_ALLOW_HTTP=1");
  }
  if (url.search || url.hash || url.username || url.password) {
    fail("WOPI_PUBLIC_URL must not contain a query, fragment, username or password");
  }
  return raw.replace(/\/+$/, "");
}

function parseMaxFileBytes(env: Record<string, string | undefined>): number {
  const raw = env.WOPI_MAX_FILE_BYTES;
  if (raw === undefined) return DEFAULT_MAX_FILE_BYTES;
  if (!/^\d+$/.test(raw)) fail("WOPI_MAX_FILE_BYTES must be a positive integer");
  const value = Number(raw);
  if (!Number.isSafeInteger(value) || value <= 0) fail("WOPI_MAX_FILE_BYTES must be a positive integer");
  return value;
}

function parseOrigin(entry: unknown, field: string, index: number, allowHttp: boolean): string {
  if (typeof entry !== "string" || entry.length === 0 || /[\s;*'"`,<>]/.test(entry)) {
    fail(`WOPI_HOSTS entry ${index} has an invalid ${field} origin`);
  }
  let url: URL;
  try {
    url = new URL(entry);
  } catch {
    fail(`WOPI_HOSTS entry ${index} has an invalid ${field} origin`);
  }
  const protocolOk = url.protocol === "https:" || (allowHttp && url.protocol === "http:");
  if (
    !protocolOk ||
    url.username ||
    url.password ||
    url.search ||
    url.hash ||
    url.pathname !== "/" ||
    url.origin === "null"
  ) {
    fail(`WOPI_HOSTS entry ${index} has an invalid ${field} origin`);
  }
  return url.origin;
}

function parseOriginList(
  value: unknown,
  field: string,
  index: number,
  allowHttp: boolean,
  nonEmpty: boolean,
): string[] {
  if (!Array.isArray(value) || (nonEmpty && value.length === 0)) {
    fail(`WOPI_HOSTS entry ${index} must have a ${field} array`);
  }
  return value.map((entry) => parseOrigin(entry, field, index, allowHttp));
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
  const allowHttp = env.WOPI_ALLOW_HTTP === "1";
  const hosts = parsed.map((entry, index) => {
    if (typeof entry !== "object" || entry === null) {
      fail(`WOPI_HOSTS entry ${index} must be an object`);
    }
    const host = entry as Record<string, unknown>;
    if (typeof host.id !== "string" || host.id.length === 0 || host.id.trim() !== host.id) {
      fail(`WOPI_HOSTS entry ${index} must have a non-empty id`);
    }
    return {
      id: host.id,
      wopiSrcOrigins: parseOriginList(host.wopiSrcOrigins, "wopiSrcOrigins", index, allowHttp, true),
      frameOrigins: parseOriginList(host.frameOrigins, "frameOrigins", index, allowHttp, false),
    };
  });
  const ids = new Set<string>();
  for (const host of hosts) {
    if (ids.has(host.id)) fail(`WOPI_HOSTS has a duplicate host id ${host.id}`);
    ids.add(host.id);
  }
  const wopiSrcOwners = new Map<string, string>();
  for (const host of hosts) {
    for (const origin of host.wopiSrcOrigins) {
      const owner = wopiSrcOwners.get(origin);
      if (owner !== undefined && owner !== host.id) {
        fail(`WOPI_HOSTS wopiSrcOrigins origin ${origin} is listed under two hosts`);
      }
      wopiSrcOwners.set(origin, host.id);
    }
  }
  return hosts;
}

function optionalFile(env: Record<string, string | undefined>, name: string): string | undefined {
  const value = env[name]?.trim();
  return value ? value : undefined;
}

function parseMaxEditors(env: Record<string, string | undefined>): number {
  const raw = env.WOPI_MAX_EDITORS_PER_FILE;
  if (raw === undefined) return 50;
  if (!/^\d+$/.test(raw)) fail("WOPI_MAX_EDITORS_PER_FILE must be a positive integer");
  const value = Number(raw);
  if (!Number.isSafeInteger(value) || value <= 0) fail("WOPI_MAX_EDITORS_PER_FILE must be a positive integer");
  return value;
}

export function loadConfig(env: Record<string, string | undefined> = process.env): WopiConfig {
  return {
    publicUrl: parsePublicUrl(env),
    proofKeyFile: optionalFile(env, "WOPI_PROOF_KEY_FILE") ?? DEFAULT_PROOF_KEY_FILE,
    proofOldKeyFile: optionalFile(env, "WOPI_PROOF_OLD_KEY_FILE"),
    maxFileBytes: parseMaxFileBytes(env),
    hosts: parseHosts(env),
    journalDir: optionalFile(env, "WOPI_JOURNAL_DIR"),
    instanceId: optionalFile(env, "WOPI_INSTANCE_ID"),
    maxEditorsPerFile: parseMaxEditors(env),
  };
}
