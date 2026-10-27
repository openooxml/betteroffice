import { loadConfig, type WopiConfig } from "./config";
import { buildDiscovery, capabilities } from "./discovery";
import {
  loadKeyPairFile,
  loadOrCreateKeyPairFile,
  type ProofKeyPair,
} from "./proof";

export interface WopiServerDeps {
  config: WopiConfig;
  currentKey: ProofKeyPair;
  oldKey?: ProofKeyPair;
  productVersion?: string;
}

const SECURITY_HEADERS = {
  "X-Content-Type-Options": "nosniff",
  "Referrer-Policy": "no-referrer",
} as const;

function respond(body: string, contentType: string, status = 200): Response {
  return new Response(body, {
    status,
    headers: { "Content-Type": contentType, ...SECURITY_HEADERS },
  });
}

export function resolvePort(env: Record<string, string | undefined> = process.env): number {
  const raw = env.PORT;
  if (raw === undefined || raw === "") return 3000;
  if (!/^\d+$/.test(raw)) throw new Error("PORT must be an integer between 1 and 65535");
  const value = Number(raw);
  if (value < 1 || value > 65535) throw new Error("PORT must be an integer between 1 and 65535");
  return value;
}

/** Request handler for the WOPI hosting endpoints. */
export function createHandler(deps: WopiServerDeps): (request: Request) => Promise<Response> {
  const discoveryXml = buildDiscovery({
    publicUrl: deps.config.publicUrl,
    currentKey: deps.currentKey,
    oldKey: deps.oldKey,
  });
  const capabilitiesJson = JSON.stringify(capabilities(deps.productVersion));
  const healthJson = JSON.stringify({ status: "ok" });
  return async (request: Request): Promise<Response> => {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname === "/hosting/discovery") {
      return respond(discoveryXml, "text/xml; charset=utf-8");
    }
    if (request.method === "GET" && url.pathname === "/hosting/capabilities") {
      return respond(capabilitiesJson, "application/json; charset=utf-8");
    }
    if (request.method === "GET" && url.pathname === "/healthz") {
      return respond(healthJson, "application/json; charset=utf-8");
    }
    return respond("Not Found", "text/plain; charset=utf-8", 404);
  };
}

export function loadServerDeps(config: WopiConfig): WopiServerDeps {
  const oldKey = config.proofOldKeyFile ? loadKeyPairFile(config.proofOldKeyFile) : undefined;
  const currentKey = loadOrCreateKeyPairFile(config.proofKeyFile);
  return { config, currentKey, oldKey };
}

export function startServer(deps: WopiServerDeps, port = resolvePort()) {
  const handler = createHandler(deps);
  return Bun.serve({ port, fetch: handler });
}

if ((import.meta as { main?: boolean }).main) {
  const config = loadConfig();
  const server = startServer(loadServerDeps(config));
  console.log(`WOPI app listening on ${server.url}`);
}
