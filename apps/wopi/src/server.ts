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

/** Request handler for the WOPI hosting endpoints. */
export function createHandler(deps: WopiServerDeps): (request: Request) => Promise<Response> {
  return async (request: Request): Promise<Response> => {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname === "/hosting/discovery") {
      const xml = buildDiscovery({
        publicUrl: deps.config.publicUrl,
        currentKey: deps.currentKey,
        oldKey: deps.oldKey,
      });
      return respond(xml, "text/xml; charset=utf-8");
    }
    if (request.method === "GET" && url.pathname === "/hosting/capabilities") {
      return respond(
        JSON.stringify(capabilities(deps.productVersion)),
        "application/json; charset=utf-8",
      );
    }
    if (request.method === "GET" && url.pathname === "/healthz") {
      return respond(JSON.stringify({ status: "ok" }), "application/json; charset=utf-8");
    }
    return respond("Not Found", "text/plain; charset=utf-8", 404);
  };
}

export function loadServerDeps(config: WopiConfig): WopiServerDeps {
  const oldKey = config.proofOldKeyFile ? loadKeyPairFile(config.proofOldKeyFile) : undefined;
  const currentKey = loadOrCreateKeyPairFile(config.proofKeyFile);
  return { config, currentKey, oldKey };
}

export function startServer(deps: WopiServerDeps, port = Number(process.env.PORT ?? 3000)) {
  const handler = createHandler(deps);
  return Bun.serve({ port, fetch: handler });
}

if ((import.meta as { main?: boolean }).main) {
  const config = loadConfig();
  const server = startServer(loadServerDeps(config));
  console.log(`WOPI app listening on ${server.url}`);
}
