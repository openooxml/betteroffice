import type { WopiConfig, WopiHostConfig } from "../config";
import { WopiHostError } from "./errors";
import { redactUrl } from "./redact";

type HostsInput = WopiConfig | WopiHostConfig[] | { hosts: WopiHostConfig[] };

function hostsOf(config: HostsInput): WopiHostConfig[] {
  return Array.isArray(config) ? config : config.hosts;
}

function notAllowed(wopiSrc: string): WopiHostError {
  return new WopiHostError("notAllowed", `WOPISrc origin is not allowed: ${redactUrl(wopiSrc)}`);
}

/** Returns the allowlisted host for a WOPISrc, or throws notAllowed. */
export function resolveHost(config: HostsInput, wopiSrc: string): WopiHostConfig {
  let url: URL;
  try {
    url = new URL(wopiSrc);
  } catch {
    throw notAllowed(wopiSrc);
  }
  if (url.hash !== "") throw notAllowed(wopiSrc);
  if (url.username !== "" || url.password !== "") throw notAllowed(wopiSrc);
  if (url.protocol !== "https:" && url.protocol !== "http:") throw notAllowed(wopiSrc);
  for (const host of hostsOf(config)) {
    for (const origin of host.wopiSrcOrigins) {
      let allowed: URL;
      try {
        allowed = new URL(origin);
      } catch {
        continue;
      }
      if (url.protocol === "http:" && allowed.protocol !== "http:") continue;
      if (url.origin === allowed.origin) return host;
    }
  }
  throw notAllowed(wopiSrc);
}
