import { getCloudflareContext } from "@opennextjs/cloudflare";
import { type CachedCount, type Registry, cachedDownloads } from "./download-stats";

/** The badge routes only read what the scheduled refresh wrote. */
export async function readDownloads(registry: Registry): Promise<CachedCount | null> {
  try {
    return await cachedDownloads(getCloudflareContext().env.STATS_KV, registry);
  } catch {
    return null;
  }
}

export const CACHE_HIT = "public, max-age=3600, s-maxage=3600";
export const CACHE_MISS = "public, max-age=60";

export function perMonth(downloads: number): string {
  return `${downloads.toLocaleString("en-US")}/month`;
}
