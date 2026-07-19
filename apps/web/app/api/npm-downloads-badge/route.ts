import { getCloudflareContext } from "@opennextjs/cloudflare";
import { flatSquareBadge } from "../../../lib/badge";
import { monthlyDownloadsTotal } from "../../../lib/downloads";

const KV_KEY = "npm-downloads";
const REFRESH_MS = 24 * 60 * 60 * 1000;
const NPM_RED = "#CB3837";

interface CachedCount {
  downloads: number;
  at: number;
}

function svg(body: string, cacheSeconds: number): Response {
  return new Response(body, {
    headers: {
      "Content-Type": "image/svg+xml; charset=utf-8",
      "Cache-Control": `public, max-age=${cacheSeconds}, s-maxage=${cacheSeconds}`,
    },
  });
}

function downloadsBadge(downloads: number): string {
  return flatSquareBadge({
    label: "downloads",
    message: `${downloads.toLocaleString("en-US")}/month`,
    color: NPM_RED,
    logo: "npm",
  });
}

export async function GET() {
  let kv: KVNamespace | undefined;
  try {
    kv = getCloudflareContext().env.STATS_KV;
  } catch {
    kv = undefined;
  }

  const cached = kv ? await kv.get<CachedCount>(KV_KEY, "json").catch(() => null) : null;
  if (cached && Date.now() - cached.at < REFRESH_MS) {
    return svg(downloadsBadge(cached.downloads), 3600);
  }

  try {
    const downloads = await monthlyDownloadsTotal();
    if (downloads === null) throw new Error("no download counts resolved");
    if (kv) {
      await kv
        .put(KV_KEY, JSON.stringify({ downloads, at: Date.now() } satisfies CachedCount))
        .catch(() => {});
    }
    return svg(downloadsBadge(downloads), 3600);
  } catch {
    if (cached) return svg(downloadsBadge(cached.downloads), 3600);
    return svg(
      flatSquareBadge({ label: "downloads", message: "unavailable", color: "#9f9f9f" }),
      60,
    );
  }
}
