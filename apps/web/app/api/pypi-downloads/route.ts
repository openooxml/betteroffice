import { getCloudflareContext } from "@opennextjs/cloudflare";
import { monthlyDownloadsTotal } from "../../../lib/pypi-downloads";

const KV_KEY = "pypi-downloads";
const REFRESH_MS = 24 * 60 * 60 * 1000;

interface CachedCount {
  downloads: number;
  at: number;
}

function payload(downloads: number, cacheSeconds: number): Response {
  return Response.json(
    {
      schemaVersion: 1,
      label: "PyPI downloads",
      message: `${downloads.toLocaleString("en-US")}/month`,
      color: "blue",
    },
    { headers: { "Cache-Control": `public, max-age=${cacheSeconds}, s-maxage=${cacheSeconds}` } },
  );
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
    return payload(cached.downloads, 3600);
  }

  try {
    const downloads = await monthlyDownloadsTotal();
    if (downloads === null) throw new Error("no download counts resolved");
    if (kv) {
      await kv
        .put(KV_KEY, JSON.stringify({ downloads, at: Date.now() } satisfies CachedCount))
        .catch(() => {});
    }
    return payload(downloads, 3600);
  } catch {
    if (cached) return payload(cached.downloads, 3600);
    return Response.json(
      { schemaVersion: 1, label: "PyPI downloads", message: "unavailable", color: "lightgrey", isError: true },
      { headers: { "Cache-Control": "public, max-age=60" } },
    );
  }
}
