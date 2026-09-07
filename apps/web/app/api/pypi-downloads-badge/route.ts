import { flatSquareBadge } from "../../../lib/badge";
import { CACHE_HIT, CACHE_MISS, perMonth, readDownloads } from "../../../lib/stats-cache";

function svg(body: string, cacheControl: string): Response {
  return new Response(body, {
    headers: { "Content-Type": "image/svg+xml; charset=utf-8", "Cache-Control": cacheControl },
  });
}

export async function GET() {
  const cached = await readDownloads("pypi");
  if (!cached) {
    return svg(
      flatSquareBadge({ label: "downloads", message: "unavailable", color: "#9f9f9f", logo: "python" }),
      CACHE_MISS,
    );
  }
  return svg(
    flatSquareBadge({ label: "downloads", message: perMonth(cached.downloads), color: "#3775A9", logo: "python" }),
    CACHE_HIT,
  );
}
