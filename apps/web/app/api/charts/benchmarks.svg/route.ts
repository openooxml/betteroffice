import { benchmarksChartSvg, themeOf } from "../../../../lib/charts";
import { publishedFidelity } from "../../../../lib/fidelity";
import { CACHE_HIT, CACHE_MISS } from "../../../../lib/stats-cache";

export function GET(request: Request) {
  const svg = benchmarksChartSvg(publishedFidelity(), themeOf(new URL(request.url).searchParams.get("theme")));
  if (!svg) return new Response("No benchmark results", { status: 404, headers: { "Cache-Control": CACHE_MISS } });
  return new Response(svg, { headers: { "Content-Type": "image/svg+xml; charset=utf-8", "Cache-Control": CACHE_HIT } });
}
