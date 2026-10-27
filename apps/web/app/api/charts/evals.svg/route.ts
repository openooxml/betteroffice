import { evalsChartSvg, publishedEvalTiers, themeOf } from "../../../../lib/charts";
import { CACHE_HIT, CACHE_MISS } from "../../../../lib/stats-cache";

export function GET(request: Request) {
  const { basic, complex } = publishedEvalTiers(process.env.BETTEROFFICE_AGENT_EVALS ?? "");
  const svg = evalsChartSvg(basic, complex, themeOf(new URL(request.url).searchParams.get("theme")));
  if (!svg) return new Response("No eval results", { status: 404, headers: { "Cache-Control": CACHE_MISS } });
  return new Response(svg, { headers: { "Content-Type": "image/svg+xml; charset=utf-8", "Cache-Control": CACHE_HIT } });
}
