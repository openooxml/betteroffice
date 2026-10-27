import { evalsChartSvg, evalsModel, publishedEvalTiers, themeOf } from "../../../../lib/charts";
import { CACHE_HIT, CACHE_MISS } from "../../../../lib/stats-cache";

export function GET(request: Request) {
  const section = process.env.BETTEROFFICE_AGENT_EVALS ?? "";
  const { basic, complex } = publishedEvalTiers(section);
  const svg = evalsChartSvg(basic, complex, themeOf(new URL(request.url).searchParams.get("theme")), evalsModel(section));
  if (!svg) return new Response("No eval results", { status: 404, headers: { "Cache-Control": CACHE_MISS } });
  return new Response(svg, { headers: { "Content-Type": "image/svg+xml; charset=utf-8", "Cache-Control": CACHE_HIT } });
}
