import { CACHE_HIT, CACHE_MISS, perMonth, readDownloads } from "../../../lib/stats-cache";

export async function GET() {
  const cached = await readDownloads("pypi");
  if (!cached) {
    return Response.json(
      { schemaVersion: 1, label: "PyPI downloads", message: "unavailable", color: "lightgrey", isError: true },
      { headers: { "Cache-Control": CACHE_MISS } },
    );
  }
  return Response.json(
    { schemaVersion: 1, label: "PyPI downloads", message: perMonth(cached.downloads), color: "blue" },
    { headers: { "Cache-Control": CACHE_HIT } },
  );
}
