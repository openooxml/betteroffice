import { CACHE_HIT, CACHE_MISS, perMonth, readDownloads } from "../../../lib/stats-cache";

export async function GET() {
  const cached = await readDownloads("npm");
  if (!cached) {
    return Response.json(
      { schemaVersion: 1, label: "npm downloads", message: "unavailable", color: "lightgrey", isError: true },
      { headers: { "Cache-Control": CACHE_MISS } },
    );
  }
  return Response.json(
    { schemaVersion: 1, label: "npm downloads", message: perMonth(cached.downloads), color: "red" },
    { headers: { "Cache-Control": CACHE_HIT } },
  );
}
