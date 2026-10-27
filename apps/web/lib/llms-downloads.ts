import type { CachedCount, Registry } from "./download-stats";
import { readDownloads } from "./stats-cache";

const NAMES: Record<Registry, string> = { npm: "npm", crates: "crates.io", pypi: "PyPI" };
const ORDER: Registry[] = ["npm", "crates", "pypi"];

export type DownloadCounts = Partial<Record<Registry, CachedCount | null>>;

/** One line with the cached 30-day totals and the date of the newest refresh; empty when nothing is cached. */
export function downloadsLine(counts: DownloadCounts): string {
  const present = ORDER.filter((registry) => counts[registry]);
  if (present.length === 0) return "";
  const newest = Math.max(...present.map((registry) => counts[registry]!.at));
  const parts = present.map((registry) => `${NAMES[registry]} ${counts[registry]!.downloads.toLocaleString("en-US")}`);
  const date = Number.isFinite(newest) ? new Date(newest).toISOString().slice(0, 10) : "";
  return `Downloads in the last 30 days: ${parts.join(", ")}${date ? ` (as of ${date})` : ""}.`;
}

/** Puts the line before the first section, after the introduction. */
export function withDownloads(body: string, line: string): string {
  if (!line) return body;
  const firstSection = body.indexOf("\n## ");
  if (firstSection < 0) return `${body.trimEnd()}\n\n${line}\n`;
  return `${body.slice(0, firstSection).trimEnd()}\n\n${line}\n${body.slice(firstSection)}`;
}

export async function liveDownloadsLine(read: (registry: Registry) => Promise<CachedCount | null> = readDownloads): Promise<string> {
  const [npm, crates, pypi] = await Promise.all(ORDER.map((registry) => read(registry)));
  return downloadsLine({ npm, crates, pypi });
}
