const OFFICIAL_REPOSITORY = "https://github.com/openooxml/betteroffice";

interface CrateSummary {
  name: string;
  repository: string | null;
}

interface DownloadEntry {
  date: string;
  downloads: number;
}

interface DownloadHistory {
  version_downloads: DownloadEntry[];
  meta: { extra_downloads: DownloadEntry[] };
}

function normalizeRepository(repository: string | null) {
  return repository?.replace(/\.git$/, "").replace(/\/$/, "");
}

export function officialCrateNames(crates: CrateSummary[]) {
  return crates
    .filter(
      (crate) =>
        crate.name.startsWith("betteroffice-") &&
        normalizeRepository(crate.repository) === OFFICIAL_REPOSITORY,
    )
    .map((crate) => crate.name);
}

export function rollingDownloads(history: DownloadHistory, now = new Date()) {
  const cutoff = new Date(now);
  cutoff.setUTCHours(0, 0, 0, 0);
  cutoff.setUTCDate(cutoff.getUTCDate() - 29);
  const cutoffDate = cutoff.toISOString().slice(0, 10);
  const entries = [...history.version_downloads, ...history.meta.extra_downloads];

  return entries.reduce((total, entry) => {
    if (
      !/^\d{4}-\d{2}-\d{2}$/.test(entry.date) ||
      !Number.isSafeInteger(entry.downloads) ||
      entry.downloads < 0
    ) {
      throw new Error("Invalid crates.io download history");
    }
    return entry.date >= cutoffDate ? total + entry.downloads : total;
  }, 0);
}
