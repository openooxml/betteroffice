const CACHE_CONTROL =
  "public, max-age=300, s-maxage=86400, stale-while-revalidate=604800";
const CRATES_API = "https://crates.io/api/v1";
const OFFICIAL_REPOSITORY = "https://github.com/openooxml/betteroffice";
const USER_AGENT =
  "betteroffice.dev downloads badge (https://github.com/openooxml/betteroffice)";

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

export function rollingDownloads(
  history: DownloadHistory,
  now = new Date(),
) {
  const cutoff = new Date(now);
  cutoff.setUTCHours(0, 0, 0, 0);
  cutoff.setUTCDate(cutoff.getUTCDate() - 29);
  const cutoffDate = cutoff.toISOString().slice(0, 10);
  const entries = [
    ...history.version_downloads,
    ...history.meta.extra_downloads,
  ];

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

export async function GET() {
  try {
    const cratesResponse = await fetch(
      `${CRATES_API}/crates?page=1&per_page=100&q=betteroffice`,
      { headers: { "User-Agent": USER_AGENT } },
    );

    if (!cratesResponse.ok) {
      throw new Error(`crates.io search failed: ${cratesResponse.status}`);
    }

    const search = (await cratesResponse.json()) as { crates: CrateSummary[] };
    const crateNames = officialCrateNames(search.crates);
    if (crateNames.length === 0) {
      throw new Error("No public BetterOffice crates found");
    }

    const counts = await Promise.all(
      crateNames.map(async (crateName) => {
        const response = await fetch(
          `${CRATES_API}/crates/${encodeURIComponent(crateName)}/downloads`,
          { headers: { "User-Agent": USER_AGENT } },
        );
        if (!response.ok) {
          throw new Error(
            `crates.io downloads request failed for ${crateName}: ${response.status}`,
          );
        }
        return rollingDownloads((await response.json()) as DownloadHistory);
      }),
    );
    const downloads = counts.reduce((total, count) => total + count, 0);

    return Response.json(
      {
        schemaVersion: 1,
        label: "crates.io downloads",
        message: `${downloads.toLocaleString("en-US")}/month`,
        color: "orange",
      },
      { headers: { "Cache-Control": CACHE_CONTROL } },
    );
  } catch {
    return Response.json(
      {
        schemaVersion: 1,
        label: "crates.io downloads",
        message: "unavailable",
        color: "lightgrey",
        isError: true,
      },
      { status: 502, headers: { "Cache-Control": "no-store" } },
    );
  }
}
