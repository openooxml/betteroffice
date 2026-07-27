const PYPISTATS_API = "https://pypistats.org/api/packages";
const USER_AGENT =
  "betteroffice.dev downloads badge (https://github.com/openooxml/betteroffice)";

/**
 * PyPI has no org listing endpoint, so the distributions are named here. Add a
 * line when a new one ships.
 */
export const PYPI_PACKAGES = ["betteroffice-xlsx"];

interface RecentDownloads {
  data?: { last_month?: number };
}

/**
 * Sum last-month downloads, tolerating per-package failures: an unpublished or
 * freshly published distribution 404s until pypistats has data for it.
 * `resolved` counts the packages that produced a value.
 */
export async function monthlyDownloads(
  packageNames: string[],
  fetchImpl: typeof fetch = fetch,
): Promise<{ downloads: number; resolved: number }> {
  const fetchCount = async (packageName: string): Promise<number> => {
    const response = await fetchImpl(
      `${PYPISTATS_API}/${encodeURIComponent(packageName)}/recent`,
      { headers: { "User-Agent": USER_AGENT } },
    );
    if (!response.ok) {
      throw new Error(
        `pypistats request failed for ${packageName}: ${response.status}`,
      );
    }
    const body = (await response.json()) as RecentDownloads;
    const downloads = body.data?.last_month;
    if (!Number.isSafeInteger(downloads) || (downloads as number) < 0) {
      throw new Error(`Invalid PyPI download count for ${packageName}`);
    }
    return downloads as number;
  };

  const results = await Promise.allSettled(packageNames.map(fetchCount));

  let downloads = 0;
  let resolved = 0;
  for (const result of results) {
    if (result.status === "fulfilled") {
      downloads += result.value;
      resolved += 1;
    }
  }
  return { downloads, resolved };
}

export async function monthlyDownloadsTotal(
  fetchImpl: typeof fetch = fetch,
): Promise<number | null> {
  const { downloads, resolved } = await monthlyDownloads(
    PYPI_PACKAGES,
    fetchImpl,
  );
  return resolved === 0 ? null : downloads;
}
