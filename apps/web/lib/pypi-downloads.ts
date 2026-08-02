import { PYPI_DISTRIBUTIONS } from "../../../scripts/python-bindings.mjs";

const PYPISTATS_API = "https://pypistats.org/api/packages";
const USER_AGENT =
  "betteroffice.dev downloads badge (https://github.com/openooxml/betteroffice)";

/** PyPI has no org listing endpoint, so the binding registry is the source. */
export const PYPI_PACKAGES: string[] = PYPI_DISTRIBUTIONS;

interface RecentDownloads {
  data?: { last_month?: number };
}

const NOT_PUBLISHED = Symbol("not published");

/**
 * Sum last-month downloads. `resolved` counts packages that produced a value
 * and `missing` counts those pypistats has no record of, which is distinct
 * from a failed request: a package with no stats yet has genuinely had no
 * downloads, whereas a rate-limited or failing API tells us nothing.
 */
export async function monthlyDownloads(
  packageNames: string[],
  fetchImpl: typeof fetch = fetch,
): Promise<{ downloads: number; resolved: number; missing: number }> {
  const fetchCount = async (
    packageName: string,
  ): Promise<number | typeof NOT_PUBLISHED> => {
    const response = await fetchImpl(
      `${PYPISTATS_API}/${encodeURIComponent(packageName)}/recent`,
      { headers: { "User-Agent": USER_AGENT } },
    );
    if (response.status === 404) return NOT_PUBLISHED;
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
  let missing = 0;
  for (const result of results) {
    if (result.status !== "fulfilled") continue;
    if (result.value === NOT_PUBLISHED) {
      missing += 1;
    } else {
      downloads += result.value;
      resolved += 1;
    }
  }
  return { downloads, resolved, missing };
}

/** `null` only when the API told us nothing; a package with no stats is 0. */
export async function monthlyDownloadsTotal(
  fetchImpl: typeof fetch = fetch,
): Promise<number | null> {
  const { downloads, resolved, missing } = await monthlyDownloads(
    PYPI_PACKAGES,
    fetchImpl,
  );
  if (resolved > 0) return downloads;
  return missing === PYPI_PACKAGES.length ? 0 : null;
}
