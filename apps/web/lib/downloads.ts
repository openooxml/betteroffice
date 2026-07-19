export function officialPackageNames(
  packageAccess: Record<string, string>,
): string[] {
  return Object.keys(packageAccess).filter((name) =>
    name.startsWith("@betteroffice/"),
  );
}

async function resolvePackageNames(fetchImpl: typeof fetch = fetch): Promise<string[]> {
  const response = await fetchImpl(
    "https://registry.npmjs.org/-/org/betteroffice/package?format=cli",
  );
  if (!response.ok) {
    throw new Error(`npm org listing failed: ${response.status}`);
  }
  const names = officialPackageNames((await response.json()) as Record<string, string>);
  if (names.length === 0) throw new Error("no @betteroffice packages found");
  return names;
}

/**
 * Sum last-month downloads across packages, tolerating per-package failures
 * (freshly published packages 404 on the stats API for up to a day; each
 * package gets one retry). `resolved` counts packages that produced a value.
 */
export async function monthlyDownloads(
  packageNames: string[],
  fetchImpl: typeof fetch = fetch,
): Promise<{ downloads: number; resolved: number }> {
  const fetchCount = async (packageName: string): Promise<number> => {
    let lastError: unknown;
    for (let attempt = 0; attempt < 2; attempt += 1) {
      try {
        const response = await fetchImpl(
          `https://api.npmjs.org/downloads/point/last-month/${encodeURIComponent(packageName)}`,
        );
        if (!response.ok) {
          throw new Error(`npm downloads request failed for ${packageName}: ${response.status}`);
        }
        const data = (await response.json()) as { downloads?: number };
        if (!Number.isSafeInteger(data.downloads)) {
          throw new Error(`Invalid npm download count for ${packageName}`);
        }
        return data.downloads as number;
      } catch (error) {
        lastError = error;
      }
    }
    throw lastError;
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
  const names = await resolvePackageNames(fetchImpl);
  const { downloads, resolved } = await monthlyDownloads(names, fetchImpl);
  return resolved === 0 ? null : downloads;
}
