const CACHE_CONTROL =
  "public, max-age=300, s-maxage=86400, stale-while-revalidate=604800";

export function officialPackageNames(
  packageAccess: Record<string, string>,
): string[] {
  return Object.keys(packageAccess).filter((name) =>
    name.startsWith("@betteroffice/"),
  );
}

/**
 * Sum last-month downloads across packages, tolerating per-package failures:
 * freshly published packages 404 on the stats API for up to a day, and one
 * flaky fetch must not turn the whole badge into "unavailable". `resolved`
 * reports how many packages produced a valid count.
 */
export async function monthlyDownloads(
  packageNames: string[],
  fetchImpl: typeof fetch = fetch,
): Promise<{ downloads: number; resolved: number }> {
  const results = await Promise.allSettled(
    packageNames.map(async (packageName) => {
      const response = await fetchImpl(
        `https://api.npmjs.org/downloads/point/last-month/${encodeURIComponent(packageName)}`,
      );

      if (!response.ok) {
        throw new Error(
          `npm downloads request failed for ${packageName}: ${response.status}`,
        );
      }

      const data = (await response.json()) as { downloads?: number };

      if (!Number.isSafeInteger(data.downloads)) {
        throw new Error(`Invalid npm download count for ${packageName}`);
      }

      return data.downloads as number;
    }),
  );

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

export async function GET() {
  try {
    const packagesResponse = await fetch(
      "https://registry.npmjs.org/-/org/betteroffice/package?format=cli",
    );

    if (!packagesResponse.ok) {
      throw new Error(`npm packages request failed: ${packagesResponse.status}`);
    }

    const packageAccess = (await packagesResponse.json()) as Record<
      string,
      string
    >;
    const packageNames = officialPackageNames(packageAccess);

    if (packageNames.length === 0) {
      throw new Error("No public @betteroffice packages found");
    }

    const { downloads, resolved } = await monthlyDownloads(packageNames);

    if (resolved === 0) {
      throw new Error("No package download counts resolved");
    }

    return Response.json(
      {
        schemaVersion: 1,
        label: "npm downloads",
        message: `${downloads.toLocaleString("en-US")}/month`,
        color: "brightgreen",
      },
      { headers: { "Cache-Control": CACHE_CONTROL } },
    );
  } catch {
    return Response.json(
      {
        schemaVersion: 1,
        label: "npm downloads",
        message: "unavailable",
        color: "lightgrey",
        isError: true,
      },
      { status: 502, headers: { "Cache-Control": "no-store" } },
    );
  }
}
