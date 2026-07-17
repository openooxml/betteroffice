const CACHE_CONTROL =
  "public, max-age=300, s-maxage=86400, stale-while-revalidate=604800";

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
    const packageNames = Object.keys(packageAccess).filter((name) =>
      name.startsWith("@betteroffice/"),
    );

    if (packageNames.length === 0) {
      throw new Error("No public @betteroffice packages found");
    }

    const counts = await Promise.all(
      packageNames.map(async (packageName) => {
        const response = await fetch(
          `https://api.npmjs.org/downloads/point/last-month/${encodeURIComponent(packageName)}`,
        );

        if (!response.ok) {
          throw new Error(
            `npm downloads request failed for ${packageName}: ${response.status}`,
          );
        }

        const data = (await response.json()) as { downloads: number };

        if (!Number.isSafeInteger(data.downloads)) {
          throw new Error(`Invalid npm download count for ${packageName}`);
        }

        return data.downloads;
      }),
    );
    const downloads = counts.reduce((total, count) => total + count, 0);

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
