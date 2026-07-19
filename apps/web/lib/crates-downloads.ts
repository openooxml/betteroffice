import { officialCrateNames, rollingDownloads } from "../app/api/crates-downloads/route";

const CRATES_API = "https://crates.io/api/v1";
const USER_AGENT =
  "betteroffice.dev badge (https://github.com/openooxml/betteroffice)";

type CrateSearch = Parameters<typeof officialCrateNames>[0];
type CrateHistory = Parameters<typeof rollingDownloads>[0];

async function fetchWithRetry(url: string): Promise<Response> {
  let lastError: unknown;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      const response = await fetch(url, { headers: { "User-Agent": USER_AGENT } });
      if (!response.ok) throw new Error(`${url} failed: ${response.status}`);
      return response;
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError;
}

export async function cratesMonthlyDownloadsTotal(): Promise<number | null> {
  const search = (await (
    await fetchWithRetry(`${CRATES_API}/crates?page=1&per_page=100&q=betteroffice`)
  ).json()) as { crates: CrateSearch };
  const crateNames = officialCrateNames(search.crates);
  if (crateNames.length === 0) throw new Error("no BetterOffice crates found");

  const results = await Promise.allSettled(
    crateNames.map(async (crateName) => {
      const response = await fetchWithRetry(
        `${CRATES_API}/crates/${encodeURIComponent(crateName)}/downloads`,
      );
      return rollingDownloads((await response.json()) as CrateHistory);
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
  return resolved === 0 ? null : downloads;
}
