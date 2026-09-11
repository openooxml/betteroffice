import { officialCrateNames, rollingDownloads } from "./crates-downloads";
import { officialPackageNames } from "./downloads";
import { PYPI_PACKAGES } from "./pypi-downloads";

export const KV_KEYS = {
  npm: "npm-downloads",
  pypi: "pypi-downloads",
  crates: "crates-downloads",
} as const;

export type Registry = keyof typeof KV_KEYS;

export interface CachedCount {
  downloads: number;
  at: number;
}

/** The subset of a KV namespace the stats need; a fake satisfies it in tests. */
export interface StatsStore {
  get(key: string, type: "json"): Promise<unknown>;
  put(key: string, value: string): Promise<void>;
}

export interface RefreshOptions {
  fetchImpl?: typeof fetch;
  sleep?: (ms: number) => Promise<void>;
  now?: () => number;
  intervalMs?: number;
  attempts?: number;
}

export type RefreshResult = Record<Registry, { downloads: number } | { error: string }>;

const REQUEST_INTERVAL_MS = 3000;
const ATTEMPTS = 3;
const USER_AGENT = "betteroffice.dev downloads badge (https://github.com/openooxml/betteroffice)";

const wait = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms));

interface Pacer {
  fetchImpl: typeof fetch;
  sleep: (ms: number) => Promise<void>;
  intervalMs: number;
  attempts: number;
  started: boolean;
}

const MAX_RETRY_AFTER_MS = 60_000;

/** Registries answer a 429 with `Retry-After`; otherwise each retry waits twice as long as the last. */
function retryDelay(pacer: Pacer, attempt: number, response?: Response): number {
  const retryAfter = Number(response?.headers?.get("Retry-After"));
  if (Number.isFinite(retryAfter) && retryAfter > 0) {
    return Math.min(retryAfter * 1000, MAX_RETRY_AFTER_MS);
  }
  return pacer.intervalMs * 2 ** attempt;
}

/**
 * One request at a time, `intervalMs` apart, each retried `attempts` times
 * with a growing delay before the whole run fails. A registry that
 * rate-limits parallel bursts answers a paced sequence reliably.
 */
async function request(pacer: Pacer, url: string): Promise<Response> {
  const { fetchImpl, sleep } = pacer;
  let lastError: unknown;
  let delay = pacer.intervalMs;
  for (let attempt = 0; attempt < pacer.attempts; attempt += 1) {
    if (pacer.started) await sleep(delay);
    pacer.started = true;
    try {
      const response = await fetchImpl(url, { headers: { "User-Agent": USER_AGENT } });
      if (response.ok || response.status === 404) return response;
      delay = retryDelay(pacer, attempt + 1, response);
      throw new Error(`${url} failed: ${response.status}`);
    } catch (error) {
      lastError = error;
      delay = Math.max(delay, retryDelay(pacer, attempt + 1));
    }
  }
  throw lastError;
}

function count(value: unknown, what: string): number {
  if (!Number.isSafeInteger(value) || (value as number) < 0) {
    throw new Error(`invalid ${what} download count`);
  }
  return value as number;
}

/** npm's point API 404s for a package published within the last day; that is zero, not a failure. */
async function npmDownloads(pacer: Pacer): Promise<number> {
  const listing = await request(pacer, "https://registry.npmjs.org/-/org/betteroffice/package?format=cli");
  if (!listing.ok) throw new Error(`npm org listing failed: ${listing.status}`);
  const names = officialPackageNames((await listing.json()) as Record<string, string>);
  if (names.length === 0) throw new Error("no @betteroffice packages found");

  let total = 0;
  for (const name of names) {
    const response = await request(
      pacer,
      `https://api.npmjs.org/downloads/point/last-month/${encodeURIComponent(name)}`,
    );
    if (response.status === 404) continue;
    total += count(((await response.json()) as { downloads?: unknown }).downloads, name);
  }
  return total;
}

/** pypistats has no record of a package nobody has downloaded yet; that is zero. */
async function pypiDownloads(pacer: Pacer): Promise<number> {
  let total = 0;
  for (const name of PYPI_PACKAGES) {
    const response = await request(
      pacer,
      `https://pypistats.org/api/packages/${encodeURIComponent(name)}/recent`,
    );
    if (response.status === 404) continue;
    const body = (await response.json()) as { data?: { last_month?: unknown } };
    total += count(body.data?.last_month, name);
  }
  return total;
}

async function cratesDownloads(pacer: Pacer, now: number): Promise<number> {
  const search = await request(pacer, "https://crates.io/api/v1/crates?page=1&per_page=100&q=betteroffice");
  if (!search.ok) throw new Error(`crates.io search failed: ${search.status}`);
  const names = officialCrateNames(
    ((await search.json()) as { crates: Parameters<typeof officialCrateNames>[0] }).crates,
  );
  if (names.length === 0) throw new Error("no BetterOffice crates found");

  let total = 0;
  for (const name of names) {
    const response = await request(
      pacer,
      `https://crates.io/api/v1/crates/${encodeURIComponent(name)}/downloads`,
    );
    if (!response.ok) throw new Error(`crates.io downloads missing for ${name}`);
    total += rollingDownloads(
      (await response.json()) as Parameters<typeof rollingDownloads>[0],
      new Date(now),
    );
  }
  return total;
}

/**
 * Refresh every registry's cached monthly total. Each registry is written
 * only when all of its requests succeeded; a failed registry keeps whatever
 * the store held before.
 */
export async function refreshDownloadStats(
  store: StatsStore,
  options: RefreshOptions = {},
): Promise<RefreshResult> {
  const now = options.now ?? Date.now;
  const pacer: Pacer = {
    fetchImpl: options.fetchImpl ?? ((input, init) => fetch(input, init)),
    sleep: options.sleep ?? wait,
    intervalMs: options.intervalMs ?? REQUEST_INTERVAL_MS,
    attempts: options.attempts ?? ATTEMPTS,
    started: false,
  };
  const collectors: Record<Registry, () => Promise<number>> = {
    npm: () => npmDownloads(pacer),
    pypi: () => pypiDownloads(pacer),
    crates: () => cratesDownloads(pacer, now()),
  };

  const result = {} as RefreshResult;
  for (const registry of Object.keys(collectors) as Registry[]) {
    try {
      const downloads = await collectors[registry]();
      await store.put(KV_KEYS[registry], JSON.stringify({ downloads, at: now() } satisfies CachedCount));
      result[registry] = { downloads };
    } catch (error) {
      result[registry] = { error: error instanceof Error ? error.message : String(error) };
    }
  }
  return result;
}

export async function cachedDownloads(store: StatsStore, registry: Registry): Promise<CachedCount | null> {
  const value = (await store.get(KV_KEYS[registry], "json").catch(() => null)) as CachedCount | null;
  return value && Number.isSafeInteger(value.downloads) ? value : null;
}
