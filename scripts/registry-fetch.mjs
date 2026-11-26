const USER_AGENT = 'betteroffice-release (https://github.com/openooxml/betteroffice)';
const ATTEMPTS = 5;
export const REQUEST_TIMEOUT_MS = Number(process.env.REGISTRY_TIMEOUT_MS ?? 15_000);
const RETRY_AFTER_CAP_MS = 60_000;

export function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function backoffMs(attempt) {
  const base = 2 ** attempt * 500;
  return base / 2 + Math.random() * (base / 2);
}

function retryAfterMs(response) {
  const header = response.headers.get('Retry-After');
  if (!header) return undefined;
  const seconds = Number(header);
  const ms = Number.isFinite(seconds) ? seconds * 1000 : Date.parse(header) - Date.now();
  if (!Number.isFinite(ms)) return undefined;
  return Math.min(Math.max(ms, 0), RETRY_AFTER_CAP_MS);
}

export async function attemptRegistry(url, headers, timeoutMs, method = 'GET', readBody = method !== 'HEAD', fetchImpl = fetch) {
  const response = await fetchImpl(url, {
    method,
    headers: { 'User-Agent': USER_AGENT, ...headers },
    cache: 'no-store',
    signal: AbortSignal.timeout(timeoutMs)
  });
  if (response.status === 404) return { body: null };
  if (response.ok && readBody) return { body: await response.json() };
  if (response.ok) {
    response.body?.cancel().catch(() => {});
    return { body: true };
  }
  const error = new Error(`${url} returned ${response.status}`);
  const retry = response.status === 429 || response.status >= 500;
  return { error, retry, wait: retry ? retryAfterMs(response) : undefined };
}

export async function fetchRegistry(url, {
  headers = {},
  timeoutMs = REQUEST_TIMEOUT_MS,
  readBody = true,
  fetchImpl = fetch
} = {}) {
  let lastError;
  for (let attempt = 0; attempt < ATTEMPTS; attempt++) {
    let result;
    try {
      result = await attemptRegistry(url, headers, timeoutMs, 'GET', readBody, fetchImpl);
    } catch (error) {
      result = { error, retry: true };
    }
    if (!result.error) return result.body;
    if (!result.retry) throw result.error;
    lastError = result.error;
    if (attempt < ATTEMPTS - 1) await sleep(result.wait ?? backoffMs(attempt));
  }
  throw lastError;
}
