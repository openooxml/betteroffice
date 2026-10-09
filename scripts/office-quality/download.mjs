export const DOWNLOAD_ATTEMPTS = 5;
export const RETRY_AFTER_CAP_MS = 30_000;

export function isTransientStatus(status) {
  return status === 408 || status === 429 || (status >= 500 && status <= 599);
}

export function retryAfterMs(response, now = Date.now()) {
  const header = response.headers?.get?.('Retry-After');
  if (!header) return undefined;
  const seconds = Number(header);
  const ms = Number.isFinite(seconds) ? seconds * 1000 : Date.parse(header) - now;
  if (!Number.isFinite(ms)) return undefined;
  return Math.min(Math.max(ms, 0), RETRY_AFTER_CAP_MS);
}

export function backoffMs(attempt, random = Math.random) {
  const base = 2 ** attempt * 500;
  return base / 2 + random() * (base / 2);
}

function timeoutMs(maximum) {
  return Math.max(30_000, Math.ceil(maximum / (1024 * 1024)) * 1000);
}

async function cancelBody(response) {
  try {
    await response.body?.cancel?.();
  } catch {}
}

async function readBody(response, maximum, url) {
  const body = response.body;
  if (body?.getReader) {
    const reader = body.getReader();
    try {
      const chunks = [];
      let size = 0;
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        size += value.length;
        if (size > maximum) {
          await reader.cancel().catch(() => {});
          throw new Error(`Download exceeds byte limit: ${url}`);
        }
        chunks.push(value);
      }
      return Buffer.concat(chunks);
    } catch (error) {
      if (String(error?.message ?? '').startsWith('Download exceeds byte limit')) throw error;
      try {
        await reader.cancel();
      } catch {}
      throw error;
    } finally {
      try {
        reader.releaseLock();
      } catch {}
    }
  }
  const chunks = [];
  let size = 0;
  try {
    for await (const chunk of body) {
      size += chunk.length;
      if (size > maximum) {
        await cancelBody(response);
        throw new Error(`Download exceeds byte limit: ${url}`);
      }
      chunks.push(chunk);
    }
  } catch (error) {
    if (String(error?.message ?? '').startsWith('Download exceeds byte limit')) throw error;
    await cancelBody(response);
    throw error;
  }
  return Buffer.concat(chunks);
}

function logAttempt(url, attempt, attempts, status, ray) {
  const detail = status ? `status ${status}` : 'connection error';
  const suffix = ray ? ` cf-ray ${ray}` : '';
  process.stderr.write(`Download attempt ${attempt + 1}/${attempts} ${detail}: ${url}${suffix}\n`);
}

export async function download(url, maximum = 32 * 1024 * 1024, options = {}) {
  const {
    attempts = DOWNLOAD_ATTEMPTS,
    fetchImpl = fetch,
    sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
    random = Math.random,
    now = Date.now,
  } = options;
  let lastError;
  for (let attempt = 0; attempt < attempts; attempt++) {
    let response;
    try {
      response = await fetchImpl(url, {
        signal: AbortSignal.timeout(timeoutMs(maximum)),
        credentials: 'omit',
        referrerPolicy: 'no-referrer',
      });
    } catch (error) {
      lastError = error;
      logAttempt(url, attempt, attempts);
      if (attempt < attempts - 1) await sleep(backoffMs(attempt, random));
      continue;
    }
    if (!response.ok) {
      const ray = response.headers?.get?.('cf-ray');
      await cancelBody(response);
      lastError = new Error(`Download failed (${response.status}): ${url}`);
      logAttempt(url, attempt, attempts, response.status, ray);
      if (!isTransientStatus(response.status)) throw lastError;
      if (attempt < attempts - 1)
        await sleep(retryAfterMs(response, now()) ?? backoffMs(attempt, random));
      continue;
    }
    try {
      return await readBody(response, maximum, url);
    } catch (error) {
      lastError = error;
      if (String(error?.message ?? '').startsWith('Download exceeds byte limit')) throw error;
      logAttempt(url, attempt, attempts);
      if (attempt < attempts - 1) await sleep(backoffMs(attempt, random));
    }
  }
  throw lastError;
}
