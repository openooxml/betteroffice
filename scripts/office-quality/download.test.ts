import { expect, test } from 'bun:test';
import { download } from './download.mjs';

const url = 'https://corpus.betteroffice.dev/sample/reference/page_0001.png';

function body(chunks: Uint8Array[], seen?: { read: boolean }) {
  return (async function* () {
    if (seen) seen.read = true;
    for (const chunk of chunks) yield chunk;
  })();
}

function ok(chunks: Uint8Array[], headers: Record<string, string> = {}) {
  return { ok: true, status: 200, headers: new Headers(headers), body: body(chunks) };
}

function failure(status: number, headers: Record<string, string> = {}, seen?: { read: boolean }) {
  return { ok: false, status, headers: new Headers(headers), body: body([], seen) };
}

function openStream(tracker: { cancelled: boolean }) {
  return new ReadableStream({
    start() {},
    cancel() {
      tracker.cancelled = true;
    },
  });
}

function oversizeStream(chunks: Uint8Array[], tracker: { cancelled: boolean }) {
  let index = 0;
  return new ReadableStream({
    pull(controller) {
      if (index < chunks.length) controller.enqueue(chunks[index++]);
    },
    cancel() {
      tracker.cancelled = true;
    },
  });
}

function streamOk(chunks: Uint8Array[], tracker: { cancelled: boolean }, headers: Record<string, string> = {}) {
  return { ok: true, status: 200, headers: new Headers(headers), body: oversizeStream(chunks, tracker) };
}

function streamFailure(status: number, tracker: { cancelled: boolean }, headers: Record<string, string> = {}) {
  return { ok: false, status, headers: new Headers(headers), body: openStream(tracker) };
}

function harness(queue: unknown[], waits: number[] = [], now?: number) {
  const calls: unknown[] = [];
  const fetchImpl = async (...args: unknown[]) => {
    calls.push(args);
    const next = queue.shift();
    if (next instanceof Error) throw next;
    return next;
  };
  const options = {
    fetchImpl: fetchImpl as typeof fetch,
    sleep: async (ms: number) => {
      waits.push(ms);
    },
    random: () => 0,
    ...(now === undefined ? {} : { now: () => now }),
  };
  return { calls, waits, options };
}

test('recovers from transient 503s then returns the exact bytes', async () => {
  const payload = [Buffer.from('png-a'), Buffer.from('png-b')];
  const { calls, waits, options } = harness([failure(503), failure(503), ok(payload)]);
  const bytes = await download(url, 1024, options);
  expect(bytes).toEqual(Buffer.concat(payload));
  expect(calls.length).toBe(3);
  expect(waits.length).toBe(2);
});

test('throws after exhausting finite attempts on a persistent 503', async () => {
  const { calls, waits, options } = harness([
    failure(503),
    failure(503),
    failure(503),
    failure(503),
    failure(503),
    failure(503),
  ]);
  await expect(download(url, 1024, options)).rejects.toThrow('Download failed (503)');
  expect(calls.length).toBe(5);
  expect(waits.length).toBe(4);
});

test('respects Retry-After seconds', async () => {
  const { waits, options } = harness([failure(429, { 'Retry-After': '2' }), ok([Buffer.from('x')])]);
  await download(url, 1024, options);
  expect(waits).toEqual([2000]);
});

test('respects Retry-After HTTP dates and caps long waits', async () => {
  const now = Date.UTC(2026, 0, 1, 0, 0, 0);
  const soon = new Date(now + 5000).toUTCString();
  const first = harness([failure(429, { 'Retry-After': soon }), ok([Buffer.from('x')])], [], now);
  await download(url, 1024, first.options);
  expect(first.waits).toEqual([5000]);

  const capped = harness([
    failure(503, { 'Retry-After': '3600' }),
    ok([Buffer.from('x')]),
  ]);
  await download(url, 1024, capped.options);
  expect(capped.waits).toEqual([30_000]);

  const far = new Date(now + 3600_000).toUTCString();
  const dated = harness([failure(503, { 'Retry-After': far }), ok([Buffer.from('x')])], [], now);
  await download(url, 1024, dated.options);
  expect(dated.waits).toEqual([30_000]);
});

test('fails permanent errors immediately without reading the error page', async () => {
  const seen = { read: false };
  const { calls, waits, options } = harness([failure(404, {}, seen)]);
  await expect(download(url, 1024, options)).rejects.toThrow('Download failed (404)');
  expect(calls.length).toBe(1);
  expect(waits).toEqual([]);
  expect(seen.read).toBe(false);
});

test('cancels rejected ReadableStream bodies for permanent errors', async () => {
  const tracker = { cancelled: false };
  const { calls, waits, options } = harness([streamFailure(404, tracker)]);
  await expect(download(url, 1024, options)).rejects.toThrow('Download failed (404)');
  expect(calls.length).toBe(1);
  expect(waits).toEqual([]);
  expect(tracker.cancelled).toBe(true);
});

test('cancels rejected ReadableStream bodies for transient errors', async () => {
  const first = { cancelled: false };
  const second = { cancelled: false };
  const { calls, options } = harness([
    streamFailure(503, first),
    streamFailure(503, second),
    ok([Buffer.from('recovered')]),
  ]);
  const bytes = await download(url, 1024, options);
  expect(bytes).toEqual(Buffer.from('recovered'));
  expect(calls.length).toBe(3);
  expect(first.cancelled).toBe(true);
  expect(second.cancelled).toBe(true);
});

test('retries a body that stalls mid-stream', async () => {
  const stalled = {
    ok: true,
    status: 200,
    headers: new Headers(),
    body: (async function* () {
      yield Buffer.from('part');
      throw new Error('socket hang up');
    })(),
  };
  const { calls, options } = harness([stalled, ok([Buffer.from('whole')])]);
  const bytes = await download(url, 1024, options);
  expect(bytes).toEqual(Buffer.from('whole'));
  expect(calls.length).toBe(2);
});

test('rejects over-limit bodies without retrying', async () => {
  const { calls, options } = harness([ok([Buffer.from('toolarge-body')])]);
  await expect(download(url, 8, options)).rejects.toThrow('exceeds byte limit');
  expect(calls.length).toBe(1);
});

test('releases oversize ReadableStream bodies without retrying', async () => {
  const tracker = { cancelled: false };
  const { calls, options } = harness([streamOk([Buffer.from('toolarge-body')], tracker)]);
  await expect(download(url, 8, options)).rejects.toThrow('exceeds byte limit');
  expect(calls.length).toBe(1);
  expect(tracker.cancelled).toBe(true);
});

test('retries connection failures with bounded credential-free requests', async () => {
  const { calls, options } = harness([
    new TypeError('fetch failed'),
    new TypeError('fetch failed'),
    ok([Buffer.from('ok')]),
  ]);
  const bytes = await download(url, 1024, options);
  expect(bytes).toEqual(Buffer.from('ok'));
  expect(calls.length).toBe(3);
  for (const [, init] of calls as [unknown, RequestInit][]) {
    expect(init.credentials).toBe('omit');
    expect(init.referrerPolicy).toBe('no-referrer');
    expect(init.signal).toBeInstanceOf(AbortSignal);
  }
});

test('logs attempt status with cf-ray and no credential material', async () => {
  const lines: string[] = [];
  const write = process.stderr.write.bind(process.stderr);
  process.stderr.write = ((chunk: unknown) => {
    lines.push(String(chunk));
    return true;
  }) as typeof process.stderr.write;
  try {
    const { options } = harness([
      failure(503, { 'cf-ray': 'abc123-ray' }),
      ok([Buffer.from('ok')]),
    ]);
    await download(url, 1024, options);
  } finally {
    process.stderr.write = write;
  }
  const logged = lines.join('');
  expect(logged).toContain('503');
  expect(logged).toContain('abc123-ray');
  expect(logged).toContain(url);
  expect(logged).not.toContain('Authorization');
});
