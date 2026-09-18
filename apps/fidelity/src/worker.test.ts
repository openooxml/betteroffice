import { expect, test } from 'bun:test';
import worker, { type Env } from './worker';

const sha = 'a'.repeat(40);
const page = `renders/${sha}/alpha/page_0001.png`;

function object(body: string | null) {
  return {
    httpEtag: '"etag"',
    writeHttpMetadata: (headers: Headers) => headers.set('content-type', 'image/png'),
    ...(body === null ? {} : { body }),
  };
}

function environment(store: Record<string, unknown> = { [page]: object('png') }): Env {
  return {
    ASSETS: { fetch: async () => new Response('asset', { status: 200 }) },
    RENDERS: { get: async (key: string) => store[key] ?? null },
  } as unknown as Env;
}

const get = (path: string, env = environment(), init?: RequestInit) =>
  worker.fetch(new Request(`https://fidelity.betteroffice.dev${path}`, init), env);

test('a malformed escape is a client error, not a crash', async () => {
  const response = await get('/renders/%ZZ');
  expect(response.status).toBe(400);
  expect(await response.text()).toBe('Invalid path');
  expect((await get('/renders/%E0%A4%A')).status).toBe(400);
});

test('traversal and unknown keys are refused without reaching the bucket', async () => {
  let reads = 0;
  const env = {
    ASSETS: { fetch: async () => new Response('asset') },
    RENDERS: {
      get: async () => {
        reads += 1;
        return null;
      },
    },
  } as unknown as Env;
  // The URL parser resolves %2e%2e away; only an encoded slash survives to reach the guard.
  expect((await get('/renders/alpha/%2e%2e%2fsecret', env)).status).toBe(404);
  expect((await get('/renders/%2F%2Fsecret', env)).status).toBe(404);
  expect((await get('/renders//secret', env)).status).toBe(404);
  expect(reads).toBe(0);
  expect((await get(`/renders/${sha}/alpha/page_0009.png`, env)).status).toBe(404);
  expect(reads).toBe(1);
});

test('only safe methods are served', async () => {
  expect((await get('/renders/latest.json', environment(), { method: 'POST' })).status).toBe(405);
  expect((await get('/', environment(), { method: 'DELETE' })).status).toBe(405);
});

test('anything outside the prefix is handed to the assets binding', async () => {
  expect(await (await get('/')).text()).toBe('asset');
  expect(await (await get('/app.js')).text()).toBe('asset');
  expect(await (await get('/rendersneak')).text()).toBe('asset');
  expect(await (await get('/renders/%2e%2e/secret')).text()).toBe('asset');
});

test('pages are immutable and pointers revalidate', async () => {
  const image = await get(`/${page}`);
  expect(image.status).toBe(200);
  expect(image.headers.get('cache-control')).toBe('public, max-age=31536000, immutable');
  expect(image.headers.get('etag')).toBe('"etag"');
  expect(image.headers.get('content-type')).toBe('image/png');

  const pointer = await get('/renders/latest.json', environment({ 'renders/latest.json': object('{}') }));
  expect(pointer.headers.get('cache-control')).toBe('public, max-age=0, must-revalidate');
});

test('a conditional hit returns 304 and an unbound bucket reports unconfigured', async () => {
  const notModified = await get(`/${page}`, environment({ [page]: object(null) }));
  expect(notModified.status).toBe(304);
  expect(notModified.headers.get('etag')).toBe('"etag"');

  const unbound = { ASSETS: { fetch: async () => new Response('asset') } } as unknown as Env;
  expect((await get(`/${page}`, unbound)).status).toBe(503);
});
