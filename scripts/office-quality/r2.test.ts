import { createHash } from 'node:crypto';
import { afterEach, expect, test } from 'bun:test';
import { S3Client } from 'bun';
import { r2Options } from './r2.mjs';

const account = 'a'.repeat(32);
const tokenId = 'b'.repeat(32);
const token = 'test-api-token';
const environment = {
  CLOUDFLARE_ACCOUNT_ID: account,
  CLOUDFLARE_API_TOKEN: token,
  QUALITY_RENDER_BUCKET: 'test-bucket',
};
const active = () => Response.json({ success: true, result: { id: tokenId, status: 'active' } });
let server: ReturnType<typeof Bun.serve> | undefined;
afterEach(() => server?.stop(true));

test('derives bucket-scoped S3 credentials from the existing account token', async () => {
  const calls: string[] = [];
  const options = await r2Options(environment, async (url: string, init: RequestInit) => {
    calls.push(url);
    expect(new Headers(init.headers).get('Authorization')).toBe(`Bearer ${token}`);
    expect(init.redirect).toBe('error');
    expect(init.signal).toBeInstanceOf(AbortSignal);
    return active();
  });
  expect(calls).toEqual([`https://api.cloudflare.com/client/v4/accounts/${account}/tokens/verify`]);
  expect(options).toEqual({
    accessKeyId: tokenId,
    secretAccessKey: createHash('sha256').update(token).digest('hex'),
    endpoint: `https://${account}.r2.cloudflarestorage.com`,
    region: 'auto',
    bucket: 'test-bucket',
  });
});

test.each([401, 403])('supports user tokens after account verification returns %i', async (status) => {
  const calls: string[] = [];
  const options = await r2Options(environment, async (url: string) => {
    calls.push(url);
    return calls.length === 1 ? new Response('unauthorized', { status }) : active();
  });
  expect(calls).toEqual([
    `https://api.cloudflare.com/client/v4/accounts/${account}/tokens/verify`,
    'https://api.cloudflare.com/client/v4/user/tokens/verify',
  ]);
  expect(options.accessKeyId).toBe(tokenId);
});

test('fails on service errors instead of treating them as another token type', async () => {
  let calls = 0;
  await expect(r2Options(environment, async () => {
    calls += 1;
    return new Response('private diagnostics', { status: 503 });
  })).rejects.toThrow('Cloudflare token verification failed (HTTP 503)');
  expect(calls).toBe(1);
});

test.each([
  { success: false, result: { id: tokenId, status: 'active' } },
  { success: true, result: { id: tokenId, status: 'disabled' } },
  { success: true, result: { id: 'invalid', status: 'active' } },
])('rejects unverified or inactive credentials', async (response) => {
  await expect(r2Options(environment, async () => Response.json(response)))
    .rejects.toThrow('did not return an active token ID');
});

test('rejects missing credentials before making requests', async () => {
  let calls = 0;
  for (const invalid of [{}, { ...environment, CLOUDFLARE_ACCOUNT_ID: '../another' },
    { ...environment, CLOUDFLARE_API_TOKEN: ' ' }]) {
    await expect(r2Options(invalid, async () => { calls += 1; return active(); }))
      .rejects.toThrow('CLOUDFLARE_ACCOUNT_ID and CLOUDFLARE_API_TOKEN are required');
  }
  expect(calls).toBe(0);
});

test('uploads binary PNG bytes with S3 signing and content metadata', async () => {
  const requests: { path: string; method: string; type: string | null; auth: string | null; bytes: Uint8Array }[] = [];
  server = Bun.serve({
    hostname: '127.0.0.1', port: 0,
    async fetch(request) {
      requests.push({
        path: new URL(request.url).pathname,
        method: request.method,
        type: request.headers.get('content-type'),
        auth: request.headers.get('authorization'),
        bytes: new Uint8Array(await request.arrayBuffer()),
      });
      return new Response('', { headers: { ETag: 'test-etag' } });
    },
  });
  const options = await r2Options(environment, async () => active());
  const client = new S3Client({ ...options, endpoint: server.url.toString() });
  const data = new Uint8Array([137, 80, 78, 71, 0, 255, 10]);
  await client.file('renders/commit/sample/page_0001.png').write(new Blob([data]), { type: 'image/png' });
  expect(requests).toHaveLength(1);
  expect(requests[0].path).toBe('/test-bucket/renders/commit/sample/page_0001.png');
  expect(requests[0].method).toBe('PUT');
  expect(requests[0].type).toBe('image/png');
  expect(requests[0].bytes).toEqual(data);
  expect(requests[0].auth).toContain(`AWS4-HMAC-SHA256 Credential=${tokenId}/`);
  expect(requests[0].auth).toContain('/auto/s3/aws4_request');
  expect(requests[0].auth).not.toContain(token);
  expect(requests[0].auth).not.toContain(options.secretAccessKey);
});
