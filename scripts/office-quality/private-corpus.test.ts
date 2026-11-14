import { expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { fetchAsset } from './asset-cache.mjs';
import { download } from './download.mjs';
import { PRIVATE_BUCKET, corpusSigner, isUnder, privateCorpus } from './private-corpus.mjs';
import { corpusOrigin } from './samples.mjs';

const ENDPOINT = `https://${'a'.repeat(32)}.r2.cloudflarestorage.com`;
const BASE = `${ENDPOINT}/${PRIVATE_BUCKET}`;
const READ = { PRIVATE_R2_READ_KEY_ID: 'read-id', PRIVATE_R2_READ_SECRET: 'read-secret' };
const PRIVATE = { QUALITY_PRIVATE: 'true', PRIVATE_R2_ENDPOINT: ENDPOINT, ...READ };

test('private runs fail closed without the private-corpus secrets', () => {
  for (const environment of [
    { QUALITY_PRIVATE: 'true' },
    { ...PRIVATE, PRIVATE_R2_ENDPOINT: '' },
    { ...PRIVATE, PRIVATE_R2_ENDPOINT: 'https://corpus.betteroffice.dev' },
    { ...PRIVATE, PRIVATE_R2_ENDPOINT: `${ENDPOINT}/` },
    { ...PRIVATE, PRIVATE_R2_ENDPOINT: ENDPOINT.replace('https:', 'http:') },
    { ...PRIVATE, PRIVATE_R2_READ_SECRET: ' ' },
  ]) {
    expect(() => corpusOrigin(environment)).toThrow('Private corpus runs need PRIVATE_R2_ENDPOINT');
    expect(() => corpusSigner(environment)).toThrow('Private corpus runs need PRIVATE_R2_ENDPOINT');
  }
  expect(() => privateCorpus(PRIVATE, 'write')).toThrow('PRIVATE_R2_WRITE_KEY_ID and PRIVATE_R2_WRITE_SECRET');
  expect(() => privateCorpus(PRIVATE, 'admin' as any)).toThrow('read or write');
});

test('private runs read the private bucket and refuse a corpus origin input', () => {
  expect(corpusOrigin(PRIVATE)).toBe(BASE);
  expect(corpusOrigin({ ...PRIVATE, PRIVATE_R2_ENDPOINT: `https://${'b'.repeat(32)}.eu.r2.cloudflarestorage.com` }))
    .toBe(`https://${'b'.repeat(32)}.eu.r2.cloudflarestorage.com/${PRIVATE_BUCKET}`);
  expect(() => corpusOrigin({ ...PRIVATE, QUALITY_CORPUS_ORIGIN: 'https://bench.example.org' }))
    .toThrow('Private corpus runs read the private bucket');
  expect(corpusOrigin({ QUALITY_PRIVATE: 'false', ...READ })).toBe('https://corpus.betteroffice.dev');
  expect(corpusSigner({ QUALITY_PRIVATE: 'false' })).toBeNull();
});

test('only URLs inside the private bucket are signed', () => {
  const sign = corpusSigner(PRIVATE)!;
  const headers = sign(`${BASE}/sample-a/metadata.json`)!;
  expect(headers.authorization).toMatch(/^AWS4-HMAC-SHA256 Credential=read-id\/\d{8}\/auto\/s3\/aws4_request, SignedHeaders=host;x-amz-content-sha256;x-amz-date, Signature=[a-f0-9]{64}$/);
  expect(JSON.stringify(headers)).not.toContain('read-secret');
  for (const url of [
    'https://registry.npmjs.org/@betteroffice%2Fdocx/latest',
    `${ENDPOINT}/other-bucket/sample-a/metadata.json`,
    `${BASE}-copy/sample-a/metadata.json`,
    `${BASE}/../other-bucket/key`,
    `${ENDPOINT}/${PRIVATE_BUCKET}`,
  ]) expect(sign(url)).toBeNull();
  expect(isUnder(BASE, `${BASE}?list-type=2`)).toBe(true);
});

test('downloads send the signature only to the private bucket and refuse redirects there', async () => {
  const calls: any[] = [];
  const fetchImpl = async (url: any, init: any) => {
    calls.push({ url: String(url), init });
    return new Response('ok');
  };
  const sign = corpusSigner(PRIVATE);
  await download(`${BASE}/sample-a/metadata.json`, 1024, { fetchImpl, sign });
  await download('https://registry.npmjs.org/x/latest', 1024, { fetchImpl, sign });
  expect(calls[0].init.headers.authorization).toStartWith('AWS4-HMAC-SHA256 ');
  expect(calls[0].init.redirect).toBe('error');
  expect(calls[1].init.headers).toBeUndefined();
  expect(calls[1].init.redirect).toBeUndefined();
});

test('assets must live under the corpus base and the sample folder', async () => {
  const bytes = Buffer.from('private asset');
  const entry = (url: string) => ({ url, bytes: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex') });
  const fetched: string[] = [];
  const get = async (url: URL) => {
    fetched.push(url.href);
    return bytes;
  };
  expect(await fetchAsset(entry(`${BASE}/sample-a/source.docx`), 'sample-a', get, { origin: BASE })).toEqual(bytes);
  for (const url of [
    `${ENDPOINT}/sample-a/source.docx`,
    `${BASE}/sample-b/source.docx`,
    `${BASE}/sample-a/../sample-b/source.docx`,
    `${ENDPOINT}/other/${PRIVATE_BUCKET}/sample-a/source.docx`,
    `https://corpus.betteroffice.dev/sample-a/source.docx`,
  ]) await expect(fetchAsset(entry(url), 'sample-a', get, { origin: BASE })).rejects.toThrow('Invalid corpus asset metadata');
  expect(fetched).toEqual([`${BASE}/sample-a/source.docx`]);
});
