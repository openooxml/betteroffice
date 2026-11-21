import { privateMode } from './private-step.mjs';
import { EMPTY_SHA256, signRequest } from './sigv4.mjs';

export const PRIVATE_BUCKET = 'betteroffice-corpus-private';

const ENDPOINT = /^https:\/\/[a-f0-9]{32}(?:\.(?:eu|fedramp))?\.r2\.cloudflarestorage\.com$/;
const KEYS = {
  read: ['PRIVATE_R2_READ_KEY_ID', 'PRIVATE_R2_READ_SECRET'],
  write: ['PRIVATE_R2_WRITE_KEY_ID', 'PRIVATE_R2_WRITE_SECRET'],
};

export function privateCorpus(environment = process.env, access = 'read') {
  const names = KEYS[access];
  if (!names) throw new Error('Private bucket access must be read or write');
  const endpoint = environment.PRIVATE_R2_ENDPOINT?.trim() ?? '';
  const [accessKeyId, secretAccessKey] = names.map((name) => environment[name]?.trim() ?? '');
  if (!ENDPOINT.test(endpoint) || !accessKeyId || !secretAccessKey)
    throw new Error(`Private corpus runs need PRIVATE_R2_ENDPOINT, ${names.join(' and ')} from the private-corpus environment`);
  return {
    base: `${endpoint}/${PRIVATE_BUCKET}`,
    credentials: { accessKeyId, secretAccessKey, region: 'auto', service: 's3' },
  };
}

export function isUnder(base, url) {
  const inside = (value) => value.startsWith(`${base}/`) || value.startsWith(`${base}?`);
  return inside(String(url)) && inside(new URL(url).href);
}

export function signedHeaders(corpus, method, url, { payloadHash = EMPTY_SHA256, headers = {}, date } = {}) {
  if (!isUnder(corpus.base, url)) throw new Error('Refusing to sign a request outside the private bucket');
  return signRequest({ ...corpus.credentials, method, url: new URL(url).href, headers, payloadHash, date }).headers;
}

export function corpusSigner(environment = process.env) {
  if (!privateMode(environment)) return null;
  const corpus = privateCorpus(environment, 'read');
  return (url) => (isUnder(corpus.base, url) ? signedHeaders(corpus, 'GET', url) : null);
}
