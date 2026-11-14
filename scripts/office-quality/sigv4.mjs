import { createHash, createHmac } from 'node:crypto';

export const EMPTY_SHA256 = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';

const UNRESERVED = /[A-Za-z0-9\-._~]/;

export function sha256Hex(data) {
  return createHash('sha256').update(data).digest('hex');
}

function hmac(key, data) {
  return createHmac('sha256', key).update(data).digest();
}

export function uriEncode(value, keepSlash = false) {
  let encoded = '';
  for (const byte of Buffer.from(value, 'utf8')) {
    const character = String.fromCharCode(byte);
    encoded += UNRESERVED.test(character) || (keepSlash && character === '/')
      ? character
      : `%${byte.toString(16).toUpperCase().padStart(2, '0')}`;
  }
  return encoded;
}

function decode(value) {
  return decodeURIComponent(value);
}

function canonicalQuery(search) {
  if (!search || search === '?') return '';
  return search.slice(1).split('&').filter(Boolean).map((pair) => {
    const separator = pair.indexOf('=');
    const key = separator < 0 ? pair : pair.slice(0, separator);
    const value = separator < 0 ? '' : pair.slice(separator + 1);
    return [uriEncode(decode(key)), uriEncode(decode(value))];
  }).sort(([a, x], [b, y]) => (a < b ? -1 : a > b ? 1 : x < y ? -1 : x > y ? 1 : 0))
    .map(([key, value]) => `${key}=${value}`).join('&');
}

function canonicalHeaders(headers) {
  const merged = new Map();
  for (const [name, value] of Object.entries(headers)) {
    const key = name.toLowerCase().trim();
    const values = (Array.isArray(value) ? value : [value]).map((item) => String(item).trim().replace(/\s+/g, ' '));
    merged.set(key, [...(merged.get(key) ?? []), ...values]);
  }
  const names = [...merged.keys()].sort();
  return {
    text: names.map((name) => `${name}:${merged.get(name).join(',')}\n`).join(''),
    signed: names.join(';'),
  };
}

export function amzDate(date) {
  return date.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
}

export function canonicalRequest({ method, url, headers, payloadHash }) {
  const target = new URL(url);
  const { text, signed } = canonicalHeaders(headers);
  return [
    method.toUpperCase(),
    uriEncode(decode(target.pathname), true) || '/',
    canonicalQuery(target.search),
    text,
    signed,
    payloadHash,
  ].join('\n');
}

export function signRequest({
  method,
  url,
  headers = {},
  payloadHash = EMPTY_SHA256,
  accessKeyId,
  secretAccessKey,
  sessionToken,
  region,
  service,
  date = new Date(),
}) {
  if (!accessKeyId || !secretAccessKey || !region || !service)
    throw new Error('Signing needs an access key, a secret, a region and a service');
  const target = new URL(url);
  const stamp = amzDate(date);
  const day = stamp.slice(0, 8);
  const signedHeaders = {
    ...headers,
    host: target.host,
    'x-amz-date': stamp,
    ...(service === 's3' ? { 'x-amz-content-sha256': payloadHash } : {}),
    ...(sessionToken ? { 'x-amz-security-token': sessionToken } : {}),
  };
  const request = canonicalRequest({ method, url, headers: signedHeaders, payloadHash });
  const scope = `${day}/${region}/${service}/aws4_request`;
  const stringToSign = ['AWS4-HMAC-SHA256', stamp, scope, sha256Hex(request)].join('\n');
  const key = [day, region, service, 'aws4_request'].reduce(hmac, `AWS4${secretAccessKey}`);
  const signature = createHmac('sha256', key).update(stringToSign).digest('hex');
  const { signed } = canonicalHeaders(signedHeaders);
  const { host, ...sent } = signedHeaders;
  return {
    canonicalRequest: request,
    stringToSign,
    headers: {
      ...sent,
      authorization: `AWS4-HMAC-SHA256 Credential=${accessKeyId}/${scope}, SignedHeaders=${signed}, Signature=${signature}`,
    },
  };
}
