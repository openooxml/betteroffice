import { expect, test } from 'bun:test';
import suite from './fixtures/sigv4-test-suite.json';
import { EMPTY_SHA256, amzDate, sha256Hex, signRequest, uriEncode } from './sigv4.mjs';

for (const vector of suite.cases) {
  test(`AWS SigV4 test suite: ${vector.name}`, () => {
    const signed = signRequest({
      method: vector.method,
      url: vector.url,
      headers: vector.headers,
      payloadHash: sha256Hex(vector.body),
      accessKeyId: suite.access_key_id,
      secretAccessKey: suite.secret_access_key,
      sessionToken: (vector as any).session_token,
      region: suite.region,
      service: suite.service,
      date: new Date(suite.date),
    });
    expect(signed.canonicalRequest).toBe(vector.canonical_request);
    expect(signed.stringToSign).toBe(vector.string_to_sign);
    expect(signed.headers.authorization).toBe(vector.authorization);
  });
}

test('the S3 GET Object example signs with the content hash header', () => {
  const signed = signRequest({
    method: 'GET',
    url: 'https://examplebucket.s3.amazonaws.com/test.txt',
    headers: { Range: 'bytes=0-9' },
    accessKeyId: 'AKIAIOSFODNN7EXAMPLE',
    secretAccessKey: 'wJalrXUtnFEMI/K7MDENG/bPxRfiCYEXAMPLEKEY',
    region: 'us-east-1',
    service: 's3',
    date: new Date('2013-05-24T00:00:00Z'),
  });
  expect(signed.canonicalRequest).toBe([
    'GET', '/test.txt', '',
    'host:examplebucket.s3.amazonaws.com', 'range:bytes=0-9',
    `x-amz-content-sha256:${EMPTY_SHA256}`, 'x-amz-date:20130524T000000Z', '',
    'host;range;x-amz-content-sha256;x-amz-date', EMPTY_SHA256,
  ].join('\n'));
  expect(signed.headers.authorization).toBe(
    'AWS4-HMAC-SHA256 Credential=AKIAIOSFODNN7EXAMPLE/20130524/us-east-1/s3/aws4_request, ' +
    'SignedHeaders=host;range;x-amz-content-sha256;x-amz-date, ' +
    'Signature=f0e8bdb87c964420e857bd35b5d6ed310bd44f0170aba48dd91039c6036bdb41');
  expect(signed.headers['x-amz-content-sha256']).toBe(EMPTY_SHA256);
  expect(signed.headers['x-amz-date']).toBe('20130524T000000Z');
  expect('host' in signed.headers).toBe(false);
});

test('S3 paths and queries are encoded once, byte by byte', () => {
  expect(uriEncode('runs/1/a b+c~é', true)).toBe('runs/1/a%20b%2Bc~%C3%A9');
  expect(uriEncode('a/b')).toBe('a%2Fb');
  const { canonicalRequest } = signRequest({
    method: 'GET', url: 'https://h.example/bucket/a%20b?prefix=runs%2F1%2F&list-type=2',
    accessKeyId: 'k', secretAccessKey: 's', region: 'auto', service: 's3', date: new Date(0),
  });
  expect(canonicalRequest.split('\n').slice(0, 3)).toEqual(['GET', '/bucket/a%20b', 'list-type=2&prefix=runs%2F1%2F']);
  expect(amzDate(new Date('2026-11-02T21:30:05.123Z'))).toBe('20261102T213005Z');
});

test('signing refuses missing credentials', () => {
  expect(() => signRequest({ method: 'GET', url: 'https://h.example/', accessKeyId: '', secretAccessKey: 's', region: 'auto', service: 's3' }))
    .toThrow('Signing needs');
});
