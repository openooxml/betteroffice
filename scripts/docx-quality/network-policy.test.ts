import { describe, expect, test } from 'bun:test';
import { version as fontsVersion } from '../../packages/fonts/package.json';
import { version as fontsCjkVersion } from '../../packages/fonts-cjk/package.json';
import { isAllowedFontRequest } from './network-policy.mjs';

const latin = (version = fontsVersion, file = 'Carlito-Regular.ttf') =>
  `https://cdn.jsdelivr.net/npm/@betteroffice/fonts@${version}/assets/${file}`;
const cjk = (version = fontsCjkVersion) =>
  `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-cjk@${version}/assets/NotoSansJP-Regular.otf`;
const next = (version: string) => version.replace(/\d+$/, (patch) => String(Number(patch) + 1));

function request(url, overrides = {}) {
  return {
    url,
    method: 'GET',
    bodyBytes: 0,
    referer: null,
    cookie: null,
    ...overrides,
  };
}

describe('font network policy', () => {
  test('allows the font package versions the capture harness requests', () => {
    expect(isAllowedFontRequest(request(latin()))).toBe(true);
    expect(isAllowedFontRequest(request(cjk()))).toBe(true);
  });

  test('blocks other versions and query strings', () => {
    const blocked = [
      latin(next(fontsVersion)),
      latin('0.0.0'),
      latin('latest'),
      latin(fontsVersion.replaceAll('.', 'x')),
      cjk(next(fontsCjkVersion)),
      `${latin()}?v=1`,
      `${latin()}?foo=bar`,
    ];
    for (const url of blocked) expect(isAllowedFontRequest(request(url))).toBe(false);
  });

  test('blocks other hosts and asset paths', () => {
    const blocked = [
      latin().replace('cdn.jsdelivr.net/npm', 'unpkg.com'),
      latin().replace('assets/Carlito-Regular.ttf', 'dist/cdn.js'),
      latin(fontsVersion, 'Carlito-Regular.woff2'),
      `${latin()}/extra`,
      'https://fonts.googleapis.com/css2?family=Carlito',
      latin().replace('https:', 'http:'),
    ];
    for (const url of blocked) expect(isAllowedFontRequest(request(url))).toBe(false);
  });

  test('blocks unsafe request properties', () => {
    const url = latin();
    expect(isAllowedFontRequest(request(url, { method: 'POST' }))).toBe(false);
    expect(isAllowedFontRequest(request(url, { bodyBytes: 12 }))).toBe(false);
    expect(
      isAllowedFontRequest(request(url, { referer: 'http://127.0.0.1:4178/' })),
    ).toBe(false);
    expect(isAllowedFontRequest(request(url, { cookie: 'a=b' }))).toBe(false);
  });
});
