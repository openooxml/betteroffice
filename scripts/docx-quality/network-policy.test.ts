import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { runInNewContext } from 'node:vm';
import { version as fontsLatinVersion } from '../../packages/fonts-latin/package.json';
import { version as fontsVersion } from '../../packages/fonts/package.json';
import { version as fontsCjkVersion } from '../../packages/fonts-cjk/package.json';
import { version as fontsIntlVersion } from '../../packages/fonts-intl/package.json';
import { isAllowedFontRequest } from './network-policy.mjs';

const latin = (version = fontsLatinVersion, file = 'Carlito-Regular.ttf.zst') =>
  `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-latin@${version}/assets/${file}`;
const cjk = (version = fontsCjkVersion) =>
  `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-cjk@${version}/assets/NotoSansJP-Regular.otf.zst`;
const intl = (version = fontsIntlVersion) =>
  `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-intl@${version}/assets/NotoSansThai-Regular.ttf.zst`;
const decoder = (version = fontsVersion) =>
  `https://cdn.jsdelivr.net/npm/@betteroffice/fonts@${version}/dist/decoder/fontpack.wasm`;
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
    expect(isAllowedFontRequest(request(intl()))).toBe(true);
    expect(isAllowedFontRequest(request(decoder()))).toBe(true);
  });

  test('blocks other versions and query strings', () => {
    const blocked = [
      latin(next(fontsLatinVersion)),
      latin('0.0.0'),
      latin('latest'),
      latin(fontsLatinVersion.replaceAll('.', 'x')),
      cjk(next(fontsCjkVersion)),
      intl(next(fontsIntlVersion)),
      decoder(next(fontsVersion)),
      decoder('latest'),
      decoder(fontsVersion.replaceAll('.', 'x')),
      `${decoder()}?v=1`,
      `${decoder()}#fragment`,
      `${intl()}?v=1`,
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
      intl().replace('.ttf.zst', '.ttf'),
      cjk().replace('.otf.zst', '.otf'),
      decoder().replace('cdn.jsdelivr.net/npm', 'unpkg.com'),
      decoder().replace('dist/decoder/', 'assets/decoder/'),
      decoder().replace('fontpack.wasm', 'other.wasm'),
      decoder().replace('/fonts@', '/fonts-intl@'),
      `${decoder()}/extra`,
      decoder().replace('https:', 'http:'),
      `${latin()}/extra`,
      'https://fonts.googleapis.com/css2?family=Carlito',
      latin().replace('https:', 'http:'),
    ];
    for (const url of blocked) expect(isAllowedFontRequest(request(url))).toBe(false);
  });

  test('blocks unsafe request properties', () => {
    for (const url of [latin(), cjk(), intl(), decoder()]) {
      expect(isAllowedFontRequest(request(url, { method: 'POST' }))).toBe(false);
      expect(isAllowedFontRequest(request(url, { bodyBytes: 12 }))).toBe(false);
      expect(
        isAllowedFontRequest(request(url, { referer: 'http://127.0.0.1:4178/' })),
      ).toBe(false);
      expect(isAllowedFontRequest(request(url, { cookie: 'a=b' }))).toBe(false);
    }
  });
});

async function serverConfiguration(env: Record<string, string>) {
  const source = readFileSync(new URL('./server.mjs', import.meta.url), 'utf8')
    .replace(/^import .*;\n/gm, '');
  let config: any;
  await runInNewContext(`(async () => { ${source} })()`, {
    createServer: async (options: unknown) => {
      config = options;
      return { listen: async () => {}, resolvedUrls: { local: ['http://127.0.0.1:4178/'] } };
    },
    readFileSync,
    resolve,
    qualityRendererPlugin: () => ({ name: 'quality-renderer' }),
    process: { env, cwd: () => process.cwd() },
    console: { log() {} },
  });
  return config;
}

test('remote captures use the pinned provider defaults', async () => {
  const config = await serverConfiguration({});
  expect(JSON.parse(config.define.__QUALITY_FONT_BASE__)).toBeNull();
  expect(JSON.parse(config.define.__QUALITY_FONT_BASE_CJK__)).toBeNull();
});

for (const format of ['docx', 'pptx', 'xlsx', 'vsdx']) {
  test(`${format} local captures serve packs and the shipped decoder from the font root`, async () => {
    for (const assets of [undefined, resolve('/tmp/quality-font-assets')]) {
      const config = await serverConfiguration({
        QUALITY_FORMAT: format,
        QUALITY_LOCAL_ONLY: '1',
        ...(assets === undefined ? {} : { QUALITY_FONT_ASSETS: assets }),
      });
      const prefix = `/@fs${assets ?? resolve('packages/fonts-latin/assets')}/`;
      expect(JSON.parse(config.define.__QUALITY_FONT_BASE__)).toBe(prefix);
      expect(JSON.parse(config.define.__QUALITY_FONT_BASE_CJK__))
        .toBe(`/@fs${resolve('packages/fonts-cjk/assets')}/`);
      let middleware!: (request: { url: string }, response: unknown, next: () => void) => void;
      config.plugins[0].configureServer({
        middlewares: { use: (handler: typeof middleware) => { middleware = handler; } },
      });
      for (const [file, target] of [
        ['decoder/fontpack.wasm', `/@fs${resolve('packages/fonts/dist/decoder/fontpack.wasm')}`],
        ['NotoSansThai-Regular.ttf.zst', `/@fs${resolve('packages/fonts-intl/assets/NotoSansThai-Regular.ttf.zst')}`],
        ['Carlito-Regular.ttf.zst', `${prefix}Carlito-Regular.ttf.zst`],
        ['decoder/other.wasm', `${prefix}decoder/other.wasm`],
      ]) {
        const request = { url: `${prefix}${file}` };
        let continued = false;
        middleware(request, {}, () => { continued = true; });
        expect(request.url).toBe(target);
        expect(continued).toBe(true);
      }
    }
  });
}
