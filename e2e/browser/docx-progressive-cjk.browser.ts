import { readFile } from 'node:fs/promises';
import { extname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test, type Browser } from 'playwright/test';
import JSZip from 'jszip';

const japanese = '鬱蒼とした森の奥に、齟齬と矛盾が潜む。';
const traditional = '繁體中文的排版測試 ♠♣♥';
const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
/** The editor already fetches the opened file from a blob: URL; data: stays refused, so inlined font assets must not be fetched. */
const strictConnect = "connect-src 'self' blob:";
const origin = 'http://127.0.0.1:4186';
const dist = fileURLToPath(new URL('../../apps/desktop/dist/', import.meta.url));
const types: Record<string, string> = {
  '.html': 'text/html', '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css',
  '.json': 'application/json', '.wasm': 'application/wasm', '.svg': 'image/svg+xml', '.png': 'image/png',
  '.woff2': 'font/woff2',
};

/** Serves the built app from disk, as the preview server does, so no response is proxied through the test process. */
async function fromDist(pathname: string): Promise<{ status: number; type: string; body: Buffer | string }> {
  const file = resolve(dist, `.${decodeURIComponent(pathname)}`);
  const page = !file.startsWith(dist) || extname(file) === '';
  const path = page ? resolve(dist, 'index.html') : file;
  const type = types[extname(path)] ?? 'application/octet-stream';
  try {
    return { status: 200, type, body: await readFile(path) };
  } catch {
    return { status: 404, type: 'text/plain', body: 'not found' };
  }
}

async function documentBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', rels(`<Relationship Id="doc" Type="${relationships}/officeDocument" Target="word/document.xml"/>`));
  const run = (family: string, lang: string, text: string) =>
    `<w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:eastAsia="${family}" w:cs="${family}"/><w:sz w:val="40"/><w:lang w:eastAsia="${lang}"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const blank = '<w:p><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r></w:p>';
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${blank.repeat(6)}${run('MS Gothic', 'ja-JP', japanese)}${run('Microsoft JhengHei', 'zh-TW', traditional)}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

interface Opened {
  requests: string[];
  violations: string[];
  errors: string[];
  pixels: string;
}

/** Opens the document under a restrictive `connect-src`; `whole` makes the sliced faces unavailable so they load whole. */
async function open(browser: Browser, whole: boolean): Promise<Opened> {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  const page = await context.newPage();
  const requests: string[] = [];
  const violations: string[] = [];
  const errors: string[] = [];
  let lastRequest = Date.now();
  try {
    await context.route('**/*', async (route) => {
      const url = new URL(route.request().url());
      if (url.origin !== origin) return route.continue();
      if (whole && /\/slices(-[\w-]+)?\.json$/.test(url.pathname)) return route.abort();
      const { status, type, body } = await fromDist(url.pathname);
      const policy = /html|javascript/.test(type) ? { 'content-security-policy': strictConnect } : {};
      return route.fulfill({ status, body, headers: { 'content-type': type, ...policy } });
    });
    context.on('request', (request) => {
      lastRequest = Date.now();
      requests.push(request.url());
    });
    page.on('pageerror', (error) => errors.push(error.message));
    page.on('console', (message) => {
      if (/Content Security Policy|Refused to connect/i.test(message.text())) violations.push(message.text());
    });
    await page.goto('/?format=docx');
    await page.evaluate(() => {
      document.addEventListener('securitypolicyviolation', (event) => console.error(`Content Security Policy ${event.violatedDirective} ${event.blockedURI}`));
    });
    await page.locator('input[type=file]').last().setInputFiles({
      name: 'progressive-cjk.docx', mimeType: 'application/octet-stream', buffer: await documentBytes(),
    });
    const canvas = page.locator('.paged-editor__page canvas').first();
    await expect(canvas).toBeVisible({ timeout: 45_000 });
    await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false', { timeout: 45_000 });
    await expect(page.locator('.canvas-page-mirror').first()).toContainText(japanese);
    await expect(page.locator('.canvas-page-mirror').first()).toContainText(traditional);
    await expect.poll(() => Date.now() - lastRequest, { timeout: 40_000 }).toBeGreaterThanOrEqual(3_000);
    const pixels = await canvas.evaluate(async (element: HTMLCanvasElement) => {
      const top = Math.floor(element.height * 0.17);
      const region = element.getContext('2d')!.getImageData(0, top, element.width, element.height - top).data;
      let ink = 0;
      for (let index = 0; index < region.length; index += 4) if (region[index + 3]! > 0 && region[index]! < 128) ink += 1;
      const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', region));
      return `${element.width}x${element.height} ink=${ink} ${Array.from(digest, (byte) => byte.toString(16).padStart(2, '0')).join('')}`;
    });
    return { requests, violations, errors, pixels };
  } finally {
    await context.close();
  }
}

test('sliced CJK faces paint the first frame under a strict connect-src exactly as whole faces do', async ({ browser }) => {
  test.setTimeout(150_000);
  const sliced = await open(browser, false);
  const paths = sliced.requests.map((url) => new URL(url).pathname);
  expect(paths.filter((path) => /\/slices(-[\w-]+)?\.json$/.test(path)).length).toBeGreaterThanOrEqual(2);
  expect(paths.filter((path) => /\/NotoSans(JP|TC)-Regular\.otf(-[\w-]+)?\.zst$/.test(path))).toEqual([]);
  expect(sliced.requests.filter((url) => new URL(url).origin !== origin && !url.startsWith('data:'))).toEqual([]);
  expect(sliced.violations).toEqual([]);
  expect(sliced.errors).toEqual([]);
  expect(sliced.pixels).toMatch(/ ink=[1-9]\d* /);

  const whole = await open(browser, true);
  expect(whole.requests.map((url) => new URL(url).pathname)
    .filter((path) => /\/NotoSans(JP|TC)-Regular\.otf(-[\w-]+)?\.zst$/.test(path))).toHaveLength(2);
  expect(whole.violations).toEqual([]);
  expect(sliced.pixels).toBe(whole.pixels);
});

test('typing a new CJK script refreshes the fonts with sliced faces', async ({ page }) => {
  test.setTimeout(90_000);
  const requests: string[] = [];
  const rejected: string[] = [];
  page.on('request', (request) => requests.push(new URL(request.url()).pathname));
  page.on('console', (message) => {
    if (/font environment was rejected/i.test(message.text())) rejected.push(message.text());
  });
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'progressive-cjk.docx', mimeType: 'application/octet-stream', buffer: await documentBytes(),
  });
  const canvas = page.locator('.paged-editor__page canvas').first();
  await expect(canvas).toBeVisible({ timeout: 45_000 });
  await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false', { timeout: 45_000 });
  const opened = requests.filter((path) => /\/slices(-[\w-]+)?\.json$/.test(path)).length;
  await canvas.click({ position: { x: 200, y: 200 } });
  await page.keyboard.insertText('\ud55c\uae00');
  await expect.poll(() => requests.filter((path) => /\/slices(-[\w-]+)?\.json$/.test(path)).length, { timeout: 30_000 })
    .toBeGreaterThan(opened);
  await expect(page.locator('.canvas-page-mirror').first()).toContainText('\ud55c\uae00');
  await expect.poll(() => requests.filter((path) => /\/NotoSansKR-Regular\.otf(-[\w-]+)?\.zst$/.test(path)).length).toBe(0);
  expect(rejected).toEqual([]);
});
