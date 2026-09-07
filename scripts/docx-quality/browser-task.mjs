import { chromium } from 'playwright';
import { mkdir, writeFile, readFile, readdir } from 'node:fs/promises';
import { basename, resolve } from 'node:path';
import { createHash } from 'node:crypto';

const [file, out, fontMode = 'cdn', base = 'http://127.0.0.1:4178'] =
  process.argv.slice(2);
if (!file || !out)
  throw new Error(
    'usage: browser-task.mjs input.docx output-directory [cdn|none] [server-url]'
  );
if (!['cdn', 'none'].includes(fontMode)) throw new Error('font mode must be cdn or none');
const server = new URL(base);
if (
  server.protocol !== 'http:' ||
  !['127.0.0.1', 'localhost', '[::1]'].includes(server.hostname)
)
  throw new Error('capture server must use local loopback HTTP');
if ((await readdir(out).catch(() => [])).length) throw new Error('output must be empty');
const source = await readFile(file);
const sha256 = createHash('sha256').update(source).digest('hex');
const metadata = {
  source: basename(file),
  sha256,
  dpi: 150,
  engine: process.env.QUALITY_ENGINE_LABEL ?? 'BetterOffice DOCX working tree',
};
const timeout = Number(process.env.QUALITY_TIMEOUT_SECONDS ?? 600);
if (!Number.isFinite(timeout) || timeout <= 0)
  throw new Error('QUALITY_TIMEOUT_SECONDS must be positive');
await mkdir(out, { recursive: true });
await writeFile(
  resolve(out, 'result.json'),
  JSON.stringify({ ...metadata, status: 'running' })
);
const start = performance.now();
let browser;
const deadline = setTimeout(async () => {
  console.error('Browser capture exceeded its deadline');
  await writeFile(
    resolve(out, 'result.json'),
    JSON.stringify({ ...metadata, status: 'timeout' })
  );
  await browser?.close();
  process.exit(1);
}, timeout * 1000);
const logs = [];
const externalRequests = [];
const networkViolations = [];
try {
  browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({
    viewport: { width: 1400, height: 1200 },
    deviceScaleFactor: 150 / 96,
    serviceWorkers: 'block',
  });
  await page.route('**/*', async (route) => {
    const request = route.request();
    const url = new URL(request.url());
    if (url.origin === server.origin || !['http:', 'https:'].includes(url.protocol)) {
      await route.continue();
      return;
    }
    const record = {
      url: url.href,
      method: request.method(),
      bodyBytes: request.postDataBuffer()?.length ?? 0,
      referer: request.headers().referer ?? null,
      cookie: request.headers().cookie ?? null,
    };
    externalRequests.push(record);
    const font =
      /^https:\/\/cdn\.jsdelivr\.net\/npm\/@betteroffice\/fonts(?:-cjk)?@0\.1\.0\/assets\/[A-Za-z0-9-]+\.(?:ttf|otf)$/.test(
        url.href
      );
    if (
      !font ||
      record.method !== 'GET' ||
      record.bodyBytes !== 0 ||
      record.referer ||
      record.cookie
    ) {
      networkViolations.push(record);
      await route.abort();
      return;
    }
    await route.continue();
  });
  page.on('console', (message) => {
    if (['error', 'warning'].includes(message.type())) {
      logs.push(message.text());
      if (logs.length <= 40) console.error(message.text());
    }
  });
  page.on('pageerror', (error) => {
    logs.push(error.message);
    console.error(error.message);
  });
  page.on('requestfailed', (request) =>
    logs.push(`${request.url()}: ${request.failure()?.errorText}`)
  );
  await page.goto(base, { timeout: 30_000 });
  await page.waitForFunction(() => window.oracleReady === true, undefined, {
    timeout: 30_000,
  });
  console.error('editor');
  const result = await page.evaluate(
    ([bytes, fonts]) => window.oracleInit(bytes, fonts),
    [Array.from(source), fontMode === 'cdn']
  );
  console.error(`paint ${result.pages}`);
  await mkdir(out, { recursive: true });
  for (let index = 0; index < result.pages; index++) {
    const png = await page.evaluate((i) => window.oraclePage(i), index);
    await writeFile(
      resolve(out, `page_${String(index + 1).padStart(4, '0')}.png`),
      Buffer.from(png.split(',')[1], 'base64')
    );
  }
  if (networkViolations.length)
    throw new Error(`Unexpected external requests: ${JSON.stringify(networkViolations)}`);
  const record = {
    status: 'ok',
    ...metadata,
    ...result,
    logs,
    externalRequests,
    ms: performance.now() - start,
    captured_at_utc: new Date().toISOString(),
    browser: browser.version(),
  };
  await writeFile(resolve(out, 'result.json'), JSON.stringify(record, null, 2));
  console.log(JSON.stringify(record));
} catch (error) {
  await writeFile(
    resolve(out, 'result.json'),
    JSON.stringify(
      { ...metadata, status: 'error', error: String(error), networkViolations },
      null,
      2
    )
  );
  console.log(
    JSON.stringify({
      status: 'error',
      error: String(error),
      logs,
      networkViolations,
      ms: performance.now() - start,
    })
  );
  process.exitCode = 1;
} finally {
  clearTimeout(deadline);
  await browser?.close();
}
