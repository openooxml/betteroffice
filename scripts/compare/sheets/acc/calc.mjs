// Univer arm: convert each cache-cleared input.xlsx with xlsx2univer.mjs, open it on univer-calc.html with forced
// initial computing, read every scored cell. Writes univer.json, univer.report.json and univer.values.json per workbook.
// Usage (node): calc.mjs <work dir>   env: CALC_WORKERS (3), CALC_CAP_MS (120000), CALC_ONLY (comma ids)
import { chromium } from 'playwright';
import { createServer } from 'node:http';
import { access, readFile, readdir, writeFile } from 'node:fs/promises';
import { extname, relative, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';

const work = resolve(process.argv[2]);
const dist = resolve(import.meta.dirname, '../dist');
const workers = Number(process.env.CALC_WORKERS ?? 4);
const cap = Number(process.env.CALC_CAP_MS ?? 60_000);
const only = process.env.CALC_ONLY?.split(',');
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json', '.wasm': 'application/wasm' };
const server = createServer(async (request, response) => {
  const pathname = decodeURIComponent(new URL(request.url, 'http://x').pathname);
  const path = pathname.startsWith('/docs/') ? resolve(work, `.${pathname.slice(5)}`) : resolve(dist, `.${pathname}`);
  if (relative(pathname.startsWith('/docs/') ? work : dist, path).startsWith('..')) { response.writeHead(400).end(); return; }
  try { const body = await readFile(path); response.writeHead(200, { 'Content-Type': types[extname(path)] ?? 'application/octet-stream' }).end(body); }
  catch { response.writeHead(404).end(); }
});
await new Promise((done) => server.listen(0, '127.0.0.1', done));
const port = server.address().port;

let executablePath = process.env.CHROME;
if (!executablePath && process.platform === 'linux') {
  for (const entry of (await readdir('/opt/ms-playwright')).filter((n) => n.startsWith('chromium-')).sort().reverse()) {
    for (const binary of ['chrome-linux/chrome', 'chrome-linux64/chrome']) {
      const path = resolve('/opt/ms-playwright', entry, binary);
      try { await access(path); executablePath = path; break; } catch {}
    }
    if (executablePath) break;
  }
}
const browser = await chromium.launch({ executablePath, headless: true, args: ['--disable-background-timer-throttling', '--disable-renderer-backgrounding'] });
const timeout = (ms, label) => new Promise((_, fail) => setTimeout(() => fail(new Error(label)), ms));

const jobs = [];
for (const set of ['sp', 'mf']) {
  for (const id of (await readdir(resolve(work, set)).catch(() => [])).sort()) {
    if ((!only || only.includes(id)) && !existsSync(resolve(work, set, id, 'univer.values.json'))) jobs.push([set, id]);
  }
}

async function one(set, id) {
  const root = resolve(work, set, id);
  const expected = JSON.parse(await readFile(resolve(root, 'expected.json'), 'utf8'));
  const record = { id, status: 'ok', logs: [] };
  const started = Date.now();
  const marker = resolve(root, 'univer.started');
  if (existsSync(marker)) {
    await writeFile(resolve(root, 'univer.values.json'), JSON.stringify({ id, status: 'driver-crash', values: null }));
    return { id, status: 'driver-crash' };
  }
  await writeFile(marker, '');
  const child = spawnSync(process.execPath, ['--max-old-space-size=6144', resolve(import.meta.dirname, '../xlsx2univer.mjs'),
    resolve(root, 'input.xlsx'), resolve(root, 'univer.json'), resolve(root, 'univer.report.json')], { timeout: 300_000, encoding: 'utf8', maxBuffer: 1 << 24 });
  record.convertMs = Date.now() - started;
  if (child.status !== 0) {
    record.status = 'convert-failed';
    record.error = `${child.signal ?? child.status}: ${(child.stderr ?? '').split('\n').filter((l) => /Error|error/.test(l)).slice(0, 2).join(' | ')}`.slice(0, 400);
    await writeFile(resolve(root, 'univer.values.json'), JSON.stringify({ ...record, values: null }));
    return record;
  }
  const groups = new Map();
  expected.cells.forEach((cell, index) => {
    if (!groups.has(cell.name)) groups.set(cell.name, { sheet: cell.name, cells: [], index: [] });
    const g = groups.get(cell.name);
    g.cells.push([cell.r, cell.c]);
    g.index.push(index);
  });
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();
  page.on('pageerror', (error) => record.logs.length < 20 && record.logs.push(`pageerror: ${error.message.slice(0, 200)}`));
  page.on('console', (m) => m.type() === 'error' && record.logs.length < 20 && record.logs.push(`console: ${m.text().slice(0, 200)}`));
  try {
    await page.goto(`http://127.0.0.1:${port}/univer-calc.html?wb=${set}/${id}/univer&cap=${cap}`, { timeout: 60_000 });
    await page.waitForFunction(() => window.calc, undefined, { timeout: 60_000 });
    record.info = await Promise.race([page.evaluate(() => window.calc.ready), timeout(cap + 60_000, 'ready timeout')]);
    const targets = [...groups.values()].map(({ sheet, cells }) => ({ sheet, cells }));
    const read = await Promise.race([page.evaluate((t) => window.calc.read(t), targets), timeout(120_000, 'read timeout')]);
    const values = new Array(expected.cells.length);
    [...groups.values()].forEach((g, gi) => g.index.forEach((index, k) => { values[index] = read[gi][k]; }));
    record.values = values;
    if (record.info?.timedOut) record.status = 'calc-timeout';
  } catch (error) {
    record.status = /timeout/i.test(String(error)) ? 'timeout' : 'failed';
    record.error = String(error).slice(0, 300);
    record.values = null;
  }
  record.errors = await page.evaluate(() => window.calc?.errors ?? []).catch(() => []);
  record.ms = Date.now() - started;
  await context.close().catch(() => {});
  await writeFile(resolve(root, 'univer.values.json'), JSON.stringify(record));
  return record;
}

const summary = [];
let done = 0;
await Promise.all(Array.from({ length: workers }, async () => {
  for (let job; (job = jobs.shift());) {
    const r = await one(...job);
    summary.push({ id: r.id, status: r.status, ms: r.ms, calcMs: r.info?.calcMs });
    if (++done % 25 === 0 || r.status !== 'ok') console.log(`univer calc ${done}: ${r.id} ${r.status} ${r.ms ?? ''}ms ${r.error?.slice(0, 120) ?? ''}`);
  }
}));
await writeFile(resolve(work, `univer-calc-${Date.now()}.json`), JSON.stringify({ chromium: browser.version(), summary }, null, 1));
console.log(`univer calc: ${summary.filter((s) => s.status === 'ok').length}/${summary.length} ok`);
await browser.close();
server.close();
process.exit(0);
