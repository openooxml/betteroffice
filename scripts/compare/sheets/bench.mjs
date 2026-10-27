// Sheets at scale: bo-xlsx vs univer-sheets in one headless Chromium, CDP screencast pixel timing.
// Per session (fresh browser): open + accepts-input, settle, memory, wheel scroll, 10 cell edits, Ctrl+Down jump,
// recalc of the J running-total chain, memory. Timeouts are kept as null + timedOut and rank as +inf.
import { chromium } from 'playwright';
import { access, mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import { readFileSync, readdirSync } from 'node:fs';
import { extname, relative, resolve } from 'node:path';
import { execSync } from 'node:child_process';
import os from 'node:os';
import { WORKBOOKS, numeric, total, value } from './gen.mjs';

const env = process.env;
const here = import.meta.dirname;
const dist = resolve(here, 'dist');
const gen = resolve(here, 'gen');
const output = resolve(here, env.CMP_OUT ?? 'out/local');
const smoke = env.CMP_SMOKE === '1';
const reps = Number(env.CMP_N ?? (smoke ? 1 : 6));
const burners = Number(env.CMP_BURN ?? (smoke ? 0 : 1));
const edits = Number(env.CMP_EDITS ?? 10);
const wbs = (env.CMP_WBS ?? 'w100k').split(',');
const arms = (env.CMP_ARMS ?? 'bo-xlsx,univer-sheets').split(',');
const budgetMs = Number(env.CMP_WB_BUDGET_MS ?? 55 * 60_000);
const edge = env.CMP_EDGE === '1' ? JSON.parse(readFileSync(resolve(gen, 'edge.json'), 'utf8')) : null;
const colName = (c) => { let s = ''; for (c += 1; c; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s; };
const viewport = { width: 1180, height: 760 };
const OPEN_CEIL = 60_000, KEY_CEIL = 30_000, RECALC_CEIL = 30_000, JUMP_CEIL = 30_000;
const P = { x: 400, y: 400 };
const MARKER = '--cmp-sheets-bench';
await mkdir(resolve(output, 'shots'), { recursive: true });

const sleep = (ms) => new Promise((done) => setTimeout(done, ms));
const within = (promise, ms, label) => Promise.race([promise, sleep(ms).then(() => { throw new Error(`timeout ${label} ${ms}ms`); })]);
const loadavg = () => os.loadavg().map((v) => v.toFixed(2)).join(' ');

const types = { '.html': 'text/html', '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css',
  '.wasm': 'application/wasm', '.json': 'application/json', '.woff2': 'font/woff2', '.ttf': 'font/ttf', '.svg': 'image/svg+xml' };
const server = Bun.serve({ hostname: '127.0.0.1', port: 0, async fetch(request) {
  const pathname = decodeURIComponent(new URL(request.url).pathname);
  const path = pathname.startsWith('/docs/') ? resolve(gen, `.${pathname.slice(5)}`) : resolve(dist, `.${pathname}`);
  if (relative(here, path).startsWith('..')) return new Response('Invalid path', { status: 400 });
  const file = Bun.file(path);
  if (!await file.exists()) return new Response('Missing', { status: 404 });
  return new Response(file, { headers: { 'Content-Type': types[extname(path)] ?? 'application/octet-stream' } });
} });

let executablePath = env.CHROME;
if (!executablePath && process.platform === 'linux') {
  for (const entry of (await readdir('/opt/ms-playwright')).filter((n) => n.startsWith('chromium-')).sort().reverse()) {
    for (const binary of ['chrome-linux/chrome', 'chrome-linux64/chrome']) {
      const path = resolve('/opt/ms-playwright', entry, binary);
      try { await access(path); executablePath = path; break; } catch {}
    }
    if (executablePath) break;
  }
}
const launchArgs = ['--disable-background-timer-throttling', '--disable-renderer-backgrounding', '--disable-backgrounding-occluded-windows'];

// Analysis runs in its own browser so the measured browser's process tree and CPU stay clean.
const analysisBrowser = await chromium.launch({ executablePath, headless: true });
const analysis = await (await analysisBrowser.newContext()).newPage();
await analysis.evaluate(() => {
  const cache = new Map();
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d', { willReadFrequently: true });
  async function pixels(source) {
    if (cache.has(source)) return cache.get(source);
    const image = new Image(); image.src = `data:image/jpeg;base64,${source}`; await image.decode();
    canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
    context.drawImage(image, 0, 0);
    const data = context.getImageData(0, 0, canvas.width, canvas.height);
    const luma = new Uint8Array(data.width * data.height);
    for (let i = 0, j = 0; j < luma.length; i += 4, j++) luma[j] = (54 * data.data[i] + 183 * data.data[i + 1] + 19 * data.data[i + 2]) >> 8;
    const out = { w: data.width, h: data.height, luma };
    cache.set(source, out);
    if (cache.size > 6) cache.delete(cache.keys().next().value);
    return out;
  }
  const clip = (img, r) => ({ x0: Math.max(0, Math.round(r.x)), y0: Math.max(0, Math.round(r.y)),
    x1: Math.min(img.w, Math.round(r.x + r.width)), y1: Math.min(img.h, Math.round(r.y + r.height)) });
  window.diff = async (a, b, r) => {
    const A = await pixels(a), B = await pixels(b);
    const { x0, y0, x1, y1 } = clip(A, r);
    let n = 0, bx0 = 1e9, by0 = 1e9, bx1 = -1, by1 = -1;
    for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) {
      const i = y * A.w + x;
      if (Math.abs(A.luma[i] - B.luma[i]) > 40) { n++; if (x < bx0) bx0 = x; if (x > bx1) bx1 = x; if (y < by0) by0 = y; if (y > by1) by1 = y; }
    }
    return { n, bw: n ? bx1 - bx0 + 1 : 0, bh: n ? by1 - by0 + 1 : 0 };
  };
  window.ink = async (a, r) => {
    const A = await pixels(a);
    const { x0, y0, x1, y1 } = clip(A, r);
    let dark = 0, light = 0, all = 0;
    for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) {
      const v = A.luma[y * A.w + x]; all++;
      if (v < 128) dark++;
      if (v < 235) light++;
    }
    return { dark: dark / Math.max(1, all), light: light / Math.max(1, all) };
  };
});
const diff = (a, b, r) => analysis.evaluate(([x, y, z]) => window.diff(x, y, z), [a.data, b.data, r]);
const inkOf = (a, r) => analysis.evaluate(([x, z]) => window.ink(x, z), [a.data, r]);

function before(frames, t) {
  let found = null;
  for (const f of frames) { if (f.ts < t) found = f; else break; }
  return found;
}
/** First frame at or after t whose region differs from the screen state just before t by more than a caret. */
async function firstChange(frames, t, region, until) {
  const base = before(frames, t);
  if (!base) return null;
  for (const f of frames) {
    if (f.ts < t) continue;
    if (f.ts > until) break;
    const d = await diff(base, f, region);
    if (d.n >= 6 && d.bw >= 3 && d.bh >= 4) return f;
  }
  return null;
}
/** First frame at or after t whose region already matches `target` (ignoring caret-sized differences). */
async function firstMatch(frames, t, region, target, until = Infinity) {
  for (const f of frames) {
    if (f.ts < t) continue;
    if (f.ts > until) break;
    const d = await diff(f, target, region);
    if (d.n <= 8 || d.bw <= 2) return f;
  }
  return null;
}

const KEYS = { Enter: ['Enter', 13, '\r'], Escape: ['Escape', 27], ArrowDown: ['ArrowDown', 40] };
function sendKey(cdp, key, modifiers = 0) {
  const [code, vk, text] = KEYS[key] ?? [`Digit${key}`, key.charCodeAt(0), key];
  const base = { key, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers };
  const sent = Date.now();
  if (modifiers) cdp.send('Input.dispatchKeyEvent', { type: 'rawKeyDown', key: 'Control', code: 'ControlLeft', windowsVirtualKeyCode: 17, modifiers }).catch(() => {});
  cdp.send('Input.dispatchKeyEvent', { ...base, type: text && !modifiers ? 'keyDown' : 'rawKeyDown', text: modifiers ? undefined : text, unmodifiedText: text }).catch(() => {});
  cdp.send('Input.dispatchKeyEvent', { ...base, type: 'keyUp' }).catch(() => {});
  if (modifiers) cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Control', code: 'ControlLeft', windowsVirtualKeyCode: 17 }).catch(() => {});
  return sent;
}
function sendClick(cdp, x, y) {
  const sent = Date.now();
  cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y }).catch(() => {});
  cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', buttons: 1, clickCount: 1 }).catch(() => {});
  cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', buttons: 0, clickCount: 1 }).catch(() => {});
  return sent;
}
const evaluate = (page, fn, arg, ms = 5000) => within(page.evaluate(fn, arg), ms, 'evaluate');

function processes() {
  const out = new Map();
  for (const name of readdirSync('/proc')) {
    if (!/^\d+$/.test(name)) continue;
    try {
      const stat = readFileSync(`/proc/${name}/stat`, 'utf8');
      const ppid = Number(stat.slice(stat.lastIndexOf(')') + 2).split(' ')[1]);
      out.set(Number(name), { ppid, cmd: readFileSync(`/proc/${name}/cmdline`, 'utf8').replace(/\0/g, ' ') });
    } catch {}
  }
  return out;
}
function pssKb(pid) {
  try { return Number(/^Pss:\s+(\d+)/m.exec(readFileSync(`/proc/${pid}/smaps_rollup`, 'utf8'))?.[1] ?? NaN); } catch { return NaN; }
}
function rendererMemory() {
  if (process.platform !== 'linux') return null;
  const all = processes();
  const roots = [...all].filter(([, p]) => p.cmd.includes(MARKER) && !p.cmd.includes('--type=')).map(([pid]) => pid);
  const tree = new Set(roots);
  for (let grew = true; grew;) {
    grew = false;
    for (const [pid, p] of all) if (!tree.has(pid) && tree.has(p.ppid)) { tree.add(pid); grew = true; }
  }
  let renderers = 0, everything = 0, count = 0;
  for (const pid of tree) {
    const kb = pssKb(pid);
    if (!Number.isFinite(kb)) continue;
    everything += kb;
    if (all.get(pid).cmd.includes('--type=renderer')) { renderers += kb; count++; }
  }
  return { rendererPssMb: +(renderers / 1024).toFixed(1), rendererCount: count, treePssMb: +(everything / 1024).toFixed(1) };
}
async function heaps(browser, cdp) {
  const out = { pageMb: null, workers: [] };
  try {
    await cdp.send('HeapProfiler.collectGarbage');
    out.pageMb = +((await cdp.send('Runtime.getHeapUsage')).usedSize / 1048576).toFixed(1);
  } catch (error) { out.pageError = String(error).slice(0, 120); }
  let session;
  try {
    session = await browser.newBrowserCDPSession();
    const { targetInfos } = await session.send('Target.getTargets');
    let id = 0;
    for (const target of targetInfos.filter((t) => t.type === 'worker' || t.type === 'shared_worker')) {
      const { sessionId } = await session.send('Target.attachToTarget', { targetId: target.targetId, flatten: false });
      const call = (method) => new Promise((done, fail) => {
        const mine = ++id;
        const listener = (event) => {
          if (event.sessionId !== sessionId) return;
          const message = JSON.parse(event.message);
          if (message.id !== mine) return;
          session.off('Target.receivedMessageFromTarget', listener);
          message.error ? fail(new Error(message.error.message)) : done(message.result);
        };
        session.on('Target.receivedMessageFromTarget', listener);
        session.send('Target.sendMessageToTarget', { sessionId, message: JSON.stringify({ id: mine, method }) }).catch(fail);
      });
      try {
        await within(call('HeapProfiler.collectGarbage'), 10_000, 'worker gc');
        const usage = await within(call('Runtime.getHeapUsage'), 10_000, 'worker heap');
        out.workers.push({ url: target.url.replace(/^.*\//, '').slice(0, 60), mb: +(usage.usedSize / 1048576).toFixed(1) });
      } catch (error) { out.workers.push({ url: target.url.slice(-60), error: String(error).slice(0, 100) }); }
      await session.send('Target.detachFromTarget', { sessionId }).catch(() => {});
    }
  } catch (error) { out.workerError = String(error).slice(0, 120); }
  await session?.detach().catch(() => {});
  out.workersMb = +out.workers.reduce((s, w) => s + (w.mb ?? 0), 0).toFixed(1);
  return out;
}
const memory = async (browser, cdp) => ({ ...rendererMemory(), heap: await heaps(browser, cdp) });

const inner = (g) => ({ x: g.x + 60, y: g.y + 30, width: g.width - 80, height: g.height - 50 });
const cropAt = (p) => ({ x: p.x - 60, y: p.y - 14, width: 220, height: 28 });
const timedOut = (extra = {}) => ({ ms: null, timedOut: true, ...extra });

async function runSession(arm, wb, rep, measured) {
  const rows = edge ? edge[wb].rows : WORKBOOKS[wb];
  const session = { arm, wb, rep, measured, started: new Date().toISOString(), loadStart: loadavg(), errors: [] };
  const modified = new Map();
  const model = (r, c) => (modified.has(`${r}:${c}`) ? modified.get(`${r}:${c}`) : value(r, c));
  const browser = await chromium.launch({ executablePath, headless: true, args: [...launchArgs, MARKER] });
  session.chromium = browser.version();
  const context = await browser.newContext({ viewport, deviceScaleFactor: 1 });
  const page = await context.newPage();
  page.on('pageerror', (error) => session.errors.push(String(error).slice(0, 200)));
  const cdp = await context.newCDPSession(page);
  const frames = [];
  cdp.on('Page.screencastFrame', ({ data, metadata, sessionId }) => {
    frames.push({ ts: metadata.timestamp * 1000, data });
    cdp.send('Page.screencastFrameAck', { sessionId }).catch(() => {});
  });
  await cdp.send('Page.startScreencast', { format: 'jpeg', quality: 85, maxWidth: viewport.width, maxHeight: viewport.height, everyNthFrame: 1 });
  const navStart = Date.now();
  page.goto(`http://127.0.0.1:${server.port}/${arm}.html?wb=${wb}`, { waitUntil: 'commit', timeout: OPEN_CEIL }).catch((error) => session.errors.push(`goto: ${error.message.slice(0, 120)}`));

  // Open: wait for the probe, then click + digit every 250 ms until the in-cell editor shows the digit.
  let t0 = null;
  while (Date.now() - navStart < OPEN_CEIL) {
    t0 = await evaluate(page, () => window.cmp?.t0Epoch ?? null, null, 2000).catch(() => null);
    if (t0) break;
    await sleep(50);
  }
  session.t0Offset = t0 ? Math.round(t0 - navStart) : null;
  const deadline = (t0 ?? navStart) + OPEN_CEIL;
  const readyPromise = t0 ? within(page.evaluate(() => window.cmp.ready.then(() => window.cmp.readyEpoch)), Math.max(1, deadline - Date.now()), 'ready') : Promise.reject(new Error('no probe'));
  let readyEpoch = null;
  readyPromise.then((v) => { readyEpoch = v; }, () => {});
  const keys = [];
  let accepted = null, editorKnown = false;
  while (t0 && Date.now() < deadline && !accepted) {
    const tick = Date.now();
    sendClick(cdp, P.x, P.y);
    await sleep(120);
    keys.push(sendKey(cdp, '7'));
    while (Date.now() - tick < 250) {
      const text = await evaluate(page, () => window.cmp.editorText(), null, 200).catch(() => undefined);
      if (text !== null && text !== undefined) editorKnown = true;
      if (typeof text === 'string' && text.includes('7')) { accepted = { at: Date.now(), sevens: text.split('7').length - 1 }; break; }
      await sleep(30);
    }
  }
  await readyPromise.catch(() => {});
  session.open = { readyMs: readyEpoch && t0 ? Math.round(readyEpoch - t0) : null, keysSent: keys.length, editorKnown };
  if (!t0 || !readyEpoch) {
    session.open.timedOut = true;
    session.skipped = 'open-timeout';
    await browser.close().catch(() => {});
    session.firstPaint = timedOut(); session.acceptsInput = timedOut();
    return session;
  }
  sendKey(cdp, 'Escape'); await sleep(100); sendKey(cdp, 'Escape');
  if (edge?.[wb].sheet) {
    session.sheet = await evaluate(page, (n) => window.cmp.activateSheet?.(n) ?? false, edge[wb].sheet, 30_000).catch((e) => String(e).slice(0, 120));
    await sleep(2000);
    sendClick(cdp, P.x, P.y); await sleep(200); sendKey(cdp, 'Escape');
  }
  await sleep(3000);
  const grid = await evaluate(page, () => window.cmp.gridRect());
  session.grid = grid;
  session.info = await evaluate(page, () => window.cmp.info).catch(() => null);
  const gridInner = inner(grid);
  const probeCell = await evaluate(page, () => window.cmp.active()).catch(() => null);
  if (probeCell) {
    const a1 = `${colName(probeCell.col)}${probeCell.row + 1}`;
    const v = await evaluate(page, (x) => window.cmp.read(x), a1).catch((e) => ({ error: String(e) }));
    session.acceptCell = { ...probeCell, a1, value: v };
    if (!edge && probeCell.col < 8) modified.set(`${probeCell.row}:${probeCell.col}`, num(v) ?? model(probeCell.row, probeCell.col));
  }
  session.memSettled = await memory(browser, cdp);
  if (rep === 1 || smoke) await page.screenshot({ path: resolve(output, 'shots', `${arm}-${wb}-open.png`) }).catch(() => {});
  const openFrames = frames.splice(0);

  // Scroll: API to the middle, then 5 s of steady wheel at P.
  await evaluate(page, (r) => window.cmp.scrollToRow(r), Math.floor(rows / 2), 30_000).catch((e) => session.errors.push(`scrollToRow: ${String(e).slice(0, 120)}`));
  await sleep(1000);
  await evaluate(page, () => window.cmp.rafStart());
  const scrollStart = Date.now();
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: P.x, y: P.y }).catch(() => {});
  for (let i = 0; i < 100; i++) {
    cdp.send('Input.dispatchMouseEvent', { type: 'mouseWheel', x: P.x, y: P.y, ...(edge?.[wb].axis === 'x' ? { deltaX: 100, deltaY: 0 } : { deltaX: 0, deltaY: 100 }) }).catch(() => {});
    await sleep(Math.max(0, scrollStart + (i + 1) * 50 - Date.now()));
  }
  const scrollEnd = Date.now();
  await sleep(500);
  const raf = await evaluate(page, () => window.cmp.rafStop(), null, 30_000).catch(() => []);
  const longtasks = await evaluate(page, () => window.cmp.longtasks).catch(() => []);
  const scrollFrames = frames.splice(0);
  if (rep === 1 || smoke) await page.screenshot({ path: resolve(output, 'shots', `${arm}-${wb}-scrolled.png`) }).catch(() => {});

  // Edits: click a mid-sheet cell, 5 digits at 100 ms, Enter.
  const editLog = [];
  for (let k = 0; k < edits; k++) {
    const point = { x: Math.round(grid.x + (k < 5 ? 250 : 430)), y: Math.round(grid.y + 140 + 45 * (k % 5)) };
    sendClick(cdp, point.x, point.y);
    await sleep(400);
    const cell = await evaluate(page, () => window.cmp.active()).catch(() => null);
    const typed = 10000 + ((rep * 977 + k * 7919) % 89999);
    const sent = [...String(typed)].map((digit, i) => ({ digit, at: 0, i }));
    for (const key of sent) { key.at = sendKey(cdp, key.digit); await sleep(100); }
    await sleep(50);
    const enter = sendKey(cdp, 'Enter');
    let committed = false, value = null;
    const a1 = cell ? `${colName(cell.col)}${cell.row + 1}` : null;
    while (a1 && Date.now() - enter < KEY_CEIL) {
      const text = await evaluate(page, () => window.cmp.editorText(), null, 1000).catch(() => undefined);
      value = await evaluate(page, (x) => window.cmp.read(x), a1, 1000).catch(() => undefined);
      if ((text === '' || text === null) && num(value) === typed) { committed = true; break; }
      await sleep(50);
    }
    await sleep(400);
    if (cell && cell.col < 8 && committed) modified.set(`${cell.row}:${cell.col}`, typed);
    editLog.push({ k, point, cell, a1, typed, keys: sent, enter, committed, value: committed ? typed : value, settledAt: Date.now() });
  }
  const editFrames = [scrollFrames.at(-1), ...frames.splice(0)].filter(Boolean);

  let jumpKey = null, jumpRow = null, jumpFrames = editFrames, recalcStart = null, recalcFrames = editFrames, a1Before, newA1, expected, setResult, jLast = null, correct = false, valueMs = null;
  if (!edge) {
  // Jump: Ctrl+Down from a cell in A-H lands on the last row with column J on screen.
  sendClick(cdp, Math.round(grid.x + 250), Math.round(grid.y + 140));
  await sleep(400);
  jumpKey = sendKey(cdp, 'ArrowDown', 2);
  while (Date.now() - jumpKey < JUMP_CEIL) {
    jumpRow = (await evaluate(page, () => window.cmp.active(), null, 1000).catch(() => null))?.row ?? null;
    if (jumpRow === rows - 1) break;
    await sleep(50);
  }
  await quiet(frames, 1000, jumpKey + JUMP_CEIL);
  jumpFrames = [editFrames.at(-1), ...frames.splice(0)].filter(Boolean);
  if (rep === 1 || smoke) await page.screenshot({ path: resolve(output, 'shots', `${arm}-${wb}-end.png`) }).catch(() => {});

  // Recalc: A1 through the arm's API, then the J chain's last value must paint and read back right.
  a1Before = model(0, 0);
  newA1 = 5000 + rep;
  modified.set('0:0', newA1);
  expected = 0;
  for (const [key, v] of modified) { const [r, c] = key.split(':').map(Number); expected += numeric(v) - numeric(value(r, c)); }
  expected += total(rows);
  recalcStart = Date.now();
  setResult = await evaluate(page, (v) => window.cmp.setA1(v), newA1, RECALC_CEIL).then(() => 'ok', (e) => String(e).slice(0, 160));
  while (Date.now() - recalcStart < RECALC_CEIL) {
    jLast = await evaluate(page, (x) => window.cmp.read(x), `J${rows}`, 2000).catch(() => undefined);
    if (num(jLast) === expected) { correct = true; valueMs = Date.now() - recalcStart; break; }
    await sleep(50);
  }
  await quiet(frames, 1000, recalcStart + RECALC_CEIL);
  recalcFrames = [jumpFrames.at(-1), ...frames.splice(0)].filter(Boolean);
  }
  session.memAfter = await memory(browser, cdp);
  session.longtasks = await evaluate(page, () => window.cmp.longtasks).catch(() => longtasks);
  session.pageErrors = await evaluate(page, () => window.cmp.errors).catch(() => []);
  await cdp.send('Page.stopScreencast').catch(() => {});
  await browser.close().catch(() => {});

  // Analysis (measured browser closed).
  const paintFrame = await (async () => {
    for (const f of openFrames) {
      if (f.ts < t0) continue;
      if ((await inkOf(f, gridInner)).dark >= 0.004) return f;
    }
    return null;
  })();
  session.firstPaint = paintFrame && paintFrame.ts - t0 <= OPEN_CEIL ? { ms: Math.round(paintFrame.ts - t0) } : timedOut();
  if (accepted) {
    const key = keys[Math.max(0, keys.length - accepted.sevens)];
    const f = await firstChange(openFrames, key, cropAt(P), accepted.at + 500);
    const at = f ? f.ts : accepted.at;
    session.acceptsInput = at - t0 <= OPEN_CEIL ? { ms: Math.round(at - t0), via: f ? 'frame' : 'poll', keyIndex: keys.indexOf(key) } : timedOut();
  } else session.acceptsInput = timedOut({ editorKnown });

  const inWindow = scrollFrames.filter((f) => f.ts >= scrollStart && f.ts <= scrollEnd);
  const frameGaps = inWindow.slice(1).map((f, i) => f.ts - inWindow[i].ts);
  const rafGaps = raf.slice(1).map((t, i) => t - raf[i]).filter((_, i) => raf[i + 1] <= scrollEnd);
  let blank = 0;
  const blankRatios = [];
  for (const f of inWindow) {
    const ink = await inkOf(f, gridInner);
    blankRatios.push(+ink.light.toFixed(4));
    if (ink.light < 0.01) blank++;
  }
  const scrollTasks = longtasks.filter(([start]) => start >= scrollStart && start <= scrollEnd);
  session.scroll = {
    durationMs: scrollEnd - scrollStart, frames: inWindow.length, fps: +(inWindow.length / ((scrollEnd - scrollStart) / 1000)).toFixed(1),
    frameGapP50: pct(frameGaps, 50), frameGapP95: pct(frameGaps, 95),
    rafP50: pct(rafGaps, 50), rafP95: pct(rafGaps, 95), rafOver50: rafGaps.filter((g) => g > 50).length, rafCount: rafGaps.length,
    longtasks: taskStats(scrollTasks), blankFrames: blank, lightRatios: blankRatios,
  };

  session.edits = [];
  for (const e of editLog) {
    const region = cropAt(e.point);
    const keys = [];
    for (const key of e.keys) {
      const f = await firstChange(editFrames, key.at, region, key.at + KEY_CEIL);
      keys.push(f ? { ms: Math.round(f.ts - key.at) } : timedOut());
    }
    const target = before(editFrames, e.settledAt);
    const f = target ? await firstMatch(editFrames, e.enter, region, target, e.enter + KEY_CEIL) : null;
    session.edits.push({ cell: e.cell, typed: e.typed, committed: e.committed, value: e.value, colOk: e.cell ? e.cell.col < 8 : false,
      keys, commit: e.committed && f ? { ms: Math.round(f.ts - e.enter) } : timedOut() });
  }

  if (edge) {
    session.ended = new Date().toISOString();
    session.loadEnd = loadavg();
    return session;
  }
  const jumpFinal = jumpFrames.at(-1);
  const jf = jumpFinal ? await firstMatch(jumpFrames, jumpKey, gridInner, jumpFinal) : null;
  session.jump = { row: jumpRow, reachedEnd: jumpRow === rows - 1, ...(jumpRow === rows - 1 && jf ? { ms: Math.round(jf.ts - jumpKey) } : timedOut()) };

  const recalcBase = before(recalcFrames, recalcStart);
  const recalcFinal = recalcFrames.at(-1);
  const changed = recalcBase && recalcFinal ? (await diff(recalcBase, recalcFinal, gridInner)).n > 6 : false;
  const rf = changed ? await firstMatch(recalcFrames, recalcStart, gridInner, recalcFinal) : null;
  session.recalc = { a1Before, newA1, expected, read: jLast, correct, setResult, painted: changed,
    value: correct ? { ms: valueMs } : timedOut(),
    ...(correct && rf && rf.ts - recalcStart <= RECALC_CEIL ? { ms: Math.round(rf.ts - recalcStart) } : timedOut()) };
  session.ended = new Date().toISOString();
  session.loadEnd = loadavg();
  return session;
}

async function quiet(frames, quietMs, until) {
  let count = frames.length, since = Date.now();
  while (Date.now() < until) {
    await sleep(100);
    if (frames.length !== count) { count = frames.length; since = Date.now(); }
    else if (Date.now() - since >= quietMs) return true;
  }
  return false;
}
function num(v) {
  if (typeof v === 'number') return v;
  if (v && typeof v === 'object' && v.kind === 'number') return v.value;
  if (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v))) return Number(v);
  return null;
}
function pct(values, p) {
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  const v = sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil((p / 100) * sorted.length) - 1))];
  return Number.isFinite(v) ? +v.toFixed(1) : v;
}
function taskStats(tasks) {
  const d = tasks.map(([, ms]) => ms);
  return { count: d.length, sumMs: d.reduce((a, b) => a + b, 0), maxMs: d.length ? Math.max(...d) : 0 };
}

const meta = {
  started: new Date().toISOString(), smoke, reps, burners, edits, wbs, arms, viewport, deviceScaleFactor: 1,
  boSha: env.CMP_BO_SHA ?? null, region: env.CMP_REGION ?? null, host: os.hostname(), nproc: os.availableParallelism(),
  loadStart: loadavg(), executablePath,
  univer: (() => { try { return JSON.parse(readFileSync(resolve(here, 'site/node_modules/@univerjs/presets/package.json'), 'utf8')).version; } catch { return null; } })(),
  ceilings: { openMs: OPEN_CEIL, keyMs: KEY_CEIL, recalcMs: RECALC_CEIL, jumpMs: JUMP_CEIL },
};
const sessions = [];
const save = () => writeFile(resolve(output, 'results.json'), JSON.stringify({ meta, sessions }, null, 1));
for (const wb of wbs) {
  const wbStart = Date.now();
  for (let rep = burners ? 0 : 1; rep <= reps; rep++) {
    const order = rep % 2 === 1 ? arms : [...arms].reverse();
    for (const arm of order) {
      const measured = rep >= 1;
      if (Date.now() - wbStart > budgetMs) { sessions.push({ arm, wb, rep, measured, skipped: 'budget' }); continue; }
      let s;
      try { s = await runSession(arm, wb, rep, measured); } catch (error) { s = { arm, wb, rep, measured, crashed: String(error).slice(0, 300) }; }
      sessions.push(s);
      console.log(`${wb} rep${rep} ${arm}: fp=${s.firstPaint?.ms ?? 'TO'} acc=${s.acceptsInput?.ms ?? 'TO'} fps=${s.scroll?.fps ?? '-'} blank=${s.scroll?.blankFrames ?? '-'} `
        + `k2p=${pct((s.edits ?? []).flatMap((e) => e.keys.map((k) => k.ms ?? Infinity)), 50)} commit=${pct((s.edits ?? []).map((e) => e.commit.ms ?? Infinity), 50)} `
        + `jump=${s.jump?.ms ?? 'TO'} recalc=${s.recalc?.ms ?? 'TO'} value=${s.recalc?.value?.ms ?? 'TO'} ok=${s.recalc?.correct} pss=${s.memSettled?.rendererPssMb}/${s.memAfter?.rendererPssMb} `
        + `${s.skipped ?? ''}${s.crashed ? `CRASH ${s.crashed}` : ''} err=${(s.errors ?? []).length + (s.pageErrors ?? []).length}`);
      await save();
    }
  }
}
meta.ended = new Date().toISOString();
meta.loadEnd = loadavg();
await save();
const { summarize } = await import('./summary.mjs');
await writeFile(resolve(output, 'summary.md'), summarize({ meta, sessions }));
await analysisBrowser.close();
server.stop();
console.log(readFileSync(resolve(output, 'summary.md'), 'utf8'));
process.exit(0);
