import { test, type Browser, type CDPSession, type Page } from 'playwright/test';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { typingDocx } from './docx-typing-fixture';

// Keystroke latency on a large synthetic document: keystroke to painted glyph and
// to moved caret (both read from screencast pixels), the longest main-thread task,
// and the resident worker's queue wait, for single keys and bursts.

const root = resolve(import.meta.dirname, '../../..');
const env = process.env;
const PAGES = Number(env.TYPING_PAGES ?? 800);
const RUNS = Number(env.TYPING_RUNS ?? 5);
const SINGLES = Number(env.TYPING_SINGLES ?? 3);
const BURST = Number(env.TYPING_BURST ?? 10);
const BURST_GAP_MS = Number(env.TYPING_BURST_GAP_MS ?? 30);
const DPR = Number(env.TYPING_DPR ?? 2);
const MODES = (env.TYPING_MODES ?? 'default,worker').split(',') as Mode[];
const POSITIONS = (env.TYPING_POSITIONS ?? 'start,middle,end').split(',') as Position[];
const OUT = env.TYPING_OUT ?? resolve(root, '.source/e2e/perf/typing');
const LABEL = env.TYPING_LABEL ?? 'working-tree';
// `name=url,name=url` runs each arm in turn per run (A, B, A, B, ...) so arms share noise.
const ARMS = (env.TYPING_ARMS ?? `${LABEL}=`).split(',').map((arm) => {
  const [name, url] = arm.split('=');
  return { name, url: url || '' };
});
const PROFILE = env.TYPING_PROFILE === '1';
const VIEWPORT = { width: 1440, height: 1000 };
const QUIET_TYPES = new Set(['applyUpdate', 'eraseCaret', 'destroy']);

type Mode = 'default' | 'worker';
type Position = 'start' | 'middle' | 'end';
type Fields = Record<string, unknown>;
interface Frame {
  t: number;
  data: string;
}
interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}
interface KeyRow {
  arm: string;
  mode: Mode;
  run: number;
  scenario: string;
  key: number;
  glyphMs: number | null;
  caretMs: number | null;
  replyMs: number | null;
  queueMs: number | null;
  preMs: number | null;
  handleMs: number | null;
  engineMs: number | null;
  replayMs: number | null;
  requests: number;
}
interface ScenarioRow {
  arm: string;
  mode: Mode;
  run: number;
  scenario: string;
  page: number;
  keys: number;
  wrapped: boolean;
  layoutPending: boolean;
  settleMs: number;
  longestMainTaskMs: number | null;
  longestWorkerTaskMs: number | null;
  longTasksOver50: number;
  requests: string[];
  workerWasmMB: number;
  profile: Fields[];
}

declare global {
  interface Window {
    __typing: {
      editor: {
        getTotalPages(): number;
        getCurrentPage(): number;
        getEditorRef(): {
          getYrsSession(): {
            paragraphSpans(story: string): { paraId: string; length?: number; text?: string }[];
          };
          yrsLocToDisplayPosition(loc: {
            story: string;
            paraId: string;
            offset: number;
          }): number | null;
          setSelection(anchor: number, head?: number): void;
          scrollToPosition(position: number): void;
          focus(): void;
        } | null;
      } | null;
      posts: Fields[];
      replies: Fields[];
      keys: { key: string; t: number }[];
      longTasks: { start: number; duration: number }[];
      profileEdits: boolean;
      workerWasmBytes: number;
      workerLog(): Promise<Fields[]>;
    };
  }
}

const sleep = (ms: number) => new Promise((done) => setTimeout(done, ms));
const begun = Date.now();
const log = (message: string) => {
  if (env.TYPING_VERBOSE === '1') console.log(`[${((Date.now() - begun) / 1000).toFixed(1)}s] ${message}`);
};

// A page whose wasm trapped never settles again; its arm run stops at the first trap line.
const TRAP =
  /RuntimeError: unreachable|RuntimeError: memory access out of bounds|RefCell already borrowed|already mutably borrowed|display-list build failed|Layout pipeline error/;
// Worker failures and main-thread fallbacks, kept per run so arms can be compared.
const NOTABLE =
  /out of memory|fresh worker|unavailable|main thread|main-thread|Building display pages failed|Layout pipeline took|destroyed/i;
const STEP_MS = Number(env.TYPING_STEP_MS ?? 180_000);
const LAYOUT_MS = Number(env.TYPING_LAYOUT_MS ?? 300_000);

class ArmFailed extends Error {}

/** Runs one step of an arm, failing the arm on a trap or after `ms`. */
async function step<T>(label: string, ms: number, poisoned: Promise<string>, run: () => Promise<T>): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      run(),
      new Promise<never>((_, fail) => {
        timer = setTimeout(() => fail(new ArmFailed(`timeout after ${ms} ms in ${label}`)), ms);
      }),
      poisoned.then((reason) => {
        throw new ArmFailed(`trap in ${label}: ${reason}`);
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}

class Screencast {
  frames: Frame[] = [];
  recording = false;
  private last: Frame | null = null;
  constructor(private readonly cdp: CDPSession) {}

  async start() {
    this.cdp.on('Page.screencastFrame', (event) => {
      void this.cdp.send('Page.screencastFrameAck', { sessionId: event.sessionId }).catch(() => {});
      const frame = { t: (event.metadata.timestamp ?? Date.now() / 1000) * 1000, data: event.data };
      if (this.recording) this.frames.push(frame);
      else this.last = frame;
    });
    await this.cdp.send('Page.startScreencast', {
      format: 'jpeg',
      quality: 80,
      maxWidth: VIEWPORT.width,
      maxHeight: VIEWPORT.height,
      everyNthFrame: 1,
    });
  }

  record() {
    this.frames = this.last ? [this.last] : [];
    this.recording = true;
  }

  take(): Frame[] {
    this.recording = false;
    const frames = this.frames;
    this.last = frames.at(-1) ?? this.last;
    this.frames = [];
    return frames;
  }
}

async function caretRect(page: Page, waitMs = 0): Promise<Rect | null> {
  return page.evaluate(async (waitMs) => {
    const deadline = performance.now() + waitMs;
    for (;;) {
      const box = document.querySelector('[data-testid="caret"]')?.getBoundingClientRect();
      if (box && box.height > 0) return { x: box.x, y: box.y, width: box.width, height: box.height };
      if (performance.now() >= deadline) return null;
      await new Promise((done) => setTimeout(done, 10));
    }
  }, waitMs);
}

/** Waits until the worker has answered everything and the page has gone quiet. */
async function settle(page: Page, quietMs = 600, timeoutMs = 120_000): Promise<number> {
  return page.evaluate(
    async ({ quietMs, timeoutMs, quiet }) => {
      const probe = window.__typing;
      const start = performance.timeOrigin + performance.now();
      const deadline = start + timeoutMs;
      for (;;) {
        const now = performance.timeOrigin + performance.now();
        const open = new Set<unknown>();
        for (const post of probe.posts) if (!quiet.includes(post.type as string)) open.add(post.id);
        for (const reply of probe.replies) open.delete(reply.id);
        const lastReply = (probe.replies.at(-1)?.t as number | undefined) ?? 0;
        const lastLong = probe.longTasks.reduce((end, task) => Math.max(end, task.start + task.duration), 0);
        if ((open.size === 0 && now - lastReply > quietMs && now - lastLong > quietMs) || now > deadline) {
          return now - start;
        }
        await new Promise((done) => setTimeout(done, 25));
      }
    },
    { quietMs, timeoutMs, quiet: [...QUIET_TYPES] }
  );
}

/** Puts the caret at the end of a short paragraph about `fraction` into the body. */
async function placeCaret(page: Page, fraction: number, nth: number): Promise<number> {
  return page.evaluate(
    async ({ fraction, nth }) => {
      const editor = window.__typing.editor!.getEditorRef()!;
      const spans = editor.getYrsSession().paragraphSpans('body');
      const short = (index: number) => {
        const length = spans[index]?.length ?? 0;
        return length >= 12 && length <= 48;
      };
      let index = Math.min(spans.length - 1, Math.floor(spans.length * fraction));
      while (index > 0 && !short(index)) index -= 1;
      for (let skip = 0; skip < nth; skip++) {
        index -= 1;
        while (index > 0 && !short(index)) index -= 1;
      }
      const span = spans[Math.max(0, index)];
      const position = editor.yrsLocToDisplayPosition({
        story: 'body',
        paraId: span.paraId,
        offset: span.length ?? 0,
      });
      if (position == null) throw new Error('no display position for the target paragraph');
      editor.scrollToPosition(position);
      await new Promise((done) => setTimeout(done, 400));
      editor.setSelection(position);
      editor.focus();
      return window.__typing.editor!.getCurrentPage();
    },
    { fraction, nth }
  );
}

async function pressKeys(page: Page, count: number, gapMs: number) {
  const start = Date.now();
  for (let key = 0; key < count; key++) {
    const due = start + key * gapMs;
    const wait = due - Date.now();
    if (wait > 0) await sleep(wait);
    await page.keyboard.down('x');
    await page.keyboard.up('x');
  }
}

/** Per key: when its glyph's ink first appears in its cell, and when the caret reaches its right edge. */
async function analyse(
  analysis: Page,
  frames: Frame[],
  keys: number[],
  before: Rect,
  after: Rect
): Promise<{ glyph: (number | null)[]; caret: (number | null)[] }> {
  return analysis.evaluate(
    async ({ frames, keys, before, after, viewport }) => {
      const advance = (after.x - before.x) / keys.length;
      const decoded: { t: number; image: ImageData; scale: number }[] = [];
      const top = Math.max(0, before.y);
      const height = before.height;
      const left = Math.max(0, before.x - 2);
      const width = after.x - before.x + 6;
      for (const frame of frames) {
        const bytes = Uint8Array.from(atob(frame.data), (char) => char.charCodeAt(0));
        const bitmap = await createImageBitmap(new Blob([bytes], { type: 'image/jpeg' }));
        const scale = bitmap.width / viewport.width;
        const canvas = new OffscreenCanvas(Math.ceil(width * scale), Math.ceil(height * scale));
        const context = canvas.getContext('2d')!;
        context.drawImage(bitmap, -left * scale, -top * scale);
        decoded.push({
          t: frame.t,
          image: context.getImageData(0, 0, canvas.width, canvas.height),
          scale,
        });
        bitmap.close();
      }
      const darkest = (frame: (typeof decoded)[number], x0: number, x1: number, y0: number, y1: number) => {
        let min = 255;
        const { image, scale } = frame;
        for (let y = Math.floor((y0 - top) * scale); y < Math.ceil((y1 - top) * scale); y++) {
          for (let x = Math.floor((x0 - left) * scale); x < Math.ceil((x1 - left) * scale); x++) {
            if (x < 0 || y < 0 || x >= image.width || y >= image.height) continue;
            const at = (y * image.width + x) * 4;
            const luma = 0.299 * image.data[at] + 0.587 * image.data[at + 1] + 0.114 * image.data[at + 2];
            if (luma < min) min = luma;
          }
        }
        return min;
      };
      // Glyph k: the middle of its advance cell in the lower half of the line (x-height).
      const clear = Math.max(advance * 0.3, before.width + 0.5);
      const glyphInk = (frame: (typeof decoded)[number], key: number) =>
        darkest(
          frame,
          before.x + advance * key + clear,
          before.x + advance * (key + 0.7),
          top + height * 0.45,
          top + height * 0.8
        ) < 170;
      // Caret after key k: a full-height line, so look above the x-height only.
      const caretInk = (frame: (typeof decoded)[number], key: number) =>
        darkest(
          frame,
          before.x + advance * (key + 1) - 1,
          before.x + advance * (key + 1) + 1,
          top + height * 0.08,
          top + height * 0.24
        ) < 140;
      const first = (key: number, test: (frame: (typeof decoded)[number], key: number) => boolean) => {
        const hit = decoded.find((frame) => frame.t >= keys[key] && test(frame, key));
        return hit ? hit.t - keys[key] : null;
      };
      return {
        glyph: keys.map((_, key) => first(key, glyphInk)),
        caret: keys.map((_, key) => first(key, caretInk)),
      };
    },
    { frames, keys, before, after, viewport: VIEWPORT }
  );
}

/** When a square the page turns black at `from` first shows in the screencast. */
async function analyseFloor(analysis: Page, frames: Frame[], from: number): Promise<number | null> {
  return analysis.evaluate(
    async ({ frames, from, viewport }) => {
      for (const frame of frames) {
        if (frame.t < from) continue;
        const bytes = Uint8Array.from(atob(frame.data), (char) => char.charCodeAt(0));
        const bitmap = await createImageBitmap(new Blob([bytes], { type: 'image/jpeg' }));
        const scale = bitmap.width / viewport.width;
        const canvas = new OffscreenCanvas(1, 1);
        const context = canvas.getContext('2d')!;
        context.drawImage(bitmap, -12 * scale, -12 * scale);
        bitmap.close();
        if (context.getImageData(0, 0, 1, 1).data[0] < 60) return frame.t - from;
      }
      return null;
    },
    { frames, from, viewport: VIEWPORT }
  );
}

interface TraceEvent {
  name: string;
  ph: string;
  ts: number;
  dur?: number;
  pid: number;
  tid: number;
  args?: { name?: string };
}

/** Longest top-level task on the renderer main thread and on the resident worker within [from, to] (epoch ms). */
function longestTasks(trace: Buffer, from: number, to: number, clock: { trace: number; epoch: number }) {
  const events = (JSON.parse(trace.toString()) as { traceEvents: TraceEvent[] }).traceEvents;
  const names = new Map<string, string>();
  for (const event of events) {
    if (event.ph === 'M' && event.name === 'thread_name') names.set(`${event.pid}:${event.tid}`, event.args?.name ?? '');
  }
  const toEpoch = (ts: number) => clock.epoch + (ts - clock.trace) / 1000;
  let main = 0;
  let worker = 0;
  let over50 = 0;
  for (const event of events) {
    if (event.ph !== 'X' || !event.dur || !/(^|::)RunTask$/.test(event.name)) continue;
    const start = toEpoch(event.ts);
    const end = start + event.dur / 1000;
    if (end < from || start > to) continue;
    const thread = names.get(`${event.pid}:${event.tid}`) ?? '';
    const ms = event.dur / 1000;
    if (thread === 'CrRendererMain') {
      main = Math.max(main, ms);
      if (ms > 50) over50 += 1;
    } else if (/DedicatedWorker/.test(thread)) worker = Math.max(worker, ms);
  }
  return { main, worker, over50 };
}

async function startTrace(browser: Browser, page: Page) {
  await browser.startTracing(page, {
    categories: ['toplevel', 'blink.user_timing'],
  });
}

/** Maps trace microseconds to epoch ms with a user-timing mark. */
async function traceClock(page: Page) {
  return page.evaluate(() => {
    const name = `typing-clock-${Math.random()}`;
    const mark = performance.mark(name);
    return { name, epoch: performance.timeOrigin + mark.startTime };
  });
}

function clockFrom(trace: Buffer, mark: { name: string; epoch: number }) {
  const events = (JSON.parse(trace.toString()) as { traceEvents: TraceEvent[] }).traceEvents;
  const event = events.find((candidate) => candidate.name === mark.name);
  return event ? { trace: event.ts, epoch: mark.epoch } : null;
}

const quantile = (values: number[], q: number) => {
  if (values.length === 0) return null;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.floor(q * (sorted.length - 1) + 0.5))];
};
const fmt = (value: number | null) => (value === null ? '–' : value.toFixed(value < 10 ? 1 : 0));

function summarise(keys: KeyRow[], scenarios: ScenarioRow[]): string {
  const groups = new Map<string, { keys: KeyRow[]; scenarios: ScenarioRow[] }>();
  const group = (row: { arm: string; mode: Mode; scenario: string }) =>
    `${row.arm}|${row.mode}|${row.scenario}`;
  for (const row of scenarios) {
    if (!groups.has(group(row))) groups.set(group(row), { keys: [], scenarios: [] });
    groups.get(group(row))!.scenarios.push(row);
  }
  for (const row of keys) groups.get(group(row))?.keys.push(row);
  const stat = (values: (number | null)[]) => {
    const present = values.filter((value): value is number => value !== null);
    return `${fmt(quantile(present, 0.5))} / ${fmt(quantile(present, 0.95))}`;
  };
  const lines = [
    `| arm | mode | scenario | n keys | glyph p50 / p95 | caret p50 / p95 | reply p50 / p95 | queue wait p50 / p95 | pre-work p50 / p95 | engine p50 / p95 | longest main task p50 / max | longest worker task p50 / max | worker wasm MB |`,
    `|---|---|---|---|---|---|---|---|---|---|---|---|---|`,
  ];
  for (const [name, { keys, scenarios }] of groups) {
    const [arm, mode, scenario] = name.split('|');
    const mains = scenarios.map((row) => row.longestMainTaskMs).filter((v): v is number => v !== null);
    const workers = scenarios.map((row) => row.longestWorkerTaskMs).filter((v): v is number => v !== null);
    lines.push(
      `| ${arm} | ${mode} | ${scenario} | ${keys.length} | ${stat(keys.map((row) => row.glyphMs))} | ${stat(keys.map((row) => row.caretMs))} | ${stat(keys.map((row) => row.replyMs))} | ${stat(keys.map((row) => row.queueMs))} | ${stat(keys.map((row) => row.preMs))} | ${stat(keys.map((row) => row.engineMs))} | ${fmt(quantile(mains, 0.5))} / ${fmt(mains.length ? Math.max(...mains) : null)} | ${fmt(quantile(workers, 0.5))} / ${fmt(workers.length ? Math.max(...workers) : null)} | ${Math.max(...scenarios.map((row) => row.workerWasmMB))} |`
    );
  }
  return lines.join('\n');
}

test.describe.configure({ mode: 'serial' });

test('docx keystroke latency', async ({ browser }) => {
  test.setTimeout(0);
  await mkdir(OUT, { recursive: true });
  const fixture = resolve(OUT, `typing-${PAGES}.docx`);
  const bytes = await readFile(fixture).catch(async () => {
    const generated = Buffer.from(await typingDocx(PAGES));
    await writeFile(fixture, generated);
    return generated;
  });
  const keyRows: KeyRow[] = [];
  const openRows: Fields[] = [];
  const scenarioRows: ScenarioRow[] = [];
  const analysisContext = await browser.newContext();
  const analysis = await analysisContext.newPage();
  const started = new Date().toISOString().replace(/[:.]/g, '-');
  // TYPING_NAME appends to an earlier result of that name, so runs can go in separate slots.
  const name = env.TYPING_NAME ?? `${LABEL}-${PAGES}p-${started}`;
  const earlier = await readFile(resolve(OUT, `${name}.json`), 'utf8').then(
    (text) => JSON.parse(text) as { opens: Fields[]; keys: KeyRow[]; scenarios: ScenarioRow[] },
    () => null
  );
  if (earlier) {
    openRows.push(...earlier.opens);
    keyRows.push(...earlier.keys);
    scenarioRows.push(...earlier.scenarios);
  }
  const firstRun = Math.max(-1, ...openRows.map((row) => row.run as number)) + 1;
  const write = async () => {
    const floors = openRows
      .flatMap((row) => (row.echoFloorMs as (number | null)[] | undefined) ?? [])
      .filter((v): v is number => v !== null);
    const table =
      summarise(keyRows, scenarioRows) +
      `\n\nDOM echo floor (style change to screencast frame): p50 ${fmt(quantile(floors, 0.5))} ms, max ${fmt(floors.length ? Math.max(...floors) : null)} ms.\n` +
      openRows
        .map((row) =>
          row.failed
            ? `- ${row.arm} ${row.mode} run ${row.run}: FAILED, ${row.failed}`
            : `- ${row.arm} ${row.mode} run ${row.run}: ${row.pages} pages, first page ${row.firstPageMs} ms, layout complete ${Math.round(row.layoutMs as number)} ms, worker wasm at layout complete ${row.layoutWasmMB ?? '–'} MB, main-thread long tasks during open: ${row.openLongTasks}`
        )
        .join('\n') +
      '\n\n' +
      [...new Set(openRows.map((row) => `${row.arm} ${row.mode}`))]
        .map((arm) => {
          const rows = openRows.filter((row) => `${row.arm} ${row.mode}` === arm);
          const trapped = rows.filter((row) => row.trapped).length;
          const failed = rows.filter((row) => row.failed).length;
          return `- ${arm}: ${rows.length} runs, ${failed} failed (${trapped} trapped)`;
        })
        .join('\n');
    await writeFile(
      resolve(OUT, `${name}.json`),
      JSON.stringify(
        { pages: PAGES, runs: RUNS, dpr: DPR, arms: ARMS, opens: openRows, keys: keyRows, scenarios: scenarioRows },
        null,
        1
      )
    );
    await writeFile(resolve(OUT, `${name}.md`), table + '\n');
    return table;
  };

  for (let run = firstRun; run < firstRun + RUNS; run++) {
    for (const mode of MODES) {
      for (const arm of ARMS) {
        const context = await browser.newContext({ viewport: VIEWPORT, deviceScaleFactor: DPR });
        await context.route('**/__typing/document.docx', (route) =>
          route.fulfill({ body: bytes, contentType: 'application/octet-stream' })
        );
        const page = await context.newPage();
        const errors: string[] = [];
        const notable: string[] = [];
        let trapped: string | null = null;
        let failure: string | null = null;
        let poison: (reason: string) => void = () => {};
        const poisoned = new Promise<string>((done) => (poison = done));
        const armStart = Date.now();
        const seen = (kind: string, text: string) => {
          if ((NOTABLE.test(text) || TRAP.test(text)) && notable.length < 60) {
            notable.push(`${((Date.now() - armStart) / 1000).toFixed(1)}s ${kind}: ${text.slice(0, 240)}`);
          }
          if (trapped === null && TRAP.test(text)) {
            trapped = text.slice(0, 300);
            log(`${arm.name} ${mode} run ${run}: trap: ${trapped}`);
            poison(trapped);
          }
        };
        page.on('pageerror', (error) => {
          errors.push(String(error));
          seen('pageerror', String(error));
        });
        page.on('console', (message) => {
          const type = message.type();
          const text = message.text();
          if (type === 'error' || type === 'warning') log(`console.${type}: ${text.slice(0, 300)}`);
          seen(type, text);
        });
        try {
          const cdp = await context.newCDPSession(page);
          const cast = new Screencast(cdp);
          await cast.start();
          const opened = Date.now();
          await step('open', STEP_MS, poisoned, () => page.goto(`${arm.url}/docx-typing.html?mode=${mode}`));
          await step('first page', STEP_MS, poisoned, () =>
            page.locator('canvas[data-page-index="0"]').first().waitFor({ timeout: STEP_MS })
          );
          const firstPageMs = Date.now() - opened;
          log(`${arm.name} ${mode} run ${run}: first page after ${firstPageMs} ms`);
          if (PROFILE) await page.evaluate(() => (window.__typing.profileEdits = true));

          const measure = (
            scenario: string,
            pageNumber: number,
            count: number,
            gapMs: number,
            layoutPending: boolean
          ) => step(scenario, STEP_MS, poisoned, async () => {
            log(`${scenario}: start`);
            const before = await caretRect(page, 1500);
            const posts = await page.evaluate(() => window.__typing.posts.length);
            const keysBefore = await page.evaluate(() => window.__typing.keys.length);
            await startTrace(browser, page);
            const mark = await traceClock(page);
            cast.record();
            await pressKeys(page, count, gapMs);
            const settleMs = await settle(page);
            log(`${scenario}: settled after ${Math.round(settleMs)} ms`);
            const frames = cast.take();
            const trace = await browser.stopTracing();
            await sleep(700);
            const after = await caretRect(page, 1500);
            const probe = await page.evaluate(
              async ({ posts, keysBefore }) => ({
                keys: window.__typing.keys.slice(keysBefore).map((key) => key.t),
                posts: window.__typing.posts.slice(posts),
                replies: window.__typing.replies,
                worker: await window.__typing.workerLog(),
                wasmBytes: window.__typing.workerWasmBytes,
              }),
              { posts, keysBefore }
            );
            if (failure !== null) return;
            const keys = probe.keys.slice(-count);
            const wrapped = !before || !after || Math.abs(after.y - before.y) > 2 || after.x <= before.x;
            const pixels =
              !wrapped && before && after
                ? await analyse(analysis, frames, keys, before, after)
                : { glyph: keys.map(() => null), caret: keys.map(() => null) };
            const workerById = new Map<unknown, Fields>();
            for (const entry of probe.worker) workerById.set(entry.id, { ...workerById.get(entry.id), ...entry });
            const replyById = new Map(probe.replies.map((reply) => [reply.id, reply]));
            const edits = probe.posts.filter((post) => post.type === 'applyInput');
            keys.forEach((t, key) => {
              const edit = edits.find((post) => (post.t as number) >= t - 1);
              const reply = edit && replyById.get(edit.id);
              const worker = edit && workerById.get(edit.id);
              const num = (value: unknown) => (typeof value === 'number' ? value : null);
              const start = num(worker?.start);
              const arrive = num(worker?.arrive);
              const replied = num(worker?.reply);
              const total = num(reply?.workerTotalMs);
              keyRows.push({
                arm: arm.name,
                mode,
                run,
                scenario,
                key,
                glyphMs: pixels.glyph[key],
                caretMs: pixels.caret[key],
                replyMs: reply ? (reply.t as number) - t : null,
                queueMs: start !== null && arrive !== null ? start - arrive : null,
                preMs: start !== null && replied !== null && total !== null ? replied - start - total : null,
                handleMs: start !== null && replied !== null ? replied - start : null,
                engineMs: num(reply?.engineMs),
                replayMs: num(reply?.replayMs),
                requests: edits.length,
              });
            });
            if (env.TYPING_TRACES === '1') {
              await mkdir(resolve(OUT, 'traces'), { recursive: true });
              await writeFile(resolve(OUT, 'traces', `${arm.name}-${mode}-${run}-${scenarioRows.length}.json`), trace);
            }
            const clock = clockFrom(trace, mark);
            const tasks = clock ? longestTasks(trace, keys[0] ?? 0, (keys.at(-1) ?? 0) + settleMs, clock) : null;
            scenarioRows.push({
              arm: arm.name,
              mode,
              run,
              scenario,
              page: pageNumber,
              keys: count,
              wrapped,
              layoutPending,
              settleMs,
              longestMainTaskMs: tasks?.main ?? null,
              longestWorkerTaskMs: tasks?.worker ?? null,
              longTasksOver50: tasks?.over50 ?? 0,
              requests: probe.posts.map((post) => String(post.type)),
              workerWasmMB: Math.round(probe.wasmBytes / 2 ** 20),
              profile: PROFILE
                ? edits.map((edit) => (replyById.get(edit.id)?.engineProfile as Fields) ?? {})
                : [],
            });
          });

          // Straight after the first page paints, while the rest of the layout is still running.
          // The background layout has finished once every completion request is answered; the
          // worker-open replica is ready once the session reads. whenLayoutComplete would also
          // build every page's display list, which a user never does.
          const layoutDone = page.evaluate(async (deadlineMs) => {
            const probe = window.__typing;
            const since = performance.now();
            for (;;) {
              if (performance.now() - since > deadlineMs) return Number.NaN;
              const completions = probe.posts.filter((post) => post.type === 'completeLayout');
              const replied = new Set(probe.replies.map((reply) => reply.id));
              const completed =
                completions.length > 0
                  ? completions.every((post) => replied.has(post.id))
                  : performance.now() - since > 5_000;
              let replica = false;
              try {
                replica = (probe.editor?.getEditorRef()?.getYrsSession().paragraphSpans('body').length ?? 0) > 0;
              } catch {
                replica = false;
              }
              if (completed && replica) return performance.timeOrigin + performance.now();
              await new Promise((done) => setTimeout(done, 50));
            }
          }, LAYOUT_MS);
          layoutDone.catch(() => undefined);
          const surface = await page.evaluate(() => {
            for (const canvas of document.querySelectorAll('canvas[data-page-index="0"]')) {
              const box = canvas.getBoundingClientRect();
              if (box.width > 0 && box.height > 0) return { x: box.x, y: box.y };
            }
            return null;
          });
          if (surface) await page.mouse.click(surface.x + 300, surface.y + 200);
          await page.keyboard.press('ControlOrMeta+Home');
          await page.keyboard.press('End');
          const pending = await Promise.race([layoutDone.then(() => false), sleep(50).then(() => true)]);
          await measure('open: first key', 1, 1, 0, pending);
          const layoutAt = await step('layout complete', LAYOUT_MS + 15_000, poisoned, () => layoutDone);
          if (Number.isNaN(layoutAt)) throw new ArmFailed(`layout incomplete after ${LAYOUT_MS} ms`);
          const layoutMs = layoutAt - opened;
          log(`layout complete after ${Math.round(layoutMs)} ms`);
          const layoutWasmMB = await step('worker memory', 30_000, poisoned, () =>
            page.evaluate(async () => {
              await window.__typing.workerLog();
              return Math.round(window.__typing.workerWasmBytes / 2 ** 20);
            })
          );
          await step('settle after layout', STEP_MS, poisoned, () => settle(page, 1500));
          const floors: (number | null)[] = [];
          await step('echo floor', STEP_MS, poisoned, async () => {
            await page.evaluate(() => {
              const square = document.createElement('div');
              square.id = 'typing-floor';
              Object.assign(square.style, {
                position: 'fixed',
                left: '0',
                top: '0',
                width: '24px',
                height: '24px',
                background: '#fff',
                zIndex: '2147483647',
              });
              document.body.append(square);
            });
            for (let sample = 0; sample < 3; sample++) {
              await sleep(300);
              cast.record();
              const from = await page.evaluate(() => {
                document.getElementById('typing-floor')!.style.background = '#000';
                return performance.timeOrigin + performance.now();
              });
              await sleep(300);
              floors.push(await analyseFloor(analysis, cast.take(), from));
              await page.evaluate(() => (document.getElementById('typing-floor')!.style.background = '#fff'));
            }
            await page.evaluate(() => document.getElementById('typing-floor')!.remove());
          });
          const openLongTasks = await step('open long tasks', 30_000, poisoned, () =>
            page.evaluate((opened) => {
              const tasks = window.__typing.longTasks.filter((task) => task.start >= opened);
              return `${tasks.length}, total ${Math.round(tasks.reduce((sum, task) => sum + task.duration, 0))} ms, max ${Math.round(Math.max(0, ...tasks.map((task) => task.duration)))} ms`;
            }, opened)
          );
          const pages = await step('page count', 30_000, poisoned, () =>
            page.evaluate(() => window.__typing.editor!.getTotalPages())
          );
          openRows.push({ arm: arm.name, mode, run, firstPageMs, layoutMs, layoutWasmMB, openLongTasks, pages, echoFloorMs: floors, notable });

          for (const position of POSITIONS) {
            const fraction = position === 'start' ? 0.002 : position === 'middle' ? 0.5 : 0.995;
            const pageNumber = await step(`${position}: place caret`, STEP_MS, poisoned, () =>
              placeCaret(page, fraction, run)
            );
            await step(`${position}: settle`, STEP_MS, poisoned, () => settle(page, 1200));
            for (let single = 0; single < SINGLES; single++) {
              await measure(`${position}: single`, pageNumber, 1, 0, false);
            }
            await measure(`${position}: burst ${BURST}@${BURST_GAP_MS}ms`, pageNumber, BURST, BURST_GAP_MS, false);
          }
        } catch (error) {
          if (!(error instanceof ArmFailed)) throw error;
          failure = error.message;
          log(`${arm.name} ${mode} run ${run}: FAILED ${failure}`);
          openRows.push({ arm: arm.name, mode, run, failed: failure, trapped, notable });
          await browser.stopTracing().catch(() => undefined);
        }
        if (failure || env.TYPING_DUMP === '1') {
          const order = await Promise.race([
            page
              .evaluate(() => ({ posts: window.__typing.posts, replies: window.__typing.replies, longTasks: window.__typing.longTasks }))
              .catch(() => null),
            sleep(15_000).then(() => null),
          ]);
          await mkdir(resolve(OUT, 'order'), { recursive: true });
          await writeFile(
            resolve(OUT, 'order', `${name}-${arm.name}-${mode}-${run}.json`),
            JSON.stringify({ opened: armStart, failure, trapped, notable, ...order })
          );
        }
        if (errors.length > 0) console.warn(`${arm.name} ${mode} run ${run}: ${errors.join('; ')}`);
        await Promise.race([context.close(), sleep(30_000)]);
        console.log(`${arm.name} ${mode} run ${run + 1}/${firstRun + RUNS} ${failure ? 'failed' : 'done'}`);
        await write();
      }
    }
  }
  await analysisContext.close();
  console.log(await write());
});
