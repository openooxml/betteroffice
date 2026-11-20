import { test, expect, type Browser, type Page } from 'playwright/test';
import JSZip from 'jszip';

const FONT = 'Times New Roman';
const BODY = 'This paragraph holds plain text in Times New Roman at eleven points so that the page reads like an ordinary' +
  ' letter, and it runs a little past one line to wrap once on an A4 page.';
const PARAGRAPHS = 25;
const SPAN_PARAGRAPHS = 9;
const MOVES = 61;
const DRAG_MS = 1000;
const HOLD_MS = 800;
const STRUCTURE_DRAGS = 2;

interface Sent { time: number; worker: number; id: number; seq: number; kind: string }
interface Reply { time: number; worker: number; id: number; covers: number | null; geometry: boolean; staged: number; version: string | null }
interface DragProbe {
  on: boolean; sig: string; downs: number[]; ups: number[]; moves: number[]; paints: number[];
  sent: Sent[]; replies: Reply[]; pending: Set<string>; lastActivity: number;
}
interface Drag { down: number; up: number; moves: number[]; paints: number[]; sent: Sent[]; replies: Reply[] }

declare global {
  interface Window { __configuredFonts: { errors: string[]; ready: boolean; fontsReady: number[]; load(bytes: number[]): void }; __dragProbe: DragProbe; __dragSignature(): string }
}

async function shortPage() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const paragraphs = Array.from({ length: PARAGRAPHS }, (_, index) => [0, 8, 16].includes(index)
    ? { bold: true, text: `Section ${index / 8 + 1}` }
    : { bold: false, text: `${index + 1}. ${BODY}` });
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    paragraphs.map(({ bold, text }) => `<w:p><w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>` +
      `${bold ? '<w:b/>' : ''}<w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('') +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>');
  return [...await zip.generateAsync({ type: 'uint8array' })];
}

async function instrument(page: Page) {
  await page.addInitScript(() => {
    const probe: DragProbe = {
      on: false, sig: '', downs: [], ups: [], moves: [], paints: [], sent: [], replies: [],
      pending: new Set(), lastActivity: performance.now(),
    };
    window.__dragProbe = probe;
    window.__dragSignature = () => [
      ...Array.from(document.querySelectorAll<HTMLElement>('[data-docx-worker-editor] div[aria-hidden="true"]'))
        .filter((element) => element.style.background.includes('59, 130, 246, 0.25')),
      ...Array.from(document.querySelectorAll<HTMLElement>('[data-testid^="selection-rect-"]')),
    ].map(({ style }) => `${style.left},${style.top},${style.width},${style.height}`).join(';');
    const caret = (node: Node) => !!(node instanceof Element ? node : node.parentElement)?.closest('[data-testid="caret"], [data-docx-caret]');
    new MutationObserver((records) => {
      if (records.some((record) => !caret(record.target))) probe.lastActivity = performance.now();
      if (!probe.on) return;
      const sig = window.__dragSignature();
      if (sig === probe.sig) return;
      probe.sig = sig;
      probe.paints.push(performance.now());
    }).observe(document, { childList: true, subtree: true, attributes: true, attributeFilter: ['style'] });
    window.addEventListener('pointerdown', () => { if (probe.on) probe.downs.push(performance.now()); }, true);
    window.addEventListener('pointerup', () => { if (probe.on) probe.ups.push(performance.now()); }, true);
    window.addEventListener('pointermove', (event) => { if (probe.on && event.buttons === 1) probe.moves.push(performance.now()); }, true);
    const NativeWorker = window.Worker;
    let workers = 0;
    window.Worker = class extends NativeWorker {
      readonly index = workers++;

      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        this.addEventListener('message', (event: MessageEvent) => {
          type Payload = { geometry?: unknown; documentVersion?: string; stagedPresentation?: { pages?: unknown[] } };
          const data = event.data as { id?: number; ok?: boolean; progress?: unknown; commandAck?: { frame?: { payload?: Payload } };
            commandFrame?: { frame?: { coversThroughSeq?: number; payload?: Payload } } } | null;
          if (!data || typeof data.id !== 'number' || data.progress !== undefined) return;
          probe.lastActivity = performance.now();
          probe.pending.delete(`${this.index}:${data.id}`);
          if (!probe.on) return;
          const payload = data.commandFrame?.frame?.payload ?? data.commandAck?.frame?.payload;
          probe.replies.push({ time: performance.now(), worker: this.index, id: data.id,
            covers: data.commandFrame?.frame?.coversThroughSeq ?? null, geometry: !!payload?.geometry,
            staged: payload?.stagedPresentation?.pages?.length ?? 0, version: payload?.documentVersion ?? null });
        });
      }

      postMessage(message: unknown, transfer?: Transferable[] | StructuredSerializeOptions) {
        const request = message as { id?: number; type?: string; command?: { kind?: string; seq?: number } } | null;
        if (request?.type === 'edit' && typeof request.id === 'number') {
          probe.pending.add(`${this.index}:${request.id}`);
          if (probe.on) probe.sent.push({ time: performance.now(), worker: this.index, id: request.id,
            seq: request.command?.seq ?? -1, kind: request.command?.kind ?? 'unknown' });
        }
        return super.postMessage(message, transfer as StructuredSerializeOptions);
      }
    };
  });
}

async function quiet(page: Page) {
  await page.waitForFunction(() => window.__dragProbe.pending.size === 0 && performance.now() - window.__dragProbe.lastActivity > 300,
    undefined, { timeout: 30_000, polling: 50 });
}

async function open(browser: Browser) {
  const context = await browser.newContext();
  const page = await context.newPage();
  await instrument(page);
  await page.goto('/docx-configured-fonts.html');
  await page.waitForFunction(() => window.__configuredFonts?.ready === true, undefined, { timeout: 60_000 });
  await page.evaluate((bytes) => window.__configuredFonts.load(bytes), await shortPage());
  await page.waitForFunction(() => document.querySelectorAll('.layout-page .layout-paragraph, .canvas-page-mirror .layout-paragraph').length >= 11,
    undefined, { timeout: 90_000, polling: 100 });
  await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false', { timeout: 60_000 });
  await page.waitForTimeout(1_000);
  await quiet(page);
  expect(await page.evaluate(() => window.__configuredFonts.errors)).toEqual([]);
  return { context, page };
}

/** A paragraph start and a point on the paragraph nine below it, on the first page that holds both in view. */
async function dragLine(page: Page) {
  const line = await page.evaluate((span) => {
    const pages = document.querySelectorAll('.layout-page').length
      ? Array.from(document.querySelectorAll('.layout-page')) : Array.from(document.querySelectorAll('.canvas-page-mirror'));
    for (const candidate of pages) {
      const lines = Array.from(candidate.querySelectorAll('.layout-paragraph'), (paragraph) => {
        const runs = paragraph.querySelectorAll('.layout-run-text').length
          ? paragraph.querySelectorAll('.layout-run-text') : paragraph.querySelectorAll('.layout-run');
        return Array.from(runs, (run) => run.getBoundingClientRect()).filter((box) => box.width > 20 && box.height > 4)
          .sort((a, b) => a.top - b.top || a.left - b.left)[0] ?? null;
      }).filter((box): box is DOMRect => !!box).sort((a, b) => a.top - b.top);
      if (lines.length <= span + 1) continue;
      const from = lines[1]!;
      const to = lines[1 + span]!;
      return { start: { x: from.left + 1, y: from.top + from.height / 2 },
        end: { x: Math.min(to.left + 150, to.right - 2), y: to.top + to.height / 2 } };
    }
    return null;
  }, SPAN_PARAGRAPHS);
  if (!line) throw new Error('No page holds the drag span');
  const viewport = page.viewportSize()!;
  expect(line.start.y, 'drag start is out of view').toBeGreaterThan(0);
  expect(line.end.y, 'drag end is out of view').toBeLessThan(viewport.height);
  return line;
}

async function drag(page: Page, line: Awaited<ReturnType<typeof dragLine>>): Promise<Drag> {
  await page.mouse.click(line.start.x, line.start.y);
  await quiet(page);
  await page.evaluate(() => {
    const probe = window.__dragProbe;
    for (const list of [probe.downs, probe.ups, probe.moves, probe.paints, probe.sent, probe.replies]) list.length = 0;
    probe.sig = window.__dragSignature();
    probe.on = true;
  });
  await page.mouse.down();
  const started = Date.now();
  for (let index = 1; index <= MOVES; index += 1) {
    const step = index / MOVES;
    await page.mouse.move(line.start.x + (line.end.x - line.start.x) * step, line.start.y + (line.end.y - line.start.y) * step);
    const wait = started + Math.round(index * DRAG_MS / MOVES) - Date.now();
    if (wait > 0) await page.waitForTimeout(wait);
  }
  await page.waitForTimeout(HOLD_MS);
  await page.mouse.up();
  await quiet(page);
  return page.evaluate(() => {
    const probe = window.__dragProbe;
    probe.on = false;
    return { down: probe.downs[0]!, up: probe.ups.at(-1)!, moves: [...probe.moves], paints: [...probe.paints],
      sent: [...probe.sent], replies: [...probe.replies] };
  });
}

/** Worker replies that answer each extendSelection step: its own reply plus pushed frames covering its sequence. */
function stepReplies(result: Drag) {
  const steps = result.sent.filter((sent) => sent.kind === 'extendSelection');
  return steps.map((step) => result.replies.filter((reply) => reply.worker === step.worker &&
    (reply.id === step.id || reply.id === 0 && reply.covers === step.seq)));
}

test('worker drag selection gets one reply per step without frame geometry, page bitmaps or a new document version', async ({ browser }) => {
  test.setTimeout(180_000);
  const { context, page } = await open(browser);
  try {
    const line = await dragLine(page);
    for (let index = 0; index < STRUCTURE_DRAGS; index += 1) {
      const result = await drag(page, line);
      const steps = result.sent.filter((sent) => sent.kind === 'extendSelection');
      const answers = stepReplies(result);
      const perStep = answers.map((replies) => replies.length);
      console.log(`[drag-cadence] structure drag=${index + 1} moves=${result.moves.length} extendSelection=${steps.length} repliesPerStep=${JSON.stringify(perStep)}`);
      expect(steps.length, 'the drag sent no extendSelection').toBeGreaterThan(0);
      expect(steps.length, 'more extendSelection steps than pointer moves').toBeLessThanOrEqual(result.moves.length + 1);
      expect(perStep.filter((count) => count !== 1), 'extendSelection steps not answered by exactly one worker message').toEqual([]);
      expect(result.replies.filter((reply) => reply.geometry).length, 'worker messages carrying frame geometry during the drag').toBe(0);
      const stepped = answers.flat();
      expect(stepped.filter((reply) => reply.staged > 0).length, 'extendSelection replies carrying a page bitmap').toBe(0);
      const versions = new Set(stepped.flatMap((reply) => reply.version === null ? [] : [reply.version]));
      expect(versions.size, 'extendSelection replies carrying more than one document version').toBeLessThanOrEqual(1);
    }
  } finally {
    await context.close();
  }
});
