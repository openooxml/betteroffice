import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

interface DragSample { source: 'raf' | 'mutation'; time: number; count: number; button: number; top: number; left: number; bottom: number; right: number }
interface DragProbe { editDelayMs: number; sent: { seq: number; time: number }[]; acked: { throughSeq: number; time: number }[]; moves: { time: number; y: number }[] }

declare global {
  interface Window { __docxDragSamples?: { active: boolean; samples: DragSample[]; stop(): void }; __docxDragProbe: DragProbe }
}

test.use({ launchOptions: { executablePath: process.env.CHROME || undefined } });

const OVERLAY = '[data-docx-worker-editor] div[aria-hidden="true"][style*="rgba(59, 130, 246"]';
const PAGE_WIDTH = 816;
const MARGIN = 96;
const COMMENT_BUTTON = 'button[aria-label="Add comment"]';

async function installProbe(page: Page, editDelayMs: number) {
  await page.addInitScript((delay) => {
    const probe: DragProbe = { editDelayMs: delay, sent: [], acked: [], moves: [] };
    window.__docxDragProbe = probe;
    document.addEventListener('pointermove', (event) => {
      if (event.buttons === 1) probe.moves.push({ time: performance.now(), y: event.clientY });
    }, true);
    const NativeWorker = window.Worker;
    let delivery = Promise.resolve();
    window.Worker = class extends NativeWorker {
      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        this.addEventListener('message', (event: MessageEvent) => {
          const ack = event.data?.ok ? event.data.commandAck : undefined;
          if (ack) probe.acked.push({ throughSeq: ack.throughSeq, time: performance.now() });
        }, true);
      }

      postMessage(message: unknown, transfer?: Transferable[] | StructuredSerializeOptions) {
        const request = message as { type?: string; command?: { kind: string; seq: number } };
        const post = () => Array.isArray(transfer) ? super.postMessage(message, transfer) : super.postMessage(message, transfer);
        if (request.type !== 'edit') return post();
        if (request.command?.kind === 'extendSelection') probe.sent.push({ seq: request.command.seq, time: performance.now() });
        if (!probe.editDelayMs) return post();
        // A worker that takes editDelayMs per command, as on a large document.
        delivery = delivery.then(() => new Promise<void>((resolve) => setTimeout(resolve, probe.editDelayMs))).then(post);
      }
    };
  }, editDelayMs);
}

async function documentBytes() {
  const sentence = 'The quick brown fox jumps over the lazy dog while the selection follows the pointer across lines. ';
  const zip = new JSZip();
  zip.file('[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>');
  zip.file('_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>');
  zip.file('word/document.xml',
    '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    Array.from({ length: 6 }, () => `<w:p><w:r><w:t xml:space="preserve">${sentence.repeat(4)}</w:t></w:r></w:p>`).join('') +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
    '</w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function frame(page: Page) {
  await page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())));
}

async function startSampling(page: Page) {
  await page.evaluate(({ selector, button }) => {
    const samples: DragSample[] = [];
    const read = (source: DragSample['source']) => {
      if (!window.__docxDragSamples?.active) return;
      const rects = Array.from(document.querySelectorAll<HTMLElement>(selector), (element) => element.getBoundingClientRect());
      const last = rects.reduce<DOMRect | undefined>((best, rect) =>
        !best || rect.bottom > best.bottom + 0.5 || Math.abs(rect.bottom - best.bottom) <= 0.5 && rect.right > best.right ? rect : best, undefined);
      const first = rects.reduce<DOMRect | undefined>((best, rect) =>
        !best || rect.top < best.top - 0.5 || Math.abs(rect.top - best.top) <= 0.5 && rect.left < best.left ? rect : best, undefined);
      samples.push({ source, time: performance.now(), count: rects.length, button: document.querySelectorAll(button).length, top: first?.top ?? 0, left: first?.left ?? 0, bottom: last?.bottom ?? 0, right: last?.right ?? 0 });
    };
    const observer = new MutationObserver(() => read('mutation'));
    observer.observe(document.querySelector('[data-docx-worker-editor]')!, { childList: true, subtree: true, attributes: true, attributeFilter: ['style'] });
    let raf = 0;
    const tick = () => { read('raf'); raf = requestAnimationFrame(tick); };
    raf = requestAnimationFrame(tick);
    window.__docxDragSamples = { active: false, samples, stop() { observer.disconnect(); cancelAnimationFrame(raf); } };
  }, { selector: OVERLAY, button: COMMENT_BUTTON });
}

async function openDocument(page: Page, editDelayMs = 0) {
  await installProbe(page, editDelayMs);
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'drag-selection.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await documentBytes(),
  });
  const canvas = page.locator('[data-docx-worker-editor] [data-page-id] canvas').first();
  await expect(canvas).toBeVisible({ timeout: 60_000 });
  const bounds = await canvas.boundingBox();
  if (!bounds) throw new Error('Document canvas is unavailable');
  const scale = bounds.width / PAGE_WIDTH;
  return { scale, start: { x: bounds.x + (MARGIN + 1) * scale, y: bounds.y + (MARGIN + 8) * scale },
    end: { x: bounds.x + (PAGE_WIDTH - MARGIN - 40) * scale, y: bounds.y + (MARGIN + 260) * scale } };
}

const STEPS = 90;
const WORKER_MS = 300;

async function drag(page: Page, start: { x: number; y: number }, end: { x: number; y: number }, steps = STEPS) {
  await startSampling(page);
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.evaluate(() => { window.__docxDragSamples!.active = true; });
  for (let step = 1; step <= steps; step += 1) {
    await page.mouse.move(start.x + (end.x - start.x) * step / steps, start.y + (end.y - start.y) * step / steps);
    await frame(page);
  }
  const released = await page.evaluate(() => performance.now());
  await page.mouse.up();
  await expect.poll(() => page.evaluate(() => {
    const probe = window.__docxDragProbe;
    const last = Math.max(-1, ...probe.sent.map((send) => send.seq));
    return probe.acked.some((ack) => ack.throughSeq >= last);
  }), { timeout: 90_000 }).toBe(true);
  await expect.poll(async () => page.locator(OVERLAY).count()).toBeGreaterThanOrEqual(5);
  for (let turn = 0; turn < 10; turn += 1) await frame(page);
  const result = await page.evaluate(() => {
    window.__docxDragSamples!.active = false;
    window.__docxDragSamples!.stop();
    return { samples: window.__docxDragSamples!.samples, probe: window.__docxDragProbe };
  });
  return { ...result, released };
}

const sameExtent = (sample: DragSample, expected: DragSample) => sample.count > 0 &&
  Math.abs(sample.bottom - expected.bottom) <= 0.5 && Math.abs(sample.right - expected.right) <= 0.5;

function percentile(values: number[], fraction: number) {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * fraction))] ?? Number.NaN;
}

/** Pointer-to-highlight latency for each move that enters a line the highlight has not reached yet. */
function latencies(samples: DragSample[], moves: DragProbe['moves']) {
  const shown = samples.filter((sample) => sample.count > 0);
  return moves.flatMap((move) => {
    if (shown.some((sample) => sample.time <= move.time && sample.bottom >= move.y)) return [];
    const reached = shown.find((sample) => sample.time > move.time && sample.bottom >= move.y);
    return reached ? [reached.time - move.time] : [];
  });
}

/** Release flushes the last pointer position at once, so only sends during the drag count. */
function inFlightPeak(probe: DragProbe, before = Infinity) {
  let peak = 0;
  for (const send of probe.sent.filter((candidate) => candidate.time < before)) {
    const acked = Math.max(-1, ...probe.acked.filter((ack) => ack.time <= send.time).map((ack) => ack.throughSeq));
    peak = Math.max(peak, probe.sent.filter((other) => other.time <= send.time && other.seq > acked).length);
  }
  return peak;
}

test('drag selection follows the pointer without flicker or snapping back', async ({ page }) => {
  const { scale, start, end } = await openDocument(page);
  const { samples, probe, released } = await drag(page, start, end);
  const first = samples.findIndex((sample) => sample.count > 0);
  expect(first, 'the drag never showed a selection').toBeGreaterThanOrEqual(0);
  const during = samples.slice(first);
  const empty = during.filter((sample) => sample.count === 0);
  const dragging = during.filter((sample) => sample.time < released);
  const backwards: { index: number; previous: DragSample; sample: DragSample }[] = [];
  let previous: DragSample | undefined;
  for (const [index, sample] of during.entries()) {
    if (sample.count === 0) continue;
    if (previous && (sample.bottom < previous.bottom - 0.5 ||
      Math.abs(sample.bottom - previous.bottom) <= 0.5 && sample.right < previous.right - 0.5)) backwards.push({ index, previous, sample });
    previous = sample;
  }
  console.log(`[drag-selection] samples=${during.length} raf=${during.filter((sample) => sample.source === 'raf').length} ` +
    `empty=${empty.length} emptyRaf=${empty.filter((sample) => sample.source === 'raf').length} backwards=${backwards.length} ` +
    `finalBottom=${during.at(-1)?.bottom.toFixed(1)} buttonToggles=${dragging.filter((sample, index) => index > 0 && sample.button !== dragging[index - 1]!.button).length} ` +
    `commands=${probe.sent.length} inFlightPeak=${inFlightPeak(probe)}`);
  console.log(`[drag-selection] frames ${JSON.stringify(during.filter((sample) => sample.source === 'raf').map((sample) => [sample.count, Math.round(sample.bottom), Math.round(sample.right)]))}`);
  expect(empty, 'the highlight emptied while the pointer moved away from the anchor').toEqual([]);
  expect(backwards, 'the highlight moved back toward the anchor').toEqual([]);
  const shown = during.filter((sample) => sample.count > 0);
  expect(shown.filter((sample) => Math.abs(sample.top - shown[0]!.top) > 0.5 || Math.abs(sample.left - shown[0]!.left) > 0.5),
    'the anchor end of the highlight moved').toEqual([]);
  const extents = new Set(during.filter((sample) => sample.source === 'raf' && sample.count > 0)
    .map((sample) => `${Math.round(sample.bottom)}:${Math.round(sample.right)}`));
  expect(extents.size, 'the highlight did not follow the pointer during the drag').toBeGreaterThanOrEqual(10);
  expect(shown.at(-1)!.bottom - shown[0]!.bottom).toBeGreaterThan(100 * scale);
  expect(new Set(dragging.map((sample) => sample.button)).size, 'the comment button toggled during the drag').toBe(1);
  await expect(page.locator(COMMENT_BUTTON)).toHaveCount(1);
});

test('a slow worker gets one drag extension at a time and the highlight keeps up', async ({ page }) => {
  const { start, end } = await openDocument(page, WORKER_MS);
  const { samples, probe, released } = await drag(page, start, end);
  // The release point's selection, from a separate one-move drag on an unthrottled worker.
  await page.evaluate(() => { window.__docxDragProbe.editDelayMs = 0; });
  const expected = (await drag(page, start, end, 1)).samples.at(-1)!;
  expect(sameExtent(samples.at(-1)!, expected), 'the drag ended away from the release point').toBe(true);
  const settled = samples.find((sample) => sample.time > released && sameExtent(sample, expected));
  const settleMs = settled ? Math.max(0, settled.time - released) : Number.NaN;
  const delays = latencies(samples, probe.moves);
  const p50 = percentile(delays, 0.5);
  const p95 = percentile(delays, 0.95);
  console.log(`[drag-selection] slow worker: commands=${probe.sent.length} moves=${probe.moves.length} inFlightPeak=${inFlightPeak(probe, released)} ` +
    `latency p50=${p50.toFixed(0)}ms p95=${p95.toFixed(0)}ms lines=${delays.length} settle=${settleMs.toFixed(0)}ms`);
  expect(inFlightPeak(probe, released), 'more than one drag extension waited on the worker').toBe(1);
  expect(probe.sent.length, 'every pointer move became a worker command').toBeLessThanOrEqual(STEPS * 0.6);
  expect(delays.length).toBeGreaterThanOrEqual(8);
  expect(settleMs, 'the highlight reached the release point late').toBeLessThan(4 * WORKER_MS + 500);
  expect(p95, 'the highlight fell behind the pointer').toBeLessThan(8 * WORKER_MS);
});

test('shift arrow selection keeps the comment button steady', async ({ page }) => {
  const { start } = await openDocument(page);
  await page.mouse.click(start.x, start.y);
  const input = page.locator('.editor-stage textarea').last();
  await expect(input).toBeFocused();
  await input.press('Shift+ArrowRight');
  await expect(page.locator(COMMENT_BUTTON)).toHaveCount(1);
  await startSampling(page);
  await page.evaluate(() => { window.__docxDragSamples!.active = true; });
  for (let step = 0; step < 40; step += 1) {
    await input.press('Shift+ArrowRight');
    await frame(page);
  }
  await page.evaluate(() => { window.__docxDragSamples!.active = false; });
  const samples = await page.evaluate(() => { window.__docxDragSamples!.stop(); return window.__docxDragSamples!.samples; });
  console.log(`[drag-selection] shift arrow: samples=${samples.length} hidden=${samples.filter((sample) => sample.button === 0).length}`);
  expect(samples.filter((sample) => sample.button === 0), 'the comment button hid during shift arrow selection').toEqual([]);
});
