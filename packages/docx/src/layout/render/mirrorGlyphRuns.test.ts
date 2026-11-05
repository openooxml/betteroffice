import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from './displayList';
import { buildMirrorPage, updateMirrorPage } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const options = { labels: { page: 'Page 1', header: 'Header', footer: 'Footer', noteBacklink: 'Back' } };
const ADVANCE = 8;

/** One glyph run per character, as the shaper emits for an exactly positioned line. */
function clusters(blockKey: string, start: number, text: string, x: number, baseline: number,
  extra: (index: number) => Partial<DisplayPrimitive> = () => ({}), rtl = false): DisplayPrimitive[] {
  return [...text].map((ch, index) => ({
    kind: 'glyphRun', fontId: 1, size: 16, color: '#000000', text: ch,
    glyphs: [{ id: 1, x: rtl ? x - (index + 1) * ADVANCE : x + index * ADVANCE, y: baseline, cluster: 0, advance: ADVANCE }],
    blockKey, docStart: start + index, docEnd: start + index + 1, logicalOrder: start + index, bidiLevel: rtl ? 1 : 0,
    ...extra(index),
  }) as DisplayPrimitive);
}

const page = (primitives: DisplayPrimitive[]): DisplayPage => ({ pageIndex: 0, width: 600, height: 800, primitives });
const runs = (mirror: HTMLElement) => Array.from(mirror.querySelectorAll<HTMLElement>('.layout-run-text'));

test('a line shaped cluster by cluster mirrors as one text run with its spaces', () => {
  const mirror = buildMirrorPage(page(clusters('p1', 10, 'Marker01 line 1 keeps', 96, 140)), options);
  const [run, ...rest] = runs(mirror);
  expect(rest).toHaveLength(0);
  expect(run.textContent).toBe('Marker01 line 1 keeps');
  expect(run.dataset.docStart).toBe('10');
  expect(run.dataset.docEnd).toBe('31');
  expect(run.style.left).toBe('96px');
  expect(parseFloat(run.style.width)).toBeCloseTo(21 * ADVANCE, 6);
  expect(mirror.querySelector('.layout-paragraph')!.textContent).toBe('Marker01 line 1 keeps');
});

test('wrapped lines, gaps and differing runs stay apart', () => {
  const mirror = buildMirrorPage(page([
    ...clusters('p1', 10, 'one', 96, 140),
    ...clusters('p1', 13, 'two', 96, 160),
    ...clusters('p1', 16, 'far', 300, 160),
    ...clusters('p1', 19, 'de', 324, 160, () => ({ lang: 'de-DE' })),
  ]), options);
  expect(runs(mirror).map((run) => run.textContent)).toEqual(['one', 'two', 'far', 'de']);
});

test('runs that show more text than their document range stay apart', () => {
  const field = (start: number, x: number): DisplayPrimitive => ({
    kind: 'glyphRun', fontId: 1, size: 16, color: '#000000', text: '2026',
    glyphs: [{ id: 1, x, y: 140, cluster: 0, advance: 4 * ADVANCE }],
    blockKey: 'p1', docStart: start, docEnd: start + 1, logicalOrder: start, bidiLevel: 0,
  }) as DisplayPrimitive;
  const mirror = buildMirrorPage(page([field(20, 100), field(21, 100 + 4 * ADVANCE)]), options);
  expect(runs(mirror).map((run) => [run.textContent, run.dataset.docStart, run.dataset.docEnd]))
    .toEqual([['2026', '20', '21'], ['2026', '21', '22']]);
});

test('right-to-left clusters join in logical order', () => {
  const mirror = buildMirrorPage(page(clusters('p1', 10, 'שלום', 200, 140, undefined, true)), options);
  const [run, ...rest] = runs(mirror);
  expect(rest).toHaveLength(0);
  expect(run.textContent).toBe('שלום');
  expect(run.dir).toBe('rtl');
  expect(run.style.left).toBe(`${200 - 4 * ADVANCE}px`);
  expect(parseFloat(run.style.width)).toBeCloseTo(4 * ADVANCE, 6);
});

test('an updated page matches a fresh build after an edit inside a joined line', () => {
  const mirror = buildMirrorPage(page([...clusters('p1', 10, 'alpha', 96, 140), ...clusters('p2', 20, 'beta', 96, 160)]), options);
  const next = page([...clusters('p1', 10, 'alphax', 96, 140), ...clusters('p2', 21, 'beta', 96, 160)]);
  expect(updateMirrorPage(mirror, next, options)).toBe(true);
  expect(mirror.innerHTML).toBe(buildMirrorPage(next, options).innerHTML);
  expect(runs(mirror).map((run) => [run.textContent, run.dataset.docStart, run.dataset.docEnd]))
    .toEqual([['alphax', '10', '16'], ['beta', '21', '25']]);
});
