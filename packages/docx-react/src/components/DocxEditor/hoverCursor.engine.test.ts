/**
 * End-to-end: the real layout wasm builds a display list, the real query
 * facade answers, and the mapping turns that into a cursor.
 *
 * The header-band case is the only probe that reaches the direct-run branch —
 * a band has no typeable area, so nothing else there can answer `text`.
 */

import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  createDisplayListQueries,
  type DisplayList,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import { buildDisplayListJson, preloadLayoutWasm } from '@betteroffice/docx/wasm/layout';
import { canvasHoverCursor } from './hoverCursor';

const root = resolve(import.meta.dir, '../../../../..');
const fixture = (path: string) => resolve(root, 'crates/docx-layout/tests/fixtures', path);
// three one-line paragraphs on an 816x1056 page whose content box is
// x 96..720, y 96..960
const BODY = fixture('single-page-multi-paragraph.input.json');
// one header band at y 48..72 and one footer band at y 984..1008, each with a
// paragraph; a band carries no content box
const BANDS = fixture('hf/hf-default-both.input.json');

const facades: DisplayListQueries[] = [];

function open(path: string): { list: DisplayList; queries: DisplayListQueries } {
  const list = JSON.parse(buildDisplayListJson(readFileSync(path, 'utf8'))) as DisplayList;
  const queries = createDisplayListQueries(list);
  facades.push(queries);
  return { list, queries };
}

let body: ReturnType<typeof open>;
let bands: ReturnType<typeof open>;

const cursorAt = (
  opened: ReturnType<typeof open>,
  x: number,
  y: number,
  readOnly = false
): string => canvasHoverCursor({ readOnly }, opened.queries.hitTestRegions(0, x, y));

beforeAll(async () => {
  await preloadLayoutWasm();
  body = open(BODY);
  bands = open(BANDS);
  await Promise.all(facades.map((facade) => facade.whenReady()));
});

afterAll(() => {
  for (const facade of facades) facade.dispose();
});

describe('hover cursor over an engine-built page', () => {
  test('every laid-out run reads as text', () => {
    const runs = body.list.pages[0].primitives.filter((primitive) => primitive.kind === 'text');
    expect(runs.length).toBeGreaterThan(0);
    for (const run of runs) {
      expect(cursorAt(body, run.x + run.width / 2, run.baselineY - 4)).toBe('text');
    }
  });

  test('the content box is typeable and the margins are not', () => {
    const page = body.list.pages[0];
    const content = page.contentBounds;
    if (!content) throw new Error('the fixture page has no content box');
    const inside = (x: number, y: number) =>
      x >= content.x &&
      x <= content.x + content.width &&
      y >= content.y &&
      y <= content.y + content.height;

    const disagreements: string[] = [];
    for (let y = 0; y <= page.height; y += 24) {
      for (let x = 0; x <= page.width; x += 24) {
        const cursor = cursorAt(body, x, y);
        const expected = inside(x, y) ? 'text' : 'default';
        if (cursor !== expected) disagreements.push(`(${x},${y}) ${cursor} != ${expected}`);
      }
    }
    expect(disagreements).toEqual([]);
  });

  test('a read-only document never shows the caret cursor', () => {
    const run = body.list.pages[0].primitives.find((primitive) => primitive.kind === 'text');
    if (!run) throw new Error('the fixture page has no run');
    expect(cursorAt(body, run.x + run.width / 2, run.baselineY - 4, true)).toBe('default');
  });

  // A band has no content box, so only a run's own box can answer here — this
  // is what separates the direct-hit branch from the typeable-area fallback.
  test('a header band reads as text over its runs and nowhere else', () => {
    const page = bands.list.pages[0];
    const band = page.header;
    if (!band) throw new Error('the fixture page has no header band');
    const run = band.primitives.find((primitive) => primitive.kind === 'text');
    if (!run) throw new Error('the header band has no run');

    const hit = bands.queries.hitTestRegions(0, run.x + run.width / 2, run.baselineY - 2);
    expect(hit?.region).toBe('header');
    expect(cursorAt(bands, run.x + run.width / 2, run.baselineY - 2)).toBe('text');

    // same band, past the run: inside the region, outside every run
    const past = run.x + run.width + 200;
    expect(bands.queries.hitTestRegions(0, past, run.baselineY - 2)?.region).toBe('header');
    expect(cursorAt(bands, past, run.baselineY - 2)).toBe('default');
  });
});
