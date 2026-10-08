import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { buildResidentRegionLayoutRequest } from '@betteroffice/docx/editor';
import {
  applyFrameDelta,
  createDisplayListQueries,
  decodeFrameDelta,
  encodeDisplayListFrameExtras,
  type DisplayListBuildInputs,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import { createRenderedDomContext } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import {
  computeAnchorDisplayTargets,
  createYrsPositionProjection,
  createYrsSession,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from '@betteroffice/docx/yrs/residentEngineSession';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  markPresented,
  stampRevisionPreviewKey,
  stampWorkerFrameVersion,
} from '../components/DocxEditor/internals/layoutProvenance';
import { bindDisplayWindow } from '../components/DocxEditor/internals/displayWindow';
import { storyAnchorsDocx } from './__fixtures__/storyAnchorsDocx';
import { createAnchorReadCache, createPluginGeometry } from './geometry';
import type { DocxAnchorGeometryResult, DocxGeometryTarget } from './types';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const FONT = new Uint8Array(
  readFileSync(
    resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
  )
);
const PAGE_GAP = 24;
const destroy: Array<{ destroy(): void }> = [];

beforeAll(async () => {
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(import.meta.dir, '../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  );
});
afterEach(() => {
  for (const session of destroy.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

interface LaidOut {
  main: YrsSession;
  engine: ResidentEngineSession;
  queries: DisplayListQueries;
}

/** The fixture laid out on the main session, with pages outside `[first, end)` left unbuilt. */
async function laidOut(first = 0, end = 2 ** 32 - 1, window?: [number, number]): Promise<LaidOut> {
  const bytes = await storyAnchorsDocx();
  const main = await createYrsSession();
  const engine = await createResidentEngineSession();
  destroy.push(main, engine);
  const host = main.openDocx(bytes, true);
  const request = buildResidentRegionLayoutRequest(host.document, PAGE_GAP, {});
  const requirements = JSON.parse(main.layoutFontRequirementsJson(JSON.stringify(request))) as Array<{
    key: string;
  }>;
  main.clearFonts();
  const font = main.registerFont(FONT);
  const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [font]]));
  request.measurement = {
    fontChains,
    defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
    compat: { noLeading: false, doNotExpandShiftReturn: false },
    authoritativeShaping: true,
  } as typeof request.measurement;
  const { headersFooters } = JSON.parse(
    main.layoutDocumentWithRegionsRetainedJson(JSON.stringify(request))
  ) as Pick<DisplayListBuildInputs, 'headersFooters'>;
  main.setDisplayWindow(first, end);
  main.setDisplayRetainBuiltPages(false);
  const extras = encodeDisplayListFrameExtras({ fontChains, headersFooters } as DisplayListBuildInputs);
  const { displayList } = applyFrameDelta(null, decodeFrameDelta(main.buildDisplayListFrame(extras, 0)));
  const queries = createDisplayListQueries(displayList, main as never);
  if (window) bindDisplayWindow(queries, { read: () => window, subscribe: () => () => {} });
  engine.loadState(main.encodeState());
  return { main, engine, queries };
}

function pagesOf(queries: DisplayListQueries) {
  const pages = document.createElement('div');
  let top = 0;
  const pageList = queries.displayList?.pages ?? [];
  const width = Math.max(...pageList.map((page) => page.width));
  for (const page of pageList) {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(page.pageIndex);
    const rect = new DOMRect(0, top, page.width, page.height);
    canvas.getBoundingClientRect = () => rect;
    pages.appendChild(canvas);
    top += page.height + PAGE_GAP;
  }
  const bounds = new DOMRect(0, 0, width, top);
  pages.getBoundingClientRect = () => bounds;
  return pages;
}

/** Plugin geometry over the main session (`viewer` false) or a worker viewer reading `engine`. */
function geometryOf({ main, engine, queries }: LaidOut, viewer: boolean) {
  const pages = pagesOf(queries);
  const layer = document.createElement('div');
  layer.getBoundingClientRect = () => pages.getBoundingClientRect();
  markPresented(pages, queries.displayList);
  stampRevisionPreviewKey(queries, '');
  stampWorkerFrameVersion(queries, engine.geometryReader.version());
  const projection = createYrsPositionProjection(main, 'body');
  return createPluginGeometry(
    { id: 'layout', version: main.version(), previewVersion: 0, zoom: 1, pageCount: queries.pageCount() },
    createRenderedDomContext(pages, 1),
    layer,
    () => true,
    () => null,
    queries,
    () =>
      viewer
        ? null
        : {
            session: main,
            presented: true,
            editor: {
              hasPendingInput: () => false,
              yrsLocToDisplayPosition: (loc) => projection?.positionForLoc(loc) ?? null,
            },
          },
    () => false,
    undefined,
    viewer
      ? {
          read: async (targets) => computeAnchorDisplayTargets(engine.geometryReader, targets, undefined),
          cache: createAnchorReadCache(),
        }
      : undefined
  );
}

function span(laid: LaidOut, story: string, start: [string, number], end: [string, number]) {
  return {
    kind: 'range',
    version: laid.main.version(),
    range: {
      story,
      start: { paraId: start[0], offset: start[1] },
      end: { paraId: end[0], offset: end[1] },
      view: 'accepted',
    },
  } as const;
}

function range(laid: LaidOut, story: string, paraId: string, text: string) {
  const offset = laid.main.paragraphs(story).find((p) => p.paraId === paraId)!.text.indexOf(text);
  return span(laid, story, [paraId, offset], [paraId, offset + text.length]);
}

function sessionParagraph(laid: LaidOut, story: string, paraId: string) {
  return {
    kind: 'paragraph',
    paragraph: { kind: 'session', sessionId: laid.main.paragraphIdentities().sessionId, story, paraId },
  } as const;
}

function ok(result: DocxAnchorGeometryResult) {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result;
}

const pageIndices = (result: DocxAnchorGeometryResult) => [
  ...new Set(ok(result).rects.map(({ pageIndex }) => pageIndex)),
];

/** Both readers answer alike once every page is built. */
async function both(laid: LaidOut, target: DocxGeometryTarget) {
  const main = ok(geometryOf(laid, false).getAnchorGeometry(target));
  const worker = ok(await geometryOf(laid, true).readAnchorGeometry(target));
  expect(worker).toEqual({ ...main, unbuiltPages: [] });
  return main;
}

describe('table cell targets', () => {
  test('a cell token resolves to its own rects inside the cell', async () => {
    const laid = await laidOut();
    const token = await both(laid, range(laid, 'body:t0:r0c0', '00000002', '{{cell}}'));
    const intro = await both(laid, range(laid, 'body', '00000001', 'Intro'));
    expect(token.rects).toHaveLength(1);
    const [rect] = token.rects;
    expect(rect!.pageIndex).toBe(0);
    expect(rect!.x).toBeGreaterThan(intro.rects[0]!.x);
    expect(rect!.y).toBeGreaterThan(intro.rects[0]!.y);
    expect(token.anchor).toMatchObject({ pageIndex: 0, x: rect!.x + rect!.width, width: 0 });

    const paragraph = await both(laid, sessionParagraph(laid, 'body:t0:r0c0', '00000002'));
    expect(paragraph.rects[0]!.x).toBeLessThan(rect!.x);
    expect(paragraph.rects[0]!.y).toBe(rect!.y);
    const search = await both(laid, {
      kind: 'search',
      paragraph: sessionParagraph(laid, 'body:t0:r0c0', '00000002').paragraph,
      text: '{{cell}}',
    });
    expect(search.rects).toEqual(token.rects);
  });

  test('a range across a split row has rects on both pages and its anchor at the end', async () => {
    const laid = await laidOut();
    const split = await both(laid, span(laid, 'body:t1:r0c0', ['20000000', 0], ['20000045', 13]));
    expect(pageIndices(split)).toEqual([1, 2]);
    expect(split.anchor.pageIndex).toBe(2);
  });
});

describe('header and footer targets', () => {
  test('a default header paints on every odd page after the first', async () => {
    const laid = await laidOut();
    const odd = await both(laid, range(laid, 'hf:rIdH1', '10000001', '{{odd}}'));
    expect(pageIndices(odd)).toEqual([2, 4]);
    expect(odd.anchor.pageIndex).toBe(2);
    const [first, second] = odd.rects;
    expect(second!.x).toBe(first!.x);
    expect(second!.y - first!.y).toBeCloseTo(2 * (laid.queries.pageSize(2)!.height + PAGE_GAP));
    const cell = await both(laid, range(laid, 'hf:rIdH1:t0:r0c0', '10000002', '{{hcell}}'));
    expect(pageIndices(cell)).toEqual([2, 4]);
    expect(cell.rects[0]!.y).toBeGreaterThan(first!.y);
  });

  test('an even header paints on even pages and a first-page footer only on the first', async () => {
    const laid = await laidOut();
    expect(pageIndices(await both(laid, sessionParagraph(laid, 'hf:rIdH2', '10000011')))).toEqual([1, 3, 5]);
    const footer = await both(laid, range(laid, 'hf:rIdF1', '10000021', '{{first}}'));
    expect(pageIndices(footer)).toEqual([0]);
    expect(footer.anchor.y).toBeGreaterThan(500);
  });

  test('an empty header paragraph anchors on its caret on the first page', async () => {
    const laid = await laidOut();
    const empty = await both(laid, sessionParagraph(laid, 'hf:rIdH1', '10000003'));
    expect(empty.rects).toEqual([]);
    expect(empty.anchor.pageIndex).toBe(2);
  });

  test('pages not built yet are listed and the anchor stays on the first painting page', async () => {
    const laid = await laidOut(0, 3);
    const pages = laid.queries.displayList!.pages;
    expect(pages.map(({ unbuilt, hfParts }) => [unbuilt ?? false, hfParts?.header ?? null])).toEqual([
      [false, null],
      [false, null],
      [false, null],
      [true, 'rIdH2'],
      [true, 'rIdH1'],
      [true, 'rIdH2'],
    ]);
    const worker = geometryOf(laid, true);
    const odd = ok(await worker.readAnchorGeometry(range(laid, 'hf:rIdH1', '10000001', '{{odd}}')));
    expect(pageIndices(odd)).toEqual([2]);
    expect(odd.unbuiltPages).toEqual([4]);
    expect(odd.anchor.pageIndex).toBe(2);
    const even = ok(await worker.readAnchorGeometry(sessionParagraph(laid, 'hf:rIdH2', '10000011')));
    expect(pageIndices(even)).toEqual([1]);
    expect(even.unbuiltPages).toEqual([3, 5]);
    const footer = ok(await worker.readAnchorGeometry(range(laid, 'hf:rIdF1', '10000021', '{{first}}')));
    expect(footer.unbuiltPages).toEqual([]);

    const late = await laidOut(3, 6);
    const lateOdd = ok(
      await geometryOf(late, true).readAnchorGeometry(range(late, 'hf:rIdH1', '10000001', '{{odd}}'))
    );
    expect(pageIndices(lateOdd)).toEqual([4]);
    expect(lateOdd.unbuiltPages).toEqual([2]);
    expect(lateOdd.anchor).toMatchObject({ pageIndex: 2, width: 0, height: 0 });
  });

  test('the main thread refuses while a visible page that paints the part is unbuilt', async () => {
    const laid = await laidOut(0, 3, [3, 5]);
    const main = geometryOf(laid, false);
    expect(main.getAnchorGeometry(range(laid, 'hf:rIdH1', '10000001', '{{odd}}'))).toMatchObject({
      ok: false,
      failure: { code: 'layout-unavailable' },
    });
    const footer = ok(main.getAnchorGeometry(range(laid, 'hf:rIdF1', '10000021', '{{first}}')));
    expect(footer.unbuiltPages).toBeUndefined();
    expect(pageIndices(footer)).toEqual([0]);
  });

  test('footnote stories stay unsupported', async () => {
    const laid = await laidOut();
    for (const viewer of [false, true]) {
      const geometry = geometryOf(laid, viewer);
      const intro = range(laid, 'body', '00000001', 'Intro');
      expect(
        await geometry.readAnchorGeometry({ ...intro, range: { ...intro.range, story: 'fn:1' } })
      ).toMatchObject({ ok: false, failure: { code: 'unsupported' } });
    }
  });
});
