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
  computeAnchorTargetGeometry,
  computeProposalGeometryMirror,
  createYrsPositionProjection,
  createYrsSession,
  proposalRevisionPreview,
  type DocxProposalSnapshot,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from '@betteroffice/docx/yrs/residentEngineSession';
import { sessionInternals } from '@betteroffice/docx/yrs/sessionInternals';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { markPresented, stampRevisionPreviewKey } from '../components/DocxEditor/internals/layoutProvenance';
import { bindDisplayWindow } from '../components/DocxEditor/internals/displayWindow';
import { sectionedDocx, sharedHeaderDocx, storyAnchorsDocx } from './__fixtures__/storyAnchorsDocx';
import { hiddenRanges } from './anchorGeometry';
import { createPluginGeometry } from './geometry';
import { currentPreviewKey } from './proposalPreview';
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
async function laidOut(
  first = 0,
  end = 2 ** 32 - 1,
  window?: [number, number],
  fixture: () => Promise<Uint8Array> = storyAnchorsDocx,
  prepare: (main: YrsSession) => void = () => {}
): Promise<LaidOut> {
  const bytes = await fixture();
  const main = await createYrsSession();
  const engine = await createResidentEngineSession();
  destroy.push(main, engine);
  const host = main.openDocx(bytes, true);
  prepare(main);
  const request = buildResidentRegionLayoutRequest(host.document, PAGE_GAP, {
    revisionPreview: proposalRevisionPreview(main.getProposals()),
  });
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

/**
 * Plugin geometry over the main session, or in worker mode over a stub session whose targets the
 * resident engine resolves, as the worker proposal authority serves them.
 */
function geometryOf({ main, engine, queries }: LaidOut, worker: boolean, complete = false) {
  const pages = pagesOf(queries);
  const layer = document.createElement('div');
  layer.getBoundingClientRect = () => pages.getBoundingClientRect();
  markPresented(pages, queries.displayList);
  stampRevisionPreviewKey(queries, currentPreviewKey(main));
  const projection = createYrsPositionProjection(main, 'body');
  const reader = engine.geometryReader;
  const snapshot: DocxProposalSnapshot = { ...main.getProposals(), version: reader.version() };
  const mirror = computeProposalGeometryMirror(reader, snapshot);
  const stub = { version: () => mirror.version, getProposals: () => snapshot } as unknown as YrsSession;
  const geometry = createPluginGeometry(
    {
      id: 'layout',
      version: worker ? mirror.version : main.version(),
      previewVersion: snapshot.previewVersion,
      zoom: 1,
      pageCount: queries.pageCount(),
    },
    createRenderedDomContext(pages, 1),
    layer,
    () => true,
    () => null,
    queries,
    () =>
      worker
        ? {
            session: stub,
            presented: true,
            proposalGeometry: mirror,
            anchorTarget: (target) =>
              computeAnchorTargetGeometry(reader, [target], () => hiddenRanges(main, main.version()))[0],
            editor: {
              hasPendingInput: () => false,
              yrsLocToDisplayPosition: () => {
                throw new Error('worker geometry must not read replica positions');
              },
            },
          }
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
    () => complete
  );
  return (target: DocxGeometryTarget) =>
    geometry.getAnchorGeometry(
      worker && target.kind === 'range' ? { ...target, version: mirror.version } : target
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

/** The main thread and worker mode answer alike. */
function both(laid: LaidOut, target: DocxGeometryTarget, complete = false) {
  const main = ok(geometryOf(laid, false, complete)(target));
  expect(ok(geometryOf(laid, true, complete)(target))).toEqual({ ...main, version: laid.engine.geometryReader.version() });
  return main;
}

describe('table cell targets', () => {
  test('a cell token resolves to its own rects inside the cell', async () => {
    const laid = await laidOut();
    const token = both(laid, range(laid, 'body:t0:r0c0', '00000002', '{{cell}}'));
    const intro = both(laid, range(laid, 'body', '00000001', 'Intro'));
    expect(token.rects).toHaveLength(1);
    const [rect] = token.rects;
    expect(rect!.pageIndex).toBe(0);
    expect(rect!.x).toBeGreaterThan(intro.rects[0]!.x);
    expect(rect!.y).toBeGreaterThan(intro.rects[0]!.y);
    expect(token.anchor).toMatchObject({ pageIndex: 0, x: rect!.x + rect!.width, width: 0 });

    const paragraph = both(laid, sessionParagraph(laid, 'body:t0:r0c0', '00000002'));
    expect(paragraph.rects[0]!.x).toBeLessThan(rect!.x);
    expect(paragraph.rects[0]!.y).toBe(rect!.y);
    const search = both(laid, {
      kind: 'search',
      paragraph: sessionParagraph(laid, 'body:t0:r0c0', '00000002').paragraph,
      text: '{{cell}}',
    });
    expect(search.rects).toEqual(token.rects);
  });

  test('a range across a split row has rects on both pages and its anchor at the end', async () => {
    const laid = await laidOut();
    const split = both(laid, span(laid, 'body:t1:r0c0', ['20000000', 0], ['20000045', 13]));
    expect(pageIndices(split)).toEqual([1, 2]);
    expect(split.anchor.pageIndex).toBe(2);
  });
});

describe('header and footer targets', () => {
  test('a default header paints on every odd page after the first', async () => {
    const laid = await laidOut();
    const odd = both(laid, range(laid, 'hf:rIdH1', '10000001', '{{odd}}'));
    expect(pageIndices(odd)).toEqual([2, 4]);
    expect(odd.anchor.pageIndex).toBe(2);
    const [first, second] = odd.rects;
    expect(second!.x).toBe(first!.x);
    expect(second!.y - first!.y).toBeCloseTo(2 * (laid.queries.pageSize(2)!.height + PAGE_GAP));
    expect(
      both(laid, {
        kind: 'search',
        paragraph: sessionParagraph(laid, 'hf:rIdH1', '10000001').paragraph,
        text: '{{odd}}',
      }).rects
    ).toEqual(odd.rects);
    const cell = both(laid, range(laid, 'hf:rIdH1:t0:r0c0', '10000002', '{{hcell}}'));
    expect(pageIndices(cell)).toEqual([2, 4]);
    expect(cell.rects[0]!.y).toBeGreaterThan(first!.y);
  });

  test('an even header paints on even pages and a first-page footer only on the first', async () => {
    const laid = await laidOut();
    expect(pageIndices(both(laid, sessionParagraph(laid, 'hf:rIdH2', '10000011')))).toEqual([1, 3, 5]);
    const revision = laid.main.listRevisions().find(({ story }) => story === 'hf:rIdH2')!;
    const deleted = both(laid, { kind: 'revision', revisionId: revision.revisionId });
    expect(pageIndices(deleted)).toEqual([1, 3, 5]);
    expect(deleted.rects[0]!.x).toBeGreaterThan(
      both(laid, sessionParagraph(laid, 'hf:rIdH2', '10000012')).rects[0]!.x
    );
    const footer = both(laid, range(laid, 'hf:rIdF1', '10000021', '{{first}}'));
    expect(pageIndices(footer)).toEqual([0]);
    expect(footer.anchor.y).toBeGreaterThan(500);
  });

  test('an empty header paragraph anchors on its caret on the first page', async () => {
    const laid = await laidOut();
    const empty = both(laid, sessionParagraph(laid, 'hf:rIdH1', '10000003'));
    expect(empty.rects).toEqual([]);
    expect(empty.anchor.pageIndex).toBe(2);
  });

  test('a hidden leading header deletion anchors before the first visible character on its first painting page', async () => {
    const laid = await laidOut(0, 2 ** 32 - 1, undefined, storyAnchorsDocx, (main) => {
      const proposed = main.proposeChanges({
        expectVersion: main.version(),
        proposals: [
          {
            id: 'header-deletion',
            paragraph: {
              kind: 'session',
              sessionId: main.paragraphIdentities().sessionId,
              story: 'hf:rIdH1',
              paraId: '10000001',
            },
            suggest: { author: 'Reviewer', date: '2026-10-01T00:00:00Z' },
            op: 'replaceText',
            search: 'Odd ',
            replaceWith: '',
          },
        ],
      });
      if (!proposed.ok) throw new Error(proposed.failure.message);
      const accepted = main.setProposalStates({
        expectVersion: main.version(),
        expectPreviewVersion: proposed.snapshot.previewVersion,
        changes: [{ id: 'header-deletion', state: 'accepted' }],
      });
      if (!accepted.ok) throw new Error(accepted.failure.message);
    });
    expect(laid.queries.displayList!.pages.some(({ unbuilt }) => unbuilt)).toBe(false);
    const firstCharacter = both(laid, span(laid, 'hf:rIdH1', ['10000001', 0], ['10000001', 1]), true);
    expect(pageIndices(firstCharacter)).toEqual([2, 4]);
    const hidden = both(
      laid,
      { kind: 'revision', revisionId: laid.main.getProposals().proposals[0]!.revisionIds[0]! },
      true
    );
    expect(hidden.rects).toEqual([]);
    expect(hidden.anchor).toEqual({ ...firstCharacter.rects[0]!, width: 0 });
  });

  test('pages outside the display window that are not built yet are left out', async () => {
    const laid = await laidOut(0, 3);
    expect(
      laid.queries.displayList!.pages.map(({ unbuilt, hfParts }) => [unbuilt ?? false, hfParts?.header ?? null])
    ).toEqual([
      [false, null],
      [false, null],
      [false, null],
      [true, 'rIdH2'],
      [true, 'rIdH1'],
      [true, 'rIdH2'],
    ]);
    const odd = both(laid, range(laid, 'hf:rIdH1', '10000001', '{{odd}}'));
    expect(pageIndices(odd)).toEqual([2]);
    expect(odd.anchor.pageIndex).toBe(2);
    expect(pageIndices(both(laid, sessionParagraph(laid, 'hf:rIdH2', '10000011')))).toEqual([1]);

    const late = await laidOut(3, 6);
    const lateOdd = both(late, range(late, 'hf:rIdH1', '10000001', '{{odd}}'));
    expect(pageIndices(lateOdd)).toEqual([4]);
    expect(lateOdd.anchor.pageIndex).toBe(4);

    const early = await laidOut(0, 1);
    const unpainted = both(early, range(early, 'hf:rIdH1', '10000001', '{{odd}}'));
    expect(unpainted.rects).toEqual([]);
    expect(unpainted.anchor).toMatchObject({ pageIndex: 2, width: 0, height: 0 });
  });

  test('a visible page that paints the part and is not built yet refuses', async () => {
    const laid = await laidOut(0, 3, [3, 5]);
    for (const worker of [false, true]) {
      const geometry = geometryOf(laid, worker);
      expect(geometry(range(laid, 'hf:rIdH1', '10000001', '{{odd}}'))).toMatchObject({
        ok: false,
        failure: { code: 'layout-unavailable' },
      });
      expect(pageIndices(geometry(range(laid, 'hf:rIdF1', '10000021', '{{first}}')))).toEqual([0]);
    }
  });

  test('unbuilt pages name the parts their build paints across sections and a parity filler', async () => {
    const full = await laidOut(0, 2 ** 32 - 1, undefined, sectionedDocx);
    const partial = await laidOut(0, 1, undefined, sectionedDocx);
    const painted = full.queries.displayList!.pages.map(({ header, footer }) => [
      header?.rId ?? null,
      footer?.rId ?? null,
    ]);
    expect(painted).toEqual([
      ['rIdA', 'rIdF'],
      ['rIdE', null],
      ['rIdA', 'rIdF'],
      [null, null],
      ['rIdB', null],
      ['rIdC', null],
      ['rIdA', 'rIdF'],
    ]);
    expect(
      partial.queries.displayList!.pages.map(({ unbuilt, hfParts }) => [
        unbuilt ?? false,
        hfParts?.header ?? null,
        hfParts?.footer ?? null,
      ])
    ).toEqual([[false, null, null], ...painted.slice(1).map(([header, footer]) => [true, header, footer])]);
    for (const [index, rId] of ['rIdA', 'rIdE', 'rIdF', 'rIdB', 'rIdC'].entries()) {
      const pages = painted.flatMap((band, page) => (band.includes(rId) ? [page] : []));
      const paraId = `1000000${index + 1}`;
      expect(pageIndices(both(full, sessionParagraph(full, `hf:${rId}`, paraId)))).toEqual(pages);
      const early = both(partial, sessionParagraph(partial, `hf:${rId}`, paraId));
      expect(early.anchor.pageIndex).toBe(pages[0]);
      expect(pageIndices(early)).toEqual(pages[0] === 0 ? [0] : []);
    }
  });

  test('a shared story is placed under every relationship routed to it', async () => {
    const laid = await laidOut(0, 2 ** 32 - 1, undefined, sharedHeaderDocx, (main) => {
      for (const story of main.storyIds()) {
        if (story === 'hf:rIdS2' || story.startsWith('hf:rIdS2:')) main.deleteStory(story);
      }
      sessionInternals(main).setHeaderFooterAliases(
        JSON.stringify([{ isHeader: true, partPath: 'word/rIdS1.xml', relationshipIds: ['rIdS1', 'rIdS2'] }])
      );
    });
    expect(laid.queries.displayList!.pages.map(({ header }) => header?.rId)).toEqual(['rIdS1', 'rIdS2']);
    const shared = both(laid, range(laid, 'hf:rIdS1', '10005001', '{{shared}}'));
    expect(pageIndices(shared)).toEqual([0, 1]);
    expect(shared.anchor.pageIndex).toBe(0);
  });

  test('a part no page paints is missing once the layout is complete, unbuilt pages included', async () => {
    const laid = await laidOut(0, 3);
    const target = range(laid, 'hf:rIdF1', '10000021', '{{first}}');
    const pages = laid.queries.displayList!.pages;
    expect(pages.some(({ unbuilt }) => unbuilt)).toBe(true);
    for (const page of pages) {
      delete page.footer;
      if (page.hfParts) delete page.hfParts.footer;
    }
    for (const worker of [false, true]) {
      expect(geometryOf(laid, worker, true)(target)).toEqual({
        ok: false,
        failure: { code: 'missing-target', message: 'No page paints this header or footer' },
      });
      expect(geometryOf(laid, worker)(target)).toEqual({
        ok: false,
        failure: { code: 'layout-unavailable', message: 'No laid-out page paints this header or footer' },
      });
    }
  });

  test('footnote stories stay unsupported', async () => {
    const laid = await laidOut();
    const intro = range(laid, 'body', '00000001', 'Intro');
    for (const worker of [false, true]) {
      expect(geometryOf(laid, worker)({ ...intro, range: { ...intro.range, story: 'fn:1' } })).toMatchObject({
        ok: false,
        failure: { code: 'unsupported' },
      });
    }
  });
});
