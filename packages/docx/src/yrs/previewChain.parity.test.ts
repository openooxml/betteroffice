import { beforeAll, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { basename, resolve } from 'node:path';

import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import type { DisplayPage, DisplayPrimitive } from '../layout/render/displayList';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';
import {
  encodeDisplayListFrameExtras,
  type DisplayListBuildInputs,
} from '../layout/render/rustDisplayList';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { FLAVOURS, syntheticDocx } from './__fixtures__/previewChain';
import { decodeDocxHostJson, type YrsLoc, type YrsRenderEnv } from './index';
import { DisplayPositionIndex } from './displayPositionIndex';
import { displayPositionToYrsLoc } from './inputPositionMap';
import { finalPreviewDisplayWindow, finalPreviewPageCount } from './previewDisplayWindow';
import { createResidentEngineSession } from './residentEngineSession';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const CORPUS = resolve(import.meta.dir, '../../../../crates/betteroffice-docx/tests/corpus/fixtures');
const FONT = new Uint8Array(
  readFileSync(
    resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
  )
);

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

function renderEnvironment(document: Document): YrsRenderEnv {
  const themeColors: Record<string, string> = {};
  for (const [name, value] of Object.entries(document.package.theme?.colorScheme ?? {})) {
    if (typeof value === 'string') themeColors[name] = value;
  }
  return {
    themeColors,
    defaultTabStopTwips: document.package.settings?.defaultTabStop ?? null,
    numericIds: {},
    showHiddenText: false,
    mediaTokens: true,
  };
}

interface Built {
  refused: boolean;
  wholeBody: boolean;
  provisional: boolean;
  layoutPages: number;
  pages: DisplayPage[];
  positions: Array<{
    pageIndex: number;
    regions: Array<{ story: string; from: number; locs: Array<YrsLoc | null> }>;
  }>;
}

async function build(bytes: Uint8Array, preview: boolean): Promise<Built> {
  const session = await createResidentEngineSession();
  try {
    const hostJson = preview ? session.openDocxPreview(bytes, 200) : session.openDocx(bytes);
    if (hostJson === null) {
      return { refused: true, wholeBody: false, provisional: false, layoutPages: 0, pages: [], positions: [] };
    }
    const host = decodeDocxHostJson(hostJson, bytes);
    const document = host.document;
    const partial = preview && host.wholeBody !== true;
    const request = buildResidentRegionLayoutRequest(document, 24, renderEnvironment(document));
    const requirements = JSON.parse(
      session.layoutFontRequirementsJson(JSON.stringify(request))
    ) as Array<{ key: string }>;
    const window: [number, number] = [0, 2];
    session.setDisplayWindow(...window);
    session.setDisplayRetainBuiltPages(false);
    session.setWindowedIncrementalBuilds(true);
    session.setPartialDocument(partial);
    session.clearFonts();
    const font = session.registerFont(FONT);
    const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [font]]));
    const compat = document.package.settings?.compatibilityFlags;
    request.measurement = {
      fontChains,
      defaults: { fontSize: 11, fontFamily: 'Calibri' },
      compat: {
        noLeading: compat?.noLeading ?? false,
        doNotExpandShiftReturn: compat?.doNotExpandShiftReturn ?? false,
      },
      authoritativeShaping: true,
    };
    const input = JSON.stringify(request);
    const layout = JSON.parse(
      preview
        ? session.layoutDocumentWithRegionsPrefixRetainedJson(input, 3)
        : session.layoutDocumentWithRegionsRetainedJson(input)
    ) as { provisional?: boolean; layout: { pages: unknown[] } };
    const finalPages = finalPreviewPageCount(
      partial,
      layout.provisional === true,
      3,
      layout.layout.pages.length
    );
    if (finalPages !== null) session.setDisplayWindow(...finalPreviewDisplayWindow(window, finalPages));
    const retained = session.retainedHeadersFootersJson();
    const extras = encodeDisplayListFrameExtras({
      fontChains,
      ...(retained === undefined ? {} : { headersFooters: JSON.parse(retained) }),
    } as DisplayListBuildInputs);
    const frame = session.buildDisplayListFrame(extras, 0);
    const pages = applyFrameDelta(null, decodeFrameDelta(frame)).displayList.pages;
    const index = new DisplayPositionIndex({
      ...session.geometryReader,
      selectionText: session.selectionText,
      resolveComment: session.resolveComment,
    });
    const positions = pages.filter((page) => !page.unbuilt).map((page) => {
      const region = (story: string, primitives: DisplayPrimitive[]) => {
        const projection = index.projection(story)!;
        const starts = primitives.flatMap((primitive) =>
          [primitive.docStart, primitive.fragmentDocStart].filter((position): position is number => position !== undefined));
        const ends = primitives.flatMap((primitive) =>
          [primitive.docEnd, primitive.fragmentDocEnd].filter((position): position is number => position !== undefined));
        const from = page.pageIndex === 0 || story !== 'body' ? 0 : Math.min(...starts);
        const to = Math.min(projection.size, Math.max(...ends) + 1);
        const locs: Array<YrsLoc | null> = [];
        for (let position = from; position <= to; position += 1) {
          const target = projection.targetAt(position);
          const map = index.inputMap(target.story);
          locs.push(map ? displayPositionToYrsLoc(map, target.displayPosition) : null);
        }
        return { story, from, locs };
      };
      const regions = [region('body', page.primitives)];
      for (const band of [page.header, page.footer]) {
        if (band) regions.push(region(`hf:${band.rId}`, band.primitives));
      }
      for (const area of page.noteAreas ?? []) {
        for (const id of area.noteIds ?? []) {
          const story = `${area.kind === 'endnote' ? 'en' : 'fn'}:${id}`;
          const paraIds = new Set(session.geometryReader.paragraphSpans(story).map((paragraph) => paragraph.paraId));
          regions.push(region(story, (area.primitives ?? []).filter((primitive) =>
            primitive.paraId !== undefined && paraIds.has(primitive.paraId))));
        }
      }
      return { pageIndex: page.pageIndex, regions };
    });
    return {
      refused: false,
      wholeBody: host.wholeBody === true,
      provisional: layout.provisional === true,
      layoutPages: layout.layout.pages.length,
      pages,
      positions,
    };
  } finally {
    session.destroy();
  }
}

async function comparePreview(
  bytes: Uint8Array
): Promise<{ built: DisplayPage[]; wholeBody: boolean; provisional: boolean; fullPages: number }> {
  const preview = await build(bytes, true);
  expect(preview.refused).toBe(false);
  const full = await build(bytes, false);
  const built = preview.pages.filter((page) => page.unbuilt !== true);
  for (const page of built) {
    const twin = full.pages[page.pageIndex];
    expect(twin).toBeDefined();
    expect(twin?.unbuilt).not.toBe(true);
    expect(page).toEqual(twin);
    expect(preview.positions.find((entry) => entry.pageIndex === page.pageIndex)).toEqual(
      full.positions.find((entry) => entry.pageIndex === page.pageIndex)
    );
  }
  if (preview.wholeBody) {
    // A cut that holds the whole body shows the window's pages, as the whole document does.
    expect(built).toHaveLength(Math.min(2, full.pages.length));
  } else {
    const finalPages = finalPreviewPageCount(true, preview.provisional, 3, preview.layoutPages);
    expect(built).toHaveLength(Math.min(2, finalPages ?? 0));
    // Only a cut seed that laid out two pages or fewer has no final page.
    if (built.length === 0) {
      expect(preview.provisional).toBe(false);
      expect(preview.layoutPages).toBeLessThanOrEqual(2);
    }
  }
  return {
    built,
    wholeBody: preview.wholeBody,
    provisional: preview.provisional,
    fullPages: full.pages.length,
  };
}

const fixtures = readdirSync(CORPUS, { recursive: true, encoding: 'utf8' })
  .filter((name) => name.endsWith('.docx'))
  .sort();
for (const name of fixtures) {
  test(`worker preview matches full layout: ${name}`, async () => {
    const bytes = new Uint8Array(readFileSync(resolve(CORPUS, name)));
    if (basename(name) === 'wordprocessingml-comprehensive.docx') {
      expect((await build(bytes, true)).refused).toBe(true);
    } else {
      const { built } = await comparePreview(bytes);
      expect(built.length).toBeGreaterThan(0);
    }
  });
}

for (const [index, flavour] of FLAVOURS.entries()) {
  for (const size of [4, 10, 12, 22]) {
    test(`worker preview matches full layout: ${flavour} sz${size}`, async () => {
      const { built, wholeBody, provisional, fullPages } = await comparePreview(
        syntheticDocx(flavour, size, index * 100 + size)
      );
      expect(wholeBody).toBe(flavour === 'whole');
      if (flavour === 'whole') {
        // A whole body laid out to its end counts the document's pages, not the saved result.
        expect(provisional).toBe(false);
        const count = built[0]?.footer?.primitives.find(
          (primitive) => 'field' in primitive && primitive.field?.category === 'NUMPAGES'
        );
        expect(count && 'text' in count ? count.text : undefined).toBe(String(fullPages));
      } else {
        if (size === 22) expect(built).toHaveLength(2);
        if (size === 4) expect(built).toHaveLength(0);
      }
    });
  }
}
