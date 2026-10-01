import { beforeAll, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { basename, resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import type { DisplayPage } from '../layout/render/displayList';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';
import {
  encodeDisplayListFrameExtras,
  type DisplayListBuildInputs,
} from '../layout/render/rustDisplayList';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { decodeDocxHostJson, type YrsRenderEnv } from './index';
import { finalPreviewDisplayWindow, finalPreviewPageCount } from './previewDisplayWindow';
import { createResidentEngineSession } from './residentEngineSession';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const CORPUS = resolve(import.meta.dir, '../../../../crates/betteroffice-docx/tests/corpus/fixtures');
const FONT = new Uint8Array(
  readFileSync(
    resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
  )
);
const BLOCKS = 320;
const WHOLE_BLOCKS = 120;
const FLAVOURS = [
  'plain',
  'keepnext',
  'keepnext-style',
  'widows',
  'footnotes',
  'endnotes',
  'floats',
  'final-sect',
  'early-sect',
  'whole',
] as const;
type Flavour = (typeof FLAVOURS)[number];

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
}

async function build(bytes: Uint8Array, preview: boolean): Promise<Built> {
  const session = await createResidentEngineSession();
  try {
    const hostJson = preview ? session.openDocxPreview(bytes, 200) : session.openDocx(bytes);
    if (hostJson === null) {
      return { refused: true, wholeBody: false, provisional: false, layoutPages: 0, pages: [] };
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
    return {
      refused: false,
      wholeBody: host.wholeBody === true,
      provisional: layout.provisional === true,
      layoutPages: layout.layout.pages.length,
      pages: applyFrameDelta(null, decodeFrameDelta(frame)).displayList.pages,
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

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const DRAW =
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const WORDS =
  'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau'.split(' ');
const PNG = new Uint8Array(
  Buffer.from(
    '89504e470d0a1a0a0000000d4948445200000001000000010806000000' +
      '1f15c4890000000d49444154789c6360f8cfc0000003010100c9fe92ef0000000049454e44ae426082',
    'hex'
  )
);

function run(text: string): string {
  return `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
}

function field(instruction: string, cached: string): string {
  return (
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    `<w:r><w:t>${cached}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function anchor(id: number): string {
  return (
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="1828800" cy="1371600"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${id}" name="Float ${id}"/>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="f${id}.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
    '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1371600"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>' +
    '</wp:anchor></w:drawing></w:r>'
  );
}

function section(landscape = false, title = false, start?: number): string {
  const size = landscape
    ? '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    : '<w:pgSz w:w="12240" w:h="15840"/>';
  const margins = landscape
    ? '<w:pgMar w:top="1080" w:right="1800" w:bottom="1080" w:left="1800" w:header="540" w:footer="540" w:gutter="0"/>'
    : '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
  return (
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>' +
    '<w:footerReference w:type="default" r:id="rIdF1"/>' +
    (title ? '<w:headerReference w:type="first" r:id="rIdH2"/>' : '') +
    size +
    margins +
    (start === undefined ? '' : `<w:pgNumType w:start="${start}"/>`) +
    (title ? '<w:titlePg/>' : '') +
    '</w:sectPr>'
  );
}

function syntheticDocx(flavour: Flavour, size: number, seed: number): Uint8Array {
  let state = seed;
  const random = (limit: number) => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state % limit;
  };
  const text = (low: number, high: number) => {
    const words = Array.from(
      { length: low + random(high - low + 1) },
      () => WORDS[random(WORDS.length)]
    );
    const sentence = words.join(' ');
    return sentence[0].toUpperCase() + sentence.slice(1) + '.';
  };
  const parts: PartsMap = new Map();
  const set = (name: string, content: string) => parts.set(name, toBytes(content));
  const rels: Array<[string, string, string]> = [
    ['rIdS', 'styles', 'styles.xml'],
    ['rIdH1', 'header', 'header1.xml'],
    ['rIdH2', 'header', 'header2.xml'],
    ['rIdF1', 'footer', 'footer1.xml'],
  ];
  const overrides: Array<[string, string]> = [
    ['/word/styles.xml', 'styles'],
    ['/word/header1.xml', 'header'],
    ['/word/header2.xml', 'header'],
    ['/word/footer1.xml', 'footer'],
  ];
  const footnotes: number[] = [];
  const endnotes: number[] = [];
  const body: string[] = [];
  const blocks = flavour === 'whole' ? WHOLE_BLOCKS : BLOCKS;
  for (let i = 1; i <= blocks; i += 1) {
    let properties = '';
    let runs = run(text(3, 9));
    if (flavour === 'keepnext' && i >= 150 && i <= 230 && i % 40 < 30) {
      properties += '<w:keepNext/>';
    }
    if (flavour === 'keepnext-style' && i % 25 < 18) {
      properties += '<w:pStyle w:val="Keep"/>';
    }
    if (flavour === 'widows' && i % 7 === 0) {
      properties += '<w:widowControl/><w:keepLines/>';
      runs = run(text(40, 90));
    }
    if (flavour === 'footnotes' && i % 9 === 0) {
      footnotes.push(footnotes.length + 1);
      runs += `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${footnotes.length}"/></w:r>`;
    }
    if (flavour === 'endnotes' && i % 13 === 0) {
      endnotes.push(endnotes.length + 1);
      runs += `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="${endnotes.length}"/></w:r>`;
    }
    if (flavour === 'floats' && i % 45 === 0) runs = anchor(i) + runs;
    if (flavour === 'early-sect' && i === 300) properties += section(true, true, 7);
    body.push(
      `<w:p><w:pPr>${properties}<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
        `<w:rPr><w:sz w:val="${size}"/></w:rPr></w:pPr>${runs}</w:p>`
    );
  }
  set(
    'word/styles.xml',
    `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/>` +
      `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:rPrDefault>` +
      '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Keep"><w:name w:val="Keep"/><w:pPr><w:keepNext/></w:pPr></w:style></w:styles>'
  );
  set('word/header1.xml', `<w:hdr ${W}><w:p>${run('Default header')}</w:p></w:hdr>`);
  set('word/header2.xml', `<w:hdr ${W}><w:p>${run('First page header')}</w:p></w:hdr>`);
  set(
    'word/footer1.xml',
    `<w:ftr ${W}><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${run('Page ')}${field('PAGE', '1')}` +
      (flavour === 'whole' ? `${run(' of ')}${field('NUMPAGES', '99')}` : '') +
      '</w:p></w:ftr>'
  );
  for (const [kind, notes, id] of [
    ['footnote', footnotes, 'rIdFn'],
    ['endnote', endnotes, 'rIdEn'],
  ] as const) {
    if (notes.length === 0) continue;
    const items = notes
      .map((note) =>
        `<w:${kind} w:id="${note}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${kind}Ref/></w:r>` +
        `${run(' ' + text(8, 30))}</w:p></w:${kind}>`
      )
      .join('');
    set(
      `word/${kind}s.xml`,
      `<w:${kind}s ${W}><w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
        `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
        `${items}</w:${kind}s>`
    );
    rels.push([id, `${kind}s`, `${kind}s.xml`]);
    overrides.push([`/word/${kind}s.xml`, `${kind}s`]);
  }
  if (flavour === 'floats') {
    parts.set('word/media/f.png', PNG);
    rels.push(['rIdImg', 'image', 'media/f.png']);
  }
  const final = flavour === 'final-sect' ? section(true, true, 5) : section();
  set(
    'word/document.xml',
    `<w:document ${W} ${R} ${DRAW}><w:body>${body.join('')}${final}</w:body></w:document>`
  );
  const contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  set(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
      `<Override PartName="/word/document.xml" ContentType="${contentType}.document.main+xml"/>` +
      overrides.map(([part, kind]) =>
        `<Override PartName="${part}" ContentType="${contentType}.${kind}+xml"/>`
      ).join('') +
      '</Types>'
  );
  set(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
  );
  set(
    'word/_rels/document.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels.map(([id, kind, target]) =>
        `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${target}"/>`
      ).join('') + '</Relationships>'
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
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
