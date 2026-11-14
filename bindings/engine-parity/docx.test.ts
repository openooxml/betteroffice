import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import initWasm, * as wasmEdit from '../../packages/docx/src/wasm/generated/edit/docx_edit.js';
import initPreflight, { font_preflight as wasmFontPreflight } from '../../packages/docx/src/wasm/generated/preflight/docx_preflight.js';
import * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeYrs from '@betteroffice/docx-native/yrs';
import { readTableFragment, type ClipboardTableReader } from '../../packages/docx/src/yrs/clipboardFragment';
import { decodeEncodedSelection } from '../../packages/docx/src/yrs/encodedSelection';
import type { YrsStorySegment } from '../../packages/docx/src/yrs/index';
import { deobfuscateFont } from '../../packages/docx/src/utils/fontDeobfuscation';
import { FORMATS, parseSurface } from '../../scripts/check-engine-surface';
import {
  coverageProblems,
  defineRow,
  FramedZipReturn,
  oneEngineExports,
  oneEngineProblems,
  printParityCounts,
  rowName,
  runKnownDifference,
  runRow,
  surfaceMethods,
  zipEntries,
  ZipReturn,
  type EngineHandle,
  type Exclusion,
  type KnownDifference,
  type OneEngineOnly,
  type Row,
} from './harness';

type DocxEngine = typeof wasmEdit & { font_preflight: typeof wasmFontPreflight };
type Session = wasmEdit.EditSession;
type Yrs = typeof wasmYrs;

interface Docx {
  engine: DocxEngine;
  session: Session;
  fresh(): Session;
  /** Saves `replica` instead, through a host opened unseeded on `source`. */
  saveFrom(replica: Session, source?: Uint8Array): void;
  saved(replica: Session, source: Uint8Array): Promise<Uint8Array>;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const ROOT = `xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${REL}"`;
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64'));
const DATE = '2026-01-01T00:00:00Z';
const GENERATION = 'engine-parity';
const HOST_CLIENT = 7;
const EDITOR_CLIENT = 8;
const FRESH_CLIENT = 9;
const EXPORT = JSON.stringify({ revisionView: 'accepted' });
const TAG_QUERY = JSON.stringify({ kind: 'tag', tag: 'name' });

const IMAGE =
  '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="95250" cy="95250"/><wp:docPr id="7" name="Picture 7"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="7" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="95250" cy="95250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
  '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
const TEXT_CONTROL = '<w:sdt><w:sdtPr><w:tag w:val="name"/><w:id w:val="8"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Old</w:t></w:r></w:sdtContent></w:sdt>';
const CHECKBOX = '<w:sdt><w:sdtPr><w:tag w:val="agree"/><w:id w:val="9"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>';
const cell = (paraId: string, text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr><w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const TABLE = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid>' +
  `<w:tr>${cell('50000001', 'A')}${cell('50000002', 'B')}</w:tr><w:tr>${cell('50000003', 'C')}${cell('50000004', 'D')}</w:tr></w:tbl>`;
const BODY = [
  '<w:p w14:paraId="10000001"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Parity heading</w:t></w:r></w:p>',
  '<w:p w14:paraId="10000002"><w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">First body paragraph with alpha text.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>',
  '<w:p w14:paraId="10000003"><w:r><w:t xml:space="preserve">Second body paragraph, alpha again.</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>',
  `<w:p w14:paraId="10000004"><w:r><w:t xml:space="preserve">Controls: </w:t></w:r>${TEXT_CONTROL}<w:r><w:t xml:space="preserve"> </w:t></w:r>${CHECKBOX}</w:p>`,
  '<w:p w14:paraId="10000005"><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>',
  `<w:p w14:paraId="10000006"><w:r><w:t xml:space="preserve">Picture </w:t></w:r><w:r>${IMAGE}</w:r></w:p>`,
  TABLE,
  '<w:p w14:paraId="10000007"><w:r><w:t>Final paragraph.</w:t></w:r></w:p>',
  '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
].join('');

async function buildFixture(overrides: Record<string, string | Uint8Array> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  const date = new Date('2024-01-01T00:00:00Z');
  const parts: Record<string, string | Uint8Array> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
      `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/><Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/>` +
      `<Override PartName="/word/header1.xml" ContentType="${OFFICE}.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="${OFFICE}.footnotes+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rIdComments" Type="${REL}/comments" Target="comments.xml"/>` +
      `<Relationship Id="rIdHeader" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdFootnotes" Type="${REL}/footnotes" Target="footnotes.xml"/>` +
      `<Relationship Id="rIdImage" Type="${REL}/image" Target="media/image1.png"/></Relationships>`,
    'word/styles.xml': `<w:styles ${ROOT}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style></w:styles>',
    'word/comments.xml': `<w:comments ${ROOT}><w:comment w:id="0" w:author="Parity" w:date="${DATE}" w:initials="P"><w:p w14:paraId="20000001"><w:r><w:t>Seeded comment</w:t></w:r></w:p></w:comment></w:comments>`,
    'word/header1.xml': `<w:hdr ${ROOT}><w:p w14:paraId="30000001"><w:r><w:t>Parity header</w:t></w:r></w:p></w:hdr>`,
    'word/footnotes.xml': `<w:footnotes ${ROOT}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
      '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="1"><w:p w14:paraId="40000001"><w:r><w:t>Parity footnote</w:t></w:r></w:p></w:footnote></w:footnotes>',
    'word/document.xml': `<w:document ${ROOT}><w:body>${BODY}</w:body></w:document>`,
    'word/media/image1.png': PNG,
    ...overrides,
  };
  for (const [name, content] of Object.entries(parts)) zip.file(name, content, { date });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const fontBytes = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
const wasmEngine: DocxEngine = { ...wasmEdit, font_preflight: wasmFontPreflight };
const nativeEngine = createRequire(import.meta.url)('../docx-native/index.js') as DocxEngine;
let fixture: Uint8Array;
let fixtureEntries: Record<string, Uint8Array>;
let twinCommentFixture: Uint8Array;
let compareOriginal: Uint8Array;
let compareRevised: Uint8Array;
let repeatingFixture: Uint8Array;
let effectFixture: Uint8Array;
let embeddedFontFixture: Uint8Array;
let pageTotalFixture: Uint8Array;

const SECOND_RUN = '<w:r><w:t xml:space="preserve">Second body paragraph, alpha again.</w:t></w:r>';

function patched(name: string, from: string, to: string): string {
  const text = new TextDecoder().decode(fixtureEntries[name]);
  if (!text.includes(from)) throw new Error(`${name} has no ${from}`);
  return text.replace(from, to);
}

// Two comments whose paragraphs share an ID that commentsExtended.xml references.
function twinCommentParts(): Record<string, string> {
  const twin = `<w:comment w:id="1" w:author="Parity" w:date="${DATE}" w:initials="P"><w:p w14:paraId="20000001"><w:r><w:t>Twin comment</w:t></w:r></w:p></w:comment>`;
  return {
    '[Content_Types].xml': patched('[Content_Types].xml', '</Types>', `<Override PartName="/word/commentsExtended.xml" ContentType="${OFFICE}.commentsExtended+xml"/></Types>`),
    'word/_rels/document.xml.rels': patched('word/_rels/document.xml.rels', '</Relationships>',
      '<Relationship Id="rIdCommentsEx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>'),
    'word/comments.xml': patched('word/comments.xml', '</w:comments>', `${twin}</w:comments>`),
    'word/commentsExtended.xml': '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="20000001" w15:done="0"/></w15:commentsEx>',
    'word/document.xml': patched('word/document.xml', SECOND_RUN, `<w:commentRangeStart w:id="1"/>${SECOND_RUN}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>`),
  };
}

const SEEDED_FOOTNOTE = '<w:footnote w:id="1"><w:p w14:paraId="40000001"><w:r><w:t>Parity footnote</w:t></w:r></w:p></w:footnote>';

function compareParts(text: string): Record<string, string> {
  return {
    'word/document.xml': `<w:document ${ROOT}><w:body><w:p w14:paraId="60000001"><w:r><w:t>Compared heading</w:t></w:r></w:p>` +
      `<w:p w14:paraId="60000002"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
    'word/comments.xml': `<w:comments ${ROOT}/>`,
    'word/footnotes.xml': patched('word/footnotes.xml', SEEDED_FOOTNOTE, ''),
  };
}

const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const FINAL_PARAGRAPH = '<w:p w14:paraId="10000007">';
const REPEATING_SECTION = `<w:sdt><w:sdtPr><w:id w:val="11"/><w15:repeatingSection ${W15}/></w:sdtPr><w:sdtContent>` +
  `<w:sdt><w:sdtPr><w:id w:val="12"/><w15:repeatingSectionItem ${W15}/></w:sdtPr><w:sdtContent>` +
  '<w:p w14:paraId="10000008"><w:r><w:t>Repeated item</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>';

function repeatingParts(): Record<string, string> {
  return { 'word/document.xml': patched('word/document.xml', FINAL_PARAGRAPH, `${REPEATING_SECTION}${FINAL_PARAGRAPH}`) };
}

const FINAL_RUN = '<w:r><w:t>Final paragraph.</w:t></w:r>';

// A legacy animated text effect, which the canvas backend never draws.
function effectParts(): Record<string, string> {
  return { 'word/document.xml': patched('word/document.xml', FINAL_RUN, '<w:r><w:rPr><w:effect w:val="shimmer"/></w:rPr><w:t>Final paragraph.</w:t></w:r>') };
}

const PAGE_FIELD = '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>';

function pageTotalParts(): Record<string, string> {
  return { 'word/document.xml': patched('word/document.xml', PAGE_FIELD, `${PAGE_FIELD}<w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple>`) };
}

const EMBEDDED_FAMILY = 'Embedded Sans';
const FONT_KEY = '{00112233-4455-6677-8899-AABBCCDDEEFF}';
const NORMAL_NAME = '<w:name w:val="Normal"/></w:style>';

// Normal text set in a font the package embeds as an obfuscated face, the way Word writes word/fonts/*.odttf.
function embeddedFontParts(): Record<string, string | Uint8Array> {
  return {
    '[Content_Types].xml': patched('[Content_Types].xml', '</Types>',
      `<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/><Override PartName="/word/fontTable.xml" ContentType="${OFFICE}.fontTable+xml"/></Types>`),
    'word/_rels/document.xml.rels': patched('word/_rels/document.xml.rels', '</Relationships>', `<Relationship Id="rIdFontTable" Type="${REL}/fontTable" Target="fontTable.xml"/></Relationships>`),
    'word/fontTable.xml': `<w:fonts ${ROOT}><w:font w:name="${EMBEDDED_FAMILY}"><w:embedRegular r:id="rIdFace" w:fontKey="${FONT_KEY}"/></w:font></w:fonts>`,
    'word/_rels/fontTable.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFace" Type="${REL}/font" Target="fonts/font1.odttf"/></Relationships>`,
    'word/fonts/font1.odttf': new Uint8Array(deobfuscateFont(fontBytes.slice().buffer, FONT_KEY)),
    'word/styles.xml': patched('word/styles.xml', NORMAL_NAME, `<w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="${EMBEDDED_FAMILY}" w:hAnsi="${EMBEDDED_FAMILY}"/></w:rPr></w:style>`),
  };
}

beforeAll(async () => {
  await initWasm({ module_or_path: readFileSync(new URL('../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)) });
  await initPreflight({ module_or_path: readFileSync(new URL('../../packages/docx/src/wasm/generated/preflight/docx_preflight_bg.wasm', import.meta.url)) });
  fixture = await buildFixture();
  fixtureEntries = Object.fromEntries((await zipEntries(fixture)).parts);
  twinCommentFixture = await buildFixture(twinCommentParts());
  compareOriginal = await buildFixture(compareParts('The quick fox jumps.'));
  compareRevised = await buildFixture(compareParts('The quick brown fox jumps.'));
  repeatingFixture = await buildFixture(repeatingParts());
  effectFixture = await buildFixture(effectParts());
  embeddedFontFixture = await buildFixture(embeddedFontParts());
  pageTotalFixture = await buildFixture(pageTotalParts());
});

// version() is "{nonce:016x}-{epoch}"; mint_nonce in crates/docx-edit/src/batch.rs mixes host entropy and the clock.
function versionNonce(session: Session): string {
  const version = session.version();
  return version.slice(0, version.lastIndexOf('-'));
}

function engineFor(engine: DocxEngine, yrs: Yrs) {
  return async (): Promise<EngineHandle<Docx>> => {
    const host = await yrs.createYrsSession({ clientId: HOST_CLIENT });
    const sessions: Session[] = [];
    const open = (clientId: number) => {
      const session = new engine.EditSession(clientId);
      sessions.push(session);
      return session;
    };
    const free = () => {
      for (const session of sessions) session.free();
      host.destroy();
    };
    let replicaSave: { replica: Session; source: Uint8Array } | null = null;
    const saved = async (replica: Session, source: Uint8Array) => {
      const replicaHost = await yrs.createYrsSession({ clientId: HOST_CLIENT });
      try {
        replicaHost.openDocx(source, false, { generation: GENERATION });
        replicaHost.applyUpdate(replica.encode_state());
        return (await yrs.saveYrsDocx(replicaHost, { updateModifiedDate: false })).bytes;
      } finally {
        replicaHost.destroy();
      }
    };
    try {
      host.seedFromDocx(fixture, { generation: GENERATION });
      const session = open(EDITOR_CLIENT);
      session.set_media_tokens(false);
      session.open_docx(fixture, false, GENERATION);
      session.apply_update(host.encodeState());
      return {
        target: {
          engine,
          session,
          fresh: () => open(FRESH_CLIENT),
          saveFrom: (replica, source = fixture) => {
            replicaSave = { replica, source };
          },
          saved,
        },
        generated: () => sessions.map(versionNonce),
        save: async () => {
          if (replicaSave) return saved(replicaSave.replica, replicaSave.source);
          host.applyUpdate(session.encode_state());
          return (await yrs.saveYrsDocx(host, { updateModifiedDate: false })).bytes;
        },
        free,
      };
    } catch (error) {
      free();
      throw error;
    }
  };
}

const makeWasm = engineFor(wasmEngine, wasmYrs);
const facadeShape = (yrs: object) => Object.entries(yrs).map(([name, value]) => `${name}:${typeof value}`).sort();
const makeNative = engineFor(nativeEngine, nativeYrs as unknown as Yrs);

function paraId(session: Session, prefix: string): string {
  const found = (JSON.parse(session.paragraphs('body')) as { paraId: string; text: string }[]).find((paragraph) => paragraph.text.startsWith(prefix));
  if (!found) throw new Error(`no body paragraph starts with ${prefix}`);
  return found.paraId;
}

function span(session: Session, prefix: string, from: number, to: number): [string, string, number, string, number] {
  const id = paraId(session, prefix);
  return ['body', id, from, id, to];
}

function trackedInsertion(session: Session): string {
  return (JSON.parse(session.insert_text('body', paraId(session, 'First body'), 0, 'Tracked ', 'Parity', DATE)) as { revisionId: string }).revisionId;
}

function editRequest(session: Session): string {
  return JSON.stringify({
    expectVersion: session.version(),
    steps: [{ op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: paraId(session, 'Second body') }, at: 'end', text: ' added' }],
  });
}

function tableFragment(session: Session) {
  const reader = {
    storySegments: (story: string) => JSON.parse(session.story_segments(story)) as YrsStorySegment[],
    tablePayload: (story: string, tableIndex: number) => {
      const payload = session.table_payload(story, tableIndex);
      return payload === undefined ? null : JSON.parse(payload) as Record<string, unknown>;
    },
  } as ClipboardTableReader;
  return JSON.stringify(readTableFragment(reader, { story: 'body', tableIndex: 0 }));
}

// The facade's ensureUndo: track, scope history to the story, then edit.
function undoableEdit(session: Session): void {
  session.track_undo();
  session.select_story('body');
  session.insert_text('body', paraId(session, 'Second body'), 0, 'Undone ');
}

function addedComment(session: Session): string {
  const id = paraId(session, 'Second body');
  const ranges = [{ story: 'body', startPara: id, startOffset: 0, endPara: id, endOffset: 6 }];
  return (JSON.parse(session.add_comment(JSON.stringify(ranges), 'Parity', DATE, '"Anchored"')) as { commentId: string }).commentId;
}

const TABLE_LOC = JSON.stringify({ story: 'body', tableIndex: 0 });
const cellAt = (row: number, column: number) => ({ story: 'body', tableIndex: 0, row, column });
const cellLoc = (row: number, column: number) => JSON.stringify(cellAt(row, column));
const cellRange = (anchor: [number, number], head: [number, number]) => JSON.stringify({ anchor: cellAt(...anchor), head: cellAt(...head) });

function seededImage(session: Session): string {
  const image = (JSON.parse(session.story_segments('body')) as YrsStorySegment[])
    .find((segment) => segment.kind === 'embed' && segment.embedKind === 'image');
  if (image?.kind !== 'embed') throw new Error('the fixture image is not seeded');
  return JSON.stringify(image.payload);
}

function layoutRequest(session: Session) {
  const font = session.register_measure_font(fontBytes);
  return { font, request: regionRequest() };
}

function regionRequest() {
  return {
    bodyStory: 'body',
    regions: { sections: [{
      sectionId: 'main',
      properties: { pageWidth: 12240, pageHeight: 15840, marginTop: 1440, marginRight: 1440, marginBottom: 1440, marginLeft: 1440 },
    }] },
    notes: { contents: [{ id: 1, noteKind: 'footnote', height: 0 }] },
    measurement: { fontChains: {} as Record<string, number[]>, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true },
    renderEnv: {},
  };
}

function layoutInput(session: Session, request: { measurement: { fontChains: Record<string, number[]> } } = regionRequest()) {
  const font = session.register_measure_font(fontBytes);
  const fontChains = request.measurement.fontChains;
  for (const { key } of JSON.parse(session.layout_font_requirements_json(JSON.stringify(request))) as { key: string }[]) fontChains[key] = [font];
  return { input: JSON.stringify(request), fontChains };
}

function layout(session: Session) {
  const { input, fontChains } = layoutInput(session);
  return { layout: session.layout_document_with_regions_json(input), fontChains };
}

function displayList(session: Session): string {
  const { layout: json, fontChains } = layout(session);
  return session.build_display_list_json(JSON.stringify({ ...JSON.parse(json), fontChains }));
}

interface Unit { paraId: string; offset: number; segment: YrsStorySegment }

// Paragraph-local offsets as the session counts them, leading block embeds included.
function units(session: Session): Unit[] {
  const found: Unit[] = [];
  let pending: Omit<Unit, 'paraId'>[] = [];
  let offset = 0;
  for (const segment of JSON.parse(session.story_segments('body')) as YrsStorySegment[]) {
    if (segment.kind === 'pilcrow') {
      found.push(...pending.map((unit) => ({ ...unit, paraId: segment.paraId })));
      pending = [];
      offset = 0;
    } else {
      pending.push({ offset, segment });
      offset += segment.kind === 'text' ? segment.text.length : 1;
    }
  }
  return found;
}

function textAt(session: Session, prefix: string, index: number): number {
  const id = paraId(session, prefix);
  const first = units(session).find((unit) => unit.paraId === id && unit.segment.kind === 'text');
  if (!first) throw new Error(`no text in the paragraph starting with ${prefix}`);
  return first.offset + index;
}

function textSpan(session: Session, prefix: string, from: number, to: number): [string, string, number, string, number] {
  const id = paraId(session, prefix);
  return ['body', id, textAt(session, prefix, from), id, textAt(session, prefix, to)];
}

function imageAt(session: Session): { paraId: string; offset: number; src: unknown } {
  const image = units(session).find(({ segment }) => segment.kind === 'embed' && segment.embedKind === 'image');
  if (image?.segment.kind !== 'embed') throw new Error('the fixture image is not seeded');
  return { paraId: image.paraId, offset: image.offset, src: image.segment.payload.src };
}

const frameExtras = (fontChains: Record<string, number[]>) => JSON.stringify({ fontChains });

function frameEpoch(session: Session): number {
  return (JSON.parse(session.resident_caret_snapshot_json()) as { frameEpoch: number }).frameEpoch;
}

function residentFrame(session: Session): number {
  session.build_display_list_frame(frameExtras(layout(session).fontChains), 0);
  return frameEpoch(session);
}

function caretIn(session: Session, prefix: string, index: number): void {
  const id = paraId(session, prefix);
  const at = textAt(session, prefix, index);
  session.set_selection('body', id, at, id, at);
}

function workerLayoutInput(session: Session) {
  const { input, fontChains } = layoutInput(session);
  const request = JSON.parse(input) as { renderEnv: unknown; measurement: unknown };
  const workerOwned = { bodyStory: 'body', renderEnv: request.renderEnv, measurement: request.measurement };
  return { input: JSON.stringify({ ...request, measured: [], retainOnly: true, workerOwned }), fontChains };
}

function peerEdit(peer: Session, prefix: string, text: string): void {
  peer.insert_text('body', paraId(peer, prefix), textAt(peer, prefix, 0), text);
}

function remoteUpdate({ session, fresh }: Docx): Uint8Array {
  const peer = fresh();
  peer.apply_update(session.encode_state());
  peerEdit(peer, 'Second body', 'Remote ');
  return peer.encode_diff(session.encode_state_vector());
}

function remoteChunks({ session, fresh }: Docx): Uint8Array[] {
  const peer = fresh();
  peer.apply_update(session.encode_state());
  peerEdit(peer, 'Second body', 'Remote ');
  const first = peer.encode_diff(session.encode_state_vector());
  const middle = peer.encode_state_vector();
  peerEdit(peer, 'Final', 'Later ');
  return [first, peer.encode_diff(middle)];
}

function trackedRemote<U>(docx: Docx, build: (docx: Docx) => U): U {
  docx.session.track_undo();
  docx.session.select_story('body');
  return build(docx);
}

function notUndoable(session: Session, apply: () => unknown) {
  const result = apply();
  const applied = session.paragraphs('body');
  const canUndo = session.can_undo();
  const undone = session.undo();
  if (canUndo || undone || session.paragraphs('body') !== applied) throw new Error('a remote update became undoable');
  return { result, canUndo, undone, after: applied };
}

function timedEdits(session: Session, gap: () => unknown) {
  const id = paraId(session, 'Second body');
  const at = textAt(session, 'Second body', 0);
  session.track_undo();
  session.select_story('body');
  session.set_undo_input_time(5_000, 5_000);
  session.insert_text('body', id, at, 'a');
  const result = gap();
  session.insert_text('body', id, at + 1, 'b');
  return { result, undone: session.undo(), after: session.paragraphs('body') };
}

function typedUndoSteps(session: Session, gap: number) {
  session.track_undo();
  session.select_story('body');
  const epoch = residentFrame(session);
  caretIn(session, 'Second body', 6);
  session.set_undo_input_time(5_000, 5_000);
  const first = session.apply_input('a', epoch);
  session.set_undo_input_time(5_000 + gap, 5_000 + gap);
  const second = session.apply_input('b', frameEpoch(session));
  const typed = session.paragraphs('body');
  let steps = 0;
  while (session.can_undo() && steps < 10) {
    session.undo();
    steps += 1;
  }
  return { first, second, typed, steps, after: session.paragraphs('body') };
}

// A seeded open is what marks a replica seeded, which priming requires.
function seededReplica({ fresh, saveFrom }: Docx): Session {
  const f = fresh();
  f.seed_from_docx(fixture, GENERATION);
  saveFrom(f);
  return f;
}

function replicaReads(replica: Session) {
  return { segments: replica.story_segments('body'), stories: replica.list_stories_json(), layout: layout(replica).layout };
}

function refusedThenKept(session: Session, refuse: () => unknown, rebuild: (epoch: number) => Uint8Array) {
  const kept = frameEpoch(session);
  let refusal: unknown;
  try {
    refuse();
  } catch (error) {
    refusal = error;
  }
  if (refusal === undefined) throw new Error('the frame call was not refused');
  return { refusal: String(refusal), kept, after: frameEpoch(session), again: rebuild(kept) };
}

function primeSlices(session: Session) {
  const before = session.prime_last_slice_cost();
  const slices: [number, number][] = [];
  let answer = 1;
  for (let slice = 0; slice < 10_000 && answer === 1; slice += 1) {
    answer = session.prime_save_seed_states_slice(1_000, 64 * 1024 * 1024, 64);
    slices.push([answer, session.prime_last_slice_cost()]);
  }
  if (!slices.some(([, cost]) => cost > 0)) throw new Error('no prime slice charged a cost');
  return { before, slices };
}

const HYPERLINK = JSON.stringify({ href: 'https://example.com/', tooltip: null, rId: null });
const GEOMETRY = JSON.stringify({ widthEmu: 190500, heightEmu: 95250, other: { width: 20, height: 10 } });
const MODULE_OPTIONS = { pageSize: { w: 816, h: 1056 }, margins: { top: 96, right: 96, bottom: 96, left: 96 }, pageGap: 20 };
const NOT_DOCX = Uint8Array.from([1, 2, 3]);
const MALFORMED_UPDATE = Uint8Array.from([255]);

function saveRequest(engine: DocxEngine): string {
  const envelope = JSON.parse(engine.parse_docx_s9(fixture, '')) as {
    document: { package: { document: { content: unknown; finalSectionProperties?: unknown } } };
  };
  const { content, finalSectionProperties } = envelope.document.package.document;
  return JSON.stringify({
    determinism: { seed: '0'.repeat(64), now: '2000-01-01T00:00:00.000Z' },
    document: { content, finalSectionProperties },
    options: { updateModifiedDate: false },
  });
}

const DETERMINISM = { seed: '0'.repeat(64), now: '2000-01-01T00:00:00.000Z' };

interface WireBlock { type: string; content?: { type: string }[] }

interface WirePackage {
  document: { content: WireBlock[]; finalSectionProperties?: unknown; comments?: unknown[] };
  headerEntries?: [string, unknown][];
  footnotes?: unknown[];
}

function wirePackage(engine: DocxEngine) {
  const { document, headerEntries, footnotes } = (JSON.parse(engine.parse_docx_s9(fixture, '')) as { document: { package: WirePackage } }).document.package;
  const paragraphs = document.content.filter((block) => block.type === 'paragraph');
  const table = document.content.find((block) => block.type === 'table');
  const run = paragraphs[0]?.content?.find((item) => item.type === 'run');
  const header = headerEntries?.[0]?.[1];
  if (!paragraphs[1] || !table || !run || !header || !footnotes?.length || !document.comments?.length || !document.finalSectionProperties) {
    throw new Error('the fixture wire lacks a paragraph, run, table, header, footnote, comment or final section');
  }
  return { document, paragraph: paragraphs[1], table, run, header, footnotes, comments: document.comments };
}

function serialized(serialize: (request: string) => string, requests: Record<string, unknown>[]): string[] {
  return requests.map((request) => serialize(JSON.stringify({ ...request, determinism: DETERMINISM })));
}

function structuredContent(engine: DocxEngine): string {
  const json = engine.export_docx_structured_json(fixture, EXPORT);
  const outcome = JSON.parse(json) as { ok: boolean; content?: unknown };
  if (!outcome.ok) throw new Error(`the structured export refused: ${json}`);
  return JSON.stringify(outcome.content);
}

function laidOut(session: Session): string {
  const { input } = layoutInput(session);
  session.layout_document_with_regions_json(input);
  return input;
}

function pagedExport(session: Session, request: string): string {
  const json = session.export_structured_with_pages_json(EXPORT, request);
  if (!(JSON.parse(json) as { ok: boolean }).ok) throw new Error(`the paged export refused: ${json}`);
  return json;
}

function pageTexts(session: Session): string {
  const json = session.page_texts_json(Uint32Array.from([0]));
  const pages = JSON.parse(json) as { paragraphs: unknown[] }[];
  if (!pages[0]?.paragraphs.length) throw new Error(`the first page shows no text: ${json}`);
  return json;
}

function savedPairs(session: Session): [string, string][] {
  const { paragraphs } = JSON.parse(session.paragraph_identities()) as {
    paragraphs: { session: { paraId: string } | null; ooxmlParaId: string | null }[];
  };
  const held = paragraphs.flatMap(({ session: anchor, ooxmlParaId }): [string, string][] =>
    anchor && ooxmlParaId ? [[anchor.paraId, ooxmlParaId]] : []);
  const [first, second] = held;
  if (!first || !second) throw new Error('fewer than two session paragraphs carry a Word paragraph ID');
  return [...held, [first[0], second[1]]];
}

const COMPARE_OPTIONS = JSON.stringify({ author: 'Parity', date: DATE });

function comparedInto(session: Session): string {
  const json = session.compare_docx_json(compareOriginal, compareRevised, COMPARE_OPTIONS);
  const outcome = JSON.parse(json) as { ok: boolean; noop?: boolean };
  if (!outcome.ok || outcome.noop !== false) throw new Error(`the comparison left nothing to save: ${json}`);
  return json;
}

interface ControlTarget { control: { ooxmlId: string | null; controlType: string }; at: { story: string; index: number } }

function controlTarget(session: Session, matches: (control: ControlTarget['control']) => boolean): ControlTarget['at'] {
  const target = (JSON.parse(session.content_control_targets_json()) as ControlTarget[]).find(({ control }) => matches(control));
  if (!target) throw new Error('no content control target matches');
  return target.at;
}

function inlineControlAt(session: Session, ordinal: number): { paraId: string; offset: number } {
  const id = paraId(session, 'Controls');
  const control = units(session).filter((unit) => unit.paraId === id && unit.segment.kind === 'embed' && unit.segment.embedKind === 'sdt')[ordinal];
  if (!control) throw new Error(`the controls paragraph has no inline control ${ordinal}`);
  return { paraId: control.paraId, offset: control.offset };
}

const WATERMARK = JSON.stringify({ kind: 'text', text: 'DRAFT', font: 'Liberation Sans', color: 'C0C0C0', semitransparent: true, layout: 'diagonal' });

const watermarkText = (watermark: string) => (JSON.parse(watermark) as { text: string }).text;

function watermarkUnits(session: Session): Unit[] {
  return units(session).filter(({ segment }) => segment.kind === 'embed' && segment.embedKind === 'watermark');
}

function sessionWatermarks(session: Session, watermark: string): Unit[] {
  const found = watermarkUnits(session);
  const text = watermarkText(watermark);
  if (!JSON.stringify(found).includes(text)) throw new Error(`the session has no watermark embed carrying ${text}`);
  return found;
}

async function savedBody(bytes: Uint8Array): Promise<string> {
  const part = new Map((await zipEntries(bytes)).parts).get('word/document.xml');
  if (!part) throw new Error('the saved package has no word/document.xml');
  return new TextDecoder().decode(part);
}

function drainEvents(session: Session): Uint8Array[] {
  const events: Uint8Array[] = [];
  for (let event = session.drain_update_event(); event.length > 0 && events.length < 100; event = session.drain_update_event()) events.push(event);
  return events;
}

// The later chunk continues the text the first inserted, so it depends on it: yrs integrates an
// out-of-order chunk with no missing dependency at once behind a skip, leaving nothing pending.
function dependentChunks({ session, fresh }: Docx): Uint8Array[] {
  const peer = fresh();
  peer.apply_update(session.encode_state());
  peerEdit(peer, 'Second body', 'Remote ');
  const first = peer.encode_diff(session.encode_state_vector());
  const middle = peer.encode_state_vector();
  peer.insert_text('body', paraId(peer, 'Remote Second body'), textAt(peer, 'Remote Second body', 'Remote '.length), 'later ');
  return [first, peer.encode_diff(middle)];
}

// Observes a local edit, then applies a remote update whose predecessor has not arrived; returns the predecessor.
function pendingUpdate(docx: Docx): Uint8Array {
  const { session: s } = docx;
  s.start_update_event_observation();
  const [first, later] = dependentChunks(docx);
  s.insert_text('body', paraId(s, 'First body'), textAt(s, 'First body', 0), 'Local ');
  s.apply_update(later!);
  return first!;
}

function alphaTarget(session: Session): string {
  return JSON.stringify({ kind: 'search', text: 'alpha', within: { kind: 'paragraph', story: 'body', paraId: paraId(session, 'Second body') }, view: 'accepted' });
}

interface ListedComment { id: unknown; author: unknown; done: unknown }

function listedComment(session: Session, id: string): ListedComment | undefined {
  return (JSON.parse(session.list_comments()) as ListedComment[]).find((comment) => String(comment.id) === id);
}

function seededCommentId(session: Session): string {
  const [comment] = JSON.parse(session.list_comments()) as ListedComment[];
  if (comment?.author !== 'Parity') throw new Error('the fixture comment is not seeded');
  return String(comment.id);
}

function exchange(left: Session, right: Session): void {
  const fromLeft = left.encode_state();
  const fromRight = right.encode_state();
  right.apply_update(fromLeft);
  left.apply_update(fromRight);
}

// Two collaborators each change the table and undo in crossing order, which deletes every payload entry.
function crossedTableUndos({ session: s, fresh }: Docx): void {
  const f = fresh();
  f.apply_update(s.encode_state());
  s.track_undo();
  f.track_undo();
  s.insert_row(cellLoc(0, 0), true);
  exchange(s, f);
  f.set_column_width(cellLoc(0, 0), 2000);
  if (!s.undo()) throw new Error('the inserted row was not undone');
  exchange(s, f);
  if (!f.undo()) throw new Error('the column width was not undone');
  exchange(s, f);
}

const GRID_STYLE = {
  type: 'table',
  styleId: 'Grid',
  pPr: { spaceAfter: 0, lineSpacing: 240 },
  tblPr: { borders: { top: { style: 'single', size: 4 } }, cellMargins: { left: { value: 108, type: 'dxa' } } },
  tblStylePr: [{ type: 'firstRow', pPr: { spaceBefore: 40 } }],
};

interface WireTable {
  formatting?: Record<string, unknown>;
  columnWidths?: number[];
  rows: { formatting?: unknown; cells: { formatting?: unknown }[] }[];
}

// The fixture's styles plus a table style, and the fixture table set in that style, as the facade passes them.
function gridStyled(engine: DocxEngine): { styles: string; table: string } {
  const { styles, document } = (JSON.parse(engine.parse_docx_s9(fixture, '')) as {
    document: { package: { styles?: { styles?: unknown[] }; document: { content: { type: string }[] } } };
  }).document.package;
  const table = document.content.find((block) => block.type === 'table') as WireTable | undefined;
  if (!styles?.styles?.length || !table) throw new Error('the fixture wire lacks styles or a table');
  return {
    styles: JSON.stringify({ ...styles, styles: [...styles.styles, GRID_STYLE] }),
    table: JSON.stringify({
      formatting: { ...table.formatting, styleId: 'Grid', look: { firstRow: true } },
      columnWidths: table.columnWidths,
      rows: table.rows.map((row) => ({ formatting: row.formatting, cells: row.cells.map((cell) => ({ formatting: cell.formatting })) })),
    }),
  };
}

interface CellInheritance { attributes: { borders?: unknown }; paragraphFormatting: unknown; runFormatting: unknown }

function gridCell(tableAttributes: string): CellInheritance {
  const cell = (JSON.parse(tableAttributes) as { rows: CellInheritance[][] }).rows[0]?.[0];
  if (!cell?.attributes.borders || !cell.paragraphFormatting) throw new Error(`the table style reached no cell: ${tableAttributes}`);
  return cell;
}

function gridInheritance({ session: s, engine }: Docx) {
  const { styles, table } = gridStyled(engine);
  const inheritance = s.run_inheritance(styles);
  return { inheritance, table, cell: gridCell(inheritance.table_attributes(table)) };
}

function tokenizedImage(fresh: () => Session): { f: Session; src: string } {
  const f = fresh();
  f.set_media_tokens(true);
  f.open_docx(fixture, true, GENERATION);
  const { src } = imageAt(f);
  if (typeof src !== 'string' || !src.startsWith('media:')) throw new Error('the seeded image carries no media token');
  return { f, src };
}

const HEADER_ALIASES = JSON.stringify([{ isHeader: true, partPath: 'word/header1.xml', relationshipIds: ['rIdHeader', 'rIdHeaderAlias'] }]);
const HEX_DIGEST = /^[0-9a-f]{16}$/;

interface MediaSourcesWire { key: string[]; parts: [string, number][] }

function mediaSourcesShape({ key, parts, ...rest }: MediaSourcesWire) {
  return {
    ...rest,
    key: key.map((half) => HEX_DIGEST.test(half)),
    parts: parts.map(([digest, ...entry]) => [HEX_DIGEST.test(digest), ...entry]),
  };
}

const engineName = (engine: DocxEngine) => (engine === wasmEngine ? 'wasm' : 'native');

function imageSources(session: Session, input: string): string[] {
  const json = session.layout_document_with_regions_json(input);
  const sources = [...json.matchAll(/"src":"((?:data:image|media:)[^"]*)"/g)].map(([, src]) => src!);
  if (sources.length === 0) throw new Error('the layout carries no image source');
  return sources;
}

function tokenSources(session: Session): string[] {
  const { input } = layoutInput(session);
  const sources = imageSources(session, JSON.stringify({ ...JSON.parse(input), renderEnv: { mediaTokens: true } }));
  if (!sources.every((src) => src.startsWith('media:'))) throw new Error(`the layout left images untokenized: ${JSON.stringify(sources)}`);
  return sources;
}

interface EmbeddedFontWire {
  embeddedFontParts: { path: string; base64: string }[];
  fontTableRelationshipsXml?: string;
  document: { package: { fontTable?: { fonts: { name: string; embedRegular?: { relId: string; fontKey?: string } }[] } } };
}

// The embedded face as the engine's own parse hands it out, de-obfuscated with its font key.
function embeddedFace(engine: DocxEngine): Uint8Array {
  const { embeddedFontParts, fontTableRelationshipsXml, document } = JSON.parse(engine.parse_docx_s9(embeddedFontFixture, '')) as EmbeddedFontWire;
  const embed = document.package.fontTable?.fonts.find(({ name }) => name === EMBEDDED_FAMILY)?.embedRegular;
  const target = embed && fontTableRelationshipsXml?.match(new RegExp(`Id="${embed.relId}"[^>]*Target="([^"]+)"`))?.[1];
  const part = target && embeddedFontParts.find(({ path }) => path === `word/${target}`);
  if (!embed?.fontKey || !part) throw new Error(`the parse hands out no embedded ${EMBEDDED_FAMILY} face`);
  const face = new Uint8Array(deobfuscateFont(Uint8Array.from(Buffer.from(part.base64, 'base64')).buffer, embed.fontKey));
  if (Buffer.compare(Buffer.from(face), Buffer.from(fontBytes)) !== 0) throw new Error('the embedded face does not de-obfuscate to the font it was built from');
  return face;
}

function privateFontRequest(engine: DocxEngine, session: Session, chain: number[]) {
  const request = regionRequest();
  const { document } = JSON.parse(engine.parse_docx_s9(embeddedFontFixture, '')) as { document: { package: { document: { finalSectionProperties?: typeof request.regions.sections[0]['properties'] } } } };
  const properties = document.package.document.finalSectionProperties;
  if (!properties) throw new Error('the embedded font fixture parses without a final section');
  request.regions.sections[0]!.properties = properties;
  const requirements = JSON.parse(session.layout_font_requirements_json(JSON.stringify(request))) as { key: string; family: string }[];
  if (!requirements.some(({ family }) => family.toLowerCase() === EMBEDDED_FAMILY.toLowerCase())) {
    throw new Error(`the layout does not ask for the embedded font: ${JSON.stringify(requirements)}`);
  }
  for (const { key } of requirements) request.measurement.fontChains[key] = chain;
  return { requirements, request: JSON.stringify(request) };
}

function embeddedFontSession(fresh: () => Session): Session {
  const f = fresh();
  f.open_docx(embeddedFontFixture, true, GENERATION);
  return f;
}

function sliceAnswers(session: Session, paragraphs: number, inflated: number): number[] {
  const answers: number[] = [];
  let answer = 1;
  for (let slice = 0; slice < 10_000 && answer === 1; slice += 1) {
    answer = session.prime_save_seed_states_slice(paragraphs, inflated, 64);
    answers.push(answer);
  }
  return answers;
}

function mediaReplica(fresh: () => Session, seeded: Session): Session {
  const r = fresh();
  r.open_docx(fixture, false, GENERATION);
  r.apply_update(seeded.encode_state());
  return r;
}

const MISSING_STORY = 'parity-missing';

function stickyAt(session: Session, prefix: string, index: number): [string, string, number] {
  return ['body', paraId(session, prefix), textAt(session, prefix, index)];
}

function paragraphEnd(session: Session, prefix: string): [string, string, number] {
  const id = paraId(session, prefix);
  const { start, end } = JSON.parse(session.locate_paragraph('body', id)) as { start: number; end: number };
  return ['body', id, end - start];
}

function encodedSelection(session: Session) {
  const cursor = decodeEncodedSelection(session.encoded_selection());
  if (!cursor) throw new Error('the session has no encoded selection');
  return cursor;
}

function otherEngineReplica(engine: DocxEngine, source: Session): Session {
  const replica = new (engine === wasmEngine ? nativeEngine : wasmEngine).EditSession(FRESH_CLIENT);
  try {
    replica.open_docx(fixture, false, GENERATION);
    replica.apply_update(source.encode_state());
    return replica;
  } catch (error) {
    replica.free();
    throw error;
  }
}

function seededSeparators(fresh: () => Session): Uint8Array {
  const f = fresh();
  f.open_docx(fixture, true, GENERATION);
  const state = f.note_separators_state();
  if (state.length === 0) throw new Error('the seeded fixture carries no separator notes');
  return state;
}

function retainedMeta(meta: wasmEdit.RetainedLayoutMeta) {
  try {
    return {
      page_count: meta.page_count,
      page_sizes: meta.page_sizes(),
      partial: meta.partial,
      provisional: meta.provisional,
      notes_converged: meta.notes_converged,
      layout_shell_json: meta.layout_shell_json(),
    };
  } finally {
    meta.free();
  }
}

function retainedPages(session: Session) {
  const { page_count, page_sizes, layout_shell_json } = retainedMeta(session.layout_document_with_regions_retained_meta(layoutInput(session).input));
  return { page_count, page_sizes, layout_shell_json };
}

function breakBeforeFinal(session: Session): void {
  session.insert_page_break('body', paraId(session, 'Final'), 0);
}

interface OpenedDisplay { display: string; handle: number }

function openedDisplay({ engine, session }: Docx): OpenedDisplay {
  const display = displayList(session);
  return { display, handle: engine.open_display_list(display) };
}

function onDisplay<T>(engine: DocxEngine, { display, handle }: OpenedDisplay, query: (display: string, handle: number) => T): T {
  try {
    return query(display, handle);
  } finally {
    engine.close_display_list(handle);
  }
}

const HIT_POINTS: [number, number, number][] = [[0, 100, 100], [0, 100, 55], [0, 400, 900], [9, 100, 100]];
const RANGES: [number, number][] = [[1, 6], [1, 60]];
const MOVES: [number, string, number][] = [[1, 'down', Number.NaN], [1, 'up', Number.NaN], [1, 'down', 200]];

function handleAnswers(engine: DocxEngine, handle: number) {
  return {
    regions: HIT_POINTS.map(([page, x, y]) => engine.hit_test_regions_by_handle(handle, page, x, y)),
    rects: RANGES.map(([from, to]) => engine.range_rects_by_handle(handle, from, to)),
    moves: MOVES.map(([position, direction, goalX]) => engine.vertical_move_by_handle(handle, position, direction, goalX)),
  };
}

function jsonAnswers(engine: DocxEngine, display: string) {
  return {
    regions: HIT_POINTS.map(([page, x, y]) => engine.hit_test_regions_json(display, page, x, y)),
    rects: RANGES.map(([from, to]) => engine.range_rects_json(display, from, to)),
    moves: MOVES.map(([position, direction, goalX]) => engine.vertical_move_json(display, position, direction, goalX)),
  };
}

const MEASURED_BLOCK = { kind: 'paragraph', id: 1, runs: [{ kind: 'text', text: 'Parity measures one paragraph that wraps onto more than one line.' }] };

function measureInput(font: number, authoritativeShaping: boolean): string {
  return JSON.stringify({
    block: MEASURED_BLOCK,
    maxWidth: 160,
    fontChains: { 'liberation sans|0|0': [font] },
    defaults: { fontSize: 12, fontFamily: 'Liberation Sans' },
    authoritativeShaping,
  });
}

// The module font store outlives every session, so ids restart from a cleared store.
function moduleFont(engine: DocxEngine): number {
  engine.clear_measure_fonts();
  return engine.register_measure_font(fontBytes);
}

type Anchor = Record<string, unknown>;

function paragraphAnchors(session: Session): Anchor[] {
  const { paragraphs } = JSON.parse(session.paragraph_identities()) as {
    paragraphs: { session: Anchor | null; persisted: Anchor | null; source: Anchor | null }[];
  };
  const held = paragraphs.find((paragraph) => paragraph.session && paragraph.persisted);
  if (!held?.session || !held.persisted) throw new Error('no paragraph carries both a session and a persisted anchor');
  return [
    held.session,
    held.persisted,
    ...(held.source ? [held.source] : []),
    { ...held.session, sessionId: 'parity-foreign' },
    { ...held.persisted, paraId: '7FFFFFFF' },
  ];
}

const PARSES =['parse_docx_s2', 'parse_docx_s3', 'parse_docx_s4', 'parse_docx_s5', 'parse_docx_s6', 'parse_docx_s7', 'parse_docx_s8'] as const;

const rows: Row<Docx>[] = [
  { method: 'EditSession.open_docx_preview', call: ({ fresh }) => fresh().open_docx_preview(fixture, 4) },
  { method: 'EditSession.open_docx_preview_with_budget', call: ({ fresh }) => fresh().open_docx_preview_with_budget(fixture, 4, 2) },
  {
    method: 'EditSession.begin_opening',
    mutates: true,
    call: ({ session: s }) => {
      const result = s.begin_opening('engine-parity-reopened');
      return { result, after: s.paragraph_identities() };
    },
  },
  { method: 'EditSession.replace_range', mutates: true, call: ({ session: s }) => s.replace_range(...span(s, 'First body', 0, 5), 'Parity') },
  { method: 'EditSession.delete_range', mutates: true, call: ({ session: s }) => s.delete_range(...span(s, 'First body', 0, 6)) },
  { method: 'EditSession.split_paragraph', mutates: true, call: ({ session: s }) => s.split_paragraph('body', paraId(s, 'First body'), 5) },
  { method: 'EditSession.merge_paragraphs', mutates: true, call: ({ session: s }) => s.merge_paragraphs('body', paraId(s, 'First body')) },
  { method: 'EditSession.apply_paragraph_style', mutates: true, call: ({ session: s }) => s.apply_paragraph_style(...span(s, 'First body', 0, 5), 'Heading1') },
  { method: 'EditSession.set_paragraph_attr', mutates: true, call: ({ session: s }) => s.set_paragraph_attr(paraId(s, 'Second body'), 'alignment', '"center"') },
  {
    method: 'EditSession.set_paragraph_attrs',
    mutates: true,
    call: ({ session: s }) => s.set_paragraph_attrs(...span(s, 'Second body', 0, 6), JSON.stringify({ alignment: 'right', spaceAfter: 240 })),
  },
  { method: 'EditSession.insert_page_break', mutates: true, call: ({ session: s }) => s.insert_page_break('body', paraId(s, 'Final'), 0) },
  defineRow({
    method: 'EditSession.insert_image',
    mutates: true,
    setup: ({ session: s }: Docx) => seededImage(s),
    call: ({ session: s }, payload) => s.insert_image('body', paraId(s, 'Final'), 0, payload),
  }),
  defineRow({
    method: 'EditSession.insert_table_fragment',
    mutates: true,
    setup: ({ session: s }: Docx) => tableFragment(s),
    call: ({ session: s }, fragment) => s.insert_table_fragment('body', paraId(s, 'Final'), 0, fragment),
  }),
  {
    method: 'EditSession.add_comment',
    mutates: true,
    call: ({ session: s }) => {
      const id = paraId(s, 'Second body');
      const ranges = [{ story: 'body', startPara: id, startOffset: 0, endPara: id, endOffset: 6 }];
      return s.add_comment(JSON.stringify(ranges), 'Parity', DATE, JSON.stringify('Added comment'));
    },
  },
  defineRow({
    method: 'EditSession.resolve_comment',
    mutates: true,
    setup: ({ session: s }: Docx) => addedComment(s),
    call: ({ session: s }, commentId) => s.resolve_comment(commentId),
  }),
  { method: 'EditSession.list_comments', call: ({ session: s }) => s.list_comments() },
  { method: 'EditSession.set_content_control_value', mutates: true, call: ({ session: s }) => s.set_content_control_value('8', JSON.stringify('Grace Hopper')) },
  defineRow({
    method: 'EditSession.clear_content_control_value',
    mutates: true,
    setup: ({ session: s }: Docx) => s.set_content_control_value('9', JSON.stringify({ kind: 'checkbox', checked: true })),
    call: ({ session: s }) => s.clear_content_control_value('9'),
  }),
  { method: 'EditSession.list_content_controls_json', call: ({ session: s }) => s.list_content_controls_json('{}') },
  { method: 'EditSession.find_content_controls_json', call: ({ session: s }) => s.find_content_controls_json(TAG_QUERY, '{}') },
  { method: 'EditSession.content_control_targets_json', call: ({ session: s }) => s.content_control_targets_json() },
  { method: 'EditSession.search_text', call: ({ session: s }) => s.search_text('alpha', false, 10) },
  {
    method: 'EditSession.find_text_json',
    call: ({ session: s }) => s.find_text_json(JSON.stringify({ text: 'alpha', within: { kind: 'story', story: 'body' }, view: 'accepted' })),
  },
  { method: 'EditSession.headings_json', call: ({ session: s }) => s.headings_json('body') },
  { method: 'EditSession.list_sections_json', call: ({ session: s }) => s.list_sections_json() },
  { method: 'EditSession.section_margins_json', call: ({ session: s }) => s.section_margins_json() },
  { method: 'EditSession.document_properties_json', call: ({ session: s }) => s.document_properties_json() },
  { method: 'EditSession.part_metadata_json', call: ({ session: s }) => s.part_metadata_json() },
  { method: 'EditSession.materialize_docx', call: ({ session: s }) => s.materialize_docx() },
  { method: 'EditSession.selection_text_json', call: ({ session: s }) => s.selection_text_json(...span(s, 'Second body', 0, 6)) },
  { method: 'EditSession.export_markdown_json', call: ({ session: s }) => s.export_markdown_json(EXPORT) },
  { method: 'EditSession.export_structured_json', call: ({ session: s }) => s.export_structured_json(EXPORT) },
  { method: 'EditSession.register_measure_font', call: ({ session: s }) => s.register_measure_font(fontBytes) },
  defineRow({
    method: 'EditSession.layout_font_requirements_json',
    setup: ({ session: s }: Docx) => JSON.stringify(layoutRequest(s).request),
    call: ({ session: s }, request) => s.layout_font_requirements_json(request),
  }),
  defineRow({
    method: 'EditSession.layout_document_with_regions_json',
    setup: ({ session: s }: Docx) => layoutInput(s).input,
    call: ({ session: s }, input) => s.layout_document_with_regions_json(input),
  }),
  defineRow({
    method: 'EditSession.build_display_list_json',
    setup: ({ session: s }: Docx) => layout(s),
    call: ({ session: s }, { layout: json, fontChains }) => s.build_display_list_json(JSON.stringify({ ...JSON.parse(json), fontChains })),
  }),
  defineRow({
    method: 'hit_test_json',
    setup: ({ session: s }: Docx) => displayList(s),
    call: ({ engine }, display) => engine.hit_test_json(display, 0, 100, 100),
  }),
  defineRow({
    method: 'range_rects_json',
    setup: ({ session: s }: Docx) => displayList(s),
    call: ({ engine }, display) => engine.range_rects_json(display, 1, 6),
  }),
  { method: 'EditSession.insert_text', mutates: true, call: ({ session: s }) => s.insert_text('body', paraId(s, 'Second body'), 0, 'Inserted ') },
  { method: 'EditSession.insert_line_break', mutates: true, call: ({ session: s }) => s.insert_line_break('body', paraId(s, 'Second body'), 3) },
  { method: 'EditSession.insert_table', mutates: true, call: ({ session: s }) => s.insert_table('body', paraId(s, 'Final'), 0, 1, 2) },
  { method: 'EditSession.format_range', mutates: true, call: ({ session: s }) => s.format_range(...span(s, 'First body', 0, 5), JSON.stringify({ bold: true })) },
  defineRow({
    method: 'EditSession.clear_formatting',
    mutates: true,
    setup: ({ session: s }: Docx) => s.format_range(...span(s, 'First body', 0, 5), JSON.stringify({ italic: true })),
    call: ({ session: s }) => s.clear_formatting(...span(s, 'First body', 0, 5)),
  }),
  defineRow({
    method: 'EditSession.accept_change',
    mutates: true,
    setup: ({ session: s }: Docx) => trackedInsertion(s),
    call: ({ session: s }, revisionId) => s.accept_change(JSON.stringify({ revisionId })),
  }),
  defineRow({
    method: 'EditSession.reject_change',
    mutates: true,
    setup: ({ session: s }: Docx) => trackedInsertion(s),
    call: ({ session: s }, revisionId) => s.reject_change(JSON.stringify({ revisionId })),
  }),
  defineRow({
    method: 'EditSession.list_revisions',
    setup: ({ session: s }: Docx) => trackedInsertion(s),
    call: ({ session: s }) => s.list_revisions(),
  }),
  { method: 'EditSession.apply_edits_json', mutates: true, call: ({ session: s }) => s.apply_edits_json(editRequest(s)) },
  { method: 'EditSession.validate_edits_json', call: ({ session: s }) => s.validate_edits_json(editRequest(s)) },
  { method: 'EditSession.read_paragraphs_json', call: ({ session: s }) => s.read_paragraphs_json(JSON.stringify({ story: 'body', view: 'accepted' })) },
  { method: 'EditSession.read_stories_json', call: ({ session: s }) => s.read_stories_json(JSON.stringify({ view: 'accepted' })) },
  { method: 'EditSession.paragraphs', call: ({ session: s }) => s.paragraphs('body') },
  { method: 'EditSession.story_segments', call: ({ session: s }) => s.story_segments('body') },
  { method: 'EditSession.paragraph_spans', call: ({ session: s }) => s.paragraph_spans('body') },
  { method: 'EditSession.story_len', call: ({ session: s }) => s.story_len('body') },
  { method: 'EditSession.has_story', call: ({ session: s }) => [s.has_story('body'), s.has_story('missing')] },
  { method: 'EditSession.story_ids', call: ({ session: s }) => s.story_ids() },
  { method: 'EditSession.list_stories_json', call: ({ session: s }) => s.list_stories_json() },
  { method: 'EditSession.paragraph_identities', call: ({ session: s }) => s.paragraph_identities() },
  { method: 'EditSession.client_id', call: ({ session: s }) => s.client_id() },
  { method: 'EditSession.version', call: ({ session: s }) => s.version() },
  { method: 'EditSession.table_payload', call: ({ session: s }) => s.table_payload('body', 0) },
  defineRow({
    method: 'EditSession.validate_table_fragment',
    voidReason: 'a valid fragment passes by not throwing',
    setup: ({ session: s }: Docx) => tableFragment(s),
    call: ({ session: s }, fragment) => s.validate_table_fragment(fragment),
  }),
  defineRow({
    method: 'EditSession.validate_table_fragment',
    name: 'EditSession.validate_table_fragment (invalid)',
    throws: { name: 'string', reason: 'a fragment whose grid is not an array is refused with a message' },
    setup: ({ session: s }: Docx) => JSON.stringify({ ...JSON.parse(tableFragment(s)) as Record<string, unknown>, grid: 'none' }),
    call: ({ session: s }, fragment) => s.validate_table_fragment(fragment),
  }),
  defineRow({
    method: 'EditSession.undo',
    mutates: true,
    setup: ({ session: s }: Docx) => undoableEdit(s),
    call: ({ session: s }) => s.undo(),
  }),
  defineRow({
    method: 'EditSession.redo',
    mutates: true,
    setup: ({ session: s }: Docx) => {
      undoableEdit(s);
      s.undo();
    },
    call: ({ session: s }) => s.redo(),
  }),
  defineRow({
    method: 'EditSession.can_undo',
    setup: ({ session: s }: Docx) => undoableEdit(s),
    call: ({ session: s }) => s.can_undo(),
  }),
  defineRow({
    method: 'EditSession.can_redo',
    setup: ({ session: s }: Docx) => {
      undoableEdit(s);
      s.undo();
    },
    call: ({ session: s }) => s.can_redo(),
  }),
  defineRow({
    method: 'EditSession.history_stories',
    setup: ({ session: s }: Docx) => {
      undoableEdit(s);
      s.undo();
    },
    call: ({ session: s }) => s.history_stories(),
  }),
  {
    method: 'EditSession.track_undo',
    mutates: true,
    call: ({ session: s }) => {
      const result = s.track_undo();
      s.select_story('body');
      s.insert_text('body', paraId(s, 'Second body'), 0, 'Undone ');
      return { result, after: s.can_undo() };
    },
  },
  {
    method: 'EditSession.add_undo_boundary',
    mutates: true,
    call: ({ session: s }) => {
      undoableEdit(s);
      const result = s.add_undo_boundary();
      s.insert_text('body', paraId(s, 'Final'), 0, 'Later ');
      return { result, after: s.undo() };
    },
  },
  { method: 'EditSession.undo_capture_mode', call: ({ session: s }) => s.undo_capture_mode() },
  {
    method: 'EditSession.set_undo_capture_mode',
    call: ({ session: s }) => {
      const result = s.set_undo_capture_mode('manual');
      return { result, after: s.undo_capture_mode() };
    },
  },
  {
    method: 'EditSession.toggle_mark',
    mutates: true,
    call: ({ session: s }) => s.toggle_mark(...span(s, 'First body', 0, 5), JSON.stringify({ type: 'bold' })),
  },
  { method: 'EditSession.insert_row', mutates: true, call: ({ session: s }) => s.insert_row(cellLoc(0, 0), true) },
  { method: 'EditSession.insert_column', mutates: true, call: ({ session: s }) => s.insert_column(cellLoc(0, 0), true) },
  { method: 'EditSession.delete_row', mutates: true, call: ({ session: s }) => s.delete_row(cellRange([0, 0], [0, 1])) },
  { method: 'EditSession.delete_column', mutates: true, call: ({ session: s }) => s.delete_column(cellRange([0, 0], [1, 0])) },
  { method: 'EditSession.delete_table', mutates: true, call: ({ session: s }) => s.delete_table(TABLE_LOC) },
  { method: 'EditSession.merge_cells', mutates: true, call: ({ session: s }) => s.merge_cells(cellRange([0, 0], [0, 1])) },
  defineRow({
    method: 'EditSession.split_cell',
    mutates: true,
    setup: ({ session: s }: Docx) => s.merge_cells(cellRange([0, 0], [0, 1])),
    call: ({ session: s }) => s.split_cell(cellLoc(0, 0)),
  }),
  { method: 'EditSession.set_cell_shading', mutates: true, call: ({ session: s }) => s.set_cell_shading(cellRange([0, 0], [0, 0]), 'FF0000') },
  {
    method: 'EditSession.set_cell_text_format',
    mutates: true,
    call: ({ session: s }) => s.set_cell_text_format(cellRange([0, 0], [1, 1]), JSON.stringify({ margins: { left: 240, right: 108 } })),
  },
  {
    method: 'EditSession.set_cell_borders',
    mutates: true,
    call: ({ session: s }) => s.set_cell_borders(cellRange([0, 0], [0, 1]), JSON.stringify({ top: { style: 'single', size: 4, color: { rgb: '000000' } } })),
  },
  { method: 'EditSession.set_column_width', mutates: true, call: ({ session: s }) => s.set_column_width(cellLoc(0, 1), 3000) },
  { method: 'EditSession.set_row_height', mutates: true, call: ({ session: s }) => s.set_row_height(cellLoc(1, 0), 400, 'atLeast') },
  { method: 'EditSession.set_table_width', mutates: true, call: ({ session: s }) => s.set_table_width(TABLE_LOC, 5000) },
  {
    method: 'EditSession.set_table_properties',
    mutates: true,
    call: ({ session: s }) => s.set_table_properties(TABLE_LOC, JSON.stringify({ justification: 'center' })),
  },
  { method: 'EditSession.insert_section_break', mutates: true, call: ({ session: s }) => s.insert_section_break('body', paraId(s, 'Final'), 0, 'nextPage') },
  { method: 'EditSession.set_section_margins', mutates: true, call: ({ session: s }) => s.set_section_margins('final', JSON.stringify({ marginLeft: 1800 })) },
  defineRow({
    method: 'EditSession.set_comment_ranges',
    mutates: true,
    setup: ({ session: s }: Docx) => addedComment(s),
    call: ({ session: s }, commentId) => {
      const id = paraId(s, 'Final');
      const result = s.set_comment_ranges(commentId, JSON.stringify([{ story: 'body', startPara: id, startOffset: 0, endPara: id, endOffset: 5 }]));
      return { result, after: s.list_comments() };
    },
  }),
  defineRow({
    method: 'EditSession.update_comment_metadata_json',
    mutates: true,
    setup: ({ session: s }: Docx) => addedComment(s),
    call: ({ session: s }, commentId) => s.update_comment_metadata_json(commentId, JSON.stringify({ done: true })),
  }),
  { method: 'EditSession.locate_paragraph', call: ({ session: s }) => s.locate_paragraph('body', paraId(s, 'Second body')) },
  { method: 'EditSession.story_checksum', call: ({ session: s }) => s.story_checksum('body') },
  { method: 'EditSession.story_paragraph_ids', call: ({ session: s }) => s.story_paragraph_ids('body') },
  { method: 'write_docx_s13_wasm', call: ({ engine }) => new ZipReturn(engine.write_docx_s13_wasm(saveRequest(engine), fixture)) },
  { method: 'write_docx_s13_wasm_with_warnings', call: ({ engine }) => new FramedZipReturn(engine.write_docx_s13_wasm_with_warnings(saveRequest(engine), fixture)) },
  { method: 'parse_docx_s9', call: ({ engine }) => engine.parse_docx_s9(fixture, '') },
  { method: 'parse_docx_relationships', call: ({ engine }) => engine.parse_docx_relationships(fixture) },
  ...PARSES.map((method): Row<Docx> => ({ method, call: ({ engine }) => engine[method](fixture) })),
  { method: 'export_docx_structured_json', call: ({ engine }) => engine.export_docx_structured_json(fixture, EXPORT) },
  { method: 'export_docx_markdown_json', call: ({ engine }) => engine.export_docx_markdown_json(fixture, EXPORT) },
  { method: 'list_docx_content_controls_json', call: ({ engine }) => engine.list_docx_content_controls_json(fixture, '{}') },
  { method: 'find_docx_content_controls_json', call: ({ engine }) => engine.find_docx_content_controls_json(fixture, TAG_QUERY, '{}') },
  { method: 'font_preflight', call: ({ engine }) => engine.font_preflight(fixture) },
  { method: 'unzip_docx', call: ({ engine }) => engine.unzip_docx(fixture) },
  { method: 'rezip_docx', call: ({ engine }) => new ZipReturn(engine.rezip_docx(fixtureEntries)) },
  { method: 'sanitizeOoxml', call: ({ engine }) => new ZipReturn(engine.sanitizeOoxml(fixture, 'docx')) },
  { method: 'paragraphPatchWriterVersion', call: ({ engine }) => engine.paragraphPatchWriterVersion() },
  { method: 'skippedEffectKinds', call: ({ engine }) => engine.skippedEffectKinds() },
  // No EditSession save exists, so the save goes through a host and the replica's own reads are compared too.
  {
    method: 'EditSession.open_docx',
    mutates: true,
    call: ({ fresh, saveFrom }) => {
      const f = fresh();
      saveFrom(f);
      return { result: f.open_docx(fixture, true, GENERATION), after: f.paragraphs('body'), reads: replicaReads(f) };
    },
  },
  {
    method: 'EditSession.open_docx',
    name: 'EditSession.open_docx (not a docx)',
    throws: { name: 'string', reason: 'bytes that are not a readable DOCX are refused' },
    call: ({ fresh }) => fresh().open_docx(NOT_DOCX, true, GENERATION),
  },
  // No EditSession save exists, so the save goes through a host and the replica's own reads are compared too.
  {
    method: 'EditSession.seed_from_docx',
    mutates: true,
    call: ({ fresh, saveFrom }) => {
      const f = fresh();
      saveFrom(f);
      return { result: f.seed_from_docx(fixture, GENERATION), after: f.paragraphs('body'), reads: replicaReads(f) };
    },
  },
  {
    method: 'EditSession.seed_from_docx',
    name: 'EditSession.seed_from_docx (not a docx)',
    throws: { name: 'string', reason: 'bytes that are not a readable DOCX are refused' },
    call: ({ fresh }) => fresh().seed_from_docx(NOT_DOCX, GENERATION),
  },
  // No EditSession save exists, so the save goes through a host and the replica's own reads are compared too.
  defineRow({
    method: 'EditSession.load_json',
    mutates: true,
    setup: ({ fresh, saveFrom }: Docx) => {
      const f = fresh();
      f.begin_opening(GENERATION);
      saveFrom(f);
      return f;
    },
    call: (_, f) => {
      const result = f.load_json(JSON.stringify([
        { storyId: 'body', paragraphs: [{ text: 'Loaded' }, { text: 'Centred', alignment: 'center' }] },
        { storyId: 'fn:1', paragraphs: [{ text: 'Parity footnote' }] },
      ]));
      return { result, after: f.paragraphs('body'), reads: replicaReads(f) };
    },
  }),
  {
    method: 'EditSession.load_json',
    name: 'EditSession.load_json (existing story)',
    throws: { name: 'string', reason: 'a story id that already exists is refused' },
    call: ({ session: s }) => s.load_json(JSON.stringify([{ storyId: 'body', paragraphs: [{ text: 'Again' }] }])),
  },
  { method: 'EditSession.encode_state', call: ({ session: s }) => s.encode_state() },
  { method: 'EditSession.encode_state_vector', call: ({ session: s }) => s.encode_state_vector() },
  defineRow({
    method: 'EditSession.encode_diff',
    setup: ({ fresh }: Docx) => fresh().encode_state_vector(),
    call: ({ session: s }, vector) => s.encode_diff(vector),
  }),
  {
    method: 'EditSession.encode_diff',
    name: 'EditSession.encode_diff (malformed state vector)',
    throws: { name: 'string', reason: 'a state vector that does not decode is refused' },
    call: ({ session: s }) => s.encode_diff(MALFORMED_UPDATE),
  },
  defineRow({
    method: 'EditSession.paragraph_save_plan',
    setup: ({ session: s }: Docx) => s.split_paragraph('body', paraId(s, 'Second body'), textAt(s, 'Second body', 6)),
    call: ({ session: s }) => s.paragraph_save_plan(),
  }),
  defineRow({
    method: 'EditSession.persist_paragraph_ids',
    mutates: true,
    setup: ({ session: s }: Docx) => s.split_paragraph('body', paraId(s, 'Second body'), textAt(s, 'Second body', 6)),
    call: ({ session: s }) => s.persist_paragraph_ids(),
  }),
  defineRow({
    method: 'EditSession.persist_paragraph_ids',
    name: 'EditSession.persist_paragraph_ids (ambiguous comment reference)',
    setup: ({ fresh }: Docx) => {
      const f = fresh();
      f.seed_from_docx(twinCommentFixture, GENERATION);
      return f;
    },
    call: (_, f) => {
      const receipt = JSON.parse(f.persist_paragraph_ids()) as { status: string; refusal?: { kind?: string } };
      if (receipt.status !== 'refused' || receipt.refusal?.kind !== 'ambiguous-comment-reference') {
        throw new Error(`persist_paragraph_ids did not refuse: ${JSON.stringify(receipt)}`);
      }
      return { receipt, after: f.paragraph_identities() };
    },
  }),
  defineRow({
    method: 'EditSession.prime_save_seed_states',
    mutates: true,
    setup: (docx: Docx) => seededReplica(docx),
    call: (_, f) => {
      const primed = [f.prime_save_seed_states(100_000, 64 * 1024 * 1024), f.prime_save_seed_states(100_000, 64 * 1024 * 1024)];
      if (!primed[0]) throw new Error('a seeded replica within budget declined priming');
      return primed;
    },
  }),
  defineRow({
    method: 'EditSession.prime_save_seed_states',
    name: 'EditSession.prime_save_seed_states (over budget)',
    mutates: true,
    setup: (docx: Docx) => seededReplica(docx),
    call: (_, f) => [f.prime_save_seed_states(1, 1), f.prime_save_seed_states(100_000, 1)],
  }),
  {
    method: 'EditSession.prime_save_seed_states',
    name: 'EditSession.prime_save_seed_states (unseeded session)',
    call: ({ session: s }) => [s.prime_save_seed_states(1, 1), s.prime_save_seed_states(100_000, 64 * 1024 * 1024)],
  },
  defineRow({
    method: 'EditSession.prime_last_slice_cost',
    mutates: true,
    setup: (docx: Docx) => seededReplica(docx),
    call: (_, f) => primeSlices(f),
  }),
  {
    method: 'EditSession.prime_last_slice_cost',
    name: 'EditSession.prime_last_slice_cost (unseeded session)',
    mutates: true,
    call: ({ session: s }) => {
      const before = s.prime_last_slice_cost();
      const slices: [number, number][] = [];
      let answer = 1;
      for (let slice = 0; slice < 10_000 && answer === 1; slice += 1) {
        answer = s.prime_save_seed_states_slice(1_000, 64 * 1024 * 1024, 64);
        slices.push([answer, s.prime_last_slice_cost()]);
      }
      return { before, slices };
    },
  },
  defineRow({
    method: 'EditSession.prime_save_seed_states_slice',
    mutates: true,
    setup: (docx: Docx) => seededReplica(docx),
    call: (_, f) => {
      const answers = sliceAnswers(f, 1_000, 64 * 1024 * 1024);
      if (answers.at(-1) !== 0) throw new Error(`the seeded replica did not finish priming: ${JSON.stringify(answers)}`);
      return { answers, again: f.prime_save_seed_states_slice(1_000, 64 * 1024 * 1024, 64), whole: f.prime_save_seed_states(1_000, 64 * 1024 * 1024) };
    },
  }),
  defineRow({
    method: 'EditSession.prime_save_seed_states_slice',
    name: 'EditSession.prime_save_seed_states_slice (over budget)',
    mutates: true,
    setup: (docx: Docx) => seededReplica(docx),
    call: (_, f) => {
      const declined = [sliceAnswers(f, 1, 1), sliceAnswers(f, 1_000, 1)];
      if (declined.some((answers) => answers.at(-1) !== 2)) throw new Error(`an over-budget slice was not declined: ${JSON.stringify(declined)}`);
      return { declined, within: sliceAnswers(f, 1_000, 64 * 1024 * 1024) };
    },
  }),
  defineRow({
    method: 'EditSession.apply_update',
    mutates: true,
    setup: (docx: Docx) => trackedRemote(docx, remoteUpdate),
    call: ({ session: s }, update) => notUndoable(s, () => s.apply_update(update)),
  }),
  {
    method: 'EditSession.apply_update',
    name: 'EditSession.apply_update (malformed update)',
    throws: { name: 'string', reason: 'an update that does not decode is refused' },
    call: ({ session: s }) => s.apply_update(MALFORMED_UPDATE),
  },
  defineRow({
    method: 'EditSession.apply_update_chunks',
    mutates: true,
    setup: (docx: Docx) => trackedRemote(docx, remoteChunks),
    call: ({ session: s }, updates) => notUndoable(s, () => s.apply_update_chunks(updates)),
  }),
  {
    method: 'EditSession.apply_update_chunks',
    name: 'EditSession.apply_update_chunks (malformed chunk)',
    throws: { name: 'string', reason: 'a chunk that does not decode is refused before any applies' },
    call: ({ session: s }) => s.apply_update_chunks([MALFORMED_UPDATE]),
  },
  defineRow({
    method: 'EditSession.apply_update_chunks',
    name: 'EditSession.apply_update_chunks (valid chunk then malformed)',
    mutates: true,
    throws: { name: 'string', reason: 'one malformed chunk refuses the whole sync, so the valid chunk before it does not apply' },
    setup: (docx: Docx) => remoteChunks(docx)[0]!,
    call: ({ session: s }, valid) => {
      const before = s.paragraphs('body');
      try {
        return s.apply_update_chunks([valid, MALFORMED_UPDATE]);
      } catch (error) {
        if (s.paragraphs('body') !== before) throw new Error('a refused sync applied its valid chunk');
        throw error;
      }
    },
  }),
  defineRow({
    method: 'EditSession.apply_local_update',
    mutates: true,
    setup: (docx: Docx) => {
      docx.session.track_undo();
      docx.session.select_story('body');
      return remoteUpdate(docx);
    },
    call: ({ session: s }, update) => ({ result: s.apply_local_update(update), after: s.can_undo() }),
  }),
  {
    method: 'EditSession.apply_local_update',
    name: 'EditSession.apply_local_update (malformed update)',
    throws: { name: 'string', reason: 'an update that does not decode is refused' },
    call: ({ session: s }) => s.apply_local_update(MALFORMED_UPDATE),
  },
  {
    method: 'EditSession.apply_raw_ops',
    mutates: true,
    call: ({ session: s }) => s.apply_raw_ops('body', JSON.stringify([{ op: 'insert', index: 0, text: 'Raw ' }, { op: 'format', index: 0, len: 3, attrs: { bold: true } }])),
  },
  {
    method: 'EditSession.apply_raw_ops',
    name: 'EditSession.apply_raw_ops (unknown op)',
    throws: { name: 'string', reason: 'an unknown op is refused and the story is left untouched' },
    call: ({ session: s }) => s.apply_raw_ops('body', JSON.stringify([{ op: 'bogus', index: 0 }])),
  },
  {
    method: 'EditSession.set_selection',
    call: ({ session: s }) => ({ result: s.set_selection(...textSpan(s, 'Second body', 0, 6)), after: s.selection() }),
  },
  {
    method: 'EditSession.set_selection',
    name: 'EditSession.set_selection (bad offset)',
    throws: { name: 'string', reason: 'an offset past the paragraph end does not resolve' },
    call: ({ session: s }) => {
      const id = paraId(s, 'Second body');
      return s.set_selection('body', id, 9_999, id, 9_999);
    },
  },
  defineRow({
    method: 'EditSession.apply_input',
    mutates: true,
    setup: ({ session: s }: Docx) => {
      const epoch = residentFrame(s);
      caretIn(s, 'Second body', 6);
      return epoch;
    },
    call: ({ session: s }, epoch) => ({ frame: s.apply_input('X', epoch), after: s.selection() }),
  }),
  defineRow({
    method: 'EditSession.apply_input',
    name: 'EditSession.apply_input (no selection)',
    throws: { name: 'string', reason: 'typing needs a resident selection' },
    setup: ({ session: s }: Docx) => residentFrame(s),
    call: ({ session: s }, epoch) => s.apply_input('X', epoch),
  }),
  defineRow({
    method: 'EditSession.apply_input',
    name: 'EditSession.apply_input (range selection)',
    throws: { name: 'string', reason: 'typing needs a collapsed selection' },
    setup: ({ session: s }: Docx) => {
      const epoch = residentFrame(s);
      s.set_selection(...textSpan(s, 'Second body', 0, 6));
      return epoch;
    },
    call: ({ session: s }, epoch) => s.apply_input('X', epoch),
  }),
  defineRow({
    method: 'EditSession.apply_delete',
    mutates: true,
    setup: ({ session: s }: Docx) => {
      const epoch = residentFrame(s);
      caretIn(s, 'Second body', 6);
      return epoch;
    },
    call: ({ session: s }, epoch) => {
      const frame = s.apply_delete('backward', epoch, 1);
      const deleted = s.resident_deleted_units();
      if (deleted !== 1) throw new Error(`apply_delete removed ${deleted} units`);
      return { frame, deleted, after: s.selection() };
    },
  }),
  defineRow({
    method: 'EditSession.apply_delete',
    name: 'EditSession.apply_delete (unknown direction)',
    throws: { name: 'string', reason: 'a direction other than backward or forward is refused' },
    setup: ({ session: s }: Docx) => {
      const epoch = residentFrame(s);
      caretIn(s, 'Second body', 6);
      return epoch;
    },
    call: ({ session: s }, epoch) => s.apply_delete('sideways', epoch, 1),
  }),
  {
    method: 'EditSession.set_hyperlink',
    mutates: true,
    call: ({ session: s }) => s.set_hyperlink(...textSpan(s, 'Second body', 0, 6), HYPERLINK),
  },
  defineRow({
    method: 'EditSession.set_hyperlink',
    name: 'EditSession.set_hyperlink (unlink)',
    mutates: true,
    setup: ({ session: s }: Docx) => s.set_hyperlink(...textSpan(s, 'Second body', 0, 6), HYPERLINK),
    call: ({ session: s }) => s.set_hyperlink(...textSpan(s, 'Second body', 0, 6), 'null'),
  }),
  {
    method: 'EditSession.set_hyperlink',
    name: 'EditSession.set_hyperlink (not an object)',
    throws: { name: 'string', reason: 'hyperlink JSON that is neither an object nor null is refused' },
    call: ({ session: s }) => s.set_hyperlink(...textSpan(s, 'Second body', 0, 6), '"https://example.com/"'),
  },
  {
    method: 'EditSession.set_image_geometry_at',
    mutates: true,
    call: ({ session: s }) => {
      const at = imageAt(s);
      return s.set_image_geometry_at('body', at.paraId, at.offset, GEOMETRY);
    },
  },
  {
    method: 'EditSession.set_image_geometry_at',
    name: 'EditSession.set_image_geometry_at (no image)',
    throws: { name: 'string', reason: 'a position that holds no image is refused' },
    call: ({ session: s }) => s.set_image_geometry_at('body', imageAt(s).paraId, textAt(s, 'Picture', 0), GEOMETRY),
  },
  {
    method: 'EditSession.set_undo_input_time',
    mutates: true,
    call: ({ session: s }) => timedEdits(s, () => s.set_undo_input_time(5_100, 5_100)),
  },
  {
    method: 'EditSession.set_undo_input_time',
    name: 'EditSession.set_undo_input_time (typing within the window)',
    mutates: true,
    call: ({ session: s }) => typedUndoSteps(s, 100),
  },
  {
    method: 'EditSession.set_undo_input_time',
    name: 'EditSession.set_undo_input_time (typing past the window)',
    mutates: true,
    call: ({ session: s }) => typedUndoSteps(s, 600),
  },
  {
    method: 'EditSession.clear_undo_input_time',
    mutates: true,
    voidReason: 'its effect, falling back to each engine default undo clock, is not observable without wall-clock time',
    call: ({ session: s }) => {
      s.track_undo();
      s.select_story('body');
      s.set_undo_input_time(5_000, 5_000);
      s.insert_text('body', paraId(s, 'Second body'), textAt(s, 'Second body', 0), 'a');
      return s.clear_undo_input_time();
    },
  },
  {
    method: 'EditSession.set_cell_selection',
    call: ({ session: s }) => ({ result: s.set_cell_selection(cellRange([0, 0], [1, 1])), after: s.cell_selection() }),
  },
  {
    method: 'EditSession.set_cell_selection',
    name: 'EditSession.set_cell_selection (missing cell)',
    throws: { name: 'string', reason: 'a cell that does not resolve is refused' },
    call: ({ session: s }) => s.set_cell_selection(cellRange([0, 0], [5, 0])),
  },
  defineRow({
    method: 'EditSession.cell_selection',
    setup: ({ session: s }: Docx) => s.set_cell_selection(cellRange([0, 1], [1, 0])),
    call: ({ session: s }) => s.cell_selection(),
  }),
  defineRow({
    method: 'EditSession.cell_selection',
    name: 'EditSession.cell_selection (deleted table)',
    throws: { name: 'string', reason: 'a selection whose table was deleted no longer resolves' },
    setup: ({ session: s }: Docx) => {
      s.set_cell_selection(cellRange([0, 0], [1, 1]));
      s.delete_table(TABLE_LOC);
    },
    call: ({ session: s }) => s.cell_selection(),
  }),
  { method: 'EditSession.selection_context', call: ({ session: s }) => s.selection_context(...textSpan(s, 'First body', 0, 5)) },
  {
    method: 'EditSession.selection_context',
    name: 'EditSession.selection_context (unknown paragraph)',
    throws: { name: 'string', reason: 'a paragraph id the story does not hold is refused' },
    call: ({ session: s }) => s.selection_context('body', '7FFFFFFF', 0, '7FFFFFFF', 1),
  },
  defineRow({
    method: 'EditSession.media_bytes',
    setup: ({ fresh }: Docx) => tokenizedImage(fresh),
    call: (_, { f, src }) => [f.media_bytes(src), f.media_bytes('media:missing')],
  }),
  defineRow({
    method: 'EditSession.stories_changed_since',
    setup: ({ session: s }: Docx) => {
      const { revision } = JSON.parse(s.stories_changed_since(Number.MAX_SAFE_INTEGER)) as { revision: number };
      s.insert_text('body', paraId(s, 'Second body'), textAt(s, 'Second body', 0), 'Changed ');
      return revision;
    },
    call: ({ session: s }, revision) => s.stories_changed_since(revision),
  }),
  // No EditSession save exists, so the save goes through a host and the replica's own reads are compared too.
  defineRow({
    method: 'EditSession.create_story',
    mutates: true,
    setup: ({ fresh, saveFrom }: Docx) => {
      const f = fresh();
      f.begin_opening(GENERATION);
      f.create_story('fn:1', 'Parity footnote', 'Normal', 'left');
      saveFrom(f);
      return f;
    },
    call: (_, f) => ({ result: f.create_story('body', 'Created', 'Heading1', 'center'), after: f.paragraphs('body'), reads: replicaReads(f) }),
  }),
  {
    method: 'EditSession.create_story',
    name: 'EditSession.create_story (existing story)',
    throws: { name: 'string', reason: 'a story id that already exists is refused' },
    call: ({ session: s }) => s.create_story('body', 'Again', 'Normal', 'left'),
  },
  defineRow({
    method: 'EditSession.set_display_window',
    setup: ({ session: s }: Docx) => layout(s).fontChains,
    call: ({ session: s }, fontChains) => ({ result: s.set_display_window(1, 2), after: s.build_display_list_frame(frameExtras(fontChains), 0) }),
  }),
  defineRow({
    method: 'EditSession.build_display_pages_frame',
    setup: ({ session: s }: Docx) => {
      const { fontChains } = layout(s);
      s.set_display_window(1, 2);
      s.build_display_list_frame(frameExtras(fontChains), 0);
      return frameEpoch(s);
    },
    call: ({ session: s }, epoch) => {
      const frame = s.build_display_pages_frame(Uint32Array.from([0]), epoch);
      const kept = frameEpoch(s);
      return { frame, kept, again: s.build_display_pages_frame(Uint32Array.from([0]), kept) };
    },
  }),
  defineRow({
    method: 'EditSession.build_display_pages_frame',
    name: 'EditSession.build_display_pages_frame (bad epoch)',
    throws: { name: 'string', reason: 'an expected frame epoch that is not a non-negative safe integer is refused' },
    setup: ({ session: s }: Docx) => residentFrame(s),
    call: ({ session: s }) => s.build_display_pages_frame(Uint32Array.from([0]), -1),
  }),
  defineRow({
    method: 'EditSession.build_display_pages_frame',
    name: 'EditSession.build_display_pages_frame (fractional epoch)',
    throws: { name: 'string', reason: 'an expected frame epoch that is not a non-negative safe integer is refused' },
    setup: ({ session: s }: Docx) => residentFrame(s),
    call: ({ session: s }, epoch) => s.build_display_pages_frame(Uint32Array.from([0]), epoch + 0.5),
  }),
  defineRow({
    method: 'EditSession.build_display_pages_frame',
    name: 'EditSession.build_display_pages_frame (kept after a refused epoch)',
    setup: ({ session: s }: Docx) => {
      const { fontChains } = layout(s);
      s.set_display_window(1, 2);
      s.build_display_list_frame(frameExtras(fontChains), 0);
    },
    call: ({ session: s }) => refusedThenKept(
      s,
      () => s.build_display_pages_frame(Uint32Array.from([0]), -1),
      (kept) => s.build_display_pages_frame(Uint32Array.from([0]), kept),
    ),
  }),
  defineRow({
    method: 'EditSession.build_display_pages_frame',
    name: 'EditSession.build_display_pages_frame (epoch one past the current)',
    setup: ({ session: s }: Docx) => {
      const { fontChains } = layout(s);
      s.set_display_window(1, 2);
      s.build_display_list_frame(frameExtras(fontChains), 0);
      return frameEpoch(s);
    },
    call: ({ session: s }, epoch) => ({ frame: s.build_display_pages_frame(Uint32Array.from([0]), epoch + 1), kept: frameEpoch(s) }),
  }),
  defineRow({
    method: 'EditSession.build_display_list_frame',
    setup: ({ session: s }: Docx) => layout(s).fontChains,
    call: ({ session: s }, fontChains) => {
      const frame = s.build_display_list_frame(frameExtras(fontChains), 0);
      const kept = frameEpoch(s);
      return { frame, kept, again: s.build_display_list_frame(frameExtras(fontChains), kept) };
    },
  }),
  defineRow({
    method: 'EditSession.build_display_list_frame',
    name: 'EditSession.build_display_list_frame (bad epoch)',
    throws: { name: 'string', reason: 'an expected frame epoch that is not a non-negative safe integer is refused' },
    setup: ({ session: s }: Docx) => layout(s).fontChains,
    call: ({ session: s }, fontChains) => s.build_display_list_frame(frameExtras(fontChains), -1),
  }),
  defineRow({
    method: 'EditSession.build_display_list_frame',
    name: 'EditSession.build_display_list_frame (fractional epoch)',
    throws: { name: 'string', reason: 'an expected frame epoch that is not a non-negative safe integer is refused' },
    setup: ({ session: s }: Docx) => layout(s).fontChains,
    call: ({ session: s }, fontChains) => s.build_display_list_frame(frameExtras(fontChains), 0.5),
  }),
  defineRow({
    method: 'EditSession.build_display_list_frame',
    name: 'EditSession.build_display_list_frame (kept after a refused epoch)',
    setup: ({ session: s }: Docx) => {
      const { fontChains } = layout(s);
      s.build_display_list_frame(frameExtras(fontChains), 0);
      return fontChains;
    },
    call: ({ session: s }, fontChains) => refusedThenKept(
      s,
      () => s.build_display_list_frame(frameExtras(fontChains), -1),
      (kept) => s.build_display_list_frame(frameExtras(fontChains), kept),
    ),
  }),
  defineRow({
    method: 'EditSession.build_display_list_frame',
    name: 'EditSession.build_display_list_frame (kept after unparseable input)',
    setup: ({ session: s }: Docx) => {
      const { fontChains } = layout(s);
      s.build_display_list_frame(frameExtras(fontChains), 0);
      return fontChains;
    },
    call: ({ session: s }, fontChains) => refusedThenKept(
      s,
      () => s.build_display_list_frame('{', frameEpoch(s)),
      (kept) => s.build_display_list_frame(frameExtras(fontChains), kept),
    ),
  }),
  defineRow({
    method: 'EditSession.build_display_list_frame',
    name: 'EditSession.build_display_list_frame (epoch one past the current)',
    setup: ({ session: s }: Docx) => {
      const { fontChains } = layout(s);
      s.build_display_list_frame(frameExtras(fontChains), 0);
      return fontChains;
    },
    call: ({ session: s }, fontChains) => {
      const epoch = frameEpoch(s);
      return { frame: s.build_display_list_frame(frameExtras(fontChains), epoch + 1), kept: frameEpoch(s) };
    },
  }),
  {
    method: 'EditSession.build_display_list_json',
    name: 'EditSession.build_display_list_json (unparseable)',
    throws: { name: 'string', reason: 'input that is not JSON is refused' },
    call: ({ session: s }) => s.build_display_list_json('{'),
  },
  {
    method: 'EditSession.layout_document_with_regions_json',
    name: 'EditSession.layout_document_with_regions_json (unparseable)',
    throws: { name: 'string', reason: 'input that is not JSON is refused' },
    call: ({ session: s }) => s.layout_document_with_regions_json('{'),
  },
  defineRow({
    method: 'EditSession.build_display_list_frame',
    name: 'EditSession.build_display_list_frame (unparseable)',
    throws: { name: 'string', reason: 'display input that is not JSON is refused' },
    setup: ({ session: s }: Docx) => layout(s).fontChains,
    call: ({ session: s }) => s.build_display_list_frame('{', 0),
  }),
  defineRow({
    method: 'EditSession.layout_document_with_regions_retained',
    setup: ({ session: s }: Docx) => layoutInput(s).input,
    call: ({ session: s }, input) => ({ result: s.layout_document_with_regions_retained(input), after: s.retained_layout_json() }),
  }),
  {
    method: 'EditSession.layout_document_with_regions_retained',
    name: 'EditSession.layout_document_with_regions_retained (unparseable)',
    throws: { name: 'string', reason: 'input that is not JSON is refused' },
    call: ({ session: s }) => s.layout_document_with_regions_retained('{'),
  },
  defineRow({
    method: 'EditSession.layout_document_with_regions_retained_json',
    setup: ({ session: s }: Docx) => layoutInput(s).input,
    call: ({ session: s }, input) => s.layout_document_with_regions_retained_json(input),
  }),
  {
    method: 'EditSession.layout_document_with_regions_retained_json',
    name: 'EditSession.layout_document_with_regions_retained_json (unparseable)',
    throws: { name: 'string', reason: 'input that is not JSON is refused' },
    call: ({ session: s }) => s.layout_document_with_regions_retained_json('{'),
  },
  defineRow({
    method: 'EditSession.layout_document_with_regions_retained_meta',
    setup: ({ session: s }: Docx) => layoutInput(s).input,
    call: ({ session: s }, input) => retainedMeta(s.layout_document_with_regions_retained_meta(input)),
  }),
  {
    method: 'EditSession.layout_document_with_regions_retained_meta',
    name: 'EditSession.layout_document_with_regions_retained_meta (unparseable)',
    throws: { name: 'string', reason: 'input that is not JSON is refused' },
    call: ({ session: s }) => s.layout_document_with_regions_retained_meta('{'),
  },
  defineRow({
    method: 'EditSession.layout_document_with_regions_after_edit',
    mutates: true,
    setup: ({ session: s }: Docx) => {
      const { input } = layoutInput(s);
      s.layout_document_with_regions_retained(input);
      return input;
    },
    call: ({ session: s }, input) => {
      const id = paraId(s, 'Second body');
      const at = textAt(s, 'Second body', 0);
      const receipt = s.edit_text_resident('body', id, at, id, at, 'Resident ');
      const result = s.layout_document_with_regions_after_edit(input);
      return { receipt, result, layout: s.retained_layout_json(), headersFooters: s.retained_headers_footers_json() };
    },
  }),
  {
    method: 'EditSession.layout_document_with_regions_after_edit',
    name: 'EditSession.layout_document_with_regions_after_edit (unparseable)',
    throws: { name: 'string', reason: 'input that is not JSON is refused' },
    call: ({ session: s }) => s.layout_document_with_regions_after_edit('{'),
  },
  {
    method: 'RetainedLayoutMeta.page_count',
    call: ({ session: s }) => {
      breakBeforeFinal(s);
      return retainedPages(s);
    },
  },
  {
    method: 'RetainedLayoutMeta.page_sizes',
    call: ({ session: s }) => {
      const onePage = retainedPages(s);
      breakBeforeFinal(s);
      return { onePage, twoPages: retainedPages(s) };
    },
  },
  {
    method: 'RetainedLayoutMeta.partial',
    call: ({ session: s }) => {
      const read = () => {
        const { partial, page_count, page_sizes } = retainedMeta(s.layout_document_with_regions_retained_meta(layoutInput(s).input));
        return { partial, page_count, page_sizes };
      };
      const whole = read();
      const marked = s.set_partial_document(true);
      const partial = read();
      if (whole.partial !== false || partial.partial !== true) throw new Error(`the retained meta does not carry the partial flag as set: ${whole.partial} then ${partial.partial}`);
      return { whole, marked, partial };
    },
  },
  defineRow({
    method: 'EditSession.retained_headers_footers_json',
    setup: ({ session: s }: Docx) => {
      const base = regionRequest();
      return layoutInput(s, { ...base, regions: { sections: [{ ...base.regions.sections[0], headerFooterRefs: { headerDefault: 'rIdHeader' } }] } }).input;
    },
    call: ({ session: s }, input) => {
      const before = s.retained_headers_footers_json();
      s.layout_document_with_regions_retained(input);
      const after = s.retained_headers_footers_json();
      if (after === undefined) throw new Error('the retained layout of a document with a header holds no headersFooters');
      return { before, after };
    },
  }),
  defineRow({
    method: 'EditSession.display_hit_test_regions_json',
    setup: ({ session: s }: Docx) => residentFrame(s),
    call: ({ session: s }) => HIT_POINTS.map(([page, x, y]) => s.display_hit_test_regions_json(page, x, y)),
  }),
  {
    method: 'EditSession.display_hit_test_regions_json',
    name: 'EditSession.display_hit_test_regions_json (no resident display list)',
    throws: { name: 'string', reason: 'a session with no resident display list is refused' },
    call: ({ session: s }) => s.display_hit_test_regions_json(0, 100, 100),
  },
  defineRow({
    method: 'EditSession.display_range_rects_json',
    setup: ({ session: s }: Docx) => residentFrame(s),
    call: ({ session: s }) => RANGES.map(([from, to]) => s.display_range_rects_json(from, to)),
  }),
  {
    method: 'EditSession.display_range_rects_json',
    name: 'EditSession.display_range_rects_json (no resident display list)',
    throws: { name: 'string', reason: 'a session with no resident display list is refused' },
    call: ({ session: s }) => s.display_range_rects_json(1, 6),
  },
  defineRow({
    method: 'EditSession.display_vertical_move_json',
    setup: ({ session: s }: Docx) => residentFrame(s),
    call: ({ session: s }) => MOVES.map(([position, direction, goalX]) => s.display_vertical_move_json(position, direction, goalX)),
  }),
  defineRow({
    method: 'EditSession.display_vertical_move_json',
    name: 'EditSession.display_vertical_move_json (unknown direction)',
    throws: { name: 'string', reason: 'a direction other than up or down is refused' },
    setup: ({ session: s }: Docx) => residentFrame(s),
    call: ({ session: s }) => s.display_vertical_move_json(1, 'sideways', 0),
  }),
  defineRow({
    method: 'open_display_list',
    setup: openedDisplay,
    call: ({ engine }, opened) => onDisplay(engine, opened, (display, handle) => ({ byHandle: handleAnswers(engine, handle), byJson: jsonAnswers(engine, display) })),
  }),
  defineRow({
    method: 'open_display_list',
    name: 'open_display_list (display list built on the other engine)',
    setup: ({ engine, session: s }: Docx) => {
      const other = otherEngineReplica(engine, s);
      try {
        return displayList(other);
      } finally {
        other.free();
      }
    },
    call: ({ engine }, display) => onDisplay(engine, { display, handle: engine.open_display_list(display) }, (_, handle) => handleAnswers(engine, handle)),
  }),
  {
    method: 'open_display_list',
    name: 'open_display_list (unparseable)',
    throws: { name: 'string', reason: 'a display list that is not JSON is refused' },
    call: ({ engine }) => engine.open_display_list('{'),
  },
  defineRow({
    method: 'hit_test_regions_json',
    setup: openedDisplay,
    call: ({ engine }, opened) => onDisplay(engine, opened, (display) => HIT_POINTS.map(([page, x, y]) => engine.hit_test_regions_json(display, page, x, y))),
  }),
  {
    method: 'hit_test_regions_json',
    name: 'hit_test_regions_json (unparseable)',
    throws: { name: 'string', reason: 'a display list that is not JSON is refused' },
    call: ({ engine }) => engine.hit_test_regions_json('{', 0, 100, 100),
  },
  defineRow({
    method: 'range_rects_by_handle',
    setup: openedDisplay,
    call: ({ engine }, opened) => onDisplay(engine, opened, (_, handle) => RANGES.map(([from, to]) => engine.range_rects_by_handle(handle, from, to))),
  }),
  {
    method: 'range_rects_by_handle',
    name: 'range_rects_by_handle (unknown handle)',
    throws: { name: 'string', reason: 'handle 0, which open_display_list never issues, is refused' },
    call: ({ engine }) => engine.range_rects_by_handle(0, 1, 6),
  },
  defineRow({
    method: 'vertical_move_json',
    setup: openedDisplay,
    call: ({ engine }, opened) => onDisplay(engine, opened, (display) => MOVES.map(([position, direction, goalX]) => engine.vertical_move_json(display, position, direction, goalX))),
  }),
  defineRow({
    method: 'vertical_move_json',
    name: 'vertical_move_json (unknown direction)',
    throws: { name: 'string', reason: 'a direction other than up or down is refused' },
    setup: openedDisplay,
    call: ({ engine }, opened) => onDisplay(engine, opened, (display) => engine.vertical_move_json(display, 1, 'sideways', 0)),
  }),
  defineRow({
    method: 'measure_paragraph_json',
    setup: ({ engine }: Docx) => moduleFont(engine),
    call: ({ engine }, font) => ({ font, plain: engine.measure_paragraph_json(measureInput(font, false)), authoritative: engine.measure_paragraph_json(measureInput(font, true)) }),
  }),
  {
    method: 'measure_paragraph_json',
    name: 'measure_paragraph_json (unparseable)',
    throws: { name: 'string', reason: 'measurement input that is not JSON is refused' },
    call: ({ engine }) => engine.measure_paragraph_json('{'),
  },
  defineRow({
    method: 'EditSession.layout_document_json',
    setup: ({ session: s }: Docx) => workerLayoutInput(s),
    call: ({ session: s }, { input, fontChains }) => ({ result: s.layout_document_json(input), after: s.build_display_list_frame(frameExtras(fontChains), 0) }),
  }),
  {
    method: 'EditSession.layout_document_json',
    name: 'EditSession.layout_document_json (unparseable)',
    throws: { name: 'string', reason: 'input that is not JSON is refused' },
    call: ({ session: s }) => s.layout_document_json('{'),
  },
  { method: 'layout_document_json', call: ({ engine }) => engine.layout_document_json(JSON.stringify({ measured: [], options: MODULE_OPTIONS })) },
  {
    method: 'layout_document_json',
    name: 'layout_document_json (unparseable)',
    throws: { name: 'string', reason: 'input that is not JSON is refused' },
    call: ({ engine }) => engine.layout_document_json('{'),
  },
  defineRow({
    method: 'build_display_list_json',
    setup: ({ engine }: Docx) => JSON.parse(engine.layout_document_json(JSON.stringify({ measured: [], options: MODULE_OPTIONS }))) as unknown,
    call: ({ engine }, layout) => engine.build_display_list_json(JSON.stringify({ measured: [], options: MODULE_OPTIONS, layout })),
  }),
  {
    method: 'build_display_list_json',
    name: 'build_display_list_json (unparseable)',
    throws: { name: 'string', reason: 'input that is not JSON is refused' },
    call: ({ engine }) => engine.build_display_list_json('{'),
  },
  {
    method: 'serialize_docx_s10',
    call: ({ engine }) => {
      const { document, table } = wirePackage(engine);
      return serialized((request) => engine.serialize_docx_s10(request), [
        { family: 'section', properties: document.finalSectionProperties },
        { family: 'tableGrid', table },
      ]);
    },
  },
  {
    method: 'serialize_docx_s11',
    call: ({ engine }) => {
      const { paragraph, run, table } = wirePackage(engine);
      return serialized((request) => engine.serialize_docx_s11(request), [
        { family: 'paragraph', paragraph },
        { family: 'run', run },
        { family: 'table', table },
      ]);
    },
  },
  {
    method: 'serialize_docx_s12',
    call: ({ engine }) => {
      const { document, header, footnotes, comments } = wirePackage(engine);
      return serialized((request) => engine.serialize_docx_s12(request), [
        { family: 'document', body: { content: document.content, finalSectionProperties: document.finalSectionProperties } },
        { family: 'headerFooter', story: header },
        { family: 'footnotes', notes: footnotes },
        { family: 'comments', comments },
      ]);
    },
  },
  { method: 'render_docx_markdown_json', call: ({ engine }) => engine.render_docx_markdown_json(structuredContent(engine), '{}') },
  defineRow({
    method: 'render_docx_markdown_with_pages_json',
    setup: ({ session: s }: Docx) => JSON.stringify((JSON.parse(pagedExport(s, laidOut(s))) as { content: unknown }).content),
    call: ({ engine }, content) => engine.render_docx_markdown_with_pages_json(content, JSON.stringify({ pageMarkers: true })),
  }),
  defineRow({
    method: 'EditSession.export_structured_with_pages_json',
    setup: ({ session: s }: Docx) => laidOut(s),
    call: ({ session: s }, request) => pagedExport(s, request),
  }),
  defineRow({
    method: 'EditSession.page_texts_json',
    setup: ({ session: s }: Docx) => layout(s),
    call: ({ session: s }) => pageTexts(s),
  }),
  { method: 'EditSession.written_paragraph_ids', call: ({ session: s }) => s.written_paragraph_ids(fixture) },
  {
    method: 'EditSession.record_saved_paragraph_ids',
    mutates: true,
    call: ({ session: s }) => {
      const stale = s.record_saved_paragraph_ids(JSON.stringify(savedPairs(s)));
      if ((JSON.parse(stale) as unknown[]).length === 0) throw new Error('the pair holding another paragraph ID was not stale');
      return { stale, after: s.paragraph_identities() };
    },
  },
  {
    method: 'EditSession.compare_docx_json',
    mutates: true,
    call: ({ fresh, saveFrom }) => {
      const f = fresh();
      saveFrom(f, compareOriginal);
      const result = comparedInto(f);
      return { result, version: f.version(), paragraphs: f.paragraphs('body'), revisions: f.list_revisions() };
    },
  },
  defineRow({
    method: 'EditSession.finish_compared_docx_json',
    setup: async ({ fresh, saved }: Docx) => {
      const f = fresh();
      comparedInto(f);
      return { f, bytes: await saved(f, compareOriginal) };
    },
    call: (_, { f, bytes }) => ({ result: f.finish_compared_docx_json(bytes), version: f.version() }),
  }),
  {
    method: 'EditSession.finish_compared_docx_json',
    name: 'EditSession.finish_compared_docx_json (no comparison pending)',
    throws: { name: 'string', reason: 'saved bytes are refused while no comparison awaits them' },
    call: ({ session: s }) => s.finish_compared_docx_json(fixture),
  },
  defineRow({
    method: 'EditSession.fail_compared_docx_json',
    setup: ({ fresh }: Docx) => {
      const f = fresh();
      comparedInto(f);
      return f;
    },
    call: (_, f) => ({ result: f.fail_compared_docx_json('the compared document could not be written'), version: f.version() }),
  }),
  {
    method: 'EditSession.fail_compared_docx_json',
    name: 'EditSession.fail_compared_docx_json (no comparison pending)',
    throws: { name: 'string', reason: 'a save failure is refused while no comparison awaits its saved bytes' },
    call: ({ session: s }) => s.fail_compared_docx_json('the compared document could not be written'),
  },
  // No EditSession save exists, so the save goes through a host and the replica's own reads are compared too.
  defineRow({
    method: 'EditSession.delete_story',
    mutates: true,
    setup: ({ fresh, saveFrom }: Docx) => {
      const f = fresh();
      f.begin_opening(GENERATION);
      f.create_story('fn:1', 'Parity footnote', 'Normal', 'left');
      f.create_story('body', 'Created', 'Heading1', 'center');
      f.create_story('fn:2', 'Deleted footnote', 'Normal', 'left');
      saveFrom(f);
      return f;
    },
    call: (_, f) => ({ result: f.delete_story('fn:2'), reads: replicaReads(f) }),
  }),
  {
    method: 'EditSession.insert_watermark',
    mutates: true,
    call: ({ session: s }) => {
      const result = s.insert_watermark('body', paraId(s, 'Final'), 0, WATERMARK);
      return { result, watermarks: sessionWatermarks(s, WATERMARK) };
    },
  },
  // Save rewrites the paragraph without the watermark embed; when it is saved, this row fails and must change.
  defineRow({
    method: 'EditSession.insert_watermark',
    name: 'EditSession.insert_watermark (insert_watermark on the body is not saved)',
    setup: ({ fresh }: Docx) => {
      const f = fresh();
      f.begin_opening(GENERATION);
      f.create_story('body', 'Created', 'Heading1', 'center');
      return f;
    },
    call: async ({ saved }, f) => {
      const result = f.insert_watermark('body', paraId(f, 'Created'), 0, WATERMARK);
      const text = watermarkText(WATERMARK);
      const inSession = JSON.stringify(watermarkUnits(f)).includes(text);
      const body = await savedBody(await saved(f, fixture));
      const inSaved = body.includes(text);
      if (result !== undefined || !inSession || inSaved || !body.includes('Created')) {
        throw new Error(`expected an undefined return, the watermark in the session and none saved, with the paragraph saved; got ${JSON.stringify({ result, inSession, inSaved, paragraphSaved: body.includes('Created') })}`);
      }
      return { result, inSession, inSaved };
    },
  }),
  {
    method: 'EditSession.move_embed',
    mutates: true,
    call: ({ session: s }) => {
      const at = imageAt(s);
      return { receipt: s.move_embed('body', at.paraId, at.offset, 'body', paraId(s, 'Final'), 0), after: s.story_segments('body') };
    },
  },
  { method: 'EditSession.set_image_geometry', mutates: true, call: ({ session: s }) => s.set_image_geometry('rIdImage', GEOMETRY) },
  defineRow({
    method: 'EditSession.set_content_control_value_at',
    mutates: true,
    setup: ({ session: s }: Docx) => inlineControlAt(s, 1),
    call: ({ session: s }, at) => {
      const result = s.set_content_control_value_at('body', at.paraId, at.offset, JSON.stringify({ kind: 'checkbox', checked: true }));
      return { result, after: s.list_content_controls_json('{}') };
    },
  }),
  defineRow({
    method: 'EditSession.set_content_control_value_index',
    mutates: true,
    setup: ({ session: s }: Docx) => controlTarget(s, (control) => control.ooxmlId === '8'),
    call: ({ session: s }, at) => {
      const result = s.set_content_control_value_index(at.story, at.index, JSON.stringify('Ada Lovelace'));
      return { result, after: s.list_content_controls_json('{}') };
    },
  }),
  // No EditSession save exists, so the save goes through a host and the replica's own reads are compared too.
  defineRow({
    method: 'EditSession.repeat_content_control_at',
    mutates: true,
    setup: ({ fresh, saveFrom }: Docx) => {
      const f = fresh();
      f.open_docx(repeatingFixture, true, GENERATION);
      saveFrom(f, repeatingFixture);
      return { f, at: controlTarget(f, (control) => control.controlType === 'repeatingSectionItem') };
    },
    call: (_, { f, at }) => {
      const result = f.repeat_content_control_at(at.story, at.index, false);
      if (!result) throw new Error('the repeating item was not repeated');
      return { result, targets: f.content_control_targets_json(), reads: replicaReads(f) };
    },
  }),
  {
    method: 'EditSession.set_document_property',
    mutates: true,
    call: ({ session: s }) => ({ result: s.set_document_property('metadata', JSON.stringify({ category: 'Parity' })), after: s.document_properties_json() }),
  },
  {
    method: 'EditSession.set_document_property',
    name: 'EditSession.set_document_property (unknown key)',
    throws: { name: 'string', reason: 'only the metadata property is accepted' },
    call: ({ session: s }) => s.set_document_property('title', JSON.stringify({ category: 'Parity' })),
  },
  {
    method: 'EditSession.set_document_property',
    name: 'EditSession.set_document_property (not an object)',
    throws: { name: 'string', reason: 'document metadata that is not a JSON object is refused' },
    call: ({ session: s }) => s.set_document_property('metadata', JSON.stringify(['Parity'])),
  },
  {
    method: 'EditSession.set_style_language',
    mutates: true,
    call: ({ session: s }) => ({ result: s.set_style_language('Normal', JSON.stringify({ lang: 'de-DE' })), after: s.style_languages_json() }),
  },
  defineRow({
    method: 'EditSession.edit_text_resident',
    mutates: true,
    setup: ({ session: s }: Docx) => layout(s),
    call: ({ session: s }) => {
      const id = paraId(s, 'Second body');
      const at = textAt(s, 'Second body', 0);
      return { receipt: s.edit_text_resident('body', id, at, id, at, 'Resident '), after: s.paragraphs('body') };
    },
  }),
  {
    method: 'EditSession.set_part_metadata',
    mutates: true,
    call: ({ session: s }) => {
      const current = JSON.parse(s.part_metadata_json()) as { body?: Record<string, unknown> };
      const result = s.set_part_metadata(JSON.stringify({ ...current, body: { ...current.body, parity: true } }), '[]');
      return { result, after: s.part_metadata_json() };
    },
  },
  defineRow({
    method: 'EditSession.drain_update_event',
    mutates: true,
    setup: (docx: Docx) => pendingUpdate(docx),
    call: ({ session: s }, first) => {
      const local = drainEvents(s);
      if (local.length === 0) throw new Error('the observed local edit queued no update event');
      s.apply_update(first);
      return { local, resolved: drainEvents(s), drained: s.drain_update_event() };
    },
  }),
  defineRow({
    method: 'EditSession.has_pending_update',
    mutates: true,
    setup: (docx: Docx) => pendingUpdate(docx),
    call: ({ session: s }, first) => {
      const pending = s.has_pending_update();
      if (!pending) throw new Error('an update missing its predecessor is not pending');
      const events = drainEvents(s);
      if (events.length === 0) throw new Error('the observed local edit queued no update event');
      s.apply_update(first);
      return { pending, events, resolved: s.has_pending_update(), after: s.paragraphs('body') };
    },
  }),
  defineRow({
    method: 'EditSession.canvasSkippedEffects',
    setup: ({ session: s, fresh }: Docx) => {
      residentFrame(s);
      const f = fresh();
      f.open_docx(effectFixture, true, GENERATION);
      residentFrame(f);
      return f;
    },
    call: ({ session: s }, f) => {
      const effects = f.canvasSkippedEffects(0, new Uint32Array());
      if (effects === 0) throw new Error('the text effect page reports no skipped effect');
      return { plain: s.canvasSkippedEffects(0, new Uint32Array()), effects };
    },
  }),
  defineRow({
    method: 'EditSession.settle_revisions_json',
    mutates: true,
    setup: ({ session: s }: Docx) => trackedInsertion(s),
    call: ({ session: s }, revisionId) => {
      const result = s.settle_revisions_json(JSON.stringify({ accept: [revisionId] }));
      if (!(JSON.parse(result) as { revisionIds?: string[] }).revisionIds?.includes(revisionId)) throw new Error(`the tracked insertion was not settled: ${result}`);
      return { result, version: s.version(), after: s.list_revisions() };
    },
  }),
  {
    method: 'EditSession.settle_revisions_json',
    name: 'EditSession.settle_revisions_json (not JSON)',
    throws: { name: 'string', reason: 'a request that is not JSON is refused' },
    call: ({ session: s }) => s.settle_revisions_json('{'),
  },
  defineRow({
    method: 'EditSession.revision_stamps_json',
    setup: ({ session: s }: Docx) => trackedInsertion(s),
    call: ({ session: s }, revisionId) => {
      const stamps = s.revision_stamps_json(JSON.stringify([revisionId]));
      if (!(JSON.parse(stamps) as Record<string, unknown[] | undefined>)[revisionId]?.length) throw new Error(`the tracked insertion has no stamp: ${stamps}`);
      return stamps;
    },
  }),
  {
    method: 'EditSession.revision_stamps_json',
    name: 'EditSession.revision_stamps_json (not JSON)',
    throws: { name: 'string', reason: 'revision ids that are not JSON are refused' },
    call: ({ session: s }) => s.revision_stamps_json('{'),
  },
  defineRow({
    method: 'EditSession.revision_ids_at',
    setup: ({ session: s }: Docx) => {
      const id = paraId(s, 'First body');
      return { id, revisionId: trackedInsertion(s) };
    },
    call: ({ session: s }, { id, revisionId }) => {
      const ids = s.revision_ids_at('body', id, 3);
      if (!(JSON.parse(ids) as string[]).includes(revisionId)) throw new Error(`the tracked insertion does not hold the caret: ${ids}`);
      return ids;
    },
  }),
  defineRow({
    method: 'EditSession.proposal_revision_ranges_json',
    setup: ({ fresh }: Docx) => {
      const f = fresh();
      f.open_docx(compareOriginal, true, GENERATION);
      const inserted = f.insert_text('body', paraId(f, 'The quick'), 0, 'Tracked ', 'Parity', DATE);
      return { f, revisionId: (JSON.parse(inserted) as { revisionId: string }).revisionId };
    },
    call: (_, { f, revisionId }) => {
      const ranges = f.proposal_revision_ranges_json(JSON.stringify([revisionId]));
      const parsed = JSON.parse(ranges) as unknown;
      if (!Array.isArray(parsed) || !parsed.some((range: { revisionId?: unknown }) => range.revisionId === revisionId)) {
        throw new Error(`the tracked insertion has no range: ${ranges}`);
      }
      return ranges;
    },
  }),
  {
    method: 'EditSession.comment_text_target_json',
    mutates: true,
    call: ({ session: s }) => {
      const result = s.comment_text_target_json(alphaTarget(s), JSON.stringify({ id: '7', author: 'Parity', date: DATE, body: 'Targeted' }));
      if (!listedComment(s, '7')) throw new Error(`the target was not commented: ${result}`);
      return { result, after: s.list_comments() };
    },
  },
  {
    method: 'EditSession.format_text_target_json',
    mutates: true,
    call: ({ session: s }) => {
      const result = s.format_text_target_json(alphaTarget(s), JSON.stringify({ bold: true }));
      if ((JSON.parse(result) as { ok: boolean }).ok !== true) throw new Error(`the target was not formatted: ${result}`);
      return { result, after: s.story_segments('body') };
    },
  },
  defineRow({
    method: 'EditSession.seed_comment_metadata_json',
    mutates: true,
    setup: ({ session: s }: Docx) => seededCommentId(s),
    call: ({ session: s }, commentId) => {
      const result = s.seed_comment_metadata_json(commentId, JSON.stringify({ done: true }));
      if (listedComment(s, commentId)?.done !== true) throw new Error('the seeded comment was not marked done');
      return { result, after: s.list_comments() };
    },
  }),
  {
    method: 'EditSession.seed_comment_metadata_json',
    name: 'EditSession.seed_comment_metadata_json (unknown comment)',
    mutates: true,
    call: ({ session: s }) => {
      const result = s.seed_comment_metadata_json('9', JSON.stringify({ author: 'Parity', date: DATE }));
      if (!listedComment(s, '9')) throw new Error('seeding an unknown comment created none');
      return { result, after: s.list_comments() };
    },
  },
  defineRow({
    method: 'EditSession.unavailable_tables_from',
    setup: (docx: Docx) => crossedTableUndos(docx),
    call: ({ session: s }) => {
      const tables = s.unavailable_tables_from(JSON.stringify(['body', 'hf:rIdHeader']));
      const found = JSON.parse(tables) as { story: string; tableIndex: number }[] | null;
      if (!found?.some((table) => table.story === 'body' && table.tableIndex === 0)) throw new Error(`the emptied table is not listed: ${tables}`);
      return tables;
    },
  }),
  {
    method: 'EditSession.unavailable_tables_from',
    name: 'EditSession.unavailable_tables_from (not JSON)',
    throws: { name: 'string', reason: 'roots that are not a JSON array of story ids are refused' },
    call: ({ session: s }) => s.unavailable_tables_from('{'),
  },
  defineRow({
    method: 'EditSession.clear_cell_selection',
    setup: ({ session: s }: Docx) => s.set_cell_selection(cellRange([0, 0], [1, 1])),
    call: ({ session: s }) => {
      const before = s.cell_selection();
      const cleared = s.clear_cell_selection();
      if (cleared === undefined) throw new Error('clearing a cell selection returned no caret');
      return { before, cleared, after: s.cell_selection(), again: s.clear_cell_selection() };
    },
  }),
  // The handle's class differs between engines, so the row compares what its methods read.
  defineRow({
    method: 'EditSession.run_inheritance',
    setup: ({ engine }: Docx) => gridStyled(engine),
    call: ({ session: s }, { styles, table }) => {
      const inheritance = s.run_inheritance(styles);
      try {
        const attributes = inheritance.attributes('Heading1', null, null);
        if (Object.keys(JSON.parse(attributes) as object).length === 0) throw new Error('a Heading1 run inherits nothing');
        const tableAttributes = inheritance.table_attributes(table);
        gridCell(tableAttributes);
        return { attributes, paragraph: inheritance.paragraph_formatting('Heading1', 'null'), tableAttributes };
      } finally {
        inheritance.free();
      }
    },
  }),
  {
    method: 'EditSession.run_inheritance',
    name: 'EditSession.run_inheritance (not JSON)',
    throws: { name: 'string', reason: 'style definitions that are not JSON are refused' },
    call: ({ session: s }) => s.run_inheritance('{'),
  },
  defineRow({
    method: 'RunInheritance.attributes',
    setup: (docx: Docx) => gridInheritance(docx),
    call: (_, { inheritance, cell }) => {
      try {
        const heading = inheritance.attributes('Heading1', null, null);
        if (Object.keys(JSON.parse(heading) as object).length === 0) throw new Error('a Heading1 run inherits nothing');
        return [heading, inheritance.attributes(null, null, null), inheritance.attributes('Heading1', null, JSON.stringify(cell.runFormatting ?? {}))];
      } finally {
        inheritance.free();
      }
    },
  }),
  defineRow({
    method: 'RunInheritance.attributes',
    name: 'RunInheritance.attributes (table run not JSON)',
    throws: { name: 'string', reason: 'table run formatting that is not JSON is refused' },
    setup: ({ session: s }: Docx) => s.run_inheritance('null'),
    call: (_, inheritance) => {
      try {
        return inheritance.attributes(null, null, '{');
      } finally {
        inheritance.free();
      }
    },
  }),
  defineRow({
    method: 'RunInheritance.paragraph_formatting',
    setup: (docx: Docx) => gridInheritance(docx),
    call: (_, { inheritance, cell }) => {
      try {
        const inCell = inheritance.paragraph_formatting(null, JSON.stringify(cell.paragraphFormatting));
        if (JSON.parse(inCell) === null) throw new Error('a first-row cell paragraph inherits nothing');
        return [inCell, inheritance.paragraph_formatting('Heading1', JSON.stringify(cell.paragraphFormatting)), inheritance.paragraph_formatting(null, 'null')];
      } finally {
        inheritance.free();
      }
    },
  }),
  defineRow({
    method: 'RunInheritance.paragraph_formatting',
    name: 'RunInheritance.paragraph_formatting (not JSON)',
    throws: { name: 'string', reason: 'table paragraph formatting that is not JSON is refused' },
    setup: ({ session: s }: Docx) => s.run_inheritance('null'),
    call: (_, inheritance) => {
      try {
        return inheritance.paragraph_formatting(null, '{');
      } finally {
        inheritance.free();
      }
    },
  }),
  defineRow({
    method: 'RunInheritance.table_attributes',
    setup: (docx: Docx) => gridInheritance(docx),
    call: (_, { inheritance, table }) => {
      try {
        const attributes = inheritance.table_attributes(table);
        gridCell(attributes);
        return attributes;
      } finally {
        inheritance.free();
      }
    },
  }),
  defineRow({
    method: 'RunInheritance.table_attributes',
    name: 'RunInheritance.table_attributes (not JSON)',
    throws: { name: 'string', reason: 'a table that is not JSON is refused' },
    setup: ({ session: s }: Docx) => s.run_inheritance('null'),
    call: (_, inheritance) => {
      try {
        return inheritance.table_attributes('{');
      } finally {
        inheritance.free();
      }
    },
  }),
  defineRow({
    method: 'EditSession.style_languages_json',
    setup: ({ session: s }: Docx) => {
      s.set_style_language('Normal', JSON.stringify({ lang: 'de-DE' }));
      s.set_style_language('Heading1', JSON.stringify({ eastAsia: 'ja-JP' }));
    },
    call: ({ session: s }) => {
      const languages = s.style_languages_json();
      if ((JSON.parse(languages) as Record<string, { lang?: string }>).Normal?.lang !== 'de-DE') throw new Error(`the Normal language is not listed: ${languages}`);
      return languages;
    },
  }),
  defineRow({
    method: 'EditSession.header_footer_aliases_json',
    setup: ({ session: s }: Docx) => s.set_header_footer_aliases(HEADER_ALIASES),
    call: ({ session: s }) => {
      const aliases = s.header_footer_aliases_json();
      if ((JSON.parse(aliases) as Record<string, string>).rIdHeaderAlias !== 'rIdHeader') throw new Error(`the header alias is not active: ${aliases}`);
      return aliases;
    },
  }),
  // The key is drawn per opening and keys every digest, so the row compares their form, the part indices and a reload, and requires a second opening to draw another digest.
  {
    method: 'EditSession.media_sources_json',
    call: ({ engine, fresh }) => {
      const f = fresh();
      const unseeded = f.media_sources_json();
      f.open_docx(fixture, true, GENERATION);
      const json = f.media_sources_json();
      const sources = JSON.parse(json) as MediaSourcesWire;
      if (sources.parts.length === 0) throw new Error(`the seeded image left no media source: ${json}`);
      const again = fresh();
      again.open_docx(fixture, true, GENERATION);
      const reopened = JSON.parse(again.media_sources_json()) as MediaSourcesWire;
      const [first, second] = [sources.parts[0]![0], reopened.parts[0]?.[0]];
      console.log(`MEDIA-DIGEST-PROBE ${engineName(engine)} ${first} ${String(second)} same=${first === second}`);
      if (first === second) throw new Error(`two openings drew the same media digest ${first}`);
      const shape = mediaSourcesShape(sources);
      if (JSON.stringify(mediaSourcesShape(reopened)) !== JSON.stringify(shape)) throw new Error(`a second opening changed the media sources beyond their digests: ${json} then ${again.media_sources_json()}`);
      const reader = fresh();
      reader.load_media_sources(json);
      if (reader.media_sources_json() !== json) throw new Error(`the loaded media sources read back as ${reader.media_sources_json()}, not ${json}`);
      return {
        unseeded,
        ...shape,
        reopened: mediaSourcesShape(reopened),
        reloaded: reader.media_sources_json() === json,
      };
    },
  },
  // Sources only carry a key and digests, so each engine lays the other's seeded images out as the same tokens and resolves them to the same part.
  {
    method: 'EditSession.load_media_sources',
    call: ({ engine, fresh }) => {
      const other = engine === wasmEngine ? nativeEngine : wasmEngine;
      const seeder = new other.EditSession(HOST_CLIENT);
      try {
        seeder.open_docx(fixture, true, GENERATION);
        const json = seeder.media_sources_json();
        if (!json) throw new Error('the other engine seeded no media source');
        const reader = fresh();
        reader.open_docx(fixture, false, GENERATION);
        reader.apply_update(seeder.encode_state());
        const { input } = layoutInput(reader);
        const tokensInput = JSON.stringify({ ...JSON.parse(input), renderEnv: { mediaTokens: true } });
        const before = imageSources(reader, tokensInput);
        if (!before.every((src) => src.startsWith('data:image'))) throw new Error(`the reader tokenized images before loading sources: ${JSON.stringify(before)}`);
        const loaded = reader.load_media_sources(json);
        const tokens = imageSources(reader, tokensInput);
        if (!tokens.every((src) => src.startsWith('media:'))) throw new Error(`the other engine's sources left images untokenized: ${JSON.stringify(tokens)}`);
        const media = tokens.map((token) => {
          const bytes = reader.media_bytes(token);
          const seeded = seeder.media_bytes(token);
          if (!bytes || !seeded || Buffer.compare(Buffer.from(bytes), Buffer.from(seeded)) !== 0) throw new Error(`${token} resolves to other bytes than on the seeding engine`);
          if (reader.media_type(token) !== seeder.media_type(token)) throw new Error(`${token} has another media type than on the seeding engine`);
          if (Buffer.compare(Buffer.from(bytes), Buffer.from(PNG)) !== 0) throw new Error(`${token} does not resolve to the fixture PNG`);
          return { token, type: reader.media_type(token), bytes, png: Buffer.compare(Buffer.from(bytes), Buffer.from(PNG)) === 0 };
        });
        const reloaded = reader.media_sources_json();
        if (reloaded !== json) throw new Error(`the loaded media sources read back as ${reloaded}, not ${json}`);
        return { loaded, before, media };
      } finally {
        seeder.free();
      }
    },
  },
  defineRow({
    method: 'EditSession.media_type',
    setup: ({ fresh }: Docx) => tokenizedImage(fresh),
    call: (_, { f, src }) => {
      const type = f.media_type(src);
      if (!type?.startsWith('image/')) throw new Error(`the seeded image has no media type: ${String(type)}`);
      return [type, f.media_type('media:missing')];
    },
  }),
  defineRow({
    method: 'EditSession.media_data_url',
    setup: ({ fresh }: Docx) => tokenizedImage(fresh),
    call: ({ fresh }, { f, src }) => {
      const url = f.media_data_url(src);
      if (url !== `data:image/png;base64,${Buffer.from(PNG).toString('base64')}`) throw new Error(`${src} stands for ${String(url)}, not the fixture PNG`);
      const untokenized = fresh();
      untokenized.open_docx(fixture, true, GENERATION);
      if (imageAt(untokenized).src !== url) throw new Error(`an untokenized opening seeds ${String(imageAt(untokenized).src)}, not ${url}`);
      return [url, f.media_data_url('media:missing'), f.media_data_url(url), fresh().media_data_url(src)];
    },
  }),
  {
    method: 'EditSession.set_media_tokens',
    call: ({ fresh }) => {
      const opened = (...settings: boolean[]) => {
        const f = fresh();
        for (const enabled of settings) f.set_media_tokens(enabled);
        f.open_docx(fixture, true, GENERATION);
        return seededImage(f);
      };
      const [unset, on, off, toggled] = [opened(), opened(true), opened(false), opened(true, false)];
      if (unset === on || !on.includes('"media:')) throw new Error(`enabling media tokens did not tokenize the seeded image: ${unset} then ${on}`);
      return { unset, on, off, toggled };
    },
  },
  {
    method: 'EditSession.rebuild_media_sources',
    call: ({ engine, fresh }) => {
      const bare = fresh().rebuild_media_sources();
      const seeded = fresh();
      seeded.open_docx(fixture, true, GENERATION);
      const kept = seeded.media_sources_json();
      const seededRebuild = seeded.rebuild_media_sources();
      if (seeded.media_sources_json() !== kept) throw new Error(`rebuilding a seeded replica changed its sources: ${kept} then ${seeded.media_sources_json()}`);
      const replica = mediaReplica(fresh, seeded);
      const before = replica.media_sources_json();
      const rebuilt = replica.rebuild_media_sources();
      const json = replica.media_sources_json();
      const again = replica.rebuild_media_sources();
      if (replica.media_sources_json() !== json) throw new Error(`a second rebuild changed the sources: ${json} then ${replica.media_sources_json()}`);
      const other = mediaReplica(fresh, seeded);
      other.rebuild_media_sources();
      const sources = JSON.parse(json) as MediaSourcesWire;
      const reopened = JSON.parse(other.media_sources_json()) as MediaSourcesWire;
      const [first, second] = [sources.parts[0]?.[0], reopened.parts[0]?.[0]];
      console.log(`MEDIA-REBUILD-PROBE ${engineName(engine)} ${String(first)} ${String(second)} same=${first === second}`);
      if (first === undefined || first === second) throw new Error(`two rebuilds drew the same media digest ${String(first)}`);
      return {
        bare,
        seededRebuild,
        before,
        rebuilt,
        again,
        ...mediaSourcesShape(sources),
        other: mediaSourcesShape(reopened),
        tokens: tokenSources(replica),
      };
    },
  },
  {
    method: 'EditSession.rebuild_media_sources',
    name: 'EditSession.rebuild_media_sources (seeded on the other engine)',
    call: ({ engine, fresh }) => {
      const seeder = new (engine === wasmEngine ? nativeEngine : wasmEngine).EditSession(HOST_CLIENT);
      try {
        seeder.open_docx(fixture, true, GENERATION);
        const reader = mediaReplica(fresh, seeder);
        const rebuilt = reader.rebuild_media_sources();
        const tokens = tokenSources(reader);
        const media = tokens.map((token) => {
          const [bytes, seeded] = [reader.media_bytes(token), seeder.media_bytes(token)];
          if (!bytes || !seeded || Buffer.compare(Buffer.from(bytes), Buffer.from(seeded)) !== 0) throw new Error(`${token} resolves to other bytes than on the seeding engine`);
          return { token, type: reader.media_type(token), url: reader.media_data_url(token), seededUrl: seeder.media_data_url(token), bytes };
        });
        return { rebuilt, tokens, media };
      } finally {
        seeder.free();
      }
    },
  },
  {
    method: 'EditSession.export_snapshot_with_private_fonts_json',
    call: ({ engine, fresh }) => {
      const f = embeddedFontSession(fresh);
      const face = embeddedFace(engine);
      const embedded = privateFontRequest(engine, f, [0]);
      const snapshot = f.export_snapshot_with_private_fonts_json(face, Uint32Array.from([face.length]), embedded.request, EXPORT);
      if (!(JSON.parse(snapshot) as { ok: boolean }).ok) throw new Error(`the private font export refused: ${snapshot}`);
      const unfonted = f.export_snapshot_with_private_fonts_json(new Uint8Array(0), new Uint32Array(0), privateFontRequest(engine, f, []).request, EXPORT);
      return { requirements: embedded.requirements, snapshot, unfonted };
    },
  },
  defineRow({
    method: 'EditSession.export_snapshot_with_private_fonts_json',
    name: 'EditSession.export_snapshot_with_private_fonts_json (lengths past the font bytes)',
    throws: { name: 'string', reason: 'font lengths that run past the joined font bytes are refused' },
    setup: ({ engine, fresh }: Docx) => {
      const f = embeddedFontSession(fresh);
      return { f, face: embeddedFace(engine), request: privateFontRequest(engine, f, [0]).request };
    },
    call: (_, { f, face, request }) => f.export_snapshot_with_private_fonts_json(face, Uint32Array.from([face.length + 1]), request, EXPORT),
  }),
  {
    method: 'EditSession.load',
    mutates: true,
    call: ({ session: s, fresh, saveFrom }) => {
      const f = fresh();
      f.open_docx(fixture, false, GENERATION);
      const result = f.load(s.encode_state());
      saveFrom(f);
      return { result, stories: f.list_stories_json(), segments: f.story_segments('body'), comments: f.list_comments() };
    },
  },
  {
    method: 'EditSession.load',
    name: 'EditSession.load (malformed update)',
    throws: { name: 'string', reason: 'an update that does not decode is refused' },
    call: ({ fresh }) => fresh().load(MALFORMED_UPDATE),
  },
  {
    method: 'EditSession.note_separators_state',
    call: ({ session: s, fresh }) => ({ session: s.note_separators_state(), seeded: seededSeparators(fresh), bare: fresh().note_separators_state() }),
  },
  defineRow({
    method: 'EditSession.load_note_separators',
    setup: ({ fresh }: Docx) => seededSeparators(fresh),
    call: ({ fresh }, state) => {
      const r = fresh();
      const before = r.note_separators_state();
      const result = r.load_note_separators(state);
      const loaded = r.note_separators_state();
      const again = r.load_note_separators(state);
      const cleared = r.load_note_separators(new Uint8Array(0));
      return { before, result, loaded, again, cleared, after: r.note_separators_state() };
    },
  }),
  {
    method: 'EditSession.load_note_separators',
    name: 'EditSession.load_note_separators (malformed state)',
    throws: { name: 'string', reason: 'separator state that does not decode as an update is refused' },
    call: ({ fresh }) => fresh().load_note_separators(MALFORMED_UPDATE),
  },
  {
    method: 'EditSession.set_partial_document',
    call: ({ fresh }) => {
      const f = fresh();
      f.open_docx(pageTotalFixture, true, GENERATION);
      const read = () => ({ layout: layout(f).layout, texts: f.page_texts_json(Uint32Array.from([0])) });
      const flagged = (json: string) => (JSON.parse(json) as { layout: { partial?: boolean } }).layout.partial;
      const whole = read();
      const marked = f.set_partial_document(true);
      const partial = read();
      const unmarked = f.set_partial_document(false);
      const restored = read();
      if (flagged(whole.layout) !== undefined || flagged(partial.layout) !== true) throw new Error('the layout does not carry the partial flag as set');
      if (!/Page 1 of \d/.test(whole.texts) || /Page 1 of \d/.test(partial.texts) || !partial.texts.includes('Page 1 of ')) {
        throw new Error(`NUMPAGES does not go blank while the document is partial: ${whole.texts} then ${partial.texts}`);
      }
      if (restored.texts !== whole.texts) throw new Error(`clearing the flag did not restore the page total: ${restored.texts}`);
      return { whole, marked, partial, unmarked, restored };
    },
  },
  {
    method: 'EditSession.selection',
    mutates: true,
    call: ({ session: s }) => {
      const before = s.selection();
      s.set_selection(...textSpan(s, 'Second body', 0, 6));
      const set = s.selection();
      peerEdit(s, 'Second body', 'Moved ');
      return { before, set, after: s.selection() };
    },
  },
  {
    method: 'EditSession.encoded_selection',
    call: ({ session: s }) => {
      const before = s.encoded_selection();
      s.set_selection(...textSpan(s, 'Second body', 0, 6));
      return { before, after: s.encoded_selection() };
    },
  },
  defineRow({
    method: 'EditSession.resolve_encoded_selection',
    mutates: true,
    setup: ({ session: s }: Docx) => {
      s.set_selection(...textSpan(s, 'Second body', 0, 6));
      return encodedSelection(s);
    },
    call: ({ session: s }, { story, anchor, head }) => {
      const before = s.resolve_encoded_selection(story, anchor, head);
      peerEdit(s, 'Second body', 'Moved ');
      return { before, after: s.resolve_encoded_selection(story, anchor, head) };
    },
  }),
  {
    method: 'EditSession.resolve_encoded_selection',
    name: 'EditSession.resolve_encoded_selection (malformed positions)',
    throws: { name: 'string', reason: 'sticky positions that do not decode are refused' },
    call: ({ session: s }) => s.resolve_encoded_selection('body', MALFORMED_UPDATE, MALFORMED_UPDATE),
  },
  {
    method: 'EditSession.resolve_encoded_selection',
    name: 'EditSession.resolve_encoded_selection (encoded on the other engine)',
    call: ({ engine, session: s }) => {
      const other = otherEngineReplica(engine, s);
      try {
        const range = textSpan(s, 'Second body', 0, 6);
        s.set_selection(...range);
        other.set_selection(...range);
        const ours = encodedSelection(s);
        const theirs = encodedSelection(other);
        return {
          ours,
          theirs,
          oursThere: other.resolve_encoded_selection(ours.story, ours.anchor, ours.head),
          theirsHere: s.resolve_encoded_selection(theirs.story, theirs.anchor, theirs.head),
        };
      } finally {
        other.free();
      }
    },
  },
  {
    method: 'EditSession.encode_sticky_position',
    call: ({ session: s }) => [s.encode_sticky_position(...stickyAt(s, 'Second body', 0)), s.encode_sticky_position(...stickyAt(s, 'Second body', 6))],
  },
  {
    method: 'EditSession.encode_sticky_position',
    name: 'EditSession.encode_sticky_position (unknown paragraph)',
    throws: { name: 'string', reason: 'a paragraph id the story does not hold is refused' },
    call: ({ session: s }) => s.encode_sticky_position('body', '7FFFFFFF', 0),
  },
  defineRow({
    method: 'EditSession.resolve_sticky_position',
    mutates: true,
    setup: ({ session: s }: Docx) => s.encode_sticky_position(...stickyAt(s, 'Second body', 6)),
    call: ({ session: s }, position) => {
      const before = s.resolve_sticky_position('body', position);
      peerEdit(s, 'Second body', 'Moved ');
      return { before, after: s.resolve_sticky_position('body', position) };
    },
  }),
  {
    method: 'EditSession.resolve_sticky_position',
    name: 'EditSession.resolve_sticky_position (malformed position)',
    throws: { name: 'string', reason: 'a sticky position that does not decode is refused' },
    call: ({ session: s }) => s.resolve_sticky_position('body', MALFORMED_UPDATE),
  },
  {
    method: 'EditSession.resolve_sticky_position',
    name: 'EditSession.resolve_sticky_position (encoded on the other engine)',
    call: ({ engine, session: s }) => {
      const other = otherEngineReplica(engine, s);
      try {
        const at = stickyAt(s, 'Second body', 6);
        const ours = s.encode_sticky_position(...at);
        const theirs = other.encode_sticky_position(...at);
        return { ours, theirs, oursThere: other.resolve_sticky_position('body', ours), theirsHere: s.resolve_sticky_position('body', theirs) };
      } finally {
        other.free();
      }
    },
  },
  {
    method: 'EditSession.resolve_paragraph_anchor',
    call: ({ session: s }) => paragraphAnchors(s).map((anchor) => ({ anchor, resolved: s.resolve_paragraph_anchor(JSON.stringify(anchor)) })),
  },
  {
    method: 'EditSession.resolve_paragraph_anchor',
    name: 'EditSession.resolve_paragraph_anchor (unknown kind)',
    throws: { name: 'string', reason: 'an anchor of an unknown kind is refused' },
    call: ({ session: s }) => s.resolve_paragraph_anchor(JSON.stringify({ kind: 'bogus' })),
  },
  {
    method: 'EditSession.select_story',
    mutates: true,
    call: ({ session: s }) => timedEdits(s, () => {
      s.select_story(MISSING_STORY);
      return s.select_story('body');
    }),
  },
  {
    method: 'EditSession.next_paragraph_style',
    call: ({ session: s }) => ['Parity heading', 'First body'].map((prefix) => ({
      end: s.next_paragraph_style(...paragraphEnd(s, prefix)),
      start: s.next_paragraph_style('body', paraId(s, prefix), 0),
    })),
  },
  {
    method: 'EditSession.next_paragraph_style',
    name: 'EditSession.next_paragraph_style (unknown paragraph)',
    throws: { name: 'string', reason: 'a paragraph id the story does not hold is refused' },
    call: ({ session: s }) => s.next_paragraph_style('body', '7FFFFFFF', 0),
  },
  {
    method: 'EditSession.paragraph_unit',
    call: ({ session: s }) => [...['Parity heading', 'First body', 'Final'].map((prefix) => paraId(s, prefix)), '7FFFFFFF'].map((id) => s.paragraph_unit('body', id)),
  },
  {
    method: 'EditSession.paragraph_unit',
    name: 'EditSession.paragraph_unit (unknown story)',
    throws: { name: 'string', reason: 'a story the document does not hold is refused' },
    call: ({ session: s }) => s.paragraph_unit(MISSING_STORY, paraId(s, 'First body')),
  },
  {
    method: 'EditSession.paragraph_id_count',
    call: ({ session: s }) => [...['Parity heading', 'First body', 'Final'].map((prefix) => paraId(s, prefix)), '7FFFFFFF'].map((id) => s.paragraph_id_count('body', id)),
  },
  {
    method: 'EditSession.paragraph_id_count',
    name: 'EditSession.paragraph_id_count (unknown story)',
    throws: { name: 'string', reason: 'a story the document does not hold is refused' },
    call: ({ session: s }) => s.paragraph_id_count(MISSING_STORY, paraId(s, 'First body')),
  },
  { method: 'EditSession.story_segment_units', call: ({ session: s }) => s.story_segment_units('body', Uint32Array.from([0, 1, 6])) },
  {
    method: 'EditSession.story_segment_units',
    name: 'EditSession.story_segment_units (unit past the end)',
    throws: { name: 'string', reason: 'a unit index past the last unit is refused' },
    call: ({ session: s }) => s.story_segment_units('body', Uint32Array.from([9_999])),
  },
  {
    method: 'EditSession.resident_caret_snapshot_json',
    call: ({ session: s }) => {
      const bare = s.resident_caret_snapshot_json();
      residentFrame(s);
      const framed = s.resident_caret_snapshot_json();
      caretIn(s, 'Second body', 6);
      const caret = s.resident_caret_snapshot_json();
      if ((JSON.parse(caret) as { caretRect: unknown }).caretRect === null) throw new Error(`the collapsed caret has no rect: ${caret}`);
      s.set_selection(...textSpan(s, 'Second body', 0, 6));
      return { bare, framed, caret, range: s.resident_caret_snapshot_json() };
    },
  },
];

const HEAP = 'heap accounting of the calling thread; the native and wasm allocators differ by design';
const OBSERVER = 'a JS callback crosses the napi and wasm-bindgen boundaries differently; what the observer delivers is compared through the drain_update_event and has_pending_update rows';
const exclusions: Exclusion[] = [
  { method: 'set_wasm_heap_limit', reason: HEAP },
  { method: 'set_wasm_heap_counting', reason: HEAP },
  { method: 'wasm_heap_counted', reason: HEAP },
  { method: 'wasm_live_bytes', reason: HEAP },
  { method: 'wasm_peak_bytes', reason: HEAP },
  { method: 'reset_wasm_peak_bytes', reason: HEAP },
  { method: 'wasm_failed_allocation_bytes', reason: HEAP },
  { method: 'install_panic_hook', reason: 'process-wide panic logging with no document result' },
  { method: 'EditSession.apply_input_profiled', reason: 'profiling twin of apply_input; only its timings differ' },
  { method: 'EditSession.apply_delete_profiled', reason: 'profiling twin of apply_delete; only its timings differ' },
  { method: 'EditSession.apply_input_profile_json', reason: 'wall-clock stage timings, different on every run' },
  { method: 'EditSession.free', reason: 'releases the object; the harness frees every session itself' },
  { method: 'RetainedLayoutMeta.free', reason: 'releases the object; nothing to compare' },
  { method: 'RunInheritance.free', reason: 'releases the object; nothing to compare' },
  { method: 'EditSession.set_update_observer', reason: OBSERVER },
  { method: 'EditSession.clear_update_observer', reason: OBSERVER },
  { method: 'EditSession.take_base_latin_miss', reason: 'a read-and-reset font-cache flag, not a document result' },
];

const knownDifferences: KnownDifference[] = [];

const notYetCovered: string[] = [
  'EditSession.apply_host_update',
  'EditSession.apply_seed_raw_ops',
  'EditSession.apply_update_with_inference', 'EditSession.begin_region_layout', 'EditSession.begin_shared_reads',
  'EditSession.bootstrap_peer',
  'EditSession.clear_measure_fonts', 'EditSession.clear_update_event_observation',
  'EditSession.decode_font',
  'EditSession.direct_batches_applied', 'EditSession.display_range_rects_batch',
  'EditSession.display_range_rects_on_pages_json', 'EditSession.display_range_rects_region_json',
  'EditSession.echo_glyph_set_json', 'EditSession.encode_diff_chunks',
  'EditSession.encode_peer_metadata',
  'EditSession.end_shared_reads',
  'EditSession.geometry_position_outline_json',
  'EditSession.layout_document_with_regions_exact_prefix_retained',
  'EditSession.layout_document_with_regions_prefix_retained_json',
  'EditSession.measure_paragraph_json',
  'EditSession.outline_glyph_json',
  'EditSession.patch_style_languages_xml',
  'EditSession.register_substitute_measure_font', 'EditSession.release_display_pages_frame', 'EditSession.reset_frame_base',
  'EditSession.resident_deleted_units', 'EditSession.resident_geometry_json', 'EditSession.resident_section_revision',
  'EditSession.resume_region_layout',
  'EditSession.retained_kernel_inputs_json', 'EditSession.retained_layout_json',
  'EditSession.seed_opaque_sequences',
  'EditSession.set_direct_batches', 'EditSession.set_display_retain_built_pages',
  'EditSession.set_header_footer_aliases',
  'EditSession.set_local_lowering',
  'EditSession.set_windowed_incremental_builds',
  'EditSession.start_update_event_observation',
  'EditSession.story_segment_unit_digests',
  'EditSession.update_measure_font_outlines', 'EditSession.yrs_blocks_for_story',
  'RetainedLayoutMeta.layout_shell_json', 'RetainedLayoutMeta.notes_converged', 'RetainedLayoutMeta.provisional',
  'clear_measure_fonts', 'close_display_list', 'decodeFontFace', 'decodeTiffPng',
  'hit_test_regions_by_handle',
  'outline_glyph_json', 'parse_relationships_xml', 'range_rects_on_pages_by_handle', 'range_rects_on_pages_json',
  'range_rects_region_by_handle', 'range_rects_region_json', 'register_measure_font', 'register_substitute_measure_font',
  'update_display_list', 'vertical_move_by_handle',
];

const NO_RASTERIZER = 'the wasm build carries no rasterizer; wasm draws PNG targets through the canvas in packages/docx/src/wasm/enginePng.ts';
const oneEngineOnly: OneEngineOnly[] = [
  { name: 'EditSession.renderPngTarget', side: 'native', reason: 'BY_DESIGN', note: NO_RASTERIZER },
  { name: 'EditSession.renderPngTargetReport', side: 'native', reason: 'BY_DESIGN', note: NO_RASTERIZER },
  { name: '__napiBindingTarget', side: 'native', reason: 'BUILD_METADATA', note: 'napi-rs build target marker' },
];

function docxSurfaces() {
  const read = (path: string) => parseSurface(readFileSync(new URL(`../../${path}`, import.meta.url), 'utf8'));
  const wasm = Object.entries(FORMATS).filter(([name]) => name === 'docx' || name.startsWith('docx-')).map(([, format]) => read(format.wasm));
  const native = read(FORMATS.docx!.native);
  return { surface: surfaceMethods(wasm, native), oneEngine: oneEngineExports(wasm, native) };
}

describe('docx native and wasm parity table', () => {
  test('every method on both engines is a row, an exclusion or not yet covered', () => {
    const { surface, oneEngine } = docxSurfaces();
    const plan = { rows, exclusions, notYetCovered, knownDifferences, oneEngineOnly };
    printParityCounts('docx', surface, plan);
    expect(coverageProblems(surface, plan)).toEqual([]);
    expect(oneEngineProblems(oneEngine, oneEngineOnly)).toEqual([]);
  });

  test('the yrs host facade exports the same names and kinds on both engines', () => {
    expect(facadeShape(nativeYrs)).toEqual(facadeShape(wasmYrs));
  });

  test('the not yet covered list only shrinks', () => {
    expect(notYetCovered.length).toBe(60);
  });

  for (const row of rows) {
    const known = knownDifferences.find((entry) => entry.method === row.method);
    test(known ? `${rowName(row)} still differs: ${known.reason}` : `${rowName(row)} matches`, async () => {
      await (known ? runKnownDifference : runRow)(row, makeWasm, makeNative);
    }, 30_000);
  }
});
