import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import initWasm, * as wasmEdit from '../../packages/docx/src/wasm/generated/edit/docx_edit.js';
import initPreflight, { font_preflight as wasmFontPreflight } from '../../packages/docx/src/wasm/generated/preflight/docx_preflight.js';
import * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeYrs from '@betteroffice/docx-native/yrs';
import { readTableFragment, type ClipboardTableReader } from '../../packages/docx/src/yrs/clipboardFragment';
import type { YrsStorySegment } from '../../packages/docx/src/yrs/index';
import { FORMATS, parseSurface } from '../../scripts/check-engine-surface';
import {
  coverageProblems,
  defineRow,
  FramedZipReturn,
  rowName,
  runKnownDifference,
  runRow,
  surfaceMethods,
  zipEntries,
  ZipReturn,
  type EngineHandle,
  type Exclusion,
  type KnownDifference,
  type Row,
} from './harness';

type DocxEngine = typeof wasmEdit & { font_preflight: typeof wasmFontPreflight };
type Session = wasmEdit.EditSession;
type Yrs = typeof wasmYrs;

interface Docx {
  engine: DocxEngine;
  session: Session;
  fresh(): Session;
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

async function buildFixture(): Promise<Uint8Array> {
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
  };
  for (const [name, content] of Object.entries(parts)) zip.file(name, content, { date });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const fontBytes = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
const wasmEngine: DocxEngine = { ...wasmEdit, font_preflight: wasmFontPreflight };
const nativeEngine = createRequire(import.meta.url)('../docx-native/index.js') as DocxEngine;
let fixture: Uint8Array;
let fixtureEntries: Record<string, Uint8Array>;

beforeAll(async () => {
  await initWasm({ module_or_path: readFileSync(new URL('../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)) });
  await initPreflight({ module_or_path: readFileSync(new URL('../../packages/docx/src/wasm/generated/preflight/docx_preflight_bg.wasm', import.meta.url)) });
  fixture = await buildFixture();
  fixtureEntries = Object.fromEntries((await zipEntries(fixture)).parts);
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
    try {
      host.seedFromDocx(fixture, { generation: GENERATION });
      const session = open(EDITOR_CLIENT);
      session.set_media_tokens(false);
      session.open_docx(fixture, false, GENERATION);
      session.apply_update(host.encodeState());
      return {
        target: { engine, session, fresh: () => open(FRESH_CLIENT) },
        generated: () => [versionNonce(session)],
        save: async () => {
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
  const request = {
    bodyStory: 'body',
    regions: { sections: [{
      sectionId: 'main',
      properties: { pageWidth: 12240, pageHeight: 15840, marginTop: 1440, marginRight: 1440, marginBottom: 1440, marginLeft: 1440 },
    }] },
    notes: { contents: [{ id: 1, noteKind: 'footnote', height: 0 }] },
    measurement: { fontChains: {} as Record<string, number[]>, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true },
    renderEnv: {},
  };
  return { font, request };
}

function layoutInput(session: Session) {
  const { font, request } = layoutRequest(session);
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

const PARSES = ['parse_docx_s2', 'parse_docx_s3', 'parse_docx_s4', 'parse_docx_s5', 'parse_docx_s6', 'parse_docx_s7', 'parse_docx_s8'] as const;

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
];

const HEAP = 'heap accounting of the calling thread; the native and wasm allocators differ by design';
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
];

const knownDifferences: KnownDifference[] = [];

const notYetCovered: string[] = [
  'EditSession.apply_delete', 'EditSession.apply_host_update', 'EditSession.apply_input',
  'EditSession.apply_local_update', 'EditSession.apply_raw_ops', 'EditSession.apply_seed_raw_ops', 'EditSession.apply_update_chunks',
  'EditSession.apply_update', 'EditSession.apply_update_with_inference', 'EditSession.begin_region_layout', 'EditSession.begin_shared_reads',
  'EditSession.bootstrap_peer', 'EditSession.build_display_list_frame', 'EditSession.build_display_pages_frame',
  'EditSession.canvasSkippedEffects', 'EditSession.cell_selection', 'EditSession.clear_cell_selection',
  'EditSession.clear_measure_fonts', 'EditSession.clear_undo_input_time', 'EditSession.clear_update_event_observation', 'EditSession.clear_update_observer',
  'EditSession.comment_text_target_json', 'EditSession.compare_docx_json', 'EditSession.create_story', 'EditSession.decode_font',
  'EditSession.delete_story',
  'EditSession.direct_batches_applied', 'EditSession.display_hit_test_regions_json', 'EditSession.display_range_rects_batch', 'EditSession.display_range_rects_json',
  'EditSession.display_range_rects_on_pages_json', 'EditSession.display_range_rects_region_json', 'EditSession.display_vertical_move_json', 'EditSession.drain_update_event',
  'EditSession.echo_glyph_set_json', 'EditSession.edit_text_resident', 'EditSession.encode_diff_chunks', 'EditSession.encode_diff',
  'EditSession.encode_peer_metadata', 'EditSession.encode_state', 'EditSession.encode_state_vector', 'EditSession.encode_sticky_position',
  'EditSession.encoded_selection', 'EditSession.end_shared_reads', 'EditSession.export_snapshot_with_private_fonts_json', 'EditSession.export_structured_with_pages_json',
  'EditSession.fail_compared_docx_json', 'EditSession.finish_compared_docx_json', 'EditSession.format_text_target_json', 'EditSession.geometry_position_outline_json',
  'EditSession.has_pending_update', 'EditSession.header_footer_aliases_json',
  'EditSession.insert_watermark', 'EditSession.layout_document_json',
  'EditSession.layout_document_with_regions_after_edit', 'EditSession.layout_document_with_regions_exact_prefix_retained',
  'EditSession.layout_document_with_regions_prefix_retained_json', 'EditSession.layout_document_with_regions_retained',
  'EditSession.layout_document_with_regions_retained_json', 'EditSession.layout_document_with_regions_retained_meta',
  'EditSession.load', 'EditSession.load_json', 'EditSession.load_media_sources', 'EditSession.load_note_separators',
  'EditSession.measure_paragraph_json', 'EditSession.media_bytes', 'EditSession.media_data_url',
  'EditSession.media_sources_json', 'EditSession.media_type', 'EditSession.move_embed',
  'EditSession.next_paragraph_style', 'EditSession.note_separators_state', 'EditSession.open_docx', 'EditSession.outline_glyph_json',
  'EditSession.page_texts_json', 'EditSession.paragraph_id_count', 'EditSession.paragraph_save_plan', 'EditSession.paragraph_unit',
  'EditSession.patch_style_languages_xml', 'EditSession.persist_paragraph_ids', 'EditSession.prime_last_slice_cost', 'EditSession.prime_save_seed_states', 'EditSession.prime_save_seed_states_slice',
  'EditSession.proposal_revision_ranges_json', 'EditSession.rebuild_media_sources', 'EditSession.record_saved_paragraph_ids',
  'EditSession.register_substitute_measure_font', 'EditSession.release_display_pages_frame', 'EditSession.repeat_content_control_at', 'EditSession.reset_frame_base',
  'EditSession.resident_caret_snapshot_json', 'EditSession.resident_deleted_units', 'EditSession.resident_geometry_json', 'EditSession.resident_section_revision',
  'EditSession.resolve_encoded_selection', 'EditSession.resolve_paragraph_anchor', 'EditSession.resolve_sticky_position', 'EditSession.resume_region_layout',
  'EditSession.retained_headers_footers_json', 'EditSession.retained_kernel_inputs_json', 'EditSession.retained_layout_json', 'EditSession.revision_ids_at',
  'EditSession.revision_stamps_json', 'EditSession.run_inheritance', 'EditSession.seed_comment_metadata_json', 'EditSession.seed_from_docx',
  'EditSession.seed_opaque_sequences', 'EditSession.select_story', 'EditSession.selection_context', 'EditSession.selection',
  'EditSession.set_cell_selection',
  'EditSession.set_content_control_value_at', 'EditSession.set_content_control_value_index',
  'EditSession.set_direct_batches', 'EditSession.set_display_retain_built_pages', 'EditSession.set_display_window', 'EditSession.set_document_property',
  'EditSession.set_header_footer_aliases', 'EditSession.set_hyperlink', 'EditSession.set_image_geometry_at', 'EditSession.set_image_geometry',
  'EditSession.set_local_lowering', 'EditSession.set_media_tokens', 'EditSession.set_part_metadata', 'EditSession.set_partial_document',
  'EditSession.set_selection', 'EditSession.set_style_language',
  'EditSession.set_undo_input_time',
  'EditSession.set_update_observer', 'EditSession.set_windowed_incremental_builds', 'EditSession.settle_revisions_json',
  'EditSession.start_update_event_observation', 'EditSession.stories_changed_since',
  'EditSession.story_segment_unit_digests', 'EditSession.story_segment_units', 'EditSession.style_languages_json', 'EditSession.take_base_latin_miss',
  'EditSession.unavailable_tables_from',
  'EditSession.update_measure_font_outlines', 'EditSession.written_paragraph_ids', 'EditSession.yrs_blocks_for_story',
  'RetainedLayoutMeta.layout_shell_json', 'RetainedLayoutMeta.page_sizes', 'RetainedLayoutMeta.notes_converged', 'RetainedLayoutMeta.page_count',
  'RetainedLayoutMeta.partial', 'RetainedLayoutMeta.provisional',
  'RunInheritance.attributes', 'RunInheritance.paragraph_formatting', 'RunInheritance.table_attributes',
  'build_display_list_json', 'clear_measure_fonts', 'close_display_list', 'decodeFontFace', 'decodeTiffPng',
  'hit_test_regions_by_handle', 'hit_test_regions_json', 'layout_document_json', 'measure_paragraph_json', 'open_display_list',
  'outline_glyph_json', 'parse_relationships_xml', 'range_rects_by_handle', 'range_rects_on_pages_by_handle', 'range_rects_on_pages_json',
  'range_rects_region_by_handle', 'range_rects_region_json', 'register_measure_font', 'register_substitute_measure_font',
  'render_docx_markdown_json', 'render_docx_markdown_with_pages_json', 'serialize_docx_s10', 'serialize_docx_s11', 'serialize_docx_s12',
  'update_display_list', 'vertical_move_by_handle', 'vertical_move_json',
];

function docxSurface(): string[] {
  const read = (path: string) => parseSurface(readFileSync(new URL(`../../${path}`, import.meta.url), 'utf8'));
  const wasm = Object.entries(FORMATS).filter(([name]) => name === 'docx' || name.startsWith('docx-')).map(([, format]) => read(format.wasm));
  return surfaceMethods(wasm, read(FORMATS.docx!.native));
}

describe('docx native and wasm parity table', () => {
  test('every method on both engines is a row, an exclusion or not yet covered', () => {
    expect(coverageProblems(docxSurface(), { rows, exclusions, notYetCovered, knownDifferences })).toEqual([]);
  });

  test('the not yet covered list only shrinks', () => {
    expect(notYetCovered.length).toBe(177);
  });

  for (const row of rows) {
    const known = knownDifferences.find((entry) => entry.method === row.method);
    test(known ? `${rowName(row)} still differs: ${known.reason}` : `${rowName(row)} matches`, async () => {
      await (known ? runKnownDifference : runRow)(row, makeWasm, makeNative);
    }, 30_000);
  }
});
