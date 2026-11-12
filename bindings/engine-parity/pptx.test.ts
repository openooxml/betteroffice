import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import initWasm, * as wasmPptx from '../../packages/pptx/src/wasm/generated/pptx_wasm.js';
import initPreflight, { font_preflight as wasmFontPreflight } from '../../packages/pptx/src/wasm/generated/preflight/pptx_preflight.js';
import { FORMATS, parseSurface } from '../../scripts/check-engine-surface';
import {
  coverageProblems,
  defineRow,
  rowName,
  runKnownDifference,
  runRow,
  surfaceMethods,
  ZipReturn,
  type EngineHandle,
  type Exclusion,
  type KnownDifference,
  type Row,
} from './harness';

type PptxEngine = typeof wasmPptx & { font_preflight: typeof wasmFontPreflight };
type Doc = wasmPptx.PptxDocument;
type Renderer = wasmPptx.PptxRenderer;

interface Pptx {
  engine: PptxEngine;
  doc: Doc;
  renderer: Renderer;
}

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const PML = 'application/vnd.openxmlformats-officedocument.presentationml';
const NS = `xmlns:a="${A}" xmlns:r="${REL}" xmlns:p="${P}"`;
const MAIN_PART = 'ppt/presentation.xml';
const PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const PNG = Uint8Array.from(Buffer.from(PNG_BASE64, 'base64'));
const DATE = '2025-01-01T00:00:00Z';
const CLIENT = 4321;
const RECT = { x: 914400, y: 1828800, width: 2743200, height: 914400 };

const rels = (entries: [string, string, string][]) =>
  `<Relationships xmlns="${PKG_REL}">${entries.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join('')}</Relationships>`;
const tree = (shapes: string) => `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree>`;
const run = (text: string) => `<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r>`;
const paragraphs = (...texts: string[]) => texts.map((text) => `<a:p>${run(text)}</a:p>`).join('');
const xfrm = (x: number, y: number, cx: number, cy: number) => `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
const placeholder = (id: number, name: string, ph: string, body: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>`;
const textShape = (id: number, name: string, y: number, body: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(457200, y, 6000000, 900000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>`;
const cell = (text: string) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphs(text)}</a:txBody><a:tcPr/></a:tc>`;

const TABLE = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="457200" y="3200400"/><a:ext cx="4000000" cy="800000"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>` +
  `<a:tr h="400000">${cell('A')}${cell('B')}</a:tr><a:tr h="400000">${cell('C')}${cell('D')}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
const PICTURE = '<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture" descr="Parity pixel"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  `<p:spPr>${xfrm(6400800, 3200400, 914400, 914400)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;

const SLIDE1 = `<p:sld ${NS}><p:cSld>${tree(
  placeholder(2, 'Title', '<p:ph type="title"/>', paragraphs('Parity title')) +
  textShape(3, 'Body', 1371600, paragraphs('First parity paragraph', 'Second line of text')) +
  TABLE + PICTURE,
)}</p:cSld></p:sld>`;
const SLIDE2 = `<p:sld ${NS}><p:cSld name="Closing">${tree(textShape(2, 'Closing', 1371600, paragraphs('Closing parity slide')))}</p:cSld></p:sld>`;
const NOTES = `<p:notes ${NS}><p:cSld>${tree(placeholder(2, 'Notes', '<p:ph type="body" idx="1"/>', paragraphs('Speaker notes for parity')))}</p:cSld></p:notes>`;
const LAYOUT = `<p:sldLayout ${NS} type="obj"><p:cSld name="Content">${tree(
  placeholder(2, 'Title', '<p:ph type="title"/>', paragraphs('Layout title')) +
  placeholder(3, 'Body', '<p:ph type="body" idx="1"/>', paragraphs('Layout body')),
)}</p:cSld></p:sldLayout>`;
const MASTER = `<p:sldMaster ${NS}><p:cSld>${tree(
  placeholder(2, 'Title', '<p:ph type="title"/>', '<a:p/>') + placeholder(3, 'Body', '<p:ph type="body" idx="1"/>', '<a:p/>'),
)}</p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
  '<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr/></p:otherStyle></p:txStyles></p:sldMaster>';
const THEME = `<a:theme xmlns:a="${A}" name="Parity"><a:themeElements><a:clrScheme name="Parity">` +
  ['dk1:000000', 'lt1:FFFFFF', 'dk2:222222', 'lt2:EEEEEE', 'accent1:4472C4', 'accent2:ED7D31', 'accent3:A5A5A5', 'accent4:FFC000', 'accent5:5B9BD5', 'accent6:70AD47', 'hlink:0563C1', 'folHlink:954F72']
    .map((entry) => entry.split(':')).map(([name, rgb]) => `<a:${name}><a:srgbClr val="${rgb}"/></a:${name}>`).join('') +
  '</a:clrScheme><a:fontScheme name="Parity"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme></a:themeElements></a:theme>';

async function buildFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  const date = new Date('2024-01-01T00:00:00Z');
  const override = (part: string, type: string) => `<Override PartName="/${part}" ContentType="${PML}.${type}"/>`;
  const parts: Record<string, string | Uint8Array> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
      override(MAIN_PART, 'presentation.main+xml') + override('ppt/slides/slide1.xml', 'slide+xml') + override('ppt/slides/slide2.xml', 'slide+xml') +
      override('ppt/notesSlides/notesSlide1.xml', 'notesSlide+xml') + override('ppt/comments/comment1.xml', 'comments+xml') + override('ppt/commentAuthors.xml', 'commentAuthors+xml') +
      override('ppt/slideLayouts/slideLayout1.xml', 'slideLayout+xml') + override('ppt/slideMasters/slideMaster1.xml', 'slideMaster+xml') +
      '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
    '_rels/.rels': rels([['rId1', 'officeDocument', MAIN_PART]]),
    [MAIN_PART]: `<p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      '<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
    'ppt/_rels/presentation.xml.rels': rels([
      ['rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'], ['rId2', 'slide', 'slides/slide1.xml'], ['rId3', 'slide', 'slides/slide2.xml'],
      ['rId4', 'commentAuthors', 'commentAuthors.xml'], ['rId5', 'theme', 'theme/theme1.xml'],
    ]),
    'ppt/slides/slide1.xml': SLIDE1,
    'ppt/slides/_rels/slide1.xml.rels': rels([
      ['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ['rId2', 'notesSlide', '../notesSlides/notesSlide1.xml'],
      ['rId3', 'comments', '../comments/comment1.xml'], ['rIdImg', 'image', '../media/image1.png'],
    ]),
    'ppt/slides/slide2.xml': SLIDE2,
    'ppt/slides/_rels/slide2.xml.rels': rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml']]),
    'ppt/notesSlides/notesSlide1.xml': NOTES,
    'ppt/notesSlides/_rels/notesSlide1.xml.rels': rels([['rId1', 'slide', '../slides/slide1.xml']]),
    'ppt/comments/comment1.xml': `<p:cmLst ${NS}><p:cm authorId="0" dt="2024-01-01T00:00:00.000" idx="1"><p:pos x="10" y="10"/><p:text>Seeded comment</p:text></p:cm></p:cmLst>`,
    'ppt/commentAuthors.xml': `<p:cmAuthorLst ${NS}><p:cmAuthor id="0" name="Parity" initials="P" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>`,
    'ppt/slideLayouts/slideLayout1.xml': LAYOUT,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']]),
    'ppt/slideMasters/slideMaster1.xml': MASTER,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ['rId2', 'theme', '../theme/theme1.xml']]),
    'ppt/theme/theme1.xml': THEME,
    'ppt/media/image1.png': PNG,
  };
  for (const [name, content] of Object.entries(parts)) zip.file(name, content, { date });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const fontBytes = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
const wasmEngine: PptxEngine = { ...wasmPptx, font_preflight: wasmFontPreflight };
const nativeEngine = createRequire(import.meta.url)('../pptx-native/index.js') as PptxEngine;
let fixture: Uint8Array;

beforeAll(async () => {
  await initWasm({ module_or_path: readFileSync(new URL('../../packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)) });
  await initPreflight({ module_or_path: readFileSync(new URL('../../packages/pptx/src/wasm/generated/preflight/pptx_preflight_bg.wasm', import.meta.url)) });
  fixture = await buildFixture();
});

const bytes = () => fixture.slice();

// documentVersion() is "{nonce:016x}-{epoch}"; the nonce is minted per open.
function versionNonce(doc: Doc): string {
  const version = doc.documentVersion();
  return version.slice(0, version.lastIndexOf('-'));
}

function engineFor(engine: PptxEngine) {
  return async (): Promise<EngineHandle<Pptx>> => {
    const doc = engine.PptxDocument.openCollaborative(bytes(), CLIENT);
    const renderer = new engine.PptxRenderer();
    const free = () => {
      renderer.free();
      doc.free();
    };
    try {
      renderer.registerFont('Arial', false, false, fontBytes);
      renderer.registerFallbackFont('Arial', false, false, fontBytes);
    } catch (error) {
      free();
      throw error;
    }
    return {
      target: { engine, doc, renderer },
      generated: () => [versionNonce(doc)],
      save: async () => doc.saveBytes(),
      free,
    };
  };
}

const makeWasm = engineFor(wasmEngine);
const makeNative = engineFor(nativeEngine);

function scoped<R>(doc: Doc, read: (doc: Doc) => R): R {
  try {
    return read(doc);
  } finally {
    doc.free();
  }
}

interface Snapshot {
  slides: { id: string; layoutPartPath: string | null; shapes: { id: string; name: string; mediaPartPath: string | null; textStories: { id: string }[] }[] }[];
  comments?: { id: string }[];
}

const json = (value: unknown) => JSON.stringify(value);
const snapshot = (doc: Doc) => JSON.parse(doc.snapshotJson()) as Snapshot;
const slideId = (doc: Doc, index: number) => snapshot(doc).slides[index]!.id;
const layoutPart = (doc: Doc) => snapshot(doc).slides[0]!.layoutPartPath;

function shape(doc: Doc, name: string) {
  for (const slide of snapshot(doc).slides) {
    const found = slide.shapes.find((candidate) => candidate.name === name);
    if (found) return { slideId: slide.id, shapeId: found.id, storyId: found.textStories[0]?.id ?? '', mediaPartPath: found.mediaPartPath };
  }
  throw new Error(`no shape named ${name}`);
}

const ids = (doc: Doc, name: string) => {
  const { slideId: slide, shapeId } = shape(doc, name);
  return { slideId: slide, shapeId };
};
const story = (doc: Doc) => shape(doc, 'Body').storyId;
const seededComment = (doc: Doc) => snapshot(doc).comments![0]!.id;
const insertText = (doc: Doc, text = 'Inserted ') => doc.insertTextJson(json({ storyId: story(doc), index: 0, text, style: {} }));
const addComment = (doc: Doc) => (JSON.parse(doc.addCommentJson(json({
  slideId: slideId(doc, 0), author: 'Parity', initials: 'P', text: 'Added comment', created: DATE, xEmu: 31750, yEmu: 63500,
}))) as { commentId: string }).commentId;
const modernComment = (doc: Doc) => {
  doc.setCommentFlavorJson(json({ flavor: 'modern' }));
  return addComment(doc);
};
const table = (doc: Doc, extra: Record<string, unknown>) => json({ ...ids(doc, 'Table'), ...extra });
const editRequest = (doc: Doc) => {
  const body = shape(doc, 'Body');
  const closing = shape(doc, 'Closing');
  const step = (target: typeof body, text: string) => ({
    op: 'insertText',
    target: { kind: 'range', slideId: target.slideId, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 },
    at: 'start',
    text,
  });
  return json({ expectVersion: doc.documentVersion(), steps: [step(body, 'Draft: '), step(closing, 'Final: ')] });
};
const layout = ({ doc, renderer }: Pptx) => renderer.layoutSlideJson(doc, 0);

const rows: Row<Pptx>[] = [
  { method: 'PptxDocument.openCollaborative', call: ({ engine }) => scoped(engine.PptxDocument.openCollaborative(bytes(), CLIENT), (d) => [d.snapshotJson(), d.clientId]) },
  {
    method: 'PptxDocument.openCollaborative',
    name: 'PptxDocument.openCollaborative (client 0)',
    throws: { name: 'string', reason: 'a client id must be positive' },
    call: ({ engine }) => engine.PptxDocument.openCollaborative(bytes(), 0),
  },
  defineRow({
    method: 'PptxDocument.openCollaborativeFromUpdate',
    setup: ({ doc }: Pptx) => {
      insertText(doc);
      return doc.encodeStateAsUpdate();
    },
    call: ({ engine }, update) => scoped(engine.PptxDocument.openCollaborativeFromUpdate(update, CLIENT + 1, bytes()), (d) => [d.snapshotJson(), new ZipReturn(d.saveBytes())]),
  }),
  { method: 'PptxDocument.saveBytes', call: ({ doc }) => new ZipReturn(doc.saveBytes()) },
  { method: 'PptxRenderer.saveBytes', call: ({ doc, renderer }) => new ZipReturn(renderer.saveBytes(doc)) },
  { method: 'PptxDocument.version', call: ({ engine }) => engine.PptxDocument.version() },
  { method: 'PptxDocument.clientId', call: ({ doc }) => doc.clientId },
  { method: 'parsePptxJson', call: ({ engine }) => engine.parsePptxJson(bytes()) },
  { method: 'parsePptxJsonWithoutMedia', call: ({ engine }) => engine.parsePptxJsonWithoutMedia(bytes()) },
  { method: 'exportPptxStructuredJson', call: ({ engine }) => engine.exportPptxStructuredJson(bytes(), json({ includeNotes: true, includeComments: true })) },
  { method: 'exportPptxMarkdownJson', call: ({ engine }) => engine.exportPptxMarkdownJson(bytes(), json({ includeNotes: true, includeComments: true })) },
  {
    method: 'renderPptxMarkdownJson',
    call: ({ engine }) => engine.renderPptxMarkdownJson(json((JSON.parse(wasmEngine.exportPptxStructuredJson(bytes(), '{}')) as { content: unknown }).content), '{}'),
  },
  { method: 'font_preflight', call: ({ engine }) => engine.font_preflight(bytes(), 0) },
  { method: 'rendererVersion', call: ({ engine }) => engine.rendererVersion() },
  { method: 'PptxDocument.snapshotJson', call: ({ doc }) => doc.snapshotJson() },
  { method: 'PptxDocument.readContentJson', call: ({ doc }) => [doc.readContentJson('{}'), doc.readContentJson(json({ slideIds: [slideId(doc, 1)] }))] },
  { method: 'PptxDocument.findTextJson', call: ({ doc }) => doc.findTextJson(json({ text: 'parity', limit: 10 })) },
  { method: 'PptxDocument.searchTextJson', call: ({ doc }) => [false, true].map((caseSensitive) => doc.searchTextJson(json({ query: 'Parity', caseSensitive, limit: 10 }))) },
  { method: 'PptxDocument.storyJson', call: ({ doc }) => doc.storyJson(json({ storyId: story(doc) })) },
  {
    method: 'PptxDocument.storyFormattingJson',
    call: ({ doc }) => {
      const { slideId: slide, shapeId, storyId } = shape(doc, 'Body');
      return doc.storyFormattingJson(json({ slideId: slide, shapeId, storyId }));
    },
  },
  { method: 'PptxDocument.commentsJson', call: ({ doc }) => doc.commentsJson() },
  { method: 'PptxDocument.mediaBytes', call: ({ doc }) => doc.mediaBytes(shape(doc, 'Picture').mediaPartPath!) },
  { method: 'PptxDocument.metafileSvg', voidReason: 'a PNG is not a metafile, so there is no SVG', call: ({ doc }) => doc.metafileSvg(shape(doc, 'Picture').mediaPartPath!) },
  { method: 'PptxDocument.insertTextJson', mutates: true, call: ({ doc }) => insertText(doc) },
  { method: 'PptxDocument.deleteTextJson', mutates: true, call: ({ doc }) => doc.deleteTextJson(json({ storyId: story(doc), start: 0, end: 6 })) },
  {
    method: 'PptxDocument.formatTextJson',
    mutates: true,
    call: ({ doc }) => doc.formatTextJson(json({ storyId: story(doc), start: 0, end: 5, patch: { bold: true, color: '#C00000', fontSizePt: 24 } })),
  },
  { method: 'PptxDocument.setParagraphAlignmentJson', mutates: true, call: ({ doc }) => doc.setParagraphAlignmentJson(json({ storyId: story(doc), start: 0, end: 5, alignment: 'ctr' })) },
  { method: 'PptxDocument.insertParagraphBreakJson', mutates: true, call: ({ doc }) => doc.insertParagraphBreakJson(json({ storyId: story(doc), index: 5 })) },
  {
    method: 'PptxDocument.setStoryParagraphsJson',
    mutates: true,
    call: ({ doc }) => doc.setStoryParagraphsJson(json({
      storyId: story(doc),
      paragraphs: [{ alignment: 'l', runs: [{ text: 'Replaced', style: { bold: true } }] }, { alignment: 'r', runs: [{ text: 'Second', style: { italic: true } }] }],
    })),
  },
  { method: 'PptxDocument.validateEditsJson', call: ({ doc }) => doc.validateEditsJson(editRequest(doc)) },
  { method: 'PptxDocument.applyEditsJson', mutates: true, call: ({ doc }) => doc.applyEditsJson(editRequest(doc)) },
  { method: 'PptxDocument.insertTableRowJson', mutates: true, call: ({ doc }) => doc.insertTableRowJson(table(doc, { index: 0, position: 'after', cells: ['E', 'F'] })) },
  { method: 'PptxDocument.deleteTableRowJson', mutates: true, call: ({ doc }) => doc.deleteTableRowJson(table(doc, { index: 1 })) },
  {
    method: 'PptxDocument.insertTableColumnJson',
    mutates: true,
    call: ({ doc }) => doc.insertTableColumnJson(table(doc, { index: 0, position: 'after', cells: ['X', 'Y'], width: 'split' })),
  },
  { method: 'PptxDocument.deleteTableColumnJson', mutates: true, call: ({ doc }) => doc.deleteTableColumnJson(table(doc, { index: 1 })) },
  {
    method: 'PptxDocument.checkTableEditJson',
    voidReason: 'a valid edit passes by not throwing',
    call: ({ doc }) => doc.checkTableEditJson(table(doc, { edit: 'deleteRow', index: 0 })),
  },
  { method: 'PptxDocument.insertSlideJson', mutates: true, call: ({ doc }) => doc.insertSlideJson(json({ index: 1, layoutPartPath: layoutPart(doc) })) },
  { method: 'PptxDocument.duplicateSlideJson', mutates: true, call: ({ doc }) => doc.duplicateSlideJson(json({ slideId: slideId(doc, 0), toIndex: 2 })) },
  { method: 'PptxDocument.deleteSlideJson', mutates: true, call: ({ doc }) => doc.deleteSlideJson(json({ slideId: slideId(doc, 1) })) },
  { method: 'PptxDocument.moveSlideJson', mutates: true, call: ({ doc }) => doc.moveSlideJson(json({ slideId: slideId(doc, 1), toIndex: 0 })) },
  {
    method: 'PptxDocument.setSlideLayoutJson',
    mutates: true,
    call: ({ doc }) => doc.setSlideLayoutJson(json({ slideId: slideId(doc, 1), layoutPartPath: layoutPart(doc), mapping: [] })),
  },
  { method: 'PptxDocument.setSlideNotesJson', mutates: true, call: ({ doc }) => doc.setSlideNotesJson(json({ slideId: slideId(doc, 0), text: 'Rewritten notes' })) },
  {
    method: 'PptxDocument.setSlideNotesJson',
    name: 'PptxDocument.setSlideNotesJson (new notes part)',
    mutates: true,
    call: ({ doc }) => doc.setSlideNotesJson(json({ slideId: slideId(doc, 1), text: 'First notes' })),
  },
  {
    method: 'PptxDocument.addTextBoxJson',
    mutates: true,
    call: ({ doc }) => doc.addTextBoxJson(json({ slideId: slideId(doc, 1), draft: { name: 'Added box', rect: RECT, text: 'Added text', style: { bold: true } } })),
  },
  {
    method: 'PptxDocument.addShapeJson',
    mutates: true,
    call: ({ doc }) => doc.addShapeJson(json({ slideId: slideId(doc, 1), draft: { name: 'Added shape', geometry: 'ellipse', rect: RECT, fill: '#123456' } })),
  },
  {
    method: 'PptxDocument.addPictureJson',
    mutates: true,
    call: ({ doc }) => doc.addPictureJson(json({ slideId: slideId(doc, 1), name: 'Added picture', rect: RECT, contentType: 'image/png', mediaBase64: PNG_BASE64 })),
  },
  { method: 'PptxDocument.removeShapeJson', mutates: true, call: ({ doc }) => doc.removeShapeJson(json(ids(doc, 'Picture'))) },
  { method: 'PptxDocument.bringShapeToFrontJson', mutates: true, call: ({ doc }) => doc.bringShapeToFrontJson(json(ids(doc, 'Title'))) },
  { method: 'PptxDocument.sendShapeToBackJson', mutates: true, call: ({ doc }) => doc.sendShapeToBackJson(json(ids(doc, 'Picture'))) },
  { method: 'PptxDocument.bringShapeForwardJson', mutates: true, call: ({ doc }) => doc.bringShapeForwardJson(json(ids(doc, 'Title'))) },
  { method: 'PptxDocument.sendShapeBackwardJson', mutates: true, call: ({ doc }) => doc.sendShapeBackwardJson(json(ids(doc, 'Picture'))) },
  { method: 'PptxDocument.moveShapeJson', mutates: true, call: ({ doc }) => doc.moveShapeJson(json({ ...ids(doc, 'Body'), x: 500000, y: 600000 })) },
  { method: 'PptxDocument.resizeShapeJson', mutates: true, call: ({ doc }) => doc.resizeShapeJson(json({ ...ids(doc, 'Body'), width: 4000000, height: 800000 })) },
  { method: 'PptxDocument.setShapeRectJson', mutates: true, call: ({ doc }) => doc.setShapeRectJson(json({ ...ids(doc, 'Body'), rect: RECT })) },
  { method: 'PptxDocument.setShapeFillJson', mutates: true, call: ({ doc }) => doc.setShapeFillJson(json({ ...ids(doc, 'Body'), color: '#3367D6' })) },
  {
    method: 'PptxDocument.setShapeStrokeJson',
    mutates: true,
    call: ({ doc }) => doc.setShapeStrokeJson(json({ ...ids(doc, 'Body'), stroke: { color: '#112233', widthPt: 2 } })),
  },
  { method: 'PptxDocument.addCommentJson', mutates: true, call: ({ doc }) => [addComment(doc), doc.commentsJson()] },
  defineRow({
    method: 'PptxDocument.replyToCommentJson',
    mutates: true,
    setup: ({ doc }: Pptx) => modernComment(doc),
    call: ({ doc }, commentId) => doc.replyToCommentJson(json({ commentId, author: 'Parity', initials: 'P', text: 'Reply', created: DATE })),
  }),
  { method: 'PptxDocument.setCommentPositionJson', mutates: true, call: ({ doc }) => doc.setCommentPositionJson(json({ commentId: seededComment(doc), xEmu: 95250, yEmu: 127000 })) },
  defineRow({
    method: 'PptxDocument.setCommentStatusJson',
    mutates: true,
    setup: ({ doc }: Pptx) => modernComment(doc),
    call: ({ doc }, commentId) => doc.setCommentStatusJson(json({ commentId, resolved: true })),
  }),
  { method: 'PptxDocument.removeCommentJson', mutates: true, call: ({ doc }) => doc.removeCommentJson(json({ commentId: seededComment(doc) })) },
  { method: 'PptxDocument.setCommentFlavorJson', mutates: true, call: ({ doc }) => doc.setCommentFlavorJson(json({ flavor: 'modern' })) },
  defineRow({
    method: 'PptxDocument.undoJson',
    mutates: true,
    setup: ({ doc }: Pptx) => insertText(doc),
    call: ({ doc }) => doc.undoJson(),
  }),
  defineRow({
    method: 'PptxDocument.redoJson',
    mutates: true,
    setup: ({ doc }: Pptx) => {
      insertText(doc);
      doc.undoJson();
    },
    call: ({ doc }) => doc.redoJson(),
  }),
  defineRow({
    method: 'PptxDocument.undoReceiptJson',
    mutates: true,
    setup: ({ doc }: Pptx) => insertText(doc),
    call: ({ doc }) => doc.undoReceiptJson(),
  }),
  defineRow({
    method: 'PptxDocument.redoReceiptJson',
    mutates: true,
    setup: ({ doc }: Pptx) => {
      insertText(doc);
      doc.undoReceiptJson();
    },
    call: ({ doc }) => doc.redoReceiptJson(),
  }),
  defineRow({
    method: 'PptxDocument.canUndo',
    setup: ({ doc }: Pptx) => insertText(doc),
    call: ({ doc }) => doc.canUndo(),
  }),
  defineRow({
    method: 'PptxDocument.canRedo',
    setup: ({ doc }: Pptx) => {
      insertText(doc);
      doc.undoJson();
    },
    call: ({ doc }) => doc.canRedo(),
  }),
  {
    method: 'PptxDocument.addUndoBoundary',
    mutates: true,
    call: ({ doc }) => {
      insertText(doc, 'One ');
      const result = doc.addUndoBoundary();
      insertText(doc, 'Two ');
      return { result, after: doc.undoJson() };
    },
  },
  { method: 'PptxDocument.undoCaptureMode', call: ({ doc }) => doc.undoCaptureMode() },
  {
    method: 'PptxDocument.setUndoCaptureMode',
    call: ({ doc }) => {
      const result = doc.setUndoCaptureMode('manual');
      return { result, after: doc.undoCaptureMode() };
    },
  },
  { method: 'PptxRenderer.layoutSlideJson', call: ({ doc, renderer }) => [0, 1].map((index) => renderer.layoutSlideJson(doc, index)) },
  { method: 'PptxRenderer.exportSlideJson', call: ({ doc, renderer }) => renderer.exportSlideJson(doc, 0) },
  defineRow({
    method: 'PptxRenderer.hitTestJson',
    setup: layout,
    call: ({ renderer }) => [[100, 60], [200, 200], [5, 5]].map(([x, y]) => renderer.hitTestJson(x!, y!)),
  }),
  {
    method: 'PptxRenderer.hitTestSlideJson',
    call: ({ doc, renderer }) => [[100, 60], [5, 5]].map(([x, y]) => renderer.hitTestSlideJson(doc, slideId(doc, 0), x!, y!)),
  },
  { method: 'PptxRenderer.slideLayoutKey', call: ({ doc, renderer }) => [0, 1].map((index) => renderer.slideLayoutKey(doc, index)) },
  { method: 'PptxRenderer.snapshotWithLayoutKeysJson', call: ({ doc, renderer }) => renderer.snapshotWithLayoutKeysJson(doc) },
  { method: 'PptxRenderer.requestedFontFacesJson', call: ({ doc, renderer }) => [renderer.requestedFontFacesJson(doc, null), renderer.requestedFontFacesJson(doc, slideId(doc, 1))] },
  { method: 'PptxRenderer.setActiveSlide', call: ({ doc, renderer }) => renderer.setActiveSlide(doc, slideId(doc, 1), renderer.slideLayoutKey(doc, 1)) },
  { method: 'PptxRenderer.registerFont', call: ({ renderer }) => renderer.registerFont('Parity Sans', true, false, fontBytes) },
  { method: 'PptxRenderer.registerFallbackFont', call: ({ renderer }) => renderer.registerFallbackFont('Parity Sans', false, true, fontBytes) },
  {
    method: 'PptxRenderer.setBrowserCoverageFallback',
    call: ({ doc, renderer }) => {
      const result = renderer.setBrowserCoverageFallback(true);
      return { result, after: renderer.layoutSlideJson(doc, 0) };
    },
  },
];

const GOLDEN = 'scripts/parity/pptx.ts drives it on both engines against parity/pptx/expected.json';
const TIMINGS = 'profiling twin of a covered call; only its wall-clock timings differ';
const exclusions: Exclusion[] = [
  { method: 'PptxDocument.free', reason: 'releases the object; the harness frees every document itself' },
  { method: 'PptxRenderer.free', reason: 'releases the object; the harness frees every renderer itself' },
  { method: 'PptxDocument.addTextBoxProfiledJson', reason: TIMINGS },
  { method: 'PptxDocument.deleteTextProfiledJson', reason: TIMINGS },
  { method: 'PptxDocument.insertSlideProfiledJson', reason: TIMINGS },
  { method: 'PptxDocument.insertTextProfiledJson', reason: TIMINGS },
  { method: 'PptxDocument.moveShapeProfiledJson', reason: TIMINGS },
  { method: 'PptxDocument.undoProfiledJson', reason: TIMINGS },
  { method: 'PptxRenderer.layoutSlideProfiledJson', reason: TIMINGS },
  { method: 'PptxDocument.documentVersion', reason: GOLDEN },
  { method: 'PptxDocument.exportMarkdownJson', reason: GOLDEN },
  { method: 'PptxDocument.exportStructuredJson', reason: GOLDEN },
  { method: 'PptxDocument.readCommandContentJson', reason: GOLDEN },
];

const knownDifferences: KnownDifference[] = [];

const notYetCovered: string[] = [
  'PptxDocument.acceptProposalJson', 'PptxDocument.adoptPeerIdentity', 'PptxDocument.anchorCaretJson', 'PptxDocument.applyUpdateJson',
  'PptxDocument.clearUpdateObservation', 'PptxDocument.drainUpdateEvent', 'PptxDocument.encodeDiff', 'PptxDocument.encodeStateAsUpdate',
  'PptxDocument.encodeStateVector', 'PptxDocument.listProposalsJson', 'PptxDocument.openPeerDeckJson', 'PptxDocument.openReplayBaseline',
  'PptxDocument.peerHydrationJson', 'PptxDocument.previewProposalJson', 'PptxDocument.proposalSlideStripJson', 'PptxDocument.proposeJson',
  'PptxDocument.registerPeerFontsJson', 'PptxDocument.rejectProposalJson', 'PptxDocument.replayJson', 'PptxDocument.resolveCaretAnchorJson',
  'PptxDocument.sessionMetadataJson', 'PptxDocument.setCollaborative', 'PptxDocument.setShapeAdjustJson', 'PptxDocument.startUpdateObservation',
  'PptxDocument.validateProposalJson',
  'PptxRenderer.layoutProposalDiffSlideByIdJson', 'PptxRenderer.layoutProposalDiffSlideJson', 'PptxRenderer.layoutProposalSlideJson',
  'compileSlideJson', 'decodeFontFace', 'decodeTiffPng',
];

function pptxSurface(): string[] {
  const read = (path: string) => parseSurface(readFileSync(new URL(`../../${path}`, import.meta.url), 'utf8'));
  const wasm = Object.entries(FORMATS).filter(([name]) => name === 'pptx' || name.startsWith('pptx-')).map(([, format]) => read(format.wasm));
  return surfaceMethods(wasm, read(FORMATS.pptx!.native));
}

describe('pptx native and wasm parity table', () => {
  test('every method on both engines is a row, an exclusion or not yet covered', () => {
    expect(coverageProblems(pptxSurface(), { rows, exclusions, notYetCovered, knownDifferences })).toEqual([]);
  });

  test('the not yet covered list only shrinks', () => {
    expect(notYetCovered.length).toBe(31);
  });

  for (const row of rows) {
    const known = knownDifferences.find((entry) => entry.method === row.method);
    test(known ? `${rowName(row)} still differs: ${known.reason}` : `${rowName(row)} matches`, async () => {
      await (known ? runKnownDifference(row, makeWasm, makeNative) : runRow(row, makeWasm, makeNative, MAIN_PART));
    }, 30_000);
  }
});
