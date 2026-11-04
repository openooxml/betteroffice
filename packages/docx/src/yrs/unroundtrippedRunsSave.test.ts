import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { saveEditorDocument } from './editorSave';
import { createYrsSession, saveYrsDocx, yrsToDocument, type YrsSession } from './index';
import { RUN_WITH_ALTERNATE_CONTENT, TITLE_WITH_PAGE_BREAK_AND_DRAWING, TOC_ENTRY_WITH_PAGE_REFERENCE, WORD_ROOT } from './__fixtures__/unroundtrippedRuns';

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
].join(' ');
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
const text = (value: string) => `<w:p><w:r><w:t>${value}</w:t></w:r></w:p>`;
const table = (content: string) => `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${content}</w:tc></w:tr></w:tbl>`;
const PAGE_BREAK_END = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
const ALTERNATE = '<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="7" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect style="width:72pt;height:36pt"/></w:pict></mc:Fallback></mc:AlternateContent></w:r>';
const TOC_ENTRY = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Toc1"><w:r><w:t>Chapter</w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>';
const TOC_END = '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p><w:bookmarkStart w:id="1" w:name="_Toc1"/><w:r><w:t>Chapter</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>';

const makeDocx = (body: string) => new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
  ['word/document.xml', toBytes(`${body.startsWith(WORD_ROOT) ? '' : `<w:document ${NS}>`}${body.startsWith(WORD_ROOT) ? body : `<w:body>${body}`}${SECT}</w:body></w:document>`)],
])));
const partOf = (bytes: Uint8Array) => new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
const count = (xml: string, pattern: RegExp) => xml.match(pattern)?.length ?? 0;

type Save = (session: YrsSession) => Promise<Uint8Array>;
const SAVES: Record<string, Save> = {
  saveYrsDocx: async session => (await saveYrsDocx(session, { updateModifiedDate: false })).bytes,
  editor: async (session) => {
    const projected = yrsToDocument(session, session.materializeDocx()!);
    return new Uint8Array(await saveEditorDocument(session, projected, projected.package.document.comments ?? []));
  },
};
const open: YrsSession[] = [];
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => { open.splice(0).forEach(session => session.destroy()); });

function wholeBody(session: YrsSession) {
  const segments = session.storySegments('body') as Array<{ kind: string; paraId?: string; text?: string }>;
  const pilcrows = segments.filter(segment => segment.kind === 'pilcrow');
  const last = pilcrows.at(-1)!.paraId!;
  const end = segments.findIndex(segment => segment.kind === 'pilcrow' && segment.paraId === last);
  let start = end;
  while (start > 0 && segments[start - 1]!.kind !== 'pilcrow') start--;
  const length = segments.slice(start, end).reduce((sum, segment) => sum + (segment.kind === 'text' ? segment.text!.length : 1), 0);
  return { story: 'body', start: { paraId: pilcrows[0]!.paraId!, offset: 0 }, end: { paraId: last, offset: length } };
}
const EDITS = {
  clear: (session: YrsSession) => session.clearFormatting(wholeBody(session)),
  bold: (session: YrsSession) => session.formatRange(wholeBody(session), { bold: true }),
};

const CASES: Array<[string, string, Array<keyof typeof EDITS>, Array<[RegExp, number]>]> = [
  ['a page reference nested in a TOC hyperlink', `${text('Before')}${TOC_ENTRY}${TOC_END}${table(text('Cell'))}${text('After')}`,
    ['clear', 'bold'], [[/w:fldCharType="begin"/g, 2], [/PAGEREF/g, 1], [/<w:hyperlink w:anchor="_Toc1"/g, 1]]],
  ['alternate content in a cell and in the body', `${text('Before')}<w:p>${ALTERNATE}</w:p>${table(`<w:p>${ALTERNATE}</w:p>`)}${text('After')}`,
    ['clear'], [[/<mc:AlternateContent>/g, 2], [/<mc:Fallback>/g, 2], [/<w:noProof\/>/g, 2]]],
  ['a collapsed bookmark inside a hyperlink in a cell', `${text('Before')}${table('<w:p><w:hyperlink w:anchor="x"><w:r><w:t>Linked</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/><w:r><w:t xml:space="preserve"> text</w:t></w:r></w:hyperlink></w:p>')}${text('After')}`,
    ['clear'], [[/<w:bookmarkStart w:id="0" w:name="_GoBack"\/>/g, 1], [/<w:bookmarkEnd w:id="0"\/>/g, 1]]],
  ['two page breaks in a cell', `${text('Before')}${table('<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:br w:type="page"/><w:t>A</w:t></w:r></w:p>')}${text('After')}`,
    ['clear'], [[/<w:br w:type="page"\/>/g, 2]]],
  ['a tracked insertion holding only a page break', `${text('Before')}<w:p><w:pPr/><w:bookmarkStart w:id="1" w:name="_Toc9"/><w:ins w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:br w:type="page"/></w:r></w:ins></w:p>${table(text('Cell'))}${text('After')}`,
    ['clear'], [[/<w:ins w:id="2"/g, 1], [/<w:br w:type="page"\/>/g, 1]]],
  ['a page break before alternate content in a body ending with a page break', `${text('Before')}<w:p><w:r><w:t xml:space="preserve">Section title </w:t></w:r><w:r><w:br w:type="page"/></w:r>${ALTERNATE}</w:p>${text('After')}${PAGE_BREAK_END}`,
    ['clear'], [[/<w:br w:type="page"\/>/g, 2], [/<mc:AlternateContent>/g, 1], [/<mc:Fallback>/g, 1]]],
  ['a real TOC entry with a nested page reference', `${WORD_ROOT}<w:body>${text('Before')}${TOC_ENTRY_WITH_PAGE_REFERENCE}${TOC_END}${table(text('Cell'))}${text('After')}`,
    ['clear', 'bold'], [[/w:fldCharType="begin"/g, 2], [/PAGEREF/g, 1]]],
  ['a real run with alternate content in a cell', `${WORD_ROOT}<w:body>${text('Before')}${table(RUN_WITH_ALTERNATE_CONTENT)}${text('After')}`,
    ['clear'], [[/<mc:AlternateContent>/g, 1], [/<mc:Fallback>/g, 1]]],
  ['a real title with a page break and a drawing in a body ending with a page break', `${WORD_ROOT}<w:body>${text('Before')}${TITLE_WITH_PAGE_BREAK_AND_DRAWING}${text('After')}${PAGE_BREAK_END}`,
    ['clear'], [[/<w:br w:type="page"\/>/g, 2], [/<mc:AlternateContent>/g, 2]]],
];
for (const [saveName, save] of Object.entries(SAVES)) {
  for (const [name, body, edits, kept] of CASES) {
    for (const edit of edits) {
      test(`${saveName}: ${edit} over the body keeps ${name}`, async () => {
        const source = makeDocx(body);
        const before = partOf(source);
        const session = await createYrsSession();
        open.push(session);
        session.seedFromDocx(source);
        EDITS[edit](session);
        const xml = partOf(await save(session));
        for (const [pattern, expected] of kept) {
          expect(count(before, pattern)).toBe(expected);
          expect([pattern.source, count(xml, pattern)]).toEqual([pattern.source, expected]);
        }
      });
    }
  }
}
