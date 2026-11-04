import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { saveEditorDocument } from './editorSave';
import { createYrsSession, saveYrsDocx, yrsToDocument, type YrsSession } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RICH = '<w:p w:rsidR="00AA0000"><w:pPr><w:pPrChange w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr/></w:pPrChange></w:pPr><w:r w:rsidR="00BB0000"><w:rPr><w:noProof/><w:specVanish/></w:rPr><w:t>Rich</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>wrd</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>';
const TARGET = '<w:p><w:r><w:t>Target</w:t></w:r></w:p>';
const PAGE_BREAK = '<w:p w:rsidR="00CC0000"><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:br w:type="page"/></w:r></w:p>';
const TABLE = '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:tcMar><w:left w:w="99" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:noProof/></w:rPr><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
const sdt = (content: string) => `<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
const STYLES = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style></w:styles>`;
const documentOf = (body: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}${SECT}</w:body></w:document>`;
const makeDocx = (document: string) => new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')],
  ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
  ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`)],
  ['word/styles.xml', toBytes(STYLES)],
  ['word/document.xml', toBytes(document)],
])));
const partOf = (bytes: Uint8Array) => new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
type Save = (session: YrsSession) => Promise<Uint8Array>;
const SAVES: Record<string, Save> = {
  saveYrsDocx: async session => (await saveYrsDocx(session, { updateModifiedDate: false })).bytes,
  editor: async (session) => {
    const projected = yrsToDocument(session, session.materializeDocx()!);
    return new Uint8Array(await saveEditorDocument(session, projected, projected.package.document.comments ?? []));
  },
};
type Range = { story: string; start: { paraId: string; offset: number }; end: { paraId: string; offset: number } };
const EDITS: Record<string, [(session: YrsSession, range: Range) => void, string]> = {
  style: [(session, range) => session.applyParagraphStyle(range, 'Heading1'), '<w:pStyle w:val="Heading1"/>'],
  bold: [(session, range) => session.formatRange(range, { bold: true }), '<w:b/>'],
  text: [(session, range) => session.replaceRange({ ...range, end: { ...range.end, offset: 2 } }, 'Xy'), '>Xyrget<'],
};
const open: YrsSession[] = [];
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => { open.splice(0).forEach(session => session.destroy()); });
const targetRange = (session: YrsSession): Range => {
  const paragraph = (session.paragraphs('body') as Array<{ paraId: string; text: string }>).find(entry => entry.text === 'Target')!;
  return { story: 'body', start: { paraId: paragraph.paraId, offset: 0 }, end: { paraId: paragraph.paraId, offset: 6 } };
};

const TAILS: Array<[string, string]> = [
  ['a page break paragraph', PAGE_BREAK],
  ['a table then a page break', `${TABLE}${PAGE_BREAK}`],
  ['a block control ending in a page break', sdt(`<w:p><w:r><w:t>Inside</w:t></w:r></w:p>${PAGE_BREAK}`)],
  ['two page break paragraphs', `${PAGE_BREAK}${PAGE_BREAK}`],
];
for (const [saveName, save] of Object.entries(SAVES)) {
  for (const [tailName, tail] of TAILS) {
    for (const [editName, [edit, effect]] of Object.entries(EDITS)) {
      test(`${saveName}: ${editName} keeps a body ending in ${tailName} byte-identical outside the edit`, async () => {
        const source = documentOf(`${RICH}${TARGET}${TABLE}${tail}`);
        const session = await createYrsSession();
        open.push(session);
        session.seedFromDocx(makeDocx(source));
        edit(session, targetRange(session));
        const saved = partOf(await save(session));
        const start = source.indexOf(TARGET);
        const suffix = source.slice(start + TARGET.length);
        expect(saved.slice(0, start)).toBe(source.slice(0, start));
        expect(saved.endsWith(suffix)).toBe(true);
        expect(saved.slice(start, saved.length - suffix.length)).toContain(effect);
      });
    }
    test(`${saveName}: a body ending in ${tailName} saves unchanged without an edit`, async () => {
      const source = documentOf(`${RICH}${TARGET}${tail}`);
      const session = await createYrsSession();
      open.push(session);
      session.seedFromDocx(makeDocx(source));
      expect(partOf(await save(session))).toBe(source);
    });
  }
}
