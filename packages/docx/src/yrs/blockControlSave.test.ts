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
const TABLE = '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:tcMar><w:left w:w="99" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:noProof/></w:rPr><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
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
  'paragraph mark': [(session, range) => {
    const paragraphs = session.paragraphs(range.story) as Array<{ paraId: string }>;
    const next = paragraphs[paragraphs.findIndex(paragraph => paragraph.paraId === range.start.paraId) + 1]!;
    session.formatRange({ ...range, end: { paraId: next.paraId, offset: 0 } }, { italic: true });
  }, '<w:i/>'],
};
const open: YrsSession[] = [];
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => { open.splice(0).forEach(session => session.destroy()); });
const targetRange = (session: YrsSession): Range => {
  for (const { story } of session.listStories().stories as Array<{ story: string }>) {
    const paragraph = (session.paragraphs(story) as Array<{ paraId: string; text: string }>).find(entry => entry.text === 'Target');
    if (paragraph) return { story, start: { paraId: paragraph.paraId, offset: 0 }, end: { paraId: paragraph.paraId, offset: 6 } };
  }
  throw new Error('no Target paragraph');
};
const SIBLING = '<w:p w:rsidR="00DD0000" w:rsidRDefault="00DD0000"><w:proofErr w:type="spellStart"/><w:r w:rsidR="00DD0001"><w:rPr><w:noProof/></w:rPr><w:t>Sibling</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>';
const sdt = (content: string, properties = '<w:id w:val="5"/><w:alias w:val="Field"/><w:tag w:val="field"/>') => `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const BOUND = '<w:id w:val="6"/><w:dataBinding w:xpath="/root[1]/name[1]" w:storeItemID="{11111111-2222-3333-4444-555555555555}"/><w:text/>';
const cellOf = (content: string) => TABLE.replace(/<w:p>[\s\S]*?<\/w:p>(?=<\/w:tc>)/, content);

const CONTROLS: Array<[string, string]> = [
  ['a block control', sdt(`${SIBLING}${TARGET}${SIBLING}`)],
  ['a data-bound block control', sdt(`${TARGET}${SIBLING}`, BOUND)],
  ['a block control in a table cell', cellOf(sdt(`${TARGET}${SIBLING}`))],
  ['a block control nested in a block control', sdt(`${SIBLING}${sdt(`${TARGET}${SIBLING}`, '<w:id w:val="7"/>')}`)],
];
for (const [saveName, save] of Object.entries(SAVES)) {
  for (const [controlName, control] of CONTROLS) {
    for (const [editName, [edit, effect]] of Object.entries(EDITS)) {
      test(`${saveName}: a ${editName} edit inside ${controlName} keeps every other byte of the body`, async () => {
        const source = documentOf(`${RICH}${control}${SIBLING}`);
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
    test(`${saveName}: ${controlName} saves unchanged without an edit`, async () => {
      const source = documentOf(`${RICH}${control}${SIBLING}`);
      const session = await createYrsSession();
      open.push(session);
      session.seedFromDocx(makeDocx(source));
      expect(partOf(await save(session))).toBe(source);
    });
  }
}
