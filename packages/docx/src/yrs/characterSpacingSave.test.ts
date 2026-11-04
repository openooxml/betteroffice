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
const NS = [
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
].join(' ');
const ALTERNATE = '<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="7" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect style="width:72pt;height:36pt"/></w:pict></mc:Fallback></mc:AlternateContent></w:r>';
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
const WRAPPERS = [
  '<w:hyperlink w:anchor="target"><w:r><w:t>Link</w:t></w:r></w:hyperlink>',
  '<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place"><w:r><w:t>Place</w:t></w:r></w:smartTag>',
  ALTERNATE,
];
const paragraph = (properties: string) => `<w:p><w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}<w:t>Target</w:t></w:r>${WRAPPERS.join('')}</w:p>`;
const makeDocx = (body: string) => new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
  ['word/document.xml', toBytes(`<w:document xmlns:w="${W}" xmlns:r="${R}" ${NS}><w:body>${body}${SECT}</w:body></w:document>`)],
])));
const partOf = (bytes: Uint8Array) => new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
const runOf = (xml: string, value: string) => (xml.match(/<w:r\b[\s\S]*?<\/w:r>/g) ?? []).find(run => run.includes(`>${value}<`)) ?? '';
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
async function seeded(properties: string) {
  const session = await createYrsSession();
  open.push(session);
  session.seedFromDocx(makeDocx(`<w:p><w:r><w:t>Before</w:t></w:r></w:p>${paragraph(properties)}<w:p><w:r><w:t>After</w:t></w:r></w:p>`));
  const target = (session.paragraphs('body') as Array<{ paraId: string; text: string }>).find(entry => entry.text.startsWith('Target'))!;
  return { session, range: { story: 'body', start: { paraId: target.paraId, offset: 0 }, end: { paraId: target.paraId, offset: 6 } } };
}
const legacyParagraphs = (session: YrsSession) => ((session.paragraphSavePlan() as { splicedParts?: Array<{ part: string; legacyParagraphs?: unknown[] }> })
  .splicedParts ?? []).find(part => part.part === 'word/document.xml')?.legacyParagraphs ?? null;

const ELEMENTS: Array<[string, string, Record<string, number>]> = [
  ['w:spacing', '<w:spacing w:val="-20"/>', { spacing: -20 }],
  ['w:position', '<w:position w:val="6"/>', { position: 6 }],
  ['w:w', '<w:w w:val="80"/>', { scale: 80 }],
  ['w:kern', '<w:kern w:val="28"/>', { kerning: 28 }],
];
for (const [saveName, save] of Object.entries(SAVES)) {
  for (const [name, element, value] of ELEMENTS) {
    test(`${saveName}: clearing ${name} patches the run and keeps the paragraph's wrappers`, async () => {
      const { session, range } = await seeded(element);
      session.clearFormatting(range);
      expect(legacyParagraphs(session)).toEqual([]);
      const xml = partOf(await save(session));
      expect(runOf(xml, 'Target')).not.toContain(`<${name} `);
      for (const wrapper of WRAPPERS) expect(xml).toContain(wrapper);
    });
    test(`${saveName}: setting ${name} patches the run and keeps the paragraph's wrappers`, async () => {
      const { session, range } = await seeded('');
      session.formatRange(range, { other: { characterSpacing: value } } as never);
      expect(legacyParagraphs(session)).toEqual([]);
      const xml = partOf(await save(session));
      expect(runOf(xml, 'Target')).toContain(element);
      for (const wrapper of WRAPPERS) expect(xml).toContain(wrapper);
    });
  }
  test(`${saveName}: clearing a theme font removes the run fonts and keeps the paragraph's wrappers`, async () => {
    const { session, range } = await seeded('<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:hint="eastAsia"/>');
    session.clearFormatting(range);
    expect(legacyParagraphs(session)).toEqual([]);
    const xml = partOf(await save(session));
    expect(runOf(xml, 'Target')).not.toContain('<w:rFonts');
    for (const wrapper of WRAPPERS) expect(xml).toContain(wrapper);
  });
}
