import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeYrs from '@betteroffice/docx-native/yrs';

type Yrs = typeof wasmYrs;
type Session = Awaited<ReturnType<Yrs['createYrsSession']>>;

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'mc:Ignorable="w14"',
].join(' ');
const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const ALTERNATE = '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Shape 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect style="width:72pt;height:36pt"/></w:pict></mc:Fallback></mc:AlternateContent>';
const CONTENT =
  '<w:r><w:rPr><w:noProof/></w:rPr><w:t>Unchecked</w:t></w:r>' +
  '<w:r><w:rPr><w:webHidden/></w:rPr><w:t>Hidden</w:t></w:r>' +
  `<w:r>${ALTERNATE}</w:r>` +
  '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>2</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const EDITED_ID = '73000001';
const UNTOUCHED_ID = '73000002';
const paragraph = (paraId: string) => `<w:p w14:paraId="${paraId}">${CONTENT}</w:p>`;
const ELEMENTS = ['w:noProof', 'w:webHidden', 'mc:AlternateContent', 'mc:Choice', 'mc:Fallback', 'w:pict', 'w:fldChar', 'w:instrText'];
const ATTRIBUTES = ['w:dirty'];

async function docx(body = `${paragraph(EDITED_ID)}${paragraph(UNTOUCHED_ID)}`, styles?: string): Promise<Uint8Array> {
  const zip = new JSZip();
  const stylesType = styles === undefined ? '' : '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>';
  const stylesRelationship = styles === undefined ? '' : '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${stylesType}</Types>`);
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${stylesRelationship}</Relationships>`);
  if (styles !== undefined) zip.file('word/styles.xml', styles);
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NAMESPACES}><w:body>${body}${SECTION}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const savedParagraph = (xml: string, paraId: string) =>
  xml.match(new RegExp(`<w:p\\b[^>]*w14:paraId="${paraId}"[^>]*>[\\s\\S]*?</w:p>`))?.[0];

const countElement = (xml: string, name: string) => (xml.match(new RegExp(`<${name}[\\s/>]`, 'g')) ?? []).length;
const countAttribute = (xml: string, name: string) => (xml.match(new RegExp(`\\s${name}=`, 'g')) ?? []).length;

async function typeIntoFirstParagraphAndSave(yrs: Yrs): Promise<string> {
  const session: Session = await yrs.createYrsSession();
  try {
    session.openDocx(await docx(), true);
    const paraId = session.paragraphs('body')[0]!.paraId;
    session.insertText({ story: 'body', paraId, offset: 1 }, 'x');
    const saved = await yrs.saveYrsDocx(session, { updateModifiedDate: false });
    return (await JSZip.loadAsync(saved.bytes)).file('word/document.xml')!.async('string');
  } finally {
    session.destroy();
  }
}

for (const [engine, yrs] of [['wasm', wasmYrs], ['native', nativeYrs as unknown as Yrs]] as const) {
  test(`${engine}: typing into one paragraph saves the untouched paragraph as its source bytes`, async () => {
    const xml = await typeIntoFirstParagraphAndSave(yrs);
    const source = paragraph(UNTOUCHED_ID);
    const saved = savedParagraph(xml, UNTOUCHED_ID);
    const failures = saved === source ? [] : [saved === undefined ? `paragraph ${UNTOUCHED_ID} missing from saved word/document.xml` : `paragraph ${UNTOUCHED_ID} source ${source}, saved ${saved}`];
    expect(failures).toEqual([]);
  });

  test(`${engine}: a paragraph edited by typing keeps its noProof and webHidden runs, field characters with w:dirty and alternate content`, async () => {
    const xml = await typeIntoFirstParagraphAndSave(yrs);
    const source = paragraph(EDITED_ID);
    const saved = savedParagraph(xml, EDITED_ID);
    const failures: string[] = [];
    if (saved === undefined) {
      failures.push(`paragraph ${EDITED_ID} missing from saved word/document.xml`);
    } else {
      for (const name of ELEMENTS) {
        const [before, after] = [countElement(source, name), countElement(saved, name)];
        if (after !== before) failures.push(`element ${name} ${after < before ? 'lost' : 'duplicated'}: source ${before}, saved ${after}`);
      }
      for (const name of ATTRIBUTES) {
        const [before, after] = [countAttribute(source, name), countAttribute(saved, name)];
        if (after !== before) failures.push(`attribute ${name} ${after < before ? 'lost' : 'duplicated'}: source ${before}, saved ${after}`);
      }
      if (failures.length > 0) failures.push(`saved paragraph ${saved}`);
    }
    expect(failures).toEqual([]);
  });
}

const CELL_STYLES = '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="CellText"><w:name w:val="Cell Text"/><w:basedOn w:val="Normal"/></w:style></w:styles>';
const REFERENCE =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> REF _Ref1 \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>Alpha</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const MARKED =
  '<w:pPr><w:pStyle w:val="CellText"/></w:pPr>' +
  '<w:bookmarkStart w:id="11" w:name="_Ref1"/><w:r><w:t xml:space="preserve">Alpha </w:t></w:r>' +
  '<w:bookmarkStart w:id="12" w:name="second"/><w:r><w:t xml:space="preserve">Beta </w:t></w:r>' +
  '<w:bookmarkStart w:id="13" w:name="third"/><w:r><w:t>Gamma</w:t></w:r><w:bookmarkEnd w:id="11"/>' +
  '<w:r><w:t xml:space="preserve"> delta </w:t></w:r><w:bookmarkEnd w:id="12"/>' +
  `<w:bookmarkStart w:id="14" w:name="fourth"/>${REFERENCE}<w:bookmarkEnd w:id="14"/>` +
  '<w:r><w:t xml:space="preserve"> omega</w:t></w:r><w:bookmarkEnd w:id="13"/>';
const TYPED_AT = 'Alpha Beta Gamma'.length;
const CELL_ID = '73000011';
const BODY_ID = '73000021';
const tableOf = (cell: string) =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${cell}</w:tc></w:tr></w:tbl>`;
const MARKED_BODY =
  tableOf(`${tableOf(`<w:p w14:paraId="${CELL_ID}">${MARKED}</w:p>`)}<w:p w14:paraId="73000012"/>`) +
  `<w:p w14:paraId="${BODY_ID}">${MARKED}</w:p>`;
const KEPT_ELEMENTS = ['w:bookmarkStart', 'w:bookmarkEnd', 'w:fldChar', 'w:instrText', 'w:pStyle'];
const FIELD_RESULT = /fldCharType="separate"(?:(?!fldCharType)[\s\S])*>Alpha<\/w:t>(?:(?!fldCharType)[\s\S])*fldCharType="end"/;

const outerCell = (xml: string) => {
  const [from, to] = [xml.indexOf('<w:tc>'), xml.lastIndexOf('</w:tc>')];
  return from < 0 || to < from ? undefined : xml.slice(from, to + '</w:tc>'.length);
};
const textOf = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((match) => match[1]).join('');

async function typeBesideMarkedField(yrs: Yrs, inCell: boolean): Promise<{ reasons: string[]; patches: number; xml: string }> {
  const session: Session = await yrs.createYrsSession();
  try {
    session.openDocx(await docx(MARKED_BODY, CELL_STYLES), true);
    const hits = session.storyIds().flatMap((story) =>
      session.paragraphs(story).filter((entry) => entry.text.startsWith('Alpha Beta')).map((entry) => ({ story, paraId: entry.paraId })));
    const target = inCell
      ? hits.find((entry) => entry.story !== 'body') ?? hits[0]
      : hits.filter((entry) => entry.story === 'body').at(-1);
    session.insertText({ ...target!, offset: TYPED_AT }, 'x');
    const plan = session.paragraphSavePlan();
    const parts = [...(plan.splicedParts ?? []), ...(plan.sectionSplicedParts ?? [])];
    const reasons = parts.flatMap((part) => (part.legacyParagraphs ?? []).map((legacy) => `${legacy.code}: ${legacy.detail}`));
    const patches = parts.flatMap((part) => part.paragraphPatches ?? []).length;
    const saved = await yrs.saveYrsDocx(session, { updateModifiedDate: false });
    return { reasons, patches, xml: await (await JSZip.loadAsync(saved.bytes)).file('word/document.xml')!.async('string') };
  } finally {
    session.destroy();
  }
}

function keptBesideField(source: string, saved: string | undefined): string[] {
  if (saved === undefined) return ['edited container missing from saved word/document.xml'];
  const failures: string[] = [];
  const expected = textOf(source).replace('Gamma', 'Gammax');
  if (textOf(saved) !== expected) failures.push(`text: expected ${JSON.stringify(expected)}, saved ${JSON.stringify(textOf(saved))}`);
  for (const name of KEPT_ELEMENTS) {
    const [before, after] = [countElement(source, name), countElement(saved, name)];
    if (after !== before) failures.push(`element ${name} ${after < before ? 'lost' : 'duplicated'}: source ${before}, saved ${after}`);
  }
  if (!saved.includes(' REF _Ref1 \\h ')) failures.push('field code lost');
  if (!FIELD_RESULT.test(saved)) failures.push('field result run no longer between separate and end');
  if (failures.length > 0) failures.push(`saved ${saved}`);
  return failures;
}

for (const [engine, yrs] of [['wasm', wasmYrs], ['native', nativeYrs as unknown as Yrs]] as const) {
  test(`${engine}: typing beside a field in a nested table cell with bookmarks keeps the cell`, async () => {
    const { reasons, patches, xml } = await typeBesideMarkedField(yrs, true);
    expect({ reasons: reasons.map((reason) => reason.startsWith('UNSUPPORTED_STRUCTURE: field:') ? 'field' : reason), patches }).toEqual({ reasons: ['field'], patches: 0 });
    expect(keptBesideField(outerCell(MARKED_BODY)!, outerCell(xml))).toEqual([]);
  });

  test(`${engine}: typing beside a field in a body paragraph with bookmarks keeps the paragraph`, async () => {
    const { reasons, patches, xml } = await typeBesideMarkedField(yrs, false);
    expect({ reasons: reasons.map((reason) => reason.startsWith('UNSUPPORTED_STRUCTURE: field:') ? 'field' : reason), patches }).toEqual({ reasons: ['field'], patches: 0 });
    expect(keptBesideField(`<w:p w14:paraId="${BODY_ID}">${MARKED}</w:p>`, savedParagraph(xml, BODY_ID))).toEqual([]);
  });
}
