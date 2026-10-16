import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

const STYLES = `<w:styles xmlns:w="${W}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="AutoSpaced"><w:name w:val="AutoSpaced"/><w:basedOn w:val="Normal"/><w:pPr><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/></w:pPr></w:style>
</w:styles>`;

function fixture(body: string): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="${OFFICE_DOC}.wordprocessingml.styles+xml"/>
</Types>`));
  parts.set('_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/>
</Relationships>`));
  parts.set('word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/>
</Relationships>`));
  parts.set('word/styles.xml', toBytes(STYLES));
  parts.set('word/document.xml', toBytes(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function documentXml(bytes: Uint8Array): string {
  const xml = readDocxContainer(bytes.buffer as ArrayBuffer).text('word/document.xml');
  expect(xml).not.toBeNull();
  return xml!;
}

async function repackAfter(bytes: Uint8Array, story: string): Promise<string> {
  const session = await createYrsSession({ clientId: 66107 });
  try {
    session.seedFromDocx(bytes);
    const base = session.materializeDocx()!;
    const paragraph = session.paragraphs(story)[0]!;
    session.insertText({ story, paraId: paragraph.paraId, offset: 0 }, 'Edited ');
    return documentXml(new Uint8Array(await repackDocx(yrsToDocument(session, base))));
  } finally {
    session.destroy();
  }
}

function styledParagraph(text: string): string {
  return `<w:p><w:pPr><w:pStyle w:val="AutoSpaced"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function contentControl(content: string): string {
  return `<w:sdt><w:sdtPr><w:id w:val="1"/><w:tag w:val="cc"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}

function table(content: string): string {
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>${content}</w:tc></w:tr></w:tbl>`;
}

function expectInherited(xml: string, text: string): void {
  expect(xml).toContain(`Edited ${text}`);
  expect(xml).toContain('<w:pStyle w:val="AutoSpaced"/>');
  expect(xml).not.toContain('<w:autoSpaceDE');
  expect(xml).not.toContain('<w:autoSpaceDN');
  expect(xml).not.toContain('<w:adjustRightInd');
}

test('a content-control paragraph keeps style adjustRightInd and auto-spacing inherited on save', async () => {
  const bytes = fixture(contentControl(styledParagraph('Controlled QA')));
  const xml = await repackAfter(bytes, 'body:sdt0');
  expectInherited(xml, 'Controlled QA');
});

test('a table-cell paragraph inside a content control keeps style adjustRightInd and auto-spacing inherited on save', async () => {
  const bytes = fixture(contentControl(table(styledParagraph('Celled QA'))));
  const xml = await repackAfter(bytes, 'body:sdt0:t0:r0c0');
  expectInherited(xml, 'Celled QA');
});

test('a content-control paragraph inside a table cell keeps style adjustRightInd and auto-spacing inherited on save', async () => {
  const bytes = fixture(table(contentControl(styledParagraph('Nested QA'))));
  const xml = await repackAfter(bytes, 'body:t0:r0c0:sdt0');
  expectInherited(xml, 'Nested QA');
});
