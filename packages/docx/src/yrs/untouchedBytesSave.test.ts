import { describe, expect, it } from 'bun:test';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { setupStructuralSaves, W, W14 } from './structuralSaveTestSupport';
import type { YrsSession } from './index';

const save = setupStructuralSaves();
const ROOT = `<w:document xmlns:w='${W}' xmlns:w14='${W14}'><w:body>`;
const TITLE = '<w:p w14:paraId="10000001" w:rsidR="00752A6D"><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Editor save title</w:t></w:r></w:p>';
const AFTER = '<w:p w14:paraId="10000002" w:rsidR="00752A6D"><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>after table</w:t></w:r></w:p>';
const CELL = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const TABLE = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${CELL('a')}${CELL('b')}</w:tr><w:tr>${CELL('c')}${CELL('d')}</w:tr></w:tbl>`;

function fixture(body: string): Uint8Array {
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    )],
    ['_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    )],
    ['word/document.xml', toBytes(`${ROOT}${body}</w:body></w:document>`)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

describe('editor save keeps untouched paragraphs byte for byte', () => {
  it('merge_cells through saveEditorDocument keeps the title and trailing paragraphs', async () => {
    const xml = await save('main-thread (hybrid and in-thread)', fixture(`${TITLE}${TABLE}${AFTER}`), (engine) => {
      (engine as unknown as YrsSession).mergeCells({
        anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 },
        head: { story: 'body', tableIndex: 0, row: 0, column: 1 },
      });
    });
    expect(xml).toContain(TITLE);
    expect(xml).toContain(AFTER);
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
  });

  it('a bookmark spanning paragraphs through saveEditorDocument keeps every paragraph', async () => {
    const first = '<w:p w14:paraId="10000001" w:rsidR="003E0150"><w:bookmarkStart w:id="0" w:name="_Toc1"/><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Bookmarked editor heading</w:t></w:r></w:p>';
    const second = '<w:p w14:paraId="10000002"><w:bookmarkEnd w:id="0"/><w:r><w:t>after table</w:t></w:r></w:p>';
    const xml = await save('main-thread (hybrid and in-thread)', fixture(`${first}${TABLE}${second}`), (engine) => {
      const read = engine.readParagraphs({ story: 'body', view: 'accepted' });
      if (!read.ok) throw new Error(JSON.stringify(read));
      const target = read.paragraphs.find(({ text }) => text === 'after table')!;
      const result = engine.applyEdits({
        expectVersion: engine.version(),
        steps: [{ op: 'insertParagraphs', target: { story: 'body', paraId: target.paraId }, at: 'end', paragraphs: [{ text: 'inserted' }] }],
      });
      if (!result.ok) throw new Error(JSON.stringify(result));
    });
    expect(xml).toContain(first);
    expect(xml).toContain(second);
    expect(xml).toContain('inserted');
  });
});
