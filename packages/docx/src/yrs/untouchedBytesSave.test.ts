import { describe, expect, it } from 'bun:test';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { setupStructuralSaves, W, W14 } from './structuralSaveTestSupport';
import type { YrsSession } from './index';

const save = setupStructuralSaves();
const ROOT = `<w:document xmlns:w='${W}' xmlns:w14='${W14}'><w:body>`;
const RICH = (id: string, text: string) =>
  `<w:p w14:paraId="${id}" w:rsidR="00752A6D" w:rsidRDefault="00593C67"><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr><w:proofErr w:type="spellStart"/><w:r w:rsidR="00761B10"><w:rPr><w:color w:val="auto"/><w:lang w:val="en-US"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:br/></w:r><w:proofErr w:type="spellEnd"/></w:p>`;
const TITLE = RICH('10000001', 'Editor save title');
const AFTER = RICH('10000002', 'after table');
const CELL = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r w:rsidR="00761B10"><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc>`;
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
    const first = '<w:p w14:paraId="10000001" w:rsidR="003E0150" w:rsidRDefault="00F77663" w:rsidP="00F97380"><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr><w:bookmarkStart w:id="0" w:name="_Toc256857534"/><w:bookmarkStart w:id="1" w:name="_Toc256859777"/><w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="2"/><w:proofErr w:type="spellStart"/><w:r w:rsidR="003E0150"><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Bookmarked editor heading</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>';
    const second = '<w:p w14:paraId="10000002" w:rsidR="003E0150"><w:bookmarkEnd w:id="0"/><w:bookmarkEnd w:id="1"/><w:r w:rsidR="00761B10"><w:rPr><w:color w:val="auto"/><w:lang w:val="en-US"/></w:rPr><w:t>after table</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:br/></w:r></w:p>';
    const third = '<w:p w14:paraId="10000003" w:rsidR="003E0150"><w:bookmarkStart w:id="3" w:name="_Hlk169608992"/><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Closing editor paragraph</w:t></w:r></w:p>';
    const bodyEnd = '<w:bookmarkEnd w:id="3"/>';
    const fourth = RICH('10000004', 'Final editor paragraph');
    const xml = await save('main-thread (hybrid and in-thread)', fixture(`${first}${TABLE}${second}${third}${bodyEnd}${fourth}`), (engine) => {
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
    expect(xml).toContain(third);
    expect(xml).toContain(fourth);
    expect(xml).toContain(bodyEnd);
    expect(xml).toContain('inserted');
  });
});
