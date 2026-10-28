import { describe, expect, it } from 'bun:test';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { SAVE_PATHS, setupStructuralSaves } from './structuralSaveTestSupport';
import { bodyMarkerBody, MC, offsetTableBody, orphanTableMarkersBody, paragraphsOf, textNodes, tocSdtBody, V, W, W14, WPS } from './__fixtures__/untouchedBytes';
import type { YrsSession } from './index';

const save = setupStructuralSaves();
const ROOT = `<w:document xmlns:w='${W}' xmlns:w14='${W14}' xmlns:mc='${MC}' xmlns:v='${V}' xmlns:wps='${WPS}'><w:body>`;
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
});

describe.each([...SAVE_PATHS])('%s source boundaries keep untouched XML', (path) => {
  it.each([
    { name: 'body bookmark end next to an empty anchor', body: bodyMarkerBody, story: 'body', anchor: '' },
    { name: 'TOC sdt after a paragraph with a page break', body: tocSdtBody, story: 'body', anchor: 'Contents anchor' },
    { name: 'orphan bookmark ends between rows and cells', body: orphanTableMarkersBody, story: 'body', anchor: 'Table anchor' },
    { name: 'synthetic table paragraph with offset and spanned rows', body: offsetTableBody, story: 'body', anchor: null },
    { name: 'cell paragraph with offset and spanned rows', body: offsetTableBody, story: 'body:t0:r0c0', anchor: 'sample' },
  ])('$name', async ({ body, story, anchor }) => {
    const source = body();
    const xml = await save(path, fixture(source), (engine) => {
      const read = engine.readParagraphs({ story, view: 'accepted' });
      if (!read.ok) throw new Error(JSON.stringify(read));
      const target = anchor === null ? read.paragraphs[1]!
        : read.paragraphs.find(({ text }) => anchor === '' ? text === '' : anchor !== null && text.startsWith(anchor))!;
      expect(target).toBeDefined();
      const result = engine.applyEdits({
        expectVersion: engine.version(),
        steps: [{ op: 'insertParagraphs', target: { story, paraId: target.paraId }, at: 'end', paragraphs: [{ text: 'inserted' }] }],
      });
      if (!result.ok) throw new Error(JSON.stringify(result));
    });
    for (const paragraph of paragraphsOf(source)) expect(xml).toContain(paragraph);
    expect(paragraphsOf(xml)).toHaveLength(paragraphsOf(source).length + 1);
    for (const tag of ['w:txbxContent', 'mc:AlternateContent', 'mc:Fallback', 'w:bookmarkStart', 'w:bookmarkEnd']) {
      expect(xml.match(new RegExp(`<${tag}\\b`, 'g'))?.length ?? 0).toBe(source.match(new RegExp(`<${tag}\\b`, 'g'))?.length ?? 0);
    }
    expect(textNodes(xml).filter(text => text !== 'inserted')).toEqual(textNodes(source));
    expect(textNodes(xml).filter(text => text === 'inserted')).toHaveLength(1);
    if (anchor === null) {
      expect(xml).toContain(source.slice(source.indexOf('<w:tbl>')));
      expect(xml.indexOf('<w:t>inserted</w:t>')).toBeGreaterThan(xml.indexOf('</w:tbl>'));
    } else if (story === 'body:t0:r0c0') {
      const sample = xml.indexOf('<w:t>sample</w:t>');
      expect(xml.indexOf('<w:t>inserted</w:t>')).toBeGreaterThan(sample);
      expect(xml.indexOf('<w:t>inserted</w:t>')).toBeLessThan(xml.indexOf('</w:tc>', sample));
    }
  });
});
