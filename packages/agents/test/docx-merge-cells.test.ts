import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx } from '../src';
import { tableCells } from '../src/docx-edits';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const cell = (text: string, extra = '') => `<w:tc><w:tcPr/>${extra || `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`}</w:tc>`;
const row = (cells: string, properties = '') => `<w:tr>${properties}${cells}</w:tr>`;
const table = (rows: string, columns: number) => `<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(columns)}</w:tblGrid>${rows}</w:tbl>`;
const merge = (left: number, right: number) => ({
  anchor: { story: 'body', tableIndex: 0, row: 0, column: left },
  head: { story: 'body', tableIndex: 0, row: 1, column: right },
});

async function source(body: string, comments = false) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>${comments ? `<Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/>` : ''}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  if (comments) {
    zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="comments" Type="${R}/comments" Target="comments.xml"/></Relationships>`);
    zip.file('word/comments.xml', `<w:comments xmlns:w="${W}"><w:comment w:id="7" w:author="Ada"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>`);
  }
  return zip.generateAsync({ type: 'uint8array' });
}

const gaps = '<w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>';
const fullRow = row(['Left', 'A', 'B', 'Right'].map(text => cell(text)).join(''));
const gapRow = row(cell('C') + cell('D'), gaps);

test('agent table positions include leading and trailing omitted columns', async () => {
  const doc = await openDocx(await source(table(gapRow, 0)));
  try {
    const cells = tableCells('body', doc.session.storySegments('body'));
    expect([...cells.values()].map(value => [value.anchor.column, value.columns])).toEqual([[1, 4], [2, 4]]);
    expect(doc.list().items.filter(item => item.table).map(item => item.table!.column)).toEqual([2, 3]);
  } finally { doc.close(); }
});

test('merge maps gridBefore selections and retains all cell text through export', async () => {
  const untouched = row(['Keep1', 'Keep2', 'Keep3', 'Keep4'].map(text => cell(text)).join(''));
  const doc = await openDocx(await source(table(fullRow + gapRow + untouched, 4)));
  try {
    doc.session.mergeCells(merge(1, 2));
    expect(doc.session.paragraphs('body:t0:r0c1').map(paragraph => paragraph.text)).toEqual(['A', 'B', 'C', 'D']);
    const anchors = [...tableCells('body', doc.session.storySegments('body')).values()];
    expect(anchors.find(value => value.anchor.story === 'body:t0:r0c1')?.anchor).toMatchObject({
      row: 0, column: 1, rowspan: 2, colspan: 2,
    });
    const saved = await doc.export();
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
    expect(xml).toContain(gaps);
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    const reopened = await openDocx(saved);
    try {
      const stories = [...tableCells('body', reopened.session.storySegments('body')).keys()];
      expect(stories.flatMap(story => reopened.session.paragraphs(story).map(paragraph => paragraph.text)).filter(Boolean)).toEqual([
        'Left', 'A', 'B', 'C', 'D', 'Right', 'Keep1', 'Keep2', 'Keep3', 'Keep4',
      ]);
    } finally { reopened.close(); }
  } finally { doc.close(); }
});

test.todo('merge maps a 2x2 block after gridBefore and exports only the selected cells as merged (base serializer rewrites unselected cells)', async () => {
  const untouched = row(['Keep1', 'Keep2', 'Keep3', 'Keep4'].map(text => cell(text)).join(''));
  const doc = await openDocx(await source(table(fullRow + gapRow + untouched, 4)));
  try {
    doc.session.mergeCells(merge(1, 2));
    expect(doc.session.paragraphs('body:t0:r0c1').map(paragraph => paragraph.text)).toEqual(['A', 'B', 'C', 'D']);
    const saved = await doc.export();
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
    const rows = xml.match(/<w:tr\b[\s\S]*?<\/w:tr>/g)!;
    const cells = rows.map(value => value.match(/<w:tc\b[\s\S]*?<\/w:tc>/g) ?? []);
    expect(cells[0]).toHaveLength(3);
    expect(cells[0][0]).toBe(cell('Left'));
    expect(cells[0][2]).toBe(cell('Right'));
    expect(rows[2]).toBe(untouched);
    expect(cells[0][1]).toContain('<w:gridSpan w:val="2"/>');
    expect(cells[0][1]).toContain('<w:vMerge w:val="restart"/>');
    expect(cells[1]).toHaveLength(1);
    expect(cells[1][0]).toContain('<w:gridSpan w:val="2"/>');
    expect(cells[1][0]).toMatch(/<w:vMerge(?: w:val="continue")?\s*\/>/);
    expect(rows[1]).toContain(gaps);
    const reopened = await openDocx(saved);
    try {
      const anchors = [...tableCells('body', reopened.session.storySegments('body')).values()];
      expect(anchors.find(value => value.anchor.row === 0 && value.anchor.column === 1)?.anchor).toMatchObject({ rowspan: 2, colspan: 2 });
    } finally { reopened.close(); }
  } finally { doc.close(); }
});

test('merge moves a source cell comment range and reference with its UTF-16 text', async () => {
  const commented = '<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>B😀</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>';
  const doc = await openDocx(await source(table(row(cell('A') + cell('', commented)) + row(cell('C') + cell('D')), 2), true));
  try {
    doc.session.mergeCells(merge(0, 1));
    expect(doc.session.resolveComment('7')).toEqual([{ story: 'body:t0:r0c0', start: 2, end: 5 }]);
    const saved = await doc.export();
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
    for (const tag of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
      expect(xml.match(new RegExp(`<w:${tag}\\b`, 'g'))).toHaveLength(1);
    }
    expect(xml).toContain('B😀');
    const reopened = await openDocx(saved);
    try {
      expect(reopened.session.resolveComment('7')).toEqual([{ story: 'body:t0:r0c0', start: 2, end: 5 }]);
    } finally { reopened.close(); }
  } finally { doc.close(); }
});

test.todo('merged comment survivor retains its full-width rowspan after reopen (requires reverted continuation reseeding)', async () => {
  const commented = '<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>B😀</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>';
  const doc = await openDocx(await source(table(row(cell('A') + cell('', commented)) + row(cell('C') + cell('D')), 2), true));
  try {
    doc.session.mergeCells(merge(0, 1));
    expect(doc.session.resolveComment('7')).toEqual([{ story: 'body:t0:r0c0', start: 2, end: 5 }]);
    const saved = await doc.export();
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
    for (const tag of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
      expect(xml.match(new RegExp(`<w:${tag}\\b`, 'g'))).toHaveLength(1);
    }
    expect(xml).toContain('B😀');
    const reopened = await openDocx(saved);
    try {
      expect(reopened.session.resolveComment('7')).toEqual([{ story: 'body:t0:r0c0', start: 2, end: 5 }]);
      const survivor = [...tableCells('body', reopened.session.storySegments('body')).values()][0];
      expect(survivor.anchor).toMatchObject({ rowspan: 2, colspan: 2 });
    } finally { reopened.close(); }
  } finally { doc.close(); }
});
