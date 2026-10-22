import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocumentToolError, type DocxAgentDocument, type DocxAgentEdit } from '../src';
import { readCommentPart, verifyCommentOutput } from '../src/comment-output';
import { COMMENT_COMPANIONS, commentFixture, editFixture, EXISTING_COMMENT, fixture, tableCell, tableFixture } from './fixture';

const opened: DocxAgentDocument[] = [];
async function open(bytes?: Uint8Array) {
  const doc = await openDocx(bytes ?? await editFixture());
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

const hit = (doc: DocxAgentDocument, query: string) => doc.grep({ query }).matches[0];
const rows = (xml: string) => xml.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) ?? [];

async function paragraphs(bytes: Uint8Array) {
  const doc = await open(bytes);
  return doc.list({ limit: 100 }).items.map(item => {
    const { text, style, runs, table } = doc.read(item.ref);
    return { text, style, runs, table };
  });
}
type Snapshot = Awaited<ReturnType<typeof paragraphs>>;
const content = (items: Snapshot) => items.map(({ text, style, runs }) => ({ text, style, runs }));

async function edited(doc: DocxAgentDocument, edits: DocxAgentEdit[]) {
  const baseline = await doc.export();
  const receipt = await doc.edit({ version: doc.list().version, edits });
  const bytes = await doc.export();
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
  expect(xml).not.toMatch(/<w:(ins|del)\b/);
  return { receipt, xml, bytes, before: await paragraphs(baseline), after: await paragraphs(bytes) };
}

async function refusal(promise: Promise<unknown>): Promise<DocumentToolError> {
  try { await promise; } catch (error) { return error as DocumentToolError; }
  throw new Error('Expected the batch to be refused.');
}

test('set_paragraph_style applies a document style by name and changes nothing else', async () => {
  const doc = await open();
  const target = hit(doc, 'Revenue grew');
  const { receipt, xml, before, after } = await edited(doc, [{ op: 'set_paragraph_style', paragraph: target.ref, style: 'heading 2' }]);
  expect(receipt).toEqual({ applied: true, version: doc.list().version, results: [{ index: 0, op: 'set_paragraph_style', ref: target.ref, style: 'Heading2' }] });
  const index = before.findIndex(p => p.text === 'Revenue grew strongly in the third quarter.');
  expect(after[index].style).toBe('Heading2');
  expect(after.filter((_, i) => i !== index)).toEqual(before.filter((_, i) => i !== index));
  expect(xml).toContain('<w:pStyle w:val="Heading2"/>');
  expect(doc.read(target.ref).style).toBe('Heading2');
  expect(doc.session.undo()).toBe(true);
  expect(doc.read(target.ref).style).not.toBe('Heading2');
});

test('format_text sets and clears marks on exactly the matched range', async () => {
  const doc = await open();
  const match = hit(doc, 'strongly');
  const { before, after } = await edited(doc, [{ op: 'format_text', match: match.match, bold: true, underline: true }]);
  const index = before.findIndex(p => p.text.startsWith('Revenue grew'));
  const spans = (key: string) => after[index].runs.filter(run => run.formatting[key]).map(run => [run.start, run.end]);
  expect(spans('bold')).toEqual([[match.start, match.end]]);
  expect(spans('underline')).toEqual([[match.start, match.end]]);
  expect(spans('italic')).toEqual([]);
  const others = (formatting: Record<string, unknown>) => Object.fromEntries(Object.entries(formatting).filter(([key]) => key !== 'bold' && key !== 'underline'));
  for (const run of after[index].runs) expect(others(run.formatting)).toEqual(before[index].runs[0].formatting);
  expect(after[index].text).toBe(before[index].text);
  expect(after.filter((_, i) => i !== index)).toEqual(before.filter((_, i) => i !== index));
  const stale = await refusal(doc.edit({ version: doc.list().version, edits: [{ op: 'format_text', match: match.match, italic: true }] }));
  expect(stale.code).toBe('STALE_TARGET');
  const fresh = doc.grep({ query: 'strongly' });
  await doc.edit({ version: fresh.version, edits: [{ op: 'format_text', match: fresh.matches[0].match, bold: false }] });
  const runs = doc.read(match.ref).runs;
  expect(runs.some(run => run.formatting.bold)).toBe(false);
  expect(runs.filter(run => run.formatting.underline).map(run => [run.start, run.end])).toEqual([[match.start, match.end]]);
});

test('insert_paragraphs adds body-default paragraphs after a heading and delete_paragraph removes one', async () => {
  const doc = await open();
  const heading = hit(doc, 'Quarterly report');
  const obsolete = hit(doc, 'Obsolete remark');
  const { receipt, before, after } = await edited(doc, [
    { op: 'insert_paragraphs', after: heading.ref, paragraphs: [{ text: 'Added line.' }, { text: 'Quoted line.', style: 'Quote' }] },
    { op: 'delete_paragraph', paragraph: obsolete.ref },
  ]);
  const refs = receipt.results[0].refs as string[];
  expect(refs.map(ref => doc.read(ref).text)).toEqual(['Added line.', 'Quoted line.']);
  expect(receipt.results[1]).toEqual({ index: 1, op: 'delete_paragraph', ref: obsolete.ref });
  expect(() => doc.read(obsolete.ref)).toThrow('Paragraph missing');
  const body = (items: Snapshot) => items.filter(p => !p.table).map(p => p.text);
  const expected = body(before).filter(text => text !== 'Obsolete remark.');
  expected.splice(1, 0, 'Added line.', 'Quoted line.');
  expect(body(after)).toEqual(expected);
  expect([null, 'Normal']).toContain(after.find(p => p.text === 'Added line.')!.style);
  expect(after.find(p => p.text === 'Quoted line.')!.style).toBe('Quote');
  expect(after.filter(p => p.text !== 'Added line.' && p.text !== 'Quoted line.')).toEqual(before.filter(p => p.text !== 'Obsolete remark.'));
});

test('insert_table_row copies the anchor row, its cell properties and first-run formatting', async () => {
  const doc = await open();
  const q1 = hit(doc, 'Q1');
  const position = { story: 'body', index: 1, row: 2, column: 1, rows: 2, columns: 2 };
  expect(doc.read(q1.ref).table).toEqual(position);
  expect(doc.list({ limit: 100 }).items.find(item => item.ref === q1.ref)?.table).toEqual(position);
  const { receipt, xml, before, after } = await edited(doc, [{ op: 'insert_table_row', row: q1.ref, position: 'after', cells: ['Q2', '1.4M'] }]);
  const [first, second] = receipt.results[0].refs as string[];
  expect(doc.read(first)).toMatchObject({ text: 'Q2', table: { row: 3, column: 1, rows: 3, columns: 2 } });
  expect(doc.read(second)).toMatchObject({ text: '1.4M', table: { row: 3, column: 2 } });
  const grid = (items: Snapshot) => items.filter(p => p.table).sort((a, b) => a.table!.row - b.table!.row || a.table!.column - b.table!.column).map(p => p.text);
  expect(grid(after)).toEqual(['Quarter', 'Revenue', 'Q1', '1.0M', 'Q2', '1.4M']);
  const run = (text: string) => after.find(p => p.text === text)!.runs[0].formatting;
  expect(run('Q2')).toEqual(run('Q1'));
  expect(run('1.4M')).toEqual(run('1.0M'));
  expect(run('Q2').italic).toBe(true);
  expect(content(after.filter(p => p.text !== 'Q2' && p.text !== '1.4M'))).toEqual(content(before));
  const tableRows = rows(xml);
  expect(tableRows).toHaveLength(3);
  expect(tableRows[2]).toMatch(/<w:trHeight [^>]*w:val="500"/);
  expect(tableRows[2]).toMatch(/w:fill="DDEEFF"[\s\S]*w:fill="FFEEDD"/i);
  expect(tableRows[2]).toMatch(/<w:jc w:val="(right|end)"\/>/);

  const outline = doc.list({ limit: 100 });
  const header = outline.items.find(item => item.text === 'Quarter')!;
  const top = await doc.edit({ version: outline.version, edits: [{ op: 'insert_table_row', row: header.ref, position: 'before', cells: ['Top', 'Row'] }] });
  const [topRef] = top.results[0].refs as string[];
  expect(doc.read(topRef)).toMatchObject({ text: 'Top', table: { row: 1, rows: 4 } });
  expect(doc.read(topRef).runs[0].formatting.bold).toBe(true);
  expect(doc.read(header.ref).table).toMatchObject({ row: 2 });
});

test('insert_table_row copies each template paragraph style, indentation and spacing', async () => {
  const first = '<w:pPr><w:pStyle w:val="Quote"/><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="720" w:right="180" w:firstLine="240"/><w:jc w:val="center"/></w:pPr>';
  const second = '<w:pPr><w:pStyle w:val="Heading2"/><w:spacing w:before="60" w:after="180"/><w:ind w:left="360" w:hanging="120"/><w:jc w:val="right"/></w:pPr>';
  const doc = await open(await tableFixture([
    tableCell('Header') + tableCell('Value'),
    tableCell('Template', '', first) + tableCell('Amount', '', second),
  ]));
  const { receipt, xml, before, after } = await edited(doc, [
    { op: 'insert_table_row', row: hit(doc, 'Template').ref, position: 'after', cells: ['Copied', 'Cell'] },
  ]);
  const refs = receipt.results[0].refs as string[];
  expect(refs.map(ref => doc.read(ref).style)).toEqual(['Quote', 'Heading2']);
  const properties = (row: string) => row.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/g);
  const tableRows = rows(xml);
  expect(properties(tableRows[1])).toHaveLength(2);
  expect(properties(tableRows[2])).toEqual(properties(tableRows[1]));
  expect(tableRows[2]).toContain('<w:ind');
  expect(tableRows[2]).toContain('<w:spacing');
  expect(after.find(p => p.text === 'Copied')?.style).toBe('Quote');
  expect(after.find(p => p.text === 'Cell')?.style).toBe('Heading2');
  expect(content(after.filter(p => p.text !== 'Copied' && p.text !== 'Cell'))).toEqual(content(before));
});

test('delete_table_row removes one row; refusals name the problem and change nothing', async () => {
  const doc = await open();
  const q1 = hit(doc, 'Q1');
  const heading = hit(doc, 'Quarterly report');
  const version = doc.list().version;
  const refused = async (edits: DocxAgentEdit[], code: string) => {
    const error = await refusal(doc.edit({ version, edits }));
    expect(error.code).toBe(code);
    return error;
  };
  await refused([{ op: 'delete_paragraph', paragraph: 'p999' }], 'UNKNOWN_REF');
  await refused([{ op: 'format_text', match: 'm999', bold: true }], 'UNKNOWN_MATCH');
  await refused([{ op: 'format_text', match: q1.match }], 'INVALID_ARGUMENT');
  const style = await refused([{ op: 'set_paragraph_style', paragraph: heading.ref, style: 'Heading 9' }], 'UNKNOWN_STYLE');
  expect(style.details?.validStyles).toEqual(expect.arrayContaining(['Heading1', 'Heading2']));
  expect(style.details?.validStyles).not.toContain('Strong');
  expect((await refused([{ op: 'insert_table_row', row: q1.ref, position: 'after', cells: ['Q2'] }], 'CELL_COUNT')).message).toContain('2 columns');
  await refused([{ op: 'insert_table_row', row: heading.ref, position: 'after', cells: ['a', 'b'] }], 'NOT_IN_TABLE');
  await refused([{ op: 'delete_paragraph', paragraph: q1.ref }], 'LAST_PARAGRAPH');
  const atomic = await refused([{ op: 'set_paragraph_style', paragraph: heading.ref, style: 'Heading2' }, { op: 'delete_paragraph', paragraph: q1.ref }], 'LAST_PARAGRAPH');
  expect(atomic.details).toMatchObject({ failedIndex: 1 });
  expect(doc.list().version).toBe(version);
  expect(doc.read(heading.ref).style).toBe('Heading1');

  const { receipt, xml, before, after } = await edited(doc, [{ op: 'delete_table_row', row: q1.ref }]);
  expect(receipt.results).toEqual([{ index: 0, op: 'delete_table_row', ref: q1.ref, deletedTable: false }]);
  expect(() => doc.read(q1.ref)).toThrow('Paragraph missing');
  expect(rows(xml)).toHaveLength(1);
  expect(xml).not.toContain('1.0M');
  expect(content(after)).toEqual(content(before.filter(p => p.text !== 'Q1' && p.text !== '1.0M')));
  expect((await refusal(doc.edit({ version, edits: [{ op: 'delete_paragraph', paragraph: heading.ref }] }))).code).toBe('STALE_VERSION');
});

test('table row edits refuse vertical merges without applying earlier batch edits', async () => {
  const variants = [
    [tableCell('Merged', '<w:vMerge w:val="restart"/>') + tableCell('Top'), tableCell('', '<w:vMerge/>') + tableCell('Bottom')],
    [tableCell('Merged', '<w:vMerge w:val="restart"/>') + tableCell('Top'), tableCell('Plain') + tableCell('Bottom')],
    [tableCell('Merged', '<w:vMerge/>') + tableCell('Top'), tableCell('Plain') + tableCell('Bottom')],
  ];
  for (const [index, tableRows] of variants.entries()) {
    const doc = await open(await tableFixture(tableRows));
    const targets = index === 0 ? ['Merged', 'Top', 'Bottom'] : ['Merged', 'Top'];
    const heading = hit(doc, 'Quarterly report');
    const version = doc.list().version;
    const baseline = await (await JSZip.loadAsync(await doc.export())).file('word/document.xml')!.async('string');
    for (const text of targets) {
      const row = hit(doc, text).ref;
      const edits: DocxAgentEdit[] = [
        { op: 'delete_table_row', row },
        { op: 'insert_table_row', row, position: 'before', cells: ['New', 'Row'] },
        { op: 'insert_table_row', row, position: 'after', cells: ['New', 'Row'] },
      ];
      for (const edit of edits) {
        const error = await refusal(doc.edit({ version, edits: [
          { op: 'set_paragraph_style', paragraph: heading.ref, style: 'Heading2' }, edit,
        ] }));
        expect(error.code).toBe('UNSUPPORTED_TARGET');
        expect(error.details).toMatchObject({ failedIndex: 1 });
        expect(doc.list().version).toBe(version);
        expect(doc.read(heading.ref).style).toBe('Heading1');
        const xml = await (await JSZip.loadAsync(await doc.export())).file('word/document.xml')!.async('string');
        expect(xml).toBe(baseline);
      }
    }
  }
});

test('delete_table_row refuses the last row without changing the document', async () => {
  const doc = await open();
  await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_table_row', row: hit(doc, 'Q1').ref }] });
  const outline = doc.list({ limit: 100 });
  const header = outline.items.find(item => item.text === 'Quarter')!;
  const version = outline.version;
  const baseline = await paragraphs(await doc.export());
  const error = await refusal(doc.edit({ version, edits: [{ op: 'delete_table_row', row: header.ref }] }));
  expect(error.code).toBe('UNSUPPORTED_TARGET');
  expect(doc.list().version).toBe(version);
  expect(await paragraphs(await doc.export())).toEqual(baseline);
});

test('outline keeps code-unit order for header stories with case and punctuation', async () => {
  const zip = await JSZip.loadAsync(await editFixture());
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ids = ['rIda', 'rId_A', 'rIdA'];
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', ids.map((_, index) =>
    `<Override PartName="/word/header${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`).join('') + '</Types>'));
  const relationships = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', relationships.replace('</Relationships>', ids.map((id, index) =>
    `<Relationship Id="${id}" Type="${r}/header" Target="header${index + 1}.xml"/>`).join('') + '</Relationships>'));
  const xml = await zip.file('word/document.xml')!.async('string');
  const variants = ['default', 'first', 'even'];
  zip.file('word/document.xml', xml.replace('<w:sectPr>', '<w:sectPr>' + ids.map((id, index) =>
    `<w:headerReference w:type="${variants[index]}" r:id="${id}"/>`).join('')));
  for (const [index, id] of ids.entries()) {
    zip.file(`word/header${index + 1}.xml`, `<w:hdr xmlns:w="${w}"><w:p><w:r><w:t>${id}</w:t></w:r></w:p></w:hdr>`);
  }
  const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  const expected = ['hf:rIdA', 'hf:rId_A', 'hf:rIda'];
  for (const document of [doc, await open(await doc.export())]) {
    expect(document.list({ limit: 100 }).items.filter(item => item.story.startsWith('hf:')).map(item => item.story)).toEqual(expected);
    expect(document.overview().stories.filter(item => item.story.startsWith('hf:')).map(item => item.story)).toEqual(expected);
  }
});

test('outline lists table cells by numeric row and column before and after insertion', async () => {
  const tableRows = Array.from({ length: 11 }, (_, row) =>
    Array.from({ length: 11 }, (_, column) => tableCell(`Row ${row + 1} Cell ${column + 1}`)).join(''));
  const doc = await open(await tableFixture(tableRows));
  const expected = Array.from({ length: 11 }, (_, row) =>
    Array.from({ length: 11 }, (_, column) => `Row ${row + 1} Cell ${column + 1}`)).flat();
  const table = (document = doc) => {
    const items: ReturnType<DocxAgentDocument['list']>['items'] = [];
    const total = document.overview().paragraphs;
    for (let offset = 0; offset < total; offset += 100) items.push(...document.list({ offset, limit: 100 }).items);
    return items.filter(item => item.table);
  };
  expect(table().map(item => item.text)).toEqual(expected);
  expect(table().map(item => [item.table!.row, item.table!.column])).toEqual(
    Array.from({ length: 11 }, (_, row) => Array.from({ length: 11 }, (_, column) => [row + 1, column + 1])).flat());
  const anchor = table().find(item => item.text === 'Row 2 Cell 1')!;
  const cells = Array.from({ length: 11 }, (_, column) => `Inserted Cell ${column + 1}`);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_table_row', row: anchor.ref, position: 'before', cells }] });
  expected.splice(11, 0, ...cells);
  expect(table().map(item => item.text)).toEqual(expected);
  const reopened = await open(await doc.export());
  expect(table(reopened).map(item => item.text)).toEqual(expected);
});

const part = async (bytes: Uint8Array, name: string) => (await JSZip.loadAsync(bytes)).file(name)?.async('string');
const texts = (xml: string) => [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join('');

function companionEntries(xml: string, entry: string) {
  const tag = entry.slice(1, entry.indexOf(' '));
  return [...xml.matchAll(new RegExp(`<${tag}\\s([^>]*?)(?:\\s*/>|>\\s*</${tag}>)`, 'g'))].map(([, attributes]) =>
    new Map([...attributes.matchAll(/([\w:]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)].map(([, name, double, single]) => [name, double ?? single])));
}

function comments(xml: string) {
  return [...xml.matchAll(/<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g)].map(([, attributes, body]) => ({
    id: attributes.match(/w:id="([^"]*)"/)![1],
    author: attributes.match(/w:author="([^"]*)"/)?.[1],
    date: attributes.match(/w:date="([^"]*)"/)?.[1],
    paragraphs: [...body.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map(([paragraph]) => texts(paragraph)),
  }));
}

function anchored(xml: string, id: string) {
  const marker = (name: string) => `<w:comment${name} w:id="${id}"/>`;
  for (const name of ['RangeStart', 'RangeEnd', 'Reference']) expect(xml.split(marker(name))).toHaveLength(2);
  const start = xml.indexOf(marker('RangeStart'));
  const end = xml.indexOf(marker('RangeEnd'));
  const reference = xml.indexOf(marker('Reference'));
  expect(reference).toBeGreaterThan(end);
  expect(xml.slice(start, reference)).not.toContain('</w:p>');
  return texts(xml.slice(start, end));
}

test('add_comment anchors a comment on exactly the matched phrase, one paragraph per line', async () => {
  const doc = await open();
  const match = hit(doc, 'strongly in the third');
  expect(await part(await doc.export(), 'word/comments.xml')).toBeUndefined();
  const { receipt, xml, bytes, before, after } = await edited(doc, [
    { op: 'add_comment', match: match.match, author: 'Reviewer', text: 'Check the source.\nCompare with Q2.', date: '2026-10-01T09:30:00Z' },
  ]);
  expect(receipt).toEqual({ applied: true, version: doc.list().version, results: [{ index: 0, op: 'add_comment', ref: match.ref, start: match.start, end: match.end, comment: expect.any(String) }] });
  expect(doc.read(match.ref).text).toBe('Revenue grew strongly in the third quarter.');
  const reference = before.findIndex(p => p.text.startsWith('Revenue grew'));
  expect(after[reference].text).toBe('Revenue grew strongly in the third\ufffc quarter.');
  expect(after.map(p => p.style)).toEqual(before.map(p => p.style));
  expect(after.filter((_, i) => i !== reference)).toEqual(before.filter((_, i) => i !== reference));
  const [comment, ...rest] = comments((await part(bytes, 'word/comments.xml'))!);
  expect(rest).toEqual([]);
  expect(comment).toMatchObject({ author: 'Reviewer', date: '2026-10-01T09:30:00Z', paragraphs: ['Check the source.', 'Compare with Q2.'] });
  expect(anchored(xml, comment.id)).toBe('strongly in the third');
});

test('add_comment keeps existing comments byte-identical and anchors in body and table cells', async () => {
  const source = await commentFixture();
  expect(await part(source, 'word/comments.xml')).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
  const doc = await open(source);
  expect(await part(await doc.export(), 'word/comments.xml')).toBe(await part(source, 'word/comments.xml'));
  const revenue = hit(doc, 'Revenue grew');
  const q1 = hit(doc, 'Q1');
  const started = Math.floor(Date.now() / 1000) * 1000;
  const { receipt, xml, bytes } = await edited(doc, [
    { op: 'add_comment', match: revenue.match, author: 'Reviewer', text: 'Body note', date: '2026-10-01T09:30:00Z' },
    { op: 'add_comment', match: q1.match, author: 'Reviewer', text: 'Cell note' },
  ]);
  expect(new Set(receipt.results.map(result => result.comment)).size).toBe(2);
  const saved = (await part(bytes, 'word/comments.xml'))!;
  expect(saved).toContain(EXISTING_COMMENT);
  const all = comments(saved);
  expect(all.map(comment => [comment.author, comment.paragraphs])).toEqual([['Existing Reviewer', ['Keep this note.']], ['Reviewer', ['Body note']], ['Reviewer', ['Cell note']]]);
  const [existing, body, cell] = all;
  expect(existing.id).toBe('0');
  expect(anchored(xml, '0')).toBe('Closing note.');
  expect(anchored(xml, body.id)).toBe('Revenue grew');
  expect(anchored(xml, cell.id)).toBe('Q1');
  expect(cell.date).toMatch(/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ$/);
  expect(Date.parse(cell.date!)).toBeGreaterThanOrEqual(started);
  expect(Date.parse(cell.date!)).toBeLessThanOrEqual(Date.now());
  expect(await part(bytes, 'word/styles.xml')).toBe(await part(source, 'word/styles.xml'));
  for (const [name, , , , entry] of COMMENT_COMPANIONS) {
    expect(companionEntries((await part(bytes, `word/${name}.xml`))!, entry)).toContainEqual(companionEntries(entry, entry)[0]);
  }
});

test('add_comment preserves prefixed relationships and unrelated bytes across batches', async () => {
  const zip = await JSZip.loadAsync(await editFixture());
  const namespace = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const name = 'word/_rels/document.xml.rels';
  const relationship = `<pkg:Relationship Id = 'rId7' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/>`;
  zip.file(name, `<Relationships xmlns="${namespace}" xmlns:pkg='${namespace}'>${relationship}</Relationships>`);
  const root = await zip.file('_rels/.rels')!.async('string');
  zip.file('_rels/.rels', root.replace('<Relationships xmlns=', '<pkg:Relationships xmlns:pkg=').replace('<Relationship ', '<pkg:Relationship ').replace('</Relationships>', '</pkg:Relationships>'));
  const payload = new Uint8Array([0, 255, 128, 32, 13, 10]);
  zip.file('customXml/preserved.bin', payload);
  const source = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const doc = await open(source);
  for (const query of ['Revenue grew', 'Closing note.']) {
    await doc.edit({ version: doc.list().version, edits: [{ op: 'add_comment', match: hit(doc, query).match, author: 'Reviewer', text: 'Note' }] });
  }
  const bytes = await doc.export();
  expect(await part(bytes, name)).toContain(relationship);
  expect(await part(bytes, '_rels/.rels')).toBe(await part(source, '_rels/.rels'));
  const saved = await JSZip.loadAsync(bytes);
  expect(await saved.file('customXml/preserved.bin')!.async('uint8array')).toEqual(payload);
  expect(comments((await part(bytes, 'word/comments.xml'))!)).toHaveLength(2);
});

test('add_comment refuses a body content control without applying any batch edits', async () => {
  const zip = await JSZip.loadAsync(await editFixture());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:p><w:r><w:t>Obsolete remark.</w:t></w:r></w:p>',
    '<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Obsolete remark.</w:t></w:r></w:p></w:sdtContent></w:sdt>'));
  const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  const match = hit(doc, 'remark');
  expect(match.story).not.toBe('body');
  const version = doc.list().version;
  const state = doc.session.encodeState();
  const baseline = await doc.export();
  const error = await refusal(doc.edit({ version, edits: [
    { op: 'format_text', match: hit(doc, 'Revenue grew').match, bold: true },
    { op: 'add_comment', match: match.match, author: 'Reviewer', text: 'Inside a control' },
  ] }));
  expect(error.code).toBe('UNSUPPORTED_TARGET');
  expect(error.message).toContain('Comments anchor in the main body text and its tables, not inside content controls, headers, footers, notes or text boxes.');
  expect(error.details).toMatchObject({ failedIndex: 1 });
  expect(doc.list().version).toBe(version);
  expect(doc.session.encodeState()).toEqual(state);
  expect(doc.session.listComments()).toEqual([]);
  const bytes = await doc.export();
  expect(await part(bytes, 'word/document.xml')).toBe(await part(baseline, 'word/document.xml'));
  expect(await part(bytes, 'word/comments.xml')).toBeUndefined();
});

test('add_comment refuses an existing comments part without a w14 root binding atomically', async () => {
  const namespace = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
  const source = await commentFixture();
  const original = (await part(source, 'word/comments.xml'))!;
  for (const existing of [
    original.replace(` ${namespace}`, '').replace(/ w14:(?:paraId|textId)="[^"]*"/g, ''),
    original.replace(` ${namespace}`, '').replace('<w:comment ', `<w:comment ${namespace} `),
  ]) {
    const zip = await JSZip.loadAsync(source);
    zip.file('word/comments.xml', existing);
    const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
    const version = doc.list().version;
    const state = doc.session.encodeState();
    const before = doc.session.listComments();
    const baseline = await doc.export();
    const error = await refusal(doc.edit({ version, edits: [
      { op: 'format_text', match: hit(doc, 'Revenue grew').match, bold: true },
      { op: 'add_comment', match: hit(doc, 'strongly').match, author: 'Reviewer', text: 'New note' },
    ] }));
    expect(error.code).toBe('UNSUPPORTED_TARGET');
    expect(error.message).toContain('The existing comments part predates paragraph ids, so a new comment cannot be added without rewriting existing comments.');
    expect(error.details).toMatchObject({ failedIndex: 1 });
    expect(doc.list().version).toBe(version);
    expect(doc.session.encodeState()).toEqual(state);
    expect(doc.session.listComments()).toEqual(before);
    const bytes = await doc.export();
    expect(await part(bytes, 'word/document.xml')).toBe(await part(baseline, 'word/document.xml'));
    expect(await part(bytes, 'word/comments.xml')).toBe(existing);
  }
});

test('add_comment refuses a comments part stored under another name atomically', async () => {
  const zip = await JSZip.loadAsync(await commentFixture());
  const rename = async (name: string, from: string, to: string) => zip.file(name, (await zip.file(name)!.async('string')).replace(from, to));
  zip.file('word/comments1.xml', await zip.file('word/comments.xml')!.async('string'));
  zip.remove('word/comments.xml');
  await rename('[Content_Types].xml', '/word/comments.xml', '/word/comments1.xml');
  await rename('word/_rels/document.xml.rels', 'Target="comments.xml"', 'Target="comments1.xml"');
  const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  const version = doc.list().version;
  const state = doc.session.encodeState();
  const error = await refusal(doc.edit({ version, edits: [
    { op: 'add_comment', match: hit(doc, 'strongly').match, author: 'Reviewer', text: 'New note' },
  ] }));
  expect(error.code).toBe('UNSUPPORTED_TARGET');
  expect(error.message).toContain('non-standard part');
  expect(error.details).toMatchObject({ failedIndex: 0 });
  expect(doc.list().version).toBe(version);
  expect(doc.session.encodeState()).toEqual(state);
});

async function refusedCommentOutput(zip: JSZip) {
  const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  const version = doc.list().version;
  const state = doc.session.encodeState();
  const before = doc.session.listComments();
  const baseline = await doc.export();
  const error = await refusal(doc.edit({ version, edits: [
    { op: 'format_text', match: hit(doc, 'Revenue grew').match, bold: true },
    { op: 'add_comment', match: hit(doc, 'strongly').match, author: 'Reviewer', text: 'New note' },
    { op: 'add_comment', match: hit(doc, 'Q1').match, author: 'Reviewer', text: 'Cell note' },
    { op: 'format_text', match: hit(doc, 'Obsolete remark').match, italic: true },
  ] }));
  expect(error.code).toBe('UNSUPPORTED_TARGET');
  expect(error.details).toMatchObject({ failedIndex: 1 });
  expect(doc.list().version).toBe(version);
  expect(doc.session.encodeState()).toEqual(state);
  expect(doc.session.listComments()).toEqual(before);
  expect(await doc.export()).toEqual(baseline);
  return error;
}

async function acceptedCommentOutput(zip: JSZip) {
  const source = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const doc = await open(source);
  const receipt = await doc.edit({ version: doc.list().version, edits: [
    { op: 'add_comment', match: hit(doc, 'strongly').match, author: 'Reviewer', text: 'New note\nSecond line' },
    { op: 'add_comment', match: hit(doc, 'Q1').match, author: 'Reviewer', text: 'Cell note' },
  ] });
  expect(receipt.applied).toBe(true);
  const bytes = await doc.export();
  const saved = (await part(bytes, 'word/comments.xml'))!;
  const originals = (await part(source, 'word/comments.xml'))!.match(/<w:comment\b[\s\S]*?<\/w:comment>/g)!;
  for (const original of originals) expect(saved).toContain(original);
  expect(comments(saved).map(comment => comment.paragraphs)).toEqual([['Keep this note.'], ['New note', 'Second line'], ['Cell note']]);
  const paraIds = [...saved.matchAll(/<w:comment\b[\s\S]*?<\/w:comment>/g)].map(([comment]) =>
    [...comment.matchAll(/<w:p\b[^>]*\bw14:paraId="([^"]+)"/g)].at(-1)![1]);
  let durableIds: string[] = [];
  for (const [name, prefix, , , entry] of COMMENT_COMPANIONS) {
    const xml = (await part(bytes, `word/${name}.xml`))!;
    for (const original of companionEntries((await part(source, `word/${name}.xml`)) ?? '', entry)) {
      expect(companionEntries(xml, entry)).toContainEqual(original);
    }
    const values = (attribute: string) => [...xml.matchAll(new RegExp(`${prefix}:${attribute}="([^"]+)"`, 'g'))].map(match => match[1]);
    const ids = values(name === 'commentsExtensible' ? 'durableId' : 'paraId');
    expect(ids).toHaveLength(3);
    expect(new Set(ids).size).toBe(3);
    expect(new Set(ids)).toEqual(new Set(name === 'commentsExtensible' ? durableIds : paraIds));
    if (name === 'commentsIds') durableIds = values('durableId');
  }
  return bytes;
}

test('add_comment accepts indented companion entries and keeps every entry', async () => {
  await acceptedCommentOutput(await JSZip.loadAsync(await commentFixture({ indentCompanions: true })));
});

test('add_comment accepts compact companion entries with reordered attributes', async () => {
  const zip = await JSZip.loadAsync(await commentFixture());
  for (const [name, , , , entry] of COMMENT_COMPANIONS) {
    const path = `word/${name}.xml`;
    const xml = await zip.file(path)!.async('string');
    const tag = entry.slice(1, entry.indexOf(' '));
    const attributes = [...entry.matchAll(/(\S+)="([^"]*)"/g)].reverse().map(([, key, value]) => `${key}='${value}'`).join(' ');
    zip.file(path, xml.replace(entry, `<${tag} ${attributes}/>`));
  }
  await acceptedCommentOutput(zip);
});

test('add_comment refuses normalized milliseconds in an existing commentExtensible atomically', async () => {
  const zip = await JSZip.loadAsync(await commentFixture());
  for (const path of ['word/comments.xml', 'word/commentsExtensible.xml']) {
    const xml = await zip.file(path)!.async('string');
    zip.file(path, xml.replace('2026-01-05T09:00:00Z', '2026-01-05T09:00:00.123Z'));
  }
  const error = await refusedCommentOutput(zip);
  expect(error.message).toContain('Saving comments would rewrite an existing entry in word/commentsExtensible.xml.');
});

test('add_comment accepts unused namespace and ignorable changes in companion roots', async () => {
  const zip = await JSZip.loadAsync(await commentFixture({ wordCompanionRoots: true }));
  for (const [name, prefix, , root] of COMMENT_COMPANIONS) {
    const path = `word/${name}.xml`;
    const xml = await zip.file(path)!.async('string');
    expect(xml).toContain('mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16du wp14"');
    expect(xml.match(/xmlns:/g)!.length).toBeGreaterThan(30);
    zip.file(path, xml.replace(`<${prefix}:${root} `, `<${prefix}:${root} xmlns:unused="urn:unused-comment-root" `));
  }
  const bytes = await acceptedCommentOutput(zip);
  for (const [name] of COMMENT_COMPANIONS) {
    const xml = (await part(bytes, `word/${name}.xml`))!;
    expect(xml).not.toContain('xmlns:unused=');
    expect(xml).toContain('mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14"');
  }
});

for (const options of [
  { companions: [] },
  { companions: ['commentsExtended', 'commentsIds'] },
  { companions: ['commentsExtended', 'commentsExtensible'], inlineDurableId: true },
]) {
  test(`add_comment synthesizes missing companion parts with ${options.companions.join(', ') || 'no companions'}`, async () => {
    await acceptedCommentOutput(await JSZip.loadAsync(await commentFixture(options)));
  });
}

for (const [name, prefix, , root, entry] of COMMENT_COMPANIONS) {
  for (const variant of ['entry attribute', 'entry extension', 'root attribute', 'extension child', 'entry attribute and extension child', 'XML comment', 'processing instruction', 'text', 'CDATA'] as const) {
    test(`add_comment refuses lost ${variant} in ${name} atomically`, async () => {
      const zip = await JSZip.loadAsync(await commentFixture());
      const path = `word/${name}.xml`;
      const xml = await zip.file(path)!.async('string');
      const extension = '<extension xmlns="urn:comment-extension"><payload>Keep this markup.</payload></extension>';
      const tag = entry.slice(1, entry.indexOf(' '));
      const attributed = xml.replace(entry, entry.replace(' ', ' extra="keep" '));
      const append = (xml: string, child: string) => xml.replace(`</${prefix}:${root}>`, `${child}</${prefix}:${root}>`);
      const changed = {
        'entry attribute': attributed,
        'entry extension': xml.replace(entry, entry.replace(/\s*\/>$/, `>${extension}</${tag}>`)),
        'root attribute': xml.replace(`<${prefix}:${root} `, `<${prefix}:${root} extra="keep" `),
        'extension child': append(xml, extension),
        'entry attribute and extension child': append(attributed, extension),
        'XML comment': append(xml, '<!-- Keep this markup. -->'),
        'processing instruction': append(xml, '<?keep this markup?>'),
        'text': append(xml, 'Keep this text.'),
        'CDATA': append(xml, '<![CDATA[ Keep this text. ]]>'),
      }[variant];
      zip.file(path, changed);
      const error = await refusedCommentOutput(zip);
      expect(error.message).toContain(path);
      if (variant === 'entry attribute' || variant === 'entry extension') {
        expect(error.message).toContain(`Saving comments would rewrite an existing entry in ${path}.`);
      }
    });
  }

  for (const variant of ['namespace removal', 'namespace remapping', 'ignorable removal'] as const) {
    test(`add_comment refuses ${variant} for kept content in ${name} atomically`, async () => {
      const zip = await JSZip.loadAsync(await commentFixture());
      const path = `word/${name}.xml`;
      const xml = await zip.file(path)!.async('string');
      const used = variant === 'namespace removal' ? 'extra' : variant === 'namespace remapping' ? 'w14' : 'wpc';
      const uri = used === 'wpc' ? 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas' : 'urn:comment-extension';
      const ignorable = variant === 'ignorable removal'
        ? ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="wpc"'
        : '';
      zip.file(path, xml.replace(`<${prefix}:${root} `, `<${prefix}:${root} xmlns:${used}="${uri}"${ignorable} `)
        .replace(entry, entry.replace(' ', ` ${used}:keep="unchanged" `)));
      const error = await refusedCommentOutput(zip);
      expect(error.message).toContain(`Saving comments would change the root metadata in ${path}.`);
    });
  }

  for (const markup of ['<!-- Keep this markup. -->', '<?keep this markup?>']) {
    for (const position of ['prolog', 'epilog']) {
      test(`add_comment refuses lost ${markup} in the ${name} ${position} atomically`, async () => {
        const zip = await JSZip.loadAsync(await commentFixture());
        const path = `word/${name}.xml`;
        const xml = await zip.file(path)!.async('string');
        const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        zip.file(path, declaration + (position === 'prolog' ? `${markup}\n${xml}` : `${xml}\n${markup}`));
        const error = await refusedCommentOutput(zip);
        expect(error.message).toContain(path);
      });
    }
  }

  test(`add_comment synthesizes a missing ${name} entry for an existing comment`, async () => {
    const zip = await JSZip.loadAsync(await commentFixture({ inlineDurableId: name === 'commentsIds' }));
    const path = `word/${name}.xml`;
    zip.file(path, (await zip.file(path)!.async('string')).replace(entry, ''));
    await acceptedCommentOutput(zip);
  });
}

async function verifySavedCommentOutput(source: JSZip, saved: JSZip) {
  const original = await readCommentPart(await source.generateAsync({ type: 'arraybuffer' }));
  await verifyCommentOutput(original, await saved.generateAsync({ type: 'uint8array' }), 0);
}

for (const [name, prefix, namespace, root, entry] of COMMENT_COMPANIONS) {
  const path = `word/${name}.xml`;
  const tag = entry.slice(1, entry.indexOf(' '));
  test(`comment output accepts equivalent attributes, prefixes and empty elements in ${name}`, async () => {
    const source = await JSZip.loadAsync(await commentFixture());
    const saved = await JSZip.loadAsync(await source.generateAsync({ type: 'uint8array' }));
    const xml = await saved.file(path)!.async('string');
    const attributes = [...entry.matchAll(/(\S+)="([^"]*)"/g)].reverse().map(([, key, value]) =>
      `${key.replace(`${prefix}:`, 'alias:')}='${value.replace(/.$/, char => `&#x${char.codePointAt(0)!.toString(16)};`)}'`).join(' ');
    const alias = tag.replace(`${prefix}:`, 'alias:');
    saved.file(path, xml.replace(entry, `<${alias} xmlns:alias="${namespace}" ${attributes}></${alias}>`));
    await verifySavedCommentOutput(source, saved);
  });

  test(`comment output accepts structurally identical nested content in ${name}`, async () => {
    const source = await JSZip.loadAsync(await commentFixture());
    const xml = await source.file(path)!.async('string');
    const original = entry.replace(/\s*\/>$/, ` xmlns:before="urn:entry-extension" before:flag="A &amp; B">
  <before:payload before:keep="yes">Keep &amp; text<before:child/> tail</before:payload>
</${tag}>`);
    const kept = entry.replace(/\s*\/>$/, ` xmlns:after="urn:entry-extension" after:flag='A &#38; B'><after:payload after:keep='yes'>Keep <![CDATA[&]]> text<after:child></after:child> tail</after:payload></${tag}>`);
    source.file(path, xml.replace(entry, original));
    const saved = await JSZip.loadAsync(await source.generateAsync({ type: 'uint8array' }));
    saved.file(path, xml.replace(entry, kept));
    await verifySavedCommentOutput(source, saved);
  });

  for (const variant of ['entry value', 'added attribute', 'child', 'child attribute', 'child namespace', 'child order', 'text', 'text spacing', 'text position', 'XML comment', 'processing instruction'] as const) {
    test(`comment output refuses changed ${variant} inside an existing entry in ${name}`, async () => {
      const source = await JSZip.loadAsync(await commentFixture());
      const xml = await source.file(path)!.async('string');
      const original = entry.replace(/\s*\/>$/, ` xmlns:extra="urn:entry-extension" extra:flag="keep"><extra:payload extra:keep="yes">Keep this text.<extra:first/><extra:second/> tail</extra:payload><!-- Keep this markup. --><?keep this markup?></${tag}>`);
      const changed = {
        'entry value': original.replace('extra:flag="keep"', 'extra:flag="changed"'),
        'added attribute': original.replace('extra:flag="keep"', 'extra:flag="keep" added="value"'),
        'child': original.replace('<extra:second/>', ''),
        'child attribute': original.replace('extra:keep="yes"', ''),
        'child namespace': original.replace('<extra:second/>', '<extra:second xmlns:extra="urn:changed-extension"/>'),
        'child order': original.replace('<extra:first/><extra:second/>', '<extra:second/><extra:first/>'),
        'text': original.replace('Keep this text.', 'Changed text.'),
        'text spacing': original.replace('Keep this text.', ' Keep this text.'),
        'text position': original.replace('Keep this text.<extra:first/>', '<extra:first/>Keep this text.'),
        'XML comment': original.replace('<!-- Keep this markup. -->', ''),
        'processing instruction': original.replace('<?keep this markup?>', ''),
      }[variant];
      source.file(path, xml.replace(entry, original));
      const saved = await JSZip.loadAsync(await source.generateAsync({ type: 'uint8array' }));
      saved.file(path, xml.replace(entry, changed));
      const error = await refusal(verifySavedCommentOutput(source, saved));
      expect(error.code).toBe('UNSUPPORTED_TARGET');
      expect(error.message).toContain(`Saving comments would rewrite an existing entry in ${path}.`);
    });
  }

  const metadata = 'xmlns:extra="urn:comment-extension" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
  for (const [usage, attributes, content] of [
    ['entry attribute', '', entry.replace(' ', ' extra:keep="unchanged" ')],
    ['element name', '', `${entry}<extra:payload/>`],
    ['attribute name', '', `${entry}<extension extra:keep="unchanged"/>`],
    ['Ignorable token', '', `${entry}<extension mc:Ignorable="extra"/>`],
    ['ProcessContent token', '', `${entry}<extension mc:ProcessContent="extra:payload"/>`],
    ['MustUnderstand token', '', `${entry}<extension mc:MustUnderstand="extra"/>`],
    ['root ProcessContent token', 'mc:ProcessContent="extra:payload"', entry],
    ['root MustUnderstand token', 'mc:MustUnderstand="extra"', entry],
  ]) {
    for (const variant of ['dropped', 'remapped']) {
      test(`comment output refuses a ${variant} namespace used by a kept ${usage} in ${name}`, async () => {
        const source = await JSZip.loadAsync(await commentFixture());
        const xml = (await source.file(path)!.async('string'))
          .replace(`<${prefix}:${root} `, `<${prefix}:${root} ${metadata} ${attributes} `).replace(entry, content);
        source.file(path, xml);
        const saved = await JSZip.loadAsync(await source.generateAsync({ type: 'uint8array' }));
        saved.file(path, xml.replace(' xmlns:extra="urn:comment-extension"', variant === 'dropped' ? '' : ' xmlns:extra="urn:changed-extension"'));
        const error = await refusal(verifySavedCommentOutput(source, saved));
        expect(error.code).toBe('UNSUPPORTED_TARGET');
        if (variant === 'remapped') expect(error.message).toContain(`root metadata in ${path}`);
      });
    }
  }

  test(`comment output refuses a lost used ignorable token with unchanged entries in ${name}`, async () => {
    const source = await JSZip.loadAsync(await commentFixture());
    const xml = (await source.file(path)!.async('string'))
      .replace(`<${prefix}:${root} `, `<${prefix}:${root} ${metadata} xmlns:unused="urn:unused" mc:Ignorable="extra unused" `)
      .replace(entry, entry.replace(' ', ' extra:keep="unchanged" '));
    source.file(path, xml);
    const saved = await JSZip.loadAsync(await source.generateAsync({ type: 'uint8array' }));
    saved.file(path, xml.replace('mc:Ignorable="extra unused"', 'mc:Ignorable="unused"'));
    const error = await refusal(verifySavedCommentOutput(source, saved));
    expect(error.code).toBe('UNSUPPORTED_TARGET');
    expect(error.message).toContain(`root metadata in ${path}`);
  });

  test(`comment output accepts unused root bindings and tokens changing in ${name}`, async () => {
    const source = await JSZip.loadAsync(await commentFixture());
    const xml = (await source.file(path)!.async('string'))
      .replace(`<${prefix}:${root} `, `<${prefix}:${root} ${metadata} mc:Ignorable="extra" `);
    source.file(path, xml);
    const saved = await JSZip.loadAsync(await source.generateAsync({ type: 'uint8array' }));
    saved.file(path, xml.replace('xmlns:extra="urn:comment-extension"', 'xmlns:extra="urn:changed-extension" xmlns:added="urn:added"')
      .replace('mc:Ignorable="extra"', 'mc:Ignorable="added"'));
    await verifySavedCommentOutput(source, saved);
  });

  test(`comment output accepts a dropped root binding shadowed inside a kept entry in ${name}`, async () => {
    const source = await JSZip.loadAsync(await commentFixture());
    const xml = (await source.file(path)!.async('string'))
      .replace(`<${prefix}:${root} `, `<${prefix}:${root} xmlns:extra="urn:unused" `)
      .replace(entry, entry.replace(' ', ' xmlns:extra="urn:local-extension" extra:keep="unchanged" '));
    source.file(path, xml);
    const saved = await JSZip.loadAsync(await source.generateAsync({ type: 'uint8array' }));
    saved.file(path, xml.replace(' xmlns:extra="urn:unused"', ''));
    await verifySavedCommentOutput(source, saved);
  });

  test(`comment output preserves ignorable prefixes used by root attributes in ${name}`, async () => {
    const source = await JSZip.loadAsync(await commentFixture());
    const xml = (await source.file(path)!.async('string'))
      .replace(`<${prefix}:${root} `, `<${prefix}:${root} ${metadata} extra:keep="unchanged" mc:Ignorable="extra" `);
    source.file(path, xml);
    const saved = await JSZip.loadAsync(await source.generateAsync({ type: 'uint8array' }));
    saved.file(path, xml.replace(' mc:Ignorable="extra"', ''));
    const error = await refusal(verifySavedCommentOutput(source, saved));
    expect(error.code).toBe('UNSUPPORTED_TARGET');
    expect(error.message).toContain(`root metadata in ${path}`);
  });

  test(`comment output preserves default namespaces used by non-entry markup in ${name}`, async () => {
    const source = await JSZip.loadAsync(await commentFixture());
    const xml = (await source.file(path)!.async('string'))
      .replace(`<${prefix}:${root} `, `<${prefix}:${root} xmlns="urn:comment-extension" `)
      .replace(entry, `${entry}<extension/>`);
    source.file(path, xml);
    const saved = await JSZip.loadAsync(await source.generateAsync({ type: 'uint8array' }));
    saved.file(path, xml.replace(' xmlns="urn:comment-extension"', ''));
    const error = await refusal(verifySavedCommentOutput(source, saved));
    expect(error.code).toBe('UNSUPPORTED_TARGET');
    expect(error.message).toContain(`root metadata in ${path}`);
  });
}

test('comment output keeps exterior markup while ignoring declarations and indentation', async () => {
  const source = await JSZip.loadAsync(await commentFixture({ indentCompanions: true }));
  const saved = await JSZip.loadAsync(await commentFixture());
  for (const [name] of COMMENT_COMPANIONS) {
    const path = `word/${name}.xml`;
    const before = await source.file(path)!.async('string');
    const after = await saved.file(path)!.async('string');
    source.file(path, `<?xml version="1.0"?>\n<!-- Keep  spacing. -->\n<?xml-stylesheet keep?>\n${before}\n<?keep spacing?>\n`);
    saved.file(path, `<?xml version="1.0" encoding="UTF-8"?> <!-- Keep  spacing. --><?xml-stylesheet keep?>${after} <?keep spacing?>`);
  }
  const original = await readCommentPart(await source.generateAsync({ type: 'arraybuffer' }));
  await verifyCommentOutput(original, await saved.generateAsync({ type: 'uint8array' }), 0);
});

for (const variant of ['unknown extended paragraph', 'unknown ids paragraph', 'wrong durable id', 'duplicate extended paragraph', 'duplicate ids paragraph', 'duplicate durable id', 'duplicate extensible durable id', 'first paragraph'] as const) {
  test(`comment output refuses synthesized entries with ${variant}`, async () => {
    const source = await JSZip.loadAsync(await commentFixture({ companions: [] }));
    const saved = await JSZip.loadAsync(await commentFixture());
    if (variant === 'first paragraph') {
      for (const zip of [source, saved]) {
        const xml = await zip.file('word/comments.xml')!.async('string');
        zip.file('word/comments.xml', xml.replace('<w:p w14:paraId=', '<w:p w14:paraId="11223344"><w:r><w:t>First line.</w:t></w:r></w:p><w:p w14:paraId='));
      }
    }
    if (variant === 'duplicate durable id') {
      const extra = EXISTING_COMMENT.replace('w:id="0"', 'w:id="1"').replace('1A2B3C4D', '11223344');
      for (const zip of [source, saved]) {
        const xml = await zip.file('word/comments.xml')!.async('string');
        zip.file('word/comments.xml', xml.replace('</w:comments>', `${extra}</w:comments>`));
      }
    }
    const [name, prefix, , root, entry] = COMMENT_COMPANIONS[
      variant.includes('extended') || variant === 'first paragraph' ? 0 : variant.includes('extensible') || variant === 'wrong durable id' ? 2 : 1
    ];
    const path = `word/${name}.xml`;
    const xml = await saved.file(path)!.async('string');
    const changed = variant === 'first paragraph' ? xml.replace('1A2B3C4D', '11223344')
      : variant.startsWith('unknown') ? xml.replace('1A2B3C4D', 'DEADBEEF')
      : variant === 'wrong durable id' ? xml.replace('2ECC71AA', 'DEADBEEF')
      : variant === 'duplicate durable id' ? xml.replace(`</${prefix}:${root}>`, `${entry.replace('1A2B3C4D', '11223344')}</${prefix}:${root}>`)
      : xml.replace(`</${prefix}:${root}>`, `${entry}</${prefix}:${root}>`);
    saved.file(path, changed);
    const original = await readCommentPart(await source.generateAsync({ type: 'arraybuffer' }));
    const error = await refusal(verifyCommentOutput(original, await saved.generateAsync({ type: 'uint8array' }), 0));
    expect(error.code).toBe('UNSUPPORTED_TARGET');
    expect(error.message).toContain(path);
    if (variant === 'duplicate durable id') expect(error.message).toContain('duplicate durable ids');
  });
}

for (const id of ["Id='rId1'", 'Id = "rId1"']) {
  test(`add_comment refuses a relationship allocation collision with ${id} atomically`, async () => {
    const zip = await JSZip.loadAsync(await editFixture());
    const name = 'word/_rels/document.xml.rels';
    zip.file(name, (await zip.file(name)!.async('string')).replace('Id="styles"', id));
    await refusedCommentOutput(zip);
  });
}

test('add_comment refuses an unrelated external URL containing comments.xml atomically', async () => {
  const zip = await JSZip.loadAsync(await editFixture());
  const name = 'word/_rels/document.xml.rels';
  zip.file(name, (await zip.file(name)!.async('string')).replace('</Relationships>',
    '<Relationship Id="link" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/comments.xml" TargetMode="External"/></Relationships>'));
  await refusedCommentOutput(zip);
});

test('add_comment refuses a comments.xml.bak override suppressing registration atomically', async () => {
  const zip = await JSZip.loadAsync(await editFixture());
  const name = '[Content_Types].xml';
  zip.file(name, (await zip.file(name)!.async('string')).replace('</Types>',
    '<Override PartName="/word/comments.xml.bak" ContentType="application/xml"/></Types>'));
  zip.file('word/comments.xml.bak', '<preserved/>');
  await refusedCommentOutput(zip);
});

test('add_comment refuses a relationships closing tag inside an XML comment atomically', async () => {
  const zip = await JSZip.loadAsync(await editFixture());
  const name = 'word/_rels/document.xml.rels';
  zip.file(name, (await zip.file(name)!.async('string')).replace('</Relationships>', '<!-- </Relationships> --></Relationships>'));
  await refusedCommentOutput(zip);
});

for (const variant of ['XML comment', 'leading-zero id']) {
  test(`add_comment refuses whole-part fallback with a ${variant} atomically`, async () => {
    const zip = await JSZip.loadAsync(await commentFixture());
    const name = 'word/comments.xml';
    const xml = await zip.file(name)!.async('string');
    const rich = EXISTING_COMMENT.replace('<w:r><w:t>Keep this note.</w:t></w:r>',
      '<w:hyperlink w:anchor="detail"><w:r><w:t>Keep this note.</w:t></w:r></w:hyperlink>');
    zip.file(name, xml.replace(EXISTING_COMMENT, variant === 'XML comment'
      ? `<!-- Preserve this markup. -->${rich}`
      : rich.replace('w:id="0"', 'w:id="00"')));
    await refusedCommentOutput(zip);
  });
}

test('add_comment refuses a renamed main part or unregistered output parts atomically', async () => {
  const relationships = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
  const renamed = await JSZip.loadAsync(await commentFixture());
  const move = async (from: string, to: string) => { renamed.file(to, await renamed.file(from)!.async('uint8array')); renamed.remove(from); };
  await move('word/document.xml', 'word/document2.xml');
  await move('word/_rels/document.xml.rels', 'word/_rels/document2.xml.rels');
  for (const name of ['[Content_Types].xml', '_rels/.rels']) {
    renamed.file(name, (await renamed.file(name)!.async('string')).replace('word/document.xml', 'word/document2.xml'));
  }
  const variants = [`<Relationships ${relationships}/>`, `<rel:Relationships xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"></rel:Relationships>`, `<Relationships ${relationships}></Relationships >`];
  const unusual = await Promise.all(variants.map(async xml => {
    const zip = await JSZip.loadAsync(await editFixture());
    zip.file('word/_rels/document.xml.rels', xml);
    return zip;
  }));
  const manifest = await JSZip.loadAsync(await editFixture());
  const types = await manifest.file('[Content_Types].xml')!.async('string');
  manifest.file('[Content_Types].xml', types.replace('<Types xmlns=', '<ct:Types xmlns:ct=').replace(/<(Default|Override)\b/g, '<ct:$1').replace('</Types>', '</ct:Types>'));
  for (const zip of [renamed, ...unusual, manifest]) {
    const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
    const version = doc.list().version;
    const state = doc.session.encodeState();
    const error = await refusal(doc.edit({ version, edits: [
      { op: 'add_comment', match: hit(doc, 'strongly').match, author: 'Reviewer', text: 'New note' },
    ] }));
    expect(error.code).toBe('UNSUPPORTED_TARGET');
    expect(error.details).toMatchObject({ failedIndex: 0 });
    if (zip === renamed) expect(error.message).toContain('non-standard part');
    expect(doc.list().version).toBe(version);
    expect(doc.session.encodeState()).toEqual(state);
  }
});

test('add_comment refusals are atomic: invalid input, stale version, stale match, non-body story', async () => {
  const doc = await open();
  const heading = hit(doc, 'Quarterly report');
  const match = hit(doc, 'strongly');
  const version = doc.list().version;
  const comment = (extra: Partial<Extract<DocxAgentEdit, { op: 'add_comment' }>> = {}): DocxAgentEdit => ({ op: 'add_comment', match: match.match, author: 'Reviewer', text: 'Note', ...extra });
  const refused = async (at: number, edits: DocxAgentEdit[], code: string) => {
    const error = await refusal(doc.edit({ version: at, edits }));
    expect(error.code).toBe(code);
    expect(doc.session.listComments()).toEqual([]);
    expect(doc.read(heading.ref).style).toBe('Heading1');
    return error;
  };
  await refused(version, [comment({ text: 'Tab\there' })], 'INVALID_ARGUMENT');
  await refused(version, [comment({ text: 'Line\r\nbreak' })], 'INVALID_ARGUMENT');
  await refused(version, [comment({ author: 'Two\nlines' })], 'INVALID_ARGUMENT');
  await refused(version, [comment({ text: 'Note\uffff' })], 'INVALID_ARGUMENT');
  await refused(version, [comment({ author: 'Reviewer\ufffe' })], 'INVALID_ARGUMENT');
  await refused(version, [comment({ text: '' })], 'INVALID_ARGUMENT');
  await refused(version, [comment({ date: '1 October 2026' })], 'INVALID_ARGUMENT');
  await refused(version, [comment({ match: 'm999' })], 'UNKNOWN_MATCH');
  expect(doc.list().version).toBe(version);
  await doc.edit({ version, edits: [{ op: 'format_text', match: match.match, italic: true }] });
  await refused(version, [comment({ match: hit(doc, 'Closing note').match })], 'STALE_VERSION');
  const fresh = doc.list().version;
  const atomic = await refused(fresh, [{ op: 'set_paragraph_style', paragraph: heading.ref, style: 'Heading2' }, comment()], 'STALE_TARGET');
  expect(atomic.details).toMatchObject({ failedIndex: 1 });
  expect(doc.list().version).toBe(fresh);
  expect(await part(await doc.export(), 'word/comments.xml')).toBeUndefined();

  const other = await open(await fixture());
  const header = other.grep({ query: 'Internal' }).matches[0];
  expect(header.story).not.toBe('body');
  const error = await refusal(other.edit({ version: other.list().version, edits: [{ op: 'add_comment', match: header.match, author: 'Reviewer', text: 'Note' }] }));
  expect(error.code).toBe('UNSUPPORTED_TARGET');
  expect(error.details).toMatchObject({ failedIndex: 0 });
  expect(other.session.listComments()).toEqual([]);
});
