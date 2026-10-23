import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DocxClipboardFragment, DocxTableFragment, DocxTableFragmentCell } from '@betteroffice/docx/yrs';
import { clipboardFragmentHtml, parseClipboardHtml } from './clipboardHtml';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const borders = {
  top: { style: 'single', size: 4, color: { rgb: '000000' } },
  bottom: { style: 'single', size: 4, color: { rgb: '000000' } },
  left: { style: 'single', size: 4, color: { rgb: '000000' } },
  right: { style: 'single', size: 4, color: { rgb: '000000' } },
};

function cell(text: string, extraTcPr: Record<string, unknown> = {}, shading?: string): DocxTableFragmentCell {
  return {
    tcPr: { borders, ...extraTcPr },
    content: text
      ? [{ kind: 'text', text, format: { bold: true } }, { kind: 'pilcrow' }]
      : [{ kind: 'pilcrow' }],
    ...(shading ? { shading } : {}),
  };
}

function tableFragment(): DocxTableFragment {
  return {
    tblPr: { width: 4680, widthType: 'dxa' },
    grid: [2340, 2340],
    rows: [
      { trPr: {}, cells: [cell('A'), cell('B')] },
      { trPr: {}, cells: [cell('C'), cell('D')] },
    ],
  };
}

test('internal round trip returns the same fragment', () => {
  const fragment: DocxClipboardFragment = {
    paragraphs: [
      { text: 'before' },
      { tables: [tableFragment()], text: 'after' },
    ],
  };
  const html = clipboardFragmentHtml(fragment);
  expect(html).toContain('<table');
  const parsed = parseClipboardHtml(html);
  expect(parsed).toEqual(fragment);
});

test('synthetic Word html parses tables with bold and col widths', () => {
  const html = `<html><head><style>p { margin: 0; }</style></head><body>`
    + `<p class="MsoNormal">Intro</p>`
    + `<table><colgroup><col style="width:117pt"><col style="width:117pt"></colgroup>`
    + `<tbody><tr><td><p class="MsoNormal"><b><span>Bold</span></b></p></td>`
    + `<td><p class="MsoNormal"><o:p>&nbsp;</o:p></p></td></tr></tbody></table>`
    + `<!--StartFragment--><p class="MsoNormal">Tail</p><!--EndFragment--></body></html>`;
  const parsed = parseClipboardHtml(html);
  expect(parsed).not.toBeNull();
  expect(parsed!.paragraphs).toHaveLength(2);
  expect(parsed!.paragraphs[0]!.text).toBe('Intro');
  const table = parsed!.paragraphs[1]!.tables![0]!;
  expect(table.grid).toEqual([2340, 2340]);
  expect(table.rows[0]!.cells[0]!.content[0]).toMatchObject({ kind: 'text', text: 'Bold', format: { bold: true } });
  expect(table.rows[0]!.cells[1]!.content).toEqual([{ kind: 'pilcrow' }]);
  expect(parsed!.paragraphs[1]!.text).toBe('Tail');
});

test('synthetic Google Docs html cancels wrapper bold and keeps strong spans', () => {
  const html = `<body><div><b style="font-weight:normal">`
    + `<table><colgroup><col style="width:100px"><col style="width:100px"></colgroup>`
    + `<tbody><tr><td><p><span style="font-weight:700">Bold</span></p></td>`
    + `<td><p>Plain</p></td></tr></tbody></table>`
    + `</b></div></body>`;
  const parsed = parseClipboardHtml(html);
  expect(parsed).not.toBeNull();
  const table = parsed!.paragraphs[0]!.tables![0]!;
  expect(table.grid).toEqual([1500, 1500]);
  expect(table.rows[0]!.cells[0]!.content[0]).toMatchObject({ kind: 'text', text: 'Bold', format: { bold: true } });
  const plain = table.rows[0]!.cells[1]!.content[0] as { kind: string; format?: unknown };
  expect(plain.kind).toBe('text');
  expect(plain.format ?? {}).not.toMatchObject({ bold: true });
});

test('colspan rowspan expand the grid and ragged rows are padded', () => {
  const html = `<body><table><tbody>`
    + `<tr><td colspan="2">Wide</td></tr>`
    + `<tr><td>Only</td></tr>`
    + `</tbody></table></body>`;
  const parsed = parseClipboardHtml(html);
  expect(parsed).not.toBeNull();
  const table = parsed!.paragraphs[0]!.tables![0]!;
  expect(table.rows[0]!.cells[0]!.tcPr).toMatchObject({ colspan: 2 });
  expect(table.rows[1]!.cells).toHaveLength(2);
  expect(table.rows[1]!.cells[1]!.content).toEqual([{ kind: 'pilcrow' }]);
  const spanHtml = `<body><table><tbody>`
    + `<tr><td rowspan="2">Tall</td><td>Top</td></tr>`
    + `<tr><td>Bottom</td></tr></tbody></table></body>`;
  const spanParsed = parseClipboardHtml(spanHtml);
  expect(spanParsed!.paragraphs[0]!.tables![0]!.rows[0]!.cells[0]!.tcPr).toMatchObject({ rowspan: 2 });
});

test('paragraph table paragraph shares the table paragraph text', () => {
  const html = `<body><p>First</p><table><tbody><tr><td><p>Cell</p></td></tr></tbody></table><p>Last</p></body>`;
  const parsed = parseClipboardHtml(html);
  expect(parsed).not.toBeNull();
  expect(parsed!.paragraphs).toHaveLength(2);
  expect(parsed!.paragraphs[0]).toMatchObject({ text: 'First' });
  expect(parsed!.paragraphs[1]!.tables).toHaveLength(1);
  expect(parsed!.paragraphs[1]!.text).toBe('Last');
  const trailing = parseClipboardHtml(`<body><p>First</p><table><tbody><tr><td><p>Cell</p></td></tr></tbody></table></body>`);
  expect(trailing!.paragraphs[1]!.text).toBe('');
});

test('rowspan occupancy places the next row cell in the first free column', () => {
  const html = `<body><table><tbody>`
    + `<tr><td rowspan="2">A</td><td>B</td></tr>`
    + `<tr><td>C</td></tr></tbody></table></body>`;
  const parsed = parseClipboardHtml(html);
  expect(parsed).not.toBeNull();
  const table = parsed!.paragraphs[0]!.tables![0]!;
  expect(table.rows).toHaveLength(2);
  expect(table.rows[0]!.cells.map((c) => c.content[0])).toEqual([
    { kind: 'text', text: 'A' },
    { kind: 'text', text: 'B' },
  ]);
  expect(table.rows[0]!.cells[0]!.tcPr).toMatchObject({ rowspan: 2 });
  expect(table.rows[1]!.cells).toHaveLength(1);
  expect(table.rows[1]!.cells[0]!.content[0]).toMatchObject({ kind: 'text', text: 'C' });
  expect(table.grid.length === 0 || table.grid.length === 2).toBe(true);
});

test('huge colspan is clamped to a bounded table', () => {
  const parsed = parseClipboardHtml(
    `<body><table><tbody><tr><td colspan="1000000000">Wide</td></tr></tbody></table></body>`
  );
  expect(parsed).not.toBeNull();
  const table = parsed!.paragraphs[0]!.tables![0]!;
  expect(table.rows).toHaveLength(1);
  expect(table.rows[0]!.cells).toHaveLength(1);
  expect(table.rows[0]!.cells[0]!.tcPr).toMatchObject({ colspan: 1000 });
});

test('text only html returns null', () => {
  expect(parseClipboardHtml(`<body><p>Hello <b>world</b></p></body>`)).toBeNull();
  expect(parseClipboardHtml(`<body><div>Just text</div></body>`)).toBeNull();
});

test('unknown payload version or invalid shape falls back to the html table', () => {
  const fragment: DocxClipboardFragment = {
    paragraphs: [{ tables: [tableFragment()], text: '' }],
  };
  const html = clipboardFragmentHtml(fragment);
  const badVersion = html.replace('&quot;v&quot;:1', '&quot;v&quot;:999');
  expect(badVersion).not.toBe(html);
  const parsedVersion = parseClipboardHtml(badVersion);
  expect(parsedVersion).not.toBeNull();
  expect(parsedVersion!.paragraphs[0]!.tables).toHaveLength(1);
  const badShape = html.replace('&quot;paragraphs&quot;', '&quot;nope&quot;');
  expect(badShape).not.toBe(html);
  const parsedShape = parseClipboardHtml(badShape);
  expect(parsedShape).not.toBeNull();
  expect(parsedShape!.paragraphs[0]!.tables).toHaveLength(1);
});


test('picture fragment round trips with ordered HTML images and escaped attributes', () => {
  const fragment = { paragraphs: [{ text: 'abcd', inlines: [
    { text: 'ab' }, { image: { src: 'data:image/png;base64,AA==', alt: 'A "picture" <&>', width: 120, height: 60 } }, { text: 'cd' },
  ] }] };
  const html = clipboardFragmentHtml(fragment);
  expect(html).toContain('<img src="data:image/png;base64,AA=="');
  expect(html).toContain('width="120" height="60"');
  expect(html).toContain('alt="A &quot;picture&quot; &lt;&amp;&gt;"');
  expect(parseClipboardHtml(html)).toEqual(fragment);
});

test('version one text fragments remain readable', () => {
  expect(parseClipboardHtml(`<div data-betteroffice-docx='{"v":1,"fragment":{"paragraphs":[{"text":"old"}]}}'></div>`))
    .toEqual({ paragraphs: [{ text: 'old' }] });
});


test('clipboard HTML emits version one without pictures and version two with pictures', () => {
  const table: DocxClipboardFragment = { paragraphs: [{ text: '', tables: [tableFragment()] }] };
  const picture: DocxClipboardFragment = { paragraphs: [{ text: '', inlines: [{ image: { src: 'data:image/png;base64,AA==' } }] }] };
  for (const fragment of [table, { paragraphs: [{ text: 'plain' }] }]) {
    const html = clipboardFragmentHtml(fragment);
    expect(html).toContain('&quot;v&quot;:1');
    expect(parseClipboardHtml(html)).toEqual(fragment);
  }
  const html = clipboardFragmentHtml(picture);
  expect(html).toContain('&quot;v&quot;:2');
  expect(parseClipboardHtml(html)).toEqual(picture);
});
