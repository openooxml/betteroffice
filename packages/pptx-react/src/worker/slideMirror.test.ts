import { expect, test } from 'bun:test';
import type { SlideDisplayList, SlidePrimitive, TextLanguage, TextLink } from '@betteroffice/pptx';
import { slideMirrorItems, slideObjectName, slideTitle } from './slideMirror';

const at = { objectId: 1, x: 10, y: 20, w: 30, h: 40 };
const text = (shapeId: string | undefined, ...lines: string[]): SlidePrimitive => ({
  ...at, kind: 'textBox', shapeId, anchor: 'top', lines: [],
  paragraphs: lines.map((line) => ({ level: 0, runs: line.split('|').map((part) => ({ text: part, fontFamily: 'Arial', fontSizePt: 12, color: '#000000' })) })),
});
const shape = (shapeId: string, name = 'Rectangle 1'): SlidePrimitive => ({ ...at, kind: 'shape', shapeId, name, geometry: 'rect', path: [] });
const image = (shapeId: string | undefined, name = 'Picture 2'): SlidePrimitive => ({ ...at, kind: 'image', shapeId, name });
const list = (primitives: SlidePrimitive[], altText?: SlideDisplayList['altText']): SlideDisplayList =>
  ({ contractVersion: 1, width: 960, height: 540, primitives, altText });
const items = (value: SlideDisplayList) => slideMirrorItems(value, 'Picture');
const languageText = (shapeId: string, value: string, languages?: TextLanguage[]): SlidePrimitive => ({
  ...at, kind: 'textBox', shapeId, anchor: 'top', lines: [],
  paragraphs: [{ level: 0, languages, runs: [{ text: value, fontFamily: 'Arial', fontSizePt: 12, color: '#000000' }] }],
});

test('a paragraph reads as one string across its runs', () => {
  const [item] = items(list([text('a', 'Quarterly |results| are in', '', 'Second\u000bline')]));
  expect(item!.paragraphs.map((paragraph) => paragraph.text)).toEqual(['Quarterly results are in', 'Second line']);
  expect(item).toMatchObject({ x: 10, y: 20, w: 30, h: 40 });
});

test('a shape and its text are one item, and a shape with neither text nor alt text is none', () => {
  expect(items(list([shape('a'), text('a', 'Label'), shape('b')])).map((item) => item.key)).toEqual(['a']);
});

test('alt text labels pictures, and an undescribed picture falls back to its name', () => {
  const value = list([image('a'), image('b'), image('c', '')], { a: { title: 'Logo', description: 'A blue square' } });
  expect(items(value).map((item) => item.label)).toEqual(['Logo: A blue square', 'Picture 2', 'Picture']);
});

test('a decorative picture is skipped, but text in a decorative shape stays', () => {
  const value = list([image('a'), shape('b'), text('b', 'Kept')], { a: { decorative: true }, b: { decorative: true, description: 'Dropped' } });
  expect(items(value)).toMatchObject([{ key: 'b', label: null, paragraphs: [{ text: 'Kept' }] }]);
});

test('a described group speaks once, and its unlabelled pictures add nothing', () => {
  const value = list([image('a'), image('b'), text('c', 'Inside')], {
    outer: { description: 'Whole diagram' }, inner: { description: 'Left half', group: 'outer' },
    a: { group: 'inner' }, b: { group: 'inner' }, c: { group: 'outer' },
  });
  expect(items(value).map((item) => [item.key, item.label])).toEqual([['outer', 'Whole diagram'], ['inner', 'Left half'], ['c', null]]);
});

test('a picture painted without a shape id keeps its alt text, its place and one item', () => {
  const op = (name: string): SlidePrimitive => ({ ...at, kind: 'shape', name, geometry: 'custom', path: [] });
  const value = list([text('a', 'Before'), op('Vector'), op('Vector'), op('Bare'), text('b', 'After')],
    { v: { description: 'A described drawing', primitive: 1 }, w: { primitive: 3 } });
  expect(items(value).map((item) => [item.key, item.label])).toEqual([['a', null], ['v', 'A described drawing'], ['w', 'Bare'], ['b', null]]);
  expect(slideObjectName(value, 'w', 'Shape')).toBe('Bare');
});

test('a frame description reads once above the shapes it expands into', () => {
  const value = list([shape('f:0'), text('f:0', 'Manager'), shape('f:1'), text('f:1', 'Report')],
    { f: { description: 'Reporting lines' }, 'f:0': { group: 'f' }, 'f:1': { group: 'f' } });
  expect(items(value).map((item) => [item.key, item.label])).toEqual([['f', 'Reporting lines'], ['f:0', null], ['f:1', null]]);
});

test('table cells and chart summaries are reachable', () => {
  const table: SlidePrimitive = { ...at, kind: 'table', shapeId: 't', name: 'Table 1', label: 'Table, 1 rows, 2 columns', primitives: [text(undefined, 'A1'), text(undefined, 'B1')] };
  const chart: SlidePrimitive = { ...at, kind: 'chart', shapeId: 'c', name: 'Chart 1', label: 'Bar chart', primitives: [text(undefined, 'Axis')] };
  expect(items(list([table, chart]))).toMatchObject([
    { label: 'Table, 1 rows, 2 columns', paragraphs: [{ text: 'A1' }, { text: 'B1' }] },
    { label: 'Bar chart', paragraphs: [] },
  ]);
});

test('layout and master shapes read after the slide, and backgrounds not at all', () => {
  const value = list([image(undefined, 'Background'), text('master:m:0', 'Confidential'), text('layout:l:0', 'Footer'), text('a', 'Title')]);
  expect(items(value).map((item) => item.key)).toEqual(['a', 'master:m:0', 'layout:l:0']);
});

test('the title placeholder reads as one heading, wherever it sits in the stack', () => {
  const value = { ...list([text('a', 'Intro'), text('b', 'Two line', 'title')]), titleShape: 'b' };
  expect(items(value).flatMap((item) => item.paragraphs)).toEqual([
    { text: 'Intro', heading: false, lang: null, segments: [{ text: 'Intro', lang: null }] },
    { text: 'Two line title', heading: true, lang: null, segments: [{ text: 'Two line title', lang: null }] },
  ]);
  expect(slideTitle(value)).toBe('Two line title');
});

test('one language and unknown languages each keep a single segment across formatting runs', () => {
  const single = languageText('a', 'Quarterly results', [
    { start: 0, end: 10, lang: 'en-US' }, { start: 10, end: 17, lang: 'en-US' },
  ]);
  expect(items(list([single, text('b', 'Whole |words')])).flatMap((item) => item.paragraphs)).toEqual([
    { text: 'Quarterly results', heading: false, lang: 'en-US', segments: [{ text: 'Quarterly results', lang: 'en-US' }] },
    { text: 'Whole words', heading: false, lang: null, segments: [{ text: 'Whole words', lang: null }] },
  ]);
});

test('mixed languages use the largest total coverage and preserve spaces after slicing raw utf16 offsets', () => {
  const raw = ' \tHello😀 \n bonjour \t world  ';
  const start = raw.indexOf('bonjour');
  const end = start + 'bonjour'.length;
  const value = languageText('a', raw, [
    { start: 0, end: start, lang: 'en-US' },
    { start, end, lang: 'fr-FR' },
    { start: end, end: raw.length, lang: 'en-US' },
  ]);
  expect(items(list([value]))[0]!.paragraphs).toEqual([{
    text: 'Hello😀 bonjour world', heading: false, lang: 'en-US', segments: [
      { text: 'Hello😀 ', lang: 'en-US' }, { text: 'bonjour', lang: 'fr-FR' }, { text: ' world', lang: 'en-US' },
    ],
  }]);
});

test('unknown gaps do not split words and x-none never becomes a language', () => {
  const value = languageText('a', 'whole words', [{ start: 0, end: 2, lang: 'en-US' }]);
  const unknown = languageText('b', 'Unmarked', [{ start: 0, end: 8, lang: 'x-none' }]);
  expect(items(list([value, unknown])).flatMap((item) => item.paragraphs)).toMatchObject([
    { lang: null, segments: [{ text: 'whole', lang: 'en-US' }, { text: ' words', lang: null }] },
    { lang: null, segments: [{ text: 'Unmarked', lang: null }] },
  ]);
});

test('language tags inside a word or on punctuation do not split the text', () => {
  const inside = languageText('a', 'Management review', [
    { start: 0, end: 4, lang: 'en-US' }, { start: 4, end: 10, lang: 'de-DE' }, { start: 10, end: 17, lang: 'en-US' },
  ]);
  const punctuation = languageText('b', 'Windows（Ctrl）', [
    { start: 0, end: 8, lang: 'ja-JP' }, { start: 8, end: 13, lang: 'ja-JP' }, { start: 13, end: 14, lang: 'en-US' },
  ]);
  const leading = languageText('c', '1. bonjour', [{ start: 0, end: 3, lang: 'en-US' }, { start: 3, end: 10, lang: 'fr-FR' }]);
  expect(items(list([inside, punctuation, leading])).map((item) => item.paragraphs[0]!.segments)).toEqual([
    [{ text: 'Management review', lang: 'en-US' }],
    [{ text: 'Windows（Ctrl）', lang: 'ja-JP' }],
    [{ text: '1. bonjour', lang: 'fr-FR' }],
  ]);
});

test('a tagged word leaves unmarked text in the language of the page', () => {
  const value = languageText('a', 'Please say bonjour to continue', [{ start: 11, end: 18, lang: 'fr-FR' }]);
  expect(items(list([value]))[0]!.paragraphs[0]).toMatchObject({
    lang: null,
    segments: [{ text: 'Please say ', lang: null }, { text: 'bonjour', lang: 'fr-FR' }, { text: ' to continue', lang: null }],
  });
});

test('a paragraph without letters keeps its language', () => {
  const value = languageText('a', '2024', [{ start: 0, end: 4, lang: 'de-DE' }]);
  expect(items(list([value]))[0]!.paragraphs[0]).toMatchObject({ lang: 'de-DE', segments: [{ text: '2024', lang: 'de-DE' }] });
});

test('a single letter in another language does not split the text', () => {
  const value = languageText('a', 'Fキーを押す', [{ start: 0, end: 1, lang: 'en-US' }, { start: 1, end: 6, lang: 'ja-JP' }]);
  expect(items(list([value]))[0]!.paragraphs[0]!.segments).toEqual([{ text: 'Fキーを押す', lang: 'ja-JP' }]);
});

test('regional variants of one language stay in one segment and scripts without spaces still change language', () => {
  const regional = languageText('a', 'colour and color', [{ start: 0, end: 10, lang: 'en-GB' }, { start: 10, end: 16, lang: 'en-US' }]);
  const unspaced = languageText('b', 'Mac版です', [{ start: 0, end: 3, lang: 'en-US' }, { start: 3, end: 6, lang: 'ja-JP' }]);
  expect(items(list([regional, unspaced])).map((item) => item.paragraphs[0]!.segments)).toEqual([
    [{ text: 'colour and color', lang: 'en-GB' }],
    [{ text: 'Mac', lang: 'en-US' }, { text: '版です', lang: 'ja-JP' }],
  ]);
});

test('a title joined from paragraphs keeps each language and table cells keep theirs', () => {
  const first = languageText('title', 'Hello world', [{ start: 0, end: 11, lang: 'en-US' }]);
  const second = languageText('title', 'Bonjour', [{ start: 0, end: 7, lang: 'fr-FR' }]);
  const table: SlidePrimitive = {
    ...at, kind: 'table', shapeId: 'table', name: 'Table', label: 'Table',
    primitives: [languageText('cell', 'Guten Tag', [{ start: 0, end: 9, lang: 'de-DE' }])],
  };
  const value = { ...list([first, second, table]), titleShape: 'title' };
  expect(items(value)[0]!.paragraphs).toEqual([{
    text: 'Hello world Bonjour', heading: true, lang: 'en-US', segments: [
      { text: 'Hello world ', lang: 'en-US' }, { text: 'Bonjour', lang: 'fr-FR' },
    ],
  }]);
  expect(items(value)[1]!.paragraphs[0]!.lang).toBe('de-DE');
  expect(slideTitle(value)).toBe('Hello world Bonjour');
});

test('spaces joining title paragraphs do not change which language covers the most text', () => {
  const value = { ...list([
    languageText('title', 'Hi', [{ start: 0, end: 2, lang: 'en-US' }]),
    languageText('title', 'Oui', [{ start: 0, end: 3, lang: 'fr-FR' }]),
  ]), titleShape: 'title' };
  expect(items(value)[0]!.paragraphs).toEqual([{
    text: 'Hi Oui', heading: true, lang: 'fr-FR', segments: [
      { text: 'Hi ', lang: 'en-US' }, { text: 'Oui', lang: 'fr-FR' },
    ],
  }]);
});

test('a slide with an empty or missing title has none', () => {
  expect(slideTitle(list([text('a', 'Body')]))).toBeNull();
  expect(slideTitle({ ...list([text('a', '')]), titleShape: 'a' })).toBeNull();
});

test('an announcement prefers alt text, then the name, then the text', () => {
  const value = list([shape('a'), text('a', 'Body'), text('b', 'Only text'), image('c', '')], { a: { description: 'Described' } });
  expect(slideObjectName(value, 'a', 'Shape')).toBe('Described');
  expect(slideObjectName(list([shape('a'), text('a', 'Body')]), 'a', 'Shape')).toBe('Rectangle 1');
  expect(slideObjectName(value, 'b', 'Shape')).toBe('Only text');
  expect(slideObjectName(value, 'c', 'Shape')).toBe('Shape');
  expect(slideObjectName(list([image('d')], { d: { decorative: true, description: 'Silent' } }), 'd', 'Shape')).toBe('Picture 2');
});

const linkedText = (value: string, links: TextLink[], languages?: TextLanguage[]): SlidePrimitive => ({
  ...at, kind: 'textBox', shapeId: 'links', anchor: 'top', lines: [],
  paragraphs: [{ level: 0, links, languages, runs: [{ text: value, fontFamily: 'Arial', fontSizePt: 12, color: '#000000' }] }],
});

test('link boundaries use raw utf16 offsets independently of language and word boundaries', () => {
  const target = { kind: 'url', href: 'https://example.test' } as const;
  const paragraph = items(list([linkedText('A😀 Management bonjour', [{ start: 4, end: 8, target }], [
    { start: 0, end: 15, lang: 'en-US' }, { start: 15, end: 22, lang: 'fr-FR' },
  ])]))[0]!.paragraphs[0]!;
  expect(paragraph.text).toBe('A😀 Management bonjour');
  expect(paragraph.segments).toEqual([
    { text: 'A😀 ', lang: 'en-US' },
    { text: 'Mana', lang: 'en-US', link: target },
    { text: 'gement ', lang: 'en-US' },
    { text: 'bonjour', lang: 'fr-FR' },
  ]);
});

test('refused link targets leave unlinked text in a single segment', () => {
  for (const href of ['javascript:alert(1)', 'JaVaScRiPt:alert(1)', 'java\tscript:alert(1)', '\x01javascript:alert(1)', 'data:text/html,bad', 'vbscript:bad', 'file:///tmp/file', 'foo.html', '//evil', 'ftp://example.test']) {
    expect(items(list([linkedText('Whole words', [{ start: 0, end: 5, target: { kind: 'url', href } }])]))[0]!.paragraphs[0]!.segments)
      .toEqual([{ text: 'Whole words', lang: null }]);
  }
});

test('safe urls and navigation targets survive paragraph and title merging', () => {
  const target = { kind: 'jump', jump: 'next' } as const;
  const first = linkedText('Next slide', [{ start: 0, end: 10, target }]);
  const value = { ...list([first]), titleShape: 'links' };
  expect(items(value)[0]!.paragraphs[0]).toMatchObject({ heading: true, segments: [{ text: 'Next slide', link: target }] });
  for (const href of ['https://example.test/path', 'HTTP://example.test', 'mailto:reader@example.test', 'tel:+123', ' \tHTTPS://example.test/\r']) {
    expect(items(list([linkedText('Read', [{ start: 0, end: 4, target: { kind: 'url', href } }])]))[0]!.paragraphs[0]!.segments[0]!.link)
      .toEqual({ kind: 'url', href });
  }
});

const grid = (cells: [number, number, number, number, 'column' | 'row' | undefined, number, number][], rows = 3, columns = 3) => ({
  rows, columns,
  cells: cells.map(([row, column, rowSpan, columnSpan, header, start, end]) =>
    ({ row, column, rowSpan, columnSpan, ...(header ? { header } : {}), primitives: [start, end] as [number, number] })),
});
const gridTable = (value: ReturnType<typeof grid> | undefined, primitives: SlidePrimitive[]): SlidePrimitive =>
  ({ ...at, kind: 'table', shapeId: 't', name: 'Table 1', label: 'Table, 3 rows, 3 columns', primitives, grid: value });
const cellTexts = ['Region', 'Total', 'North', 'Q1', 'Q2', '9'];
const merged = () => grid([
  [0, 0, 1, 2, 'column', 1, 2], [0, 2, 1, 1, 'column', 2, 3],
  [1, 0, 2, 1, 'row', 3, 4], [1, 1, 1, 1, undefined, 4, 5], [1, 2, 1, 1, undefined, 5, 5],
  [2, 1, 1, 1, undefined, 5, 6], [2, 2, 1, 1, undefined, 6, 7],
]);
const cellPrimitives = () => [shape('t'), ...cellTexts.map((value) => text(undefined, value))];

test('a table with a grid reads as rows of header and data cells with their spans', () => {
  const [item] = items(list([gridTable(merged(), cellPrimitives())]));
  expect(item!.label).toBe('Table, 3 rows, 3 columns');
  expect(item!.paragraphs).toEqual([]);
  expect(item!.table!.columns).toBe(3);
  expect(item!.table!.rows.map((row) => row.map((cell) => [
    cell.row, cell.column, cell.rowSpan, cell.columnSpan, cell.header, cell.paragraphs.map((paragraph) => paragraph.text).join(' '),
  ]))).toEqual([
    [[0, 0, 1, 2, 'column', 'Region'], [0, 2, 1, 1, 'column', 'Total']],
    [[1, 0, 2, 1, 'row', 'North'], [1, 1, 1, 1, null, 'Q1'], [1, 2, 1, 1, null, '']],
    [[2, 1, 1, 1, null, 'Q2'], [2, 2, 1, 1, null, '9']],
  ]);
});

test('an empty header cell reads as a plain cell', () => {
  const value = grid([[0, 0, 1, 1, 'column', 0, 1], [0, 1, 1, 1, 'column', 1, 1]], 1, 2);
  const [item] = items(list([gridTable(value, [text(undefined, 'Name')])]));
  expect(item!.table!.rows[0]!.map((cell) => cell.header)).toEqual(['column', null]);
});

test('a row whose cells all continue spans from above is kept empty', () => {
  const value = grid([[0, 0, 2, 1, undefined, 0, 1], [0, 1, 2, 1, undefined, 1, 2]], 2, 2);
  const [item] = items(list([gridTable(value, [text(undefined, 'Left'), text(undefined, 'Right')])]));
  expect(item!.table!.rows.map((row) => row.length)).toEqual([2, 0]);
});

test('cell text keeps its links and languages', () => {
  const linked: SlidePrimitive = {
    ...at, kind: 'textBox', anchor: 'top', lines: [],
    paragraphs: [{ level: 0, languages: [{ start: 0, end: 7, lang: 'fr-FR' }],
      links: [{ start: 0, end: 7, target: { kind: 'url', href: 'https://example.test/' } }],
      runs: [{ text: 'Bonjour', fontFamily: 'Arial', fontSizePt: 12, color: '#000000' }] }],
  };
  const value = grid([[0, 0, 1, 1, undefined, 0, 1]], 1, 1);
  const [item] = items(list([gridTable(value, [linked])]));
  expect(item!.table!.rows[0]![0]!.paragraphs).toEqual([{
    text: 'Bonjour', heading: false, lang: 'fr-FR',
    segments: [{ text: 'Bonjour', lang: 'fr-FR', link: { kind: 'url', href: 'https://example.test/' } }],
  }]);
});

test('a table without a grid, or with one that does not fit it, reads as flat paragraphs', () => {
  const broken = [
    undefined,
    grid([[0, 0, 1, 1, undefined, 0, 99]]),
    grid([[0, 0, 1, 2, undefined, 0, 1], [0, 1, 1, 1, undefined, 1, 2]]),
    grid([[2, 2, 2, 1, undefined, 0, 1]]),
    grid([[0, 3, 1, 1, undefined, 0, 1]]),
    grid([[0, 0, 0, 1, undefined, 0, 1]]),
    grid([[0, 0, 1, 1, undefined, 2, 1]]),
    grid([[0.5, 0, 1, 1, undefined, 0, 1]]),
    grid([], 0, 3),
    grid([]),
    { ...merged(), cells: [null, ...merged().cells] },
    { ...merged(), rows: 2_000_000 },
    grid(merged().cells.map((cell, index) => [cell.row, cell.column, cell.rowSpan, cell.columnSpan, cell.header, index ? 2 : 1, index ? 3 : 2])),
    grid(merged().cells.slice(0, 6).map((cell) => [cell.row, cell.column, cell.rowSpan, cell.columnSpan, cell.header, ...cell.primitives])),
  ] as unknown as (ReturnType<typeof grid> | undefined)[];
  for (const value of broken) {
    const [item] = items(list([gridTable(value, cellPrimitives())]));
    expect(item!.table).toBeUndefined();
    expect(item!.paragraphs.map((paragraph) => paragraph.text)).toEqual(cellTexts);
  }
});

test('titles and object names still read table text', () => {
  const value = { ...list([gridTable(merged(), cellPrimitives())]), titleShape: 't' };
  expect(slideTitle(value)).toBe(cellTexts.join(' '));
  expect(slideObjectName(value, 't', 'Object')).toBe('Table 1');
});

test('a table of thousands of cells with spans reads in full in under a second', () => {
  const rows = 60;
  const columns = 50;
  const cells: Parameters<typeof grid>[0] = [];
  const primitives: SlidePrimitive[] = [];
  for (let row = 0; row < rows; row++) {
    const spanned = row % 2 === 0;
    for (let column = 0; column < columns; column += spanned && column === 0 ? 2 : 1) {
      const span = spanned && column === 0 ? 2 : 1;
      cells.push([row, column, 1, span, row === 0 ? 'column' : undefined, primitives.length, primitives.length + 1]);
      primitives.push(text(undefined, `r${row}c${column}`));
    }
  }
  const started = performance.now();
  const [item] = items(list([gridTable(grid(cells, rows, columns), primitives)]));
  const elapsed = performance.now() - started;
  expect(cells.length).toBe(30 * 49 + 30 * 50);
  expect(item!.table!.rows).toHaveLength(rows);
  expect(item!.table!.rows.flat()).toHaveLength(cells.length);
  expect(item!.table!.rows[0]!.every((cell) => cell.header === 'column')).toBe(true);
  const last = item!.table!.rows[59]!;
  expect(last[last.length - 1]!.paragraphs[0]!.text).toBe('r59c49');
  expect(elapsed).toBeLessThan(1_000);
});

const chartData = (overrides: Partial<{ categories: unknown; series: unknown }> = {}) => ({
  categories: ['Q1', 'Q2', 'Q3'],
  series: [{ name: 'North', values: ['12', '19', '7'] }, { name: 'South', values: ['8', null, '21'] }],
  ...overrides,
});
const chartPrimitive = (data: unknown): SlidePrimitive =>
  ({ ...at, kind: 'chart', shapeId: 'c', name: 'Chart 1', label: 'Revenue, column chart, 2 series, 3 categories', primitives: [], data } as SlidePrimitive);

test('a chart with data reads as its series by category under its summary', () => {
  const [item] = items(list([chartPrimitive(chartData())]));
  expect(item!.label).toBe('Revenue, column chart, 2 series, 3 categories');
  expect(item!.chart).toEqual({
    series: ['North', 'South'],
    rows: [
      { category: 'Q1', values: ['12', '8'] },
      { category: 'Q2', values: ['19', null] },
      { category: 'Q3', values: ['7', '21'] },
    ],
  });
});

test('a series without a name keeps its column with no header text', () => {
  const [item] = items(list([chartPrimitive(chartData({ series: [{ values: ['1', '2', '3'] }, { name: ' ', values: ['4', '5', '6'] }] }))]));
  expect(item!.chart!.series).toEqual([null, null]);
});

test('a chart without data, or with data that does not fit, reads as its summary only', () => {
  const broken = [
    undefined,
    null,
    chartData({ categories: [] }),
    chartData({ series: [] }),
    chartData({ categories: ['Q1', 'Q2'] }),
    chartData({ categories: ['Q1', 2, 'Q3'] }),
    chartData({ series: [null] }),
    chartData({ series: [{ name: 'N', values: ['1', 2, '3'] }] }),
    chartData({ series: [{ name: 5, values: ['1', '2', '3'] }] }),
    chartData({ categories: Array.from({ length: 5_001 }, (_, index) => `c${index}`),
      series: [0, 1].map(() => ({ values: Array.from({ length: 5_001 }, () => '1') })) }),
    chartData({ categories: ['x'.repeat(600_000), 'y'.repeat(600_000), 'z'], series: [{ values: ['1', '2', '3'] }] }),
  ];
  for (const data of broken) {
    const [item] = items(list([chartPrimitive(data)]));
    expect(item!.chart).toBeUndefined();
    expect(item!.label).toBe('Revenue, column chart, 2 series, 3 categories');
  }
});

test('a decorative chart reads as nothing', () => {
  expect(items(list([chartPrimitive(chartData())], { c: { decorative: true } }))).toEqual([]);
});
