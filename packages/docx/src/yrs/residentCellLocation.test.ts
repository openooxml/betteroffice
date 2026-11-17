import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { cellLocationCacheCounts, resolveLiveCell, setCellLocationShadowCheck } from './residentCommandTableSelection';
import type { DocxWorkerCommandEngine } from './residentCommandProtocol';
import { closeCommandEngines, openCommandEngine, TIMEOUT } from './residentCommandTestSupport';

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string, properties = '') => `<w:tc><w:tcPr>${properties}</w:tcPr>${content}</w:tc>`;
const grid = (columns: number) => `<w:tblGrid>${'<w:gridCol w:w="1500"/>'.repeat(columns)}</w:tblGrid>`;
const table = (columns: number, rows: string[]) => `<w:tbl><w:tblPr/>${grid(columns)}${rows.map((row) => `<w:tr>${row}</w:tr>`).join('')}</w:tbl>`;
const square = (texts: string[][]) => table(texts[0]!.length, texts.map((row) => row.map((text) => cell(paragraph(text))).join('')));

const plain = square([['A', 'B'], ['C', 'D']]);
const merged = table(3, [
  cell(paragraph('span'), '<w:gridSpan w:val="2"/>') + cell(paragraph('top'), '<w:vMerge w:val="restart"/>'),
  cell(paragraph('a')) + cell(paragraph('b')) + cell(paragraph(''), '<w:vMerge/>'),
  cell(paragraph('tall'), '<w:vMerge w:val="restart"/>') + cell(paragraph('c')) + cell(paragraph('d')),
  cell(paragraph(''), '<w:vMerge/>') + cell(paragraph('e'), '<w:gridSpan w:val="2"/>'),
]);
const nested = table(2, [cell(square([['inner A', 'inner B'], ['inner C', 'inner D']]) + paragraph('after inner')) + cell(paragraph('outer B'))]);
const control = `<w:sdt><w:sdtPr><w:alias w:val="Box"/></w:sdtPr><w:sdtContent>${paragraph('in control')}${square([['X', 'Y']])}</w:sdtContent></w:sdt>`;

beforeAll(() => setCellLocationShadowCheck(true));
afterAll(() => setCellLocationShadowCheck(false));
afterEach(closeCommandEngines);

/** The engine without its structure key: resolveLiveCell then projects on every call, as before the cache. */
function uncached(engine: DocxWorkerCommandEngine): DocxWorkerCommandEngine {
  return new Proxy(engine, {
    get: (target, key) => (key === 'cellStructureKey' || key === 'cellLocation' ? undefined : Reflect.get(target, key)),
  });
}

function expectProjectionEquivalent(engine: DocxWorkerCommandEngine, extra: string[] = []) {
  const legacy = uncached(engine);
  const stories = [...engine.geometryReader.storyIds(), ...extra];
  for (const story of stories) {
    expect({ story, location: resolveLiveCell(engine, story) }).toEqual({ story, location: resolveLiveCell(legacy, story) });
  }
  return stories;
}

const range = (row: number, column: number, toRow = row, toColumn = column) => ({
  anchor: { story: 'body', tableIndex: 0, row, column }, head: { story: 'body', tableIndex: 0, row: toRow, column: toColumn },
});

for (const [name, bodyXml] of [
  ['plain', paragraph('Before') + plain + paragraph('After')],
  ['merged and vertically merged', merged + paragraph('After')],
  ['nested', paragraph('Before') + nested + paragraph('After')],
  ['several tables and a block control', plain + paragraph('Between') + merged + control + paragraph('After')],
] as const) {
  test(`cached cell location matches the projection for every story of a ${name} document`, async () => {
    const arm = await openCommandEngine([], { bodyXml });
    expect(typeof arm.engine.cellStructureKey).toBe('function');
    const stories = expectProjectionEquivalent(arm.engine, ['body:t99:r0c0', 'hf:rId9:t0:r0c0', 'custom:t0:r0c0']);
    const cells = stories.filter((story) => /:t\d+:r\d+c\d+/.test(story) && arm.engine.geometryReader.storyIds().includes(story));
    expect(cells.length).toBeGreaterThan(0);
    expect(cells.every((story) => resolveLiveCell(arm.engine, story) !== null)).toBe(true);
  }, TIMEOUT);
}

test('typing in a cell reuses the projected tables and a structural edit rebuilds them', async () => {
  const arm = await openCommandEngine([], { bodyXml: paragraph('Before') + square([['A', 'B'], ['C', 'D']]) + paragraph('After') });
  const story = arm.engine.geometryReader.storyIds().find((id) => /^body:t0:r1c1$/.test(id))!;
  resolveLiveCell(arm.engine, story);
  const typed = cellLocationCacheCounts();
  const [para] = arm.engine.paragraphs(story);
  for (const text of ['x', 'y', 'z']) {
    arm.engine.insertText({ story, paraId: para!.paraId, offset: 0 }, text);
    expect(resolveLiveCell(arm.engine, story)).toEqual({ story: 'body', tableIndex: 0, row: 1, column: 1 });
  }
  const afterTyping = cellLocationCacheCounts();
  expect(afterTyping.built).toBe(typed.built);
  expect(afterTyping.reused - typed.reused).toBeGreaterThanOrEqual(3);
  arm.engine.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'above');
  expect(resolveLiveCell(arm.engine, story)).toEqual({ story: 'body', tableIndex: 0, row: 2, column: 1 });
  expect(cellLocationCacheCounts().built).toBeGreaterThan(afterTyping.built);
}, TIMEOUT);

test('a nested table a merge moved keeps the projection result', async () => {
  const inner = table(1, [cell(paragraph('Inner'))]);
  const arm = await openCommandEngine([], { bodyXml: table(2, [cell(paragraph('CellA')) + cell(inner + paragraph(''))]) });
  const moved = 'body:t0:r0c1:t0:r0c0';
  resolveLiveCell(arm.engine, moved);
  arm.engine.mergeCells(range(0, 0, 0, 1));
  expect(resolveLiveCell(arm.engine, moved)).toEqual({ story: 'body:t0:r0c0', tableIndex: 0, row: 0, column: 0 });
  expectProjectionEquivalent(arm.engine, [moved]);
}, TIMEOUT);

test('cells of a repeated section item clone keep the projection result', async () => {
  const w15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
  const repeat = `<w:sdt><w:sdtPr><w:id w:val="8"/><w15:repeatingSection ${w15}/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="9"/><w:alias w:val="Line"/><w15:repeatingSectionItem ${w15}/></w:sdtPr><w:sdtContent>${square([['R1', 'R2']])}${paragraph('Item')}</w:sdtContent></w:sdt></w:sdtContent></w:sdt>`;
  const arm = await openCommandEngine([], { bodyXml: paragraph('Before') + repeat + paragraph('After') });
  const before = new Set(arm.engine.geometryReader.storyIds());
  expectProjectionEquivalent(arm.engine);
  const item = arm.engine.controlTargets().find((target) => target.control.controlType === 'repeatingSectionItem');
  expect(arm.engine.repeatControlAt(item!.at, false)).toBe(true);
  const clones = arm.engine.geometryReader.storyIds().filter((story) => !before.has(story) &&
    arm.engine.paragraphs(story).some((p) => p.text === 'R1' || p.text === 'R2'));
  expect(clones.length).toBe(2);
  expectProjectionEquivalent(arm.engine);
  expect(clones.every((story) => resolveLiveCell(arm.engine, story) !== null)).toBe(true);
}, TIMEOUT);

test('a table in a block control whose embed was removed is located by neither path', async () => {
  const arm = await openCommandEngine([], { bodyXml: control + paragraph('After') });
  const detached = arm.engine.geometryReader.storyIds().filter((story) => /:sdt\d+:t\d+:r\d+c\d+$/.test(story));
  expect(detached.length).toBe(2);
  expect(detached.every((story) => resolveLiveCell(arm.engine, story) !== null)).toBe(true);
  arm.engine.applyRawOps('body', [{ op: 'delete', index: 0, len: 1 }]);
  expect(detached.map((story) => resolveLiveCell(arm.engine, story))).toEqual([null, null]);
  expectProjectionEquivalent(arm.engine);
}, TIMEOUT);

test('a cell story two tables reference resolves to the deeper table on both paths', async () => {
  const arm = await openCommandEngine([], { bodyXml: square([['A', 'B']]) + paragraph('After') });
  const [a, b] = ['body:t0:r0c0', 'body:t0:r0c1'];
  expect(resolveLiveCell(arm.engine, a)).toEqual({ story: 'body', tableIndex: 0, row: 0, column: 0 });
  arm.engine.applyRawOps(b, [{ op: 'insertEmbed', index: 0, kind: 'table', payload: {
    tblPr: {}, grid: [1500], rows: [{ trPr: {}, cells: [{ tcPr: { rowspan: 1, colspan: 1 }, story: a }] }],
  } }]);
  expect(resolveLiveCell(arm.engine, a)).toEqual({ story: b, tableIndex: 0, row: 0, column: 0 });
  expectProjectionEquivalent(arm.engine);
}, TIMEOUT);

for (const [name, edit] of [
  ['a row inserted above', (engine: DocxWorkerCommandEngine) => engine.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'above')],
  ['a deleted column', (engine: DocxWorkerCommandEngine) => engine.deleteColumn(range(0, 0, 2, 0))],
  ['a split cell', (engine: DocxWorkerCommandEngine) => engine.splitCell({ story: 'body', tableIndex: 0, row: 1, column: 1 }, 2, 2)],
  ['merged then inserted rows', (engine: DocxWorkerCommandEngine) => { engine.mergeCells(range(0, 1, 1, 2)); engine.insertRow({ story: 'body', tableIndex: 0, row: 1, column: 0 }, 'below'); }],
] as const) {
  test(`cell positions after ${name} match the projection for every story`, async () => {
    const arm = await openCommandEngine([], { bodyXml: paragraph('Before') + square([['A', 'B', 'C'], ['D', 'E', 'F'], ['G', 'H', 'I']]) + paragraph('After') });
    expectProjectionEquivalent(arm.engine);
    edit(arm.engine);
    expectProjectionEquivalent(arm.engine);
  }, TIMEOUT);
}
