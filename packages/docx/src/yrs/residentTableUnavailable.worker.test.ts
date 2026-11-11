import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadOpcWasm } from '../wasm/opc';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { DocxSourceWorkerClient } from './residentSourceWorkerClient';
import { createDocxCommandConnection } from './residentCommandConnection';
import { createDocxWorkerEditSession } from './createDocxWorkerEditSession';
import { createYrsSession, type YrsCellLoc, type YrsSession } from './index';
import { tableUnavailableError } from './tableUnavailable';
import { closeCommandEngines, openCommandEngine, TIMEOUT } from './residentCommandTestSupport';
import type { DocxWorkerCommandEngine } from './residentCommandProtocol';
import { resolveLiveCell, setCellLocationShadowCheck } from './residentCommandTableSelection';

let start: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => {
  await openCommandEngine();
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(process.cwd(), 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm'))));
  start = await residentWorkerFactory({ transferBuffers: true });
}, 30_000);
afterEach(closeCommandEngines);

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;

const table = (...cells: string[]) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${cells.map(() => '<w:gridCol w:w="2340"/>').join('')}</w:tblGrid><w:tr>${cells.join('')}</w:tr></w:tbl>`;

/** A body with one two-cell table, or, `nested`, that table inside the first cell of a healthy outer table, or `body`. */
function documentBytes(nested: boolean, body?: string): Uint8Array {
  const inner = table(cell('First'), cell('Second'));
  const content = body ?? (nested
    ? table(`<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${inner}${paragraph('After inner')}</w:tc>`, cell('Outer two'))
    : inner);
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraph('Before')}${content}${paragraph('After')}<w:sectPr/></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function open(clientId: number, nested = false, body?: string) {
  const bytes = documentBytes(nested, body);
  const worker = start(clientId);
  const client = new DocxSourceWorkerClient(worker);
  await client.openSource({ bytes, collaboration: { clientId } });
  const engine = worker.sessions[0]!;
  const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 0 };
  const initial = await client.commandInit({ session: `replica-${clientId}`, incarnation: 1, epoch: 1,
    readOnly: false, selection: { anchor: at, focus: at },
    fonts: [new Uint8Array(readFileSync('crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))],
    layoutInput: JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } }, renderEnv: {} }),
    layoutWithRegions: true, layoutRevision: 1, extras: '{}',
  });
  const session = createDocxWorkerEditSession({ source: bytes,
    connection: await createDocxCommandConnection(client, initial),
    echo: { paint() {}, drop() {}, present() {} }, maxRecoveryAttempts: 0,
    recovery: { reopen: () => Promise.reject(new Error('Unexpected reopen')) },
  });
  return { engine, session, client, bytes, nested };
}

const merge = { lane: 'collab', fence: true } as const;

function sync(a: YrsSession, b: YrsSession): void {
  a.applyUpdate(b.encodeStateAsUpdate(a.encodeStateVector()));
  b.applyUpdate(a.encodeStateAsUpdate(b.encodeStateVector()));
}

type Opened = Awaited<ReturnType<typeof open>>;

async function peer(replica: Opened, clientId: number): Promise<YrsSession> {
  const created = await createYrsSession({ clientId });
  created.openDocx(replica.bytes, false);
  created.loadState(replica.engine.encodeState());
  return created;
}

async function receive(replica: Opened, from: YrsSession): Promise<void> {
  const update = from.encodeStateAsUpdate(replica.engine.encodeStateVector()).slice().buffer;
  const ack = await replica.session.submit('applyRemoteUpdate', { update }, merge).ack;
  expect(ack.status, JSON.stringify({ status: ack.status, reason: 'reason' in ack ? ack.reason : undefined })).toBe('applied');
}

function send(replica: Opened, to: YrsSession): void {
  to.applyUpdate(replica.engine.encodeStateAsUpdate(to.encodeStateVector()));
}

const firstCell = (replica: Opened, story: string) =>
  (replica.engine.tablePayload(story, 0)!.rows as Array<{ cells: Array<{ story: string }> }>)[0]!.cells[0]!.story;

/** The story holding the table crossing undos empty: the body, or the healthy outer table's first cell. */
const tableStory = (replica: Opened) => (replica.nested ? firstCell(replica, 'body') : 'body');
const cellStory = (replica: Opened) => firstCell(replica, tableStory(replica));
const at = (story: string, row: number, column: number): YrsCellLoc => ({ story, tableIndex: 0, row, column });

const cellTexts = (replica: Opened, story: string) => replica.engine.paragraphs(story).map(({ text }) => text);
const located = (replica: Opened, story: string) =>
  resolveLiveCell(replica.engine as unknown as DocxWorkerCommandEngine, story);

for (const nested of [false, true]) {
  const where = nested ? 'a table inside a healthy outer table' : 'a body table';
  const base = nested ? 9700 : 9600;

  test(`after a remote crossing undo on ${where}, a caret left in its cell types into the cell, saves refuse coded, and the restored table shows the text`, async () => {
    const replica = await open(base + 1, nested);
    const first = await peer(replica, base + 2);
    const second = await peer(replica, base + 3);
    try {
      const table = tableStory(replica);
      const story = cellStory(replica);
      const inCell = { story, paraId: replica.engine.paragraphs(story)[0]!.paraId, offset: 2 };
      await replica.session.submit('select', { selection: { anchor: inCell, focus: inCell } }).ack;
      expect(replica.session.state.selection?.focus.story).toBe(story);

      first.insertRow(at(table, 0, 0), 'below');
      sync(first, second);
      second.setColumnWidth(at(table, 0, 0), 2000);
      expect(first.undo()).toBe(true);
      sync(first, second);
      expect(second.undo()).toBe(true);
      sync(first, second);
      await receive(replica, second);
      expect(Array.isArray(replica.engine.tablePayload(table, 0)?.rows)).toBe(false);

      expect(replica.session.state.selection?.focus.story).toBe(story);
      expect((await replica.session.submit('insertText', { text: 'Z' }).ack).status).toBe('applied');
      expect(cellTexts(replica, story).join('')).toContain('Z');
      const unsaved = replica.engine.encodeState();
      await expect(replica.client.save({ comments: [] })).rejects.toThrow(tableUnavailableError().message);
      expect(replica.engine.encodeState()).toEqual(unsaved);

      send(replica, second);
      expect(second.redo()).toBe(true);
      await receive(replica, second);
      expect(located(replica, story)).toEqual(at(table, 0, 0));
      expect(cellTexts(replica, story).join('')).toContain('Z');
      const body = { story: 'body', paraId: replica.engine.paragraphs('body').at(-1)!.paraId, offset: 0 };
      await replica.session.submit('select', { selection: { anchor: body, focus: body } }).ack;
      expect((await replica.session.submit('insertText', { text: 'Y' }).ack).status).toBe('applied');
    } finally {
      first.destroy();
      second.destroy();
      replica.session.dispose();
    }
  }, TIMEOUT);
}

/** Counts every call this engine instance receives, by method name. */
const border = { style: 'single', size: 4, color: { rgb: '000000' } };
const CARET_TABLE_COMMANDS = [
  ['insertTableRow', { side: 'below' }], ['insertTableColumn', { side: 'right' }], ['deleteTableRow', {}],
  ['deleteTableColumn', {}], ['mergeTableCells', {}], ['splitTableCell', { rows: 1, columns: 2 }],
  ['setTableCellShading', { color: 'FF0000' }], ['setTableBorders', { borders: { top: border } }],
  ['setTableProperties', { op: 'patch', patch: { justification: 'center' } }], ['deleteTable', {}],
] as const;

test('with the caret in a cell of an inner table that lost its rows, every table command refuses coded and the outer table is untouched', async () => {
  const replica = await open(9801, true);
  const first = await peer(replica, 9802);
  const second = await peer(replica, 9803);
  try {
    const table = tableStory(replica);
    const story = cellStory(replica);
    const inCell = { story, paraId: replica.engine.paragraphs(story)[0]!.paraId, offset: 0 };
    first.insertRow(at(table, 0, 0), 'below');
    sync(first, second);
    second.setColumnWidth(at(table, 0, 0), 2000);
    expect(first.undo()).toBe(true);
    sync(first, second);
    expect(second.undo()).toBe(true);
    sync(first, second);
    await receive(replica, second);
    expect(Array.isArray(replica.engine.tablePayload(table, 0)?.rows)).toBe(false);

    const outer = () => JSON.stringify(replica.engine.tablePayload('body', 0));
    const intact = outer();
    const outerCells = (replica.engine.tablePayload('body', 0) as { rows: Array<{ cells: Array<{ story: string }> }> })
      .rows[0]!.cells.map((entry) => entry.story);
    const selection = { anchor: inCell, focus: inCell };
    for (const [kind, payload] of CARET_TABLE_COMMANDS) {
      const ack = await replica.session.submit(kind as 'deleteTable', { ...payload, selection } as never).ack;
      expect([kind, ack.status, (ack as { reason?: string }).reason]).toEqual([kind, 'refused', 'table-unavailable']);
      expect(outer()).toBe(intact);
    }
    expect(replica.engine.tablePayload(table, 0)).not.toBeNull();
    expect(cellTexts(replica, outerCells[1]!)).toEqual(['Outer two']);
  } finally {
    first.destroy();
    second.destroy();
    replica.session.dispose();
  }
}, TIMEOUT);

function countEngineCalls(raw: object): () => Record<string, number> {
  let counts: Record<string, number> = {};
  for (let proto = Object.getPrototypeOf(raw); proto && proto !== Object.prototype; proto = Object.getPrototypeOf(proto)) {
    for (const name of Object.getOwnPropertyNames(proto)) {
      const method = Reflect.get(proto, name);
      if (name === 'constructor' || typeof method !== 'function' || Object.hasOwn(raw, name)) continue;
      Reflect.set(raw, name, (...args: unknown[]) => {
        counts[name] = (counts[name] ?? 0) + 1;
        return Reflect.apply(method, raw, args);
      });
    }
  }
  return () => {
    const taken = counts;
    counts = {};
    return taken;
  };
}

for (const where of ['body paragraph', 'healthy table cell'] as const) {
  test(`typing in a ${where} makes the same engine calls on every key`, async () => {
    const { command, execute, selection } = await import('./residentCommandTestSupport');
    const arm = where === 'body paragraph'
      ? await openCommandEngine(['Alpha beta', 'Second'])
      : await openCommandEngine([], { bodyXml: paragraph('Before') + table(cell('First'), cell('Second')) + paragraph('After') });
    const story = where === 'body paragraph'
      ? 'body'
      : (arm.engine.tablePayload('body', 0)!.rows as Array<{ cells: Array<{ story: string }> }>)[0]!.cells[0]!.story;
    const paraId = arm.engine.paragraphs(story)[0]!.paraId;
    const take = countEngineCalls(arm.native);
    const at = (offset: number) => ({ ...selection(paraId, offset, offset, story) });
    expect(execute(arm, command(arm, 'insertText', { text: 'w', selection: at(0) })).status).toBe('applied');
    take();
    const perKey: number[] = [];
    let names: Record<string, number> = {};
    for (const [index, text] of ['x', 'y', 'z'].entries()) {
      expect(execute(arm, command(arm, 'insertText', { text, selection: at(index + 1) })).status).toBe('applied');
      names = take();
      perKey.push(Object.values(names).reduce((sum, count) => sum + count, 0));
    }
    console.log(`WORKCOUNT ${where} engine calls per key ${perKey.join(',')} ${JSON.stringify(Object.fromEntries(Object.entries(names).sort()))}`);
    expect(new Set(perKey).size).toBe(1);
  }, TIMEOUT);
}

test('the cached cell location drops a table that lost only its properties and finds it again after the redo', async () => {
  setCellLocationShadowCheck(true);
  const replica = await open(9631);
  const first = await peer(replica, 9632);
  const second = await peer(replica, 9633);
  try {
    const story = cellStory(replica);
    expect(located(replica, story)).toEqual(at('body', 0, 0));
    first.setTableProperties({ story: 'body', tableIndex: 0 }, { width: 4000, widthType: 'dxa' });
    sync(first, second);
    second.setTableProperties({ story: 'body', tableIndex: 0 }, { justification: 'center' });
    expect(first.undo()).toBe(true);
    sync(first, second);
    expect(second.undo()).toBe(true);
    sync(first, second);
    const payload = second.tablePayload('body', 0) ?? {};
    expect(['tblPr', 'grid', 'rows'].filter((key) => !(key in payload))).toEqual(['tblPr']);
    await receive(replica, second);
    expect(located(replica, story)).toBeNull();
    expect(second.redo()).toBe(true);
    await receive(replica, second);
    expect(located(replica, story)).toEqual(at('body', 0, 0));
  } finally {
    setCellLocationShadowCheck(false);
    first.destroy();
    second.destroy();
    replica.session.dispose();
  }
}, TIMEOUT);

const cellRange = (row: number, column: number) => ({ anchor: at('body', row, column), head: at('body', row, column) });

const LOST: Array<{ name: string; key: string; edit: (session: YrsSession, first: boolean) => void }> = [
  { name: 'its properties', key: 'tblPr', edit: (session, first) =>
    session.setTableProperties({ story: 'body', tableIndex: 0 }, first ? { width: 4000, widthType: 'dxa' } : { justification: 'center' }) },
  { name: 'its grid', key: 'grid', edit: (session, first) => session.setColumnWidth(at('body', 0, first ? 0 : 1), first ? 2000 : 3000) },
  { name: 'its rows', key: 'rows', edit: (session, first) => session.setCellShading(cellRange(0, first ? 0 : 1), first ? 'FF0000' : '00FF00') },
];

for (const [index, lost] of LOST.entries()) {
  test(`Tab in the last cell of a table that lost ${lost.name} is refused or moves out cleanly and the session keeps working`, async () => {
    const base = 9800 + index * 10;
    const replica = await open(base + 1);
    const first = await peer(replica, base + 2);
    const second = await peer(replica, base + 3);
    try {
      const last = (replica.engine.tablePayload('body', 0)!.rows as Array<{ cells: Array<{ story: string }> }>)[0]!.cells[1]!.story;
      const inLast = { story: last, paraId: replica.engine.paragraphs(last)[0]!.paraId, offset: 0 };
      lost.edit(first, true);
      sync(first, second);
      lost.edit(second, false);
      expect(first.undo()).toBe(true);
      sync(first, second);
      expect(second.undo()).toBe(true);
      sync(first, second);
      const payload = second.tablePayload('body', 0) ?? {};
      expect(['tblPr', 'grid', 'rows'].filter((key) => !(key in payload))).toContain(lost.key);
      await receive(replica, second);
      const before = replica.engine.encodeState();

      const outcome = await replica.session.submit('navigateCell', {
        selection: { anchor: inLast, focus: inLast }, direction: 'forward', extend: false, createLastRow: true,
      }).ack;
      expect(['refused', 'noop', 'applied']).toContain(outcome.status);
      expect(replica.engine.encodeState()).toEqual(before);

      const bodyParaId = replica.engine.paragraphs('body')[0]!.paraId;
      const body = { story: 'body', paraId: bodyParaId, offset: 0 };
      expect((await replica.session.submit('insertText', { text: 'Z', selection: { anchor: body, focus: body } }).ack).status).toBe('applied');
      expect(replica.engine.paragraphs('body')[0]!.text.startsWith('Z')).toBe(true);
    } finally {
      first.destroy();
      second.destroy();
      replica.session.dispose();
    }
  }, TIMEOUT);
}

const INNER_LOST: Array<{ name: string; key: string; edit: (session: YrsSession, story: string, first: boolean) => void }> = [
  { name: 'its properties', key: 'tblPr', edit: (session, story, first) =>
    session.setTableProperties({ story, tableIndex: 0 }, first ? { width: 4000, widthType: 'dxa' } : { justification: 'center' }) },
  { name: 'its grid', key: 'grid', edit: (session, story, first) => session.setColumnWidth(at(story, 0, first ? 0 : 1), first ? 2000 : 3000) },
];

for (const [index, lost] of INNER_LOST.entries()) {
  test(`with the caret in a cell of an inner table that lost ${lost.name} but kept its rows, deleteTable removes it and saves, and every other table command refuses coded`, async () => {
    const base = 9900 + index * 10;
    const replica = await open(base + 1, true);
    const first = await peer(replica, base + 2);
    const second = await peer(replica, base + 3);
    try {
      const table = tableStory(replica);
      const story = cellStory(replica);
      const inCell = { story, paraId: replica.engine.paragraphs(story)[0]!.paraId, offset: 0 };
      lost.edit(first, table, true);
      sync(first, second);
      lost.edit(second, table, false);
      expect(first.undo()).toBe(true);
      sync(first, second);
      expect(second.undo()).toBe(true);
      sync(first, second);
      await receive(replica, second);
      const payload = replica.engine.tablePayload(table, 0) ?? {};
      expect(['tblPr', 'grid', 'rows'].filter((key) => !(key in payload))).toContain(lost.key);
      expect(Array.isArray((payload as { rows?: unknown }).rows)).toBe(true);
      expect(located(replica, story)).toBeNull();

      const outerCells = (replica.engine.tablePayload('body', 0) as { rows: Array<{ cells: Array<{ story: string }> }> })
        .rows[0]!.cells.map((entry) => entry.story);
      const outerTexts = outerCells.map((entry) => cellTexts(replica, entry));
      const selection = { anchor: inCell, focus: inCell };
      const untouched = replica.engine.encodeState();
      for (const [kind, extra] of CARET_TABLE_COMMANDS.filter(([kind]) => kind !== 'deleteTable')) {
        const ack = await replica.session.submit(kind as 'deleteTable', { ...extra, selection } as never).ack;
        expect([kind, ack.status, (ack as { reason?: string }).reason]).toEqual([kind, 'refused', 'table-unavailable']);
        expect(replica.engine.encodeState()).toEqual(untouched);
      }

      const removed = await replica.session.submit('deleteTable', { selection } as never).ack;
      expect([removed.status, (removed as { reason?: string }).reason]).toEqual(['applied', undefined]);
      expect(replica.engine.tablePayload(table, 0)).toBeNull();
      expect((replica.engine.tablePayload('body', 0) as { rows: Array<{ cells: Array<{ story: string }> }> })
        .rows[0]!.cells.map((entry) => entry.story)).toEqual(outerCells);
      expect(cellTexts(replica, outerCells[0]!)).toEqual(outerTexts[0]);
      expect(cellTexts(replica, outerCells[1]!)).toEqual(['Outer two']);
      const saved = await replica.client.save({ comments: [] });
      expect(saved.bytes.byteLength).toBeGreaterThan(0);
    } finally {
      first.destroy();
      second.destroy();
      replica.session.dispose();
    }
  }, TIMEOUT);
}

test('with the caret in a cell of an inner table that lost its properties and whose cell ids name no parent story, every table command refuses coded and the outer table is untouched', async () => {
  const replica = await open(9951, true);
  const host = await peer(replica, 9952);
  const first = await peer(replica, 9953);
  const second = await peer(replica, 9954);
  try {
    const table = tableStory(replica);
    const foreign = ['foreign-cell-1', 'ghost:t0:r0c1'];
    const minted = replica.engine.tablePayload(table, 0) as { rows: Array<{ cells: Array<{ story: string }> }> };
    host.createStory(foreign[0]!, 'Foreign one');
    host.createStory(foreign[1]!, 'Foreign two');
    host.applyRawOps(table, [{ op: 'insertEmbed', index: 0, kind: 'table', payload: {
      ...minted, rows: minted.rows.map((row) => ({ ...row, cells: row.cells.map((entry, column) => ({ ...entry, story: foreign[column]! })) })),
    } }]);
    sync(host, first);
    sync(host, second);
    await receive(replica, host);
    expect(firstCell(replica, table)).toBe(foreign[0]!);

    first.setTableProperties({ story: table, tableIndex: 0 }, { width: 4000, widthType: 'dxa' });
    sync(first, second);
    second.setTableProperties({ story: table, tableIndex: 0 }, { justification: 'center' });
    expect(first.undo()).toBe(true);
    sync(first, second);
    expect(second.undo()).toBe(true);
    sync(first, second);
    await receive(replica, second);
    const payload = (replica.engine.tablePayload(table, 0) ?? {}) as { rows?: Array<{ cells: Array<{ story: string }> }> };
    expect('tblPr' in payload).toBe(false);
    expect(payload.rows?.[0]?.cells.map((entry) => entry.story)).toEqual(foreign);
    expect(foreign.map((story) => located(replica, story))).toEqual([null, null]);

    const outer = () => JSON.stringify(replica.engine.tablePayload('body', 0));
    const intact = outer();
    const outerCells = (replica.engine.tablePayload('body', 0) as { rows: Array<{ cells: Array<{ story: string }> }> })
      .rows[0]!.cells.map((entry) => entry.story);
    const untouched = replica.engine.encodeState();
    for (const story of foreign) {
      const inCell = { story, paraId: replica.engine.paragraphs(story)[0]!.paraId, offset: 0 };
      const selection = { anchor: inCell, focus: inCell };
      for (const [kind, extra] of CARET_TABLE_COMMANDS) {
        const ack = await replica.session.submit(kind as 'deleteTable', { ...extra, selection } as never).ack;
        expect([story, kind, ack.status, (ack as { reason?: string }).reason]).toEqual([story, kind, 'refused', 'invalid-cell-target']);
        expect(outer()).toBe(intact);
        expect(replica.engine.encodeState()).toEqual(untouched);
      }
    }
    expect(replica.engine.tablePayload(table, 0)).not.toBeNull();
    expect(foreign.map((story) => cellTexts(replica, story))).toEqual([['Foreign one'], ['Foreign two']]);

    const inOuter = { story: outerCells[1]!, paraId: replica.engine.paragraphs(outerCells[1]!)[0]!.paraId, offset: 0 };
    const typed = await replica.session.submit('insertText', { text: 'Z', selection: { anchor: inOuter, focus: inOuter } }).ack;
    expect(typed.status).toBe('applied');
    expect(cellTexts(replica, outerCells[1]!)).toEqual(['ZOuter two']);
  } finally {
    host.destroy();
    first.destroy();
    second.destroy();
    replica.session.dispose();
  }
}, TIMEOUT);

const siblingTables = table(cell('A one'), cell('A two')) + paragraph('Between') + table(cell('B one'), cell('B two'));
const siblingRange = (tableIndex: number, column: number) => {
  const loc = { story: 'body', tableIndex, row: 0, column };
  return { anchor: loc, head: loc };
};

test('after deleteTable removes an unavailable table, a second deleteTable from its old cell refuses and its sibling that lost its rows is untouched', async () => {
  const replica = await open(9961, false, siblingTables);
  const first = await peer(replica, 9962);
  const second = await peer(replica, 9963);
  const third = await peer(replica, 9964);
  const fourth = await peer(replica, 9965);
  try {
    const story = firstCell(replica, 'body');
    const inCell = { story, paraId: replica.engine.paragraphs(story)[0]!.paraId, offset: 0 };
    first.setTableProperties({ story: 'body', tableIndex: 0 }, { width: 4000, widthType: 'dxa' });
    sync(first, second);
    second.setTableProperties({ story: 'body', tableIndex: 0 }, { justification: 'center' });
    expect(first.undo()).toBe(true);
    sync(first, second);
    expect(second.undo()).toBe(true);
    sync(first, second);
    third.setCellShading(siblingRange(1, 0), 'FF0000');
    sync(third, fourth);
    fourth.setCellShading(siblingRange(1, 1), '00FF00');
    expect(third.undo()).toBe(true);
    sync(third, fourth);
    expect(fourth.undo()).toBe(true);
    sync(third, fourth);
    await receive(replica, second);
    await receive(replica, fourth);

    const lostTblPr = (replica.engine.tablePayload('body', 0) ?? {}) as { rows?: Array<{ cells: Array<{ story: string }> }> };
    expect('tblPr' in lostTblPr).toBe(false);
    expect(lostTblPr.rows?.[0]?.cells[0]?.story).toBe(story);
    expect('rows' in (replica.engine.tablePayload('body', 1) ?? {})).toBe(false);
    expect(located(replica, story)).toBeNull();

    const sibling = JSON.stringify(replica.engine.tablePayload('body', 1));
    const selection = { anchor: inCell, focus: inCell };
    const removed = await replica.session.submit('deleteTable', { selection } as never).ack;
    expect([removed.status, (removed as { reason?: string }).reason]).toEqual(['applied', undefined]);
    expect(replica.engine.tablePayload('body', 1)).toBeNull();
    expect(JSON.stringify(replica.engine.tablePayload('body', 0))).toBe(sibling);
    expect(cellTexts(replica, story)).toEqual(['A one']);

    const untouched = replica.engine.encodeState();
    const again = await replica.session.submit('deleteTable', { selection } as never).ack;
    expect([again.status, (again as { reason?: string }).reason]).toEqual(['refused', 'table-unavailable']);
    expect(JSON.stringify(replica.engine.tablePayload('body', 0))).toBe(sibling);
    expect(replica.engine.encodeState()).toEqual(untouched);
  } finally {
    first.destroy();
    second.destroy();
    third.destroy();
    fourth.destroy();
    replica.session.dispose();
  }
}, TIMEOUT);

test('with the caret in a cell of an inner table a merge moved into another outer cell and that then lost its properties, every table command refuses coded and the outer table is untouched', async () => {
  const replica = await open(9971, true);
  const host = await peer(replica, 9972);
  const first = await peer(replica, 9973);
  const second = await peer(replica, 9974);
  try {
    const source = tableStory(replica);
    const outerCells = (replica.engine.tablePayload('body', 0) as { rows: Array<{ cells: Array<{ story: string }> }> })
      .rows[0]!.cells.map((entry) => entry.story);
    const target = outerCells[1]!;
    const moved = replica.engine.tablePayload(source, 0) as { rows: Array<{ cells: Array<{ story: string }> }> };
    const inner = moved.rows[0]!.cells.map((entry) => entry.story);
    expect(inner.every((story) => story.startsWith(`${source}:`))).toBe(true);
    host.applyRawOps(source, [{ op: 'delete', index: 0, len: 1 }]);
    host.applyRawOps(target, [{ op: 'insertEmbed', index: 0, kind: 'table', payload: { ...moved } }]);
    sync(host, first);
    sync(host, second);
    await receive(replica, host);
    expect(replica.engine.tablePayload(source, 0)).toBeNull();
    expect(firstCell(replica, target)).toBe(inner[0]!);

    first.setTableProperties({ story: target, tableIndex: 0 }, { width: 4000, widthType: 'dxa' });
    sync(first, second);
    second.setTableProperties({ story: target, tableIndex: 0 }, { justification: 'center' });
    expect(first.undo()).toBe(true);
    sync(first, second);
    expect(second.undo()).toBe(true);
    sync(first, second);
    await receive(replica, second);
    const payload = (replica.engine.tablePayload(target, 0) ?? {}) as { rows?: Array<{ cells: Array<{ story: string }> }> };
    expect('tblPr' in payload).toBe(false);
    expect(payload.rows?.[0]?.cells.map((entry) => entry.story)).toEqual(inner);
    expect(inner.map((story) => located(replica, story))).toEqual([null, null]);

    const outer = () => JSON.stringify(replica.engine.tablePayload('body', 0));
    const intact = outer();
    const untouched = replica.engine.encodeState();
    for (const story of inner) {
      const inCell = { story, paraId: replica.engine.paragraphs(story)[0]!.paraId, offset: 0 };
      const selection = { anchor: inCell, focus: inCell };
      for (const [kind, extra] of CARET_TABLE_COMMANDS) {
        const ack = await replica.session.submit(kind as 'deleteTable', { ...extra, selection } as never).ack;
        expect([story, kind, ack.status, (ack as { reason?: string }).reason]).toEqual([story, kind, 'refused', 'invalid-cell-target']);
        expect(outer()).toBe(intact);
        expect(replica.engine.encodeState()).toEqual(untouched);
      }
    }
    expect(replica.engine.tablePayload(target, 0)).not.toBeNull();
    expect(inner.map((story) => cellTexts(replica, story))).toEqual([['First'], ['Second']]);

    const body = { story: 'body', paraId: replica.engine.paragraphs('body').at(-1)!.paraId, offset: 0 };
    const typed = await replica.session.submit('insertText', { text: 'Z', selection: { anchor: body, focus: body } }).ack;
    expect(typed.status).toBe('applied');
  } finally {
    host.destroy();
    first.destroy();
    second.destroy();
    replica.session.dispose();
  }
}, TIMEOUT);
