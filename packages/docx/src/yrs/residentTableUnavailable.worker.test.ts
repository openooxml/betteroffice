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

function documentBytes(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraph('Before')}<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid><w:tr>${cell('First')}${cell('Second')}</w:tr></w:tbl>${paragraph('After')}<w:sectPr/></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function open(clientId: number) {
  const bytes = documentBytes();
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
  return { engine, session, bytes };
}

const merge = { lane: 'collab', fence: true } as const;
const at = (row: number, column: number): YrsCellLoc => ({ story: 'body', tableIndex: 0, row, column });

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
  expect((await replica.session.submit('applyRemoteUpdate', { update }, merge).ack).status).toBe('applied');
}

function send(replica: Opened, to: YrsSession): void {
  to.applyUpdate(replica.engine.encodeStateAsUpdate(to.encodeStateVector()));
}

function cellStory(replica: Opened): string {
  return (replica.engine.tablePayload('body', 0)!.rows as Array<{ cells: Array<{ story: string }> }>)[0]!.cells[0]!.story;
}

const cellTexts = (replica: Opened, story: string) => replica.engine.paragraphs(story).map(({ text }) => text);
const located = (replica: Opened, story: string) =>
  resolveLiveCell(replica.engine as unknown as DocxWorkerCommandEngine, story);

async function expectVisibleTyping(replica: Opened, story: string): Promise<void> {
  const selection = replica.session.state.selection!;
  expect(selection.anchor.story).toBe('body');
  expect(selection.focus.story).toBe('body');
  expect((await replica.session.submit('insertText', { text: 'Z' }).ack).status).toBe('applied');
  expect(replica.engine.paragraphs('body').some(({ text }) => text.includes('Z'))).toBe(true);
  expect(cellTexts(replica, story).join('')).not.toContain('Z');
}

test('a remote crossing undo moves a caret out of the cell it hides; typing lands visibly and the redo leaves the caret outside', async () => {
  setCellLocationShadowCheck(true);
  const replica = await open(9601);
  const first = await peer(replica, 9602);
  const second = await peer(replica, 9603);
  try {
    const story = cellStory(replica);
    const inCell = { story, paraId: replica.engine.paragraphs(story)[0]!.paraId, offset: 2 };
    await replica.session.submit('select', { selection: { anchor: inCell, focus: inCell } }).ack;
    expect(replica.session.state.selection?.focus.story).toBe(story);
    const before = cellTexts(replica, story);

    first.insertRow(at(0, 0), 'below');
    sync(first, second);
    second.setColumnWidth(at(0, 0), 2000);
    expect(first.undo()).toBe(true);
    sync(first, second);
    expect(second.undo()).toBe(true);
    sync(first, second);
    await receive(replica, second);
    expect(located(replica, story)).toBeNull();
    await expectVisibleTyping(replica, story);

    send(replica, second);
    expect(second.redo()).toBe(true);
    await receive(replica, second);
    expect(located(replica, story)).not.toBeNull();
    expect(cellTexts(replica, story)).toEqual(before);
    expect(replica.session.state.selection?.focus.story).toBe('body');
  } finally {
    setCellLocationShadowCheck(false);
    first.destroy();
    second.destroy();
    replica.session.dispose();
  }
}, TIMEOUT);

for (const range of [false, true]) {
  test(`the second undoer with a ${range ? 'range selected' : 'caret'} in the cell is moved next to the table; typing lands visibly and the redo restores the cell unchanged without pulling the caret back`, async () => {
    setCellLocationShadowCheck(true);
    const replica = await open(9611 + (range ? 10 : 0));
    const first = await peer(replica, 9612 + (range ? 10 : 0));
    try {
      const story = cellStory(replica);
      const paraId = replica.engine.paragraphs(story)[0]!.paraId;
      const anchor = { story, paraId, offset: range ? 0 : 2 };
      const focus = { story, paraId, offset: range ? 4 : 2 };
      await replica.session.submit('select', { selection: { anchor, focus } }).ack;
      const before = cellTexts(replica, story);

      first.insertRow(at(0, 0), 'below');
      await receive(replica, first);
      expect((await replica.session.submit('setTableCellShading', { color: 'FFCC00' }).ack).status).toBe('applied');
      send(replica, first);
      expect(first.undo()).toBe(true);
      await receive(replica, first);
      expect((await replica.session.submit('undo', {}).ack).status).toBe('applied');
      const payload = replica.engine.tablePayload('body', 0) ?? {};
      expect(['tblPr', 'grid', 'rows'].some((key) => !(key in payload))).toBe(true);
      expect(located(replica, story)).toBeNull();
      const moved = replica.session.state.selection!;
      expect(moved.anchor).toEqual(moved.focus);
      await expectVisibleTyping(replica, story);

      expect((await replica.session.submit('redo', {}).ack).status).toBe('applied');
      expect(located(replica, story)).not.toBeNull();
      expect(cellTexts(replica, story)).toEqual(before);
      expect(replica.session.state.selection?.focus.story).toBe('body');
    } finally {
      setCellLocationShadowCheck(false);
      first.destroy();
      replica.session.dispose();
    }
  }, TIMEOUT);
}

test('the cached cell location drops a table that lost only its properties and finds it again after the redo', async () => {
  setCellLocationShadowCheck(true);
  const replica = await open(9631);
  const first = await peer(replica, 9632);
  const second = await peer(replica, 9633);
  try {
    const story = cellStory(replica);
    expect(located(replica, story)).toEqual(at(0, 0));
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
    expect(located(replica, story)).toEqual(at(0, 0));
  } finally {
    setCellLocationShadowCheck(false);
    first.destroy();
    second.destroy();
    replica.session.dispose();
  }
}, TIMEOUT);
