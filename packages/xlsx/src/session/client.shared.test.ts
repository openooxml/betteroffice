import { expect, test } from 'bun:test';
import { isClientMessage, type ClientMessage } from '../../../../shared/office-session/protocol';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { createWorkbookSession, type WorkbookSession } from './client';
import type { WorkbookCommand, WorkbookCommandWireFrame, WorkbookSharedWireAck } from './commandProtocol';
import type { WorkbookSessionOpenReply } from './methods';

const identity = { session: 'book', incarnation: 'opaque:worker' };
const viewport = { x: 0, y: 0, width: 80, height: 40 };
const displayList = { width: 80, height: 40, commands: [] };
const opened: WorkbookSessionOpenReply = {
  format: 'xlsx', stage: 'ready', version: 0, dirty: false, activeSheet: 0,
  sheets: [{ id: 'sheet:0', index: 0, name: 'Data' }],
  commandContext: { ...identity, basis: { epoch: 1, version: 'opaque:0' } },
};

function paint(seq: number): WorkbookCommandWireFrame {
  return { ...identity, coversThroughSeq: seq, epoch: seq, sequence: seq, version: `opaque:${seq}`,
    sheet: 0, viewport, zoom: 1, dpr: 1, requestGeneration: 1,
    displayList: new TextEncoder().encode(JSON.stringify(displayList)).buffer };
}

function controlled(value: (throughSeq: number) => unknown = () => undefined) {
  const calls: Extract<ClientMessage, { kind: 'call' }>[] = [];
  const snapshots = new Map<number, unknown>();
  const events: string[] = [];
  let listener: (message: unknown) => void = () => {};
  let throughSeq = 0;
  const transport: SessionTransport = {
    post(message) {
      if (!isClientMessage(message) || message.kind !== 'call') return;
      if (message.method === 'open') {
        queueMicrotask(() => { listener({ protocol: 1, kind: 'reply', id: message.id, ok: true, value: opened }); });
        return;
      }
      calls.push(message);
      if (message.method === 'command' || message.method === 'edit') {
        events.push(`send:${(message.args[0] as { seq: number }).seq}`);
      } else {
        events.push(`${message.method}:${throughSeq}`);
        snapshots.set(message.id, value(throughSeq));
      }
    },
    listen(next) { listener = next; return () => { listener = () => {}; }; },
    onError: () => () => {},
    close: () => {},
  };
  return { transport, calls, events,
    ack(index: number, sharedFrame = false, compatibilityFrame = false) {
      const call = calls[index];
      const input = call.args[0] as { seq: number };
      throughSeq = input.seq;
      const status = { seq: throughSeq, status: 'applied' as const, committed: true };
      const compatibility = { throughSeq, version: `opaque:${throughSeq}`, statuses: [status],
        ...(compatibilityFrame ? { frame: paint(throughSeq) } : {}) };
      const ack: WorkbookSharedWireAck = {
        session: identity.session, incarnation: 1, epoch: 1, throughSeq,
        version: throughSeq, engineVersion: `opaque:${throughSeq}`,
        outcomes: [{ seq: throughSeq, status: 'applied', workbook: status }], compatibility,
        ...(sharedFrame ? { frame: { coversThroughSeq: throughSeq, revision: throughSeq, payload: paint(throughSeq) } } : {}),
      };
      listener({ protocol: 1, kind: 'reply', id: call.id, ok: true,
        value: call.method === 'command' ? { ...compatibility, shared: ack } : ack });
    },
    reply(index: number) {
      const call = calls[index];
      listener({ protocol: 1, kind: 'reply', id: call.id, ok: true, value: snapshots.get(call.id) });
    },
  };
}

function command(seq: number): WorkbookCommand {
  return { ...identity, seq, basis: { epoch: 1, version: 'opaque:0' },
    kind: 'editCell', payload: { args: [0, 0, 0, String(seq)] } };
}

async function drain() {
  for (let step = 0; step < 24; step += 1) await Promise.resolve();
}

const reads: { name: string; method: string; call(session: WorkbookSession): Promise<unknown>;
  value(seq: number): unknown; expected(seq: number): unknown }[] = [
  { name: 'save', method: 'save', call: (session) => session.save(),
    value: (seq) => Uint8Array.of(seq).buffer, expected: (seq) => Uint8Array.of(seq) },
  { name: 'call.save', method: 'save', call: (session) => session.call.save(),
    value: (seq) => Uint8Array.of(seq).buffer, expected: (seq) => Uint8Array.of(seq).buffer },
  { name: 'version', method: 'version', call: (session) => session.call.version(),
    value: (seq) => `opaque:${seq}`, expected: (seq) => `opaque:${seq}` },
  { name: 'cellInputs', method: 'cellInputs', call: (session) => session.call.cellInputs(0, 'A1'),
    value: (seq) => ({ sheet: 0, version: `opaque:${seq}`, cells: [[{ input: String(seq) }]] }),
    expected: (seq) => ({ sheet: 0, version: `opaque:${seq}`, cells: [[{ input: String(seq) }]] }) },
  { name: 'frame', method: 'frame', call: (session) => session.call.frame(viewport),
    value: paint, expected: (seq) => ({ ...paint(seq), displayList }) },
  ...['renderPng', 'renderRangePng'].map((method) => ({
    name: method, method,
    call: (session: WorkbookSession) => method === 'renderPng'
      ? session.call.renderPng(viewport, { sheet: 0 })
      : session.call.renderRangePng({ range: 'A1', scale: 2 }, { sheet: 0 }),
    value: (seq: number) => Uint8Array.of(137, 80, 78, 71, seq).buffer,
    expected: (seq: number) => Uint8Array.of(137, 80, 78, 71, seq),
  })),
  { name: 'printDisplayList', method: 'printDisplayList',
    call: (session) => session.call.printDisplayList(0, 'A1', {
      dpi: 96, maxDigitWidth: 7, defaultRowHeightPt: 15, fontSizePt: 11,
      fontFamily: 'Arial', fontAscent: 10, fontDescent: 3,
    }, true),
    value: (seq) => new TextEncoder().encode(JSON.stringify({ ...displayList, width: seq })).buffer,
    expected: (seq) => ({ ...displayList, width: seq }) },
  ...['exportStructured', 'exportMarkdown'].map((method) => ({
    name: method, method,
    call: (session: WorkbookSession) => method === 'exportStructured'
      ? session.call.exportStructured({ scope: [{ sheet: 0, range: 'A1' }] })
      : session.call.exportMarkdown({ scope: [{ sheet: 0, range: 'A1' }] }, { maxRows: 1 }),
    value: (seq: number) => ({ ok: true, version: `opaque:${seq}`, content: { complete: true } }),
    expected: (seq: number) => ({ ok: true, version: `opaque:${seq}`, content: { complete: true } }),
  })),
  ...[
    { name: 'readCells', call: (session: WorkbookSession) => session.call.readCells({ ranges: [] }),
      value: (seq: number) => ({ ok: true, version: `opaque:${seq}`, sheets: [], ranges: [],
        calculation: { cycleCells: [], limitedCells: [], truncated: false } }) },
    { name: 'findText', call: (session: WorkbookSession) => session.call.findText({ text: '65' }),
      value: (seq: number) => ({ ok: true, version: `opaque:${seq}`, matches: [], truncated: false }) },
    { name: 'validateEdits', call: (session: WorkbookSession) => session.call.validateEdits({ expectVersion: 'opaque:0', steps: [] }),
      value: (seq: number) => ({ ok: true, baseVersion: `opaque:${seq}`, wouldApply: false, previews: [] }) },
    { name: 'sheetView', call: (session: WorkbookSession) => session.call.sheetView(0),
      value: (seq: number) => ({ sheet: 0, version: `opaque:${seq}`, contentWidth: seq, contentHeight: 40,
        frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0, initialScrollX: 0, initialScrollY: 0 }) },
    { name: 'selectionData', call: (session: WorkbookSession) => session.call.selectionData(0, 0, 0, 'down'),
      value: (seq: number) => ({ edge: { row: seq, col: 0 }, region: null }) },
    { name: 'cellGeometry', call: (session: WorkbookSession) => session.call.cellGeometry(0, 0, 0),
      value: (seq: number) => ({ sheet: 0, version: `opaque:${seq}`, rect: { x: 0, y: 0, width: seq, height: 40 },
        scrollPosition: { x: 0, y: 0 } }) },
    { name: 'captureFormat', call: (session: WorkbookSession) => session.call.captureFormat(0, 'A1'),
      value: (seq: number) => ({ rows: 1, columns: 1, formats: [{ fontSize: seq }] }) },
    { name: 'sheets', call: (session: WorkbookSession) => session.call.sheets(),
      value: (seq: number) => [{ id: 'sheet:0', index: 0, name: String(seq) }] },
    { name: 'calculationStatus', call: (session: WorkbookSession) => session.call.calculationStatus(),
      value: (seq: number) => ({ limitedCells: [`A${seq}`] }) },
  ].map((read) => ({ ...read, method: read.name, expected: read.value })),
];

test('exports capture sheet, viewport, scope and print metrics before the command fence', async () => {
  const port = controlled(() => new ArrayBuffer(0));
  const session = await createWorkbookSession(new Uint8Array(), {}, port.transport);
  const window = { ...viewport };
  const range = { range: 'A1', scale: 2 };
  const options = { scope: [{ sheet: 0, range: 'A1' }], maxCells: 8 };
  const markdown = { maxRows: 2 };
  const metrics = { dpi: 96, maxDigitWidth: 7, defaultRowHeightPt: 15,
    fontSizePt: 11, fontFamily: 'Arial', fontAscent: 10, fontDescent: 3 };
  const sheet = { sheet: 0 };
  try {
    const earlier = session.submit(command(1));
    const exports = [session.call.renderPng(window, sheet), session.call.renderRangePng(range, sheet),
      session.call.exportStructured(options), session.call.exportMarkdown(options, markdown),
      session.call.printDisplayList(0, 'A1', metrics, true)];
    for (const pending of exports) void pending.catch(() => {});
    window.x = 900; range.range = 'B2'; options.scope[0].sheet = 1;
    options.maxCells = 1; markdown.maxRows = 1; metrics.dpi = 300; sheet.sheet = 1;
    port.ack(0);
    await earlier;
    await drain();
    expect(port.calls.slice(1).map((call) => [call.method, call.args])).toEqual([
      ['renderPng', [viewport, { sheet: 0 }]],
      ['renderRangePng', [{ range: 'A1', scale: 2 }, { sheet: 0 }]],
      ['exportStructured', [{ scope: [{ sheet: 0, range: 'A1' }], maxCells: 8 }]],
      ['exportMarkdown', [{ scope: [{ sheet: 0, range: 'A1' }], maxCells: 8 }, { maxRows: 2 }]],
      ['printDisplayList', [0, 'A1', { ...metrics, dpi: 96 }, true]],
    ]);
  } finally { await session.dispose(); }
});

for (const entry of ['compatibility', 'native'] as const) {
  for (const read of reads) {
    test(`${read.name} fences 65 ${entry} edits before any ack and holds later dispatch`, async () => {
      const port = controlled(read.value);
      const session = await createWorkbookSession(new Uint8Array(), {}, port.transport);
      const submit = (seq: number) => entry === 'compatibility' ? session.submit(command(seq)) :
        session.editSession!.submit('editCell', { payload: { args: [0, 0, 0, String(seq)] } }).ack;
      try {
        const earlier = Array.from({ length: 65 }, (_, index) => submit(index + 1));
        expect(port.events).toEqual(Array.from({ length: 64 }, (_, index) => `send:${index + 1}`));
        const result = read.call(session);
        const later = submit(66);
        expect(port.calls).toHaveLength(64);
        for (let index = 0; index < 64; index += 1) port.ack(index);
        await drain();
        expect(port.events).toEqual(Array.from({ length: 65 }, (_, index) => `send:${index + 1}`));
        port.ack(64);
        await Promise.all(earlier);
        await drain();
        expect(port.events).toEqual([...Array.from({ length: 65 }, (_, index) => `send:${index + 1}`), `${read.method}:65`]);
        port.reply(65);
        expect(await result).toEqual(read.expected(65));
        await drain();
        expect(port.events).toEqual([...Array.from({ length: 65 }, (_, index) => `send:${index + 1}`),
          `${read.method}:65`, 'send:66']);
        port.ack(66);
        await later;
      } finally { await session.dispose(); }
    });
  }
}

for (const frames of ['compatibility', 'shared', 'both', 'neither'] as const) {
  test(`decoded acknowledgement types preserve ${frames} frames`, async () => {
    const port = controlled();
    const session = await createWorkbookSession(new Uint8Array(), {}, port.transport);
    try {
      const pending = session.call.command(command(1));
      expect(port.events).toEqual(['send:1']);
      port.ack(0, frames === 'shared' || frames === 'both', frames === 'compatibility' || frames === 'both');
      const ack = await pending;
      expect(ack.statuses).toEqual([{ seq: 1, status: 'applied', committed: true }]);
      if (frames === 'neither') expect(ack.frame).toBeUndefined();
      else {
        expect(ack.frame).toEqual({ ...paint(1), displayList });
        expect(ack.frame!.displayList).not.toBeInstanceOf(ArrayBuffer);
      }
      expect(session.throughSeq).toBe(1);
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });
}

test('an ordinary RPC before the first edit holds later dispatch', async () => {
  const port = controlled((seq) => Uint8Array.of(seq).buffer);
  const session = await createWorkbookSession(new Uint8Array(), {}, port.transport);
  try {
    const saved = session.save();
    const later = session.submit(command(1));
    expect(port.events).toEqual(['save:0']);
    port.reply(0);
    expect(await saved).toEqual(Uint8Array.of(0));
    await drain();
    expect(port.events).toEqual(['save:0', 'send:1']);
    port.ack(1);
    expect(await later).toMatchObject({ seq: 1, status: 'applied' });
  } finally { await session.dispose(); }
});
