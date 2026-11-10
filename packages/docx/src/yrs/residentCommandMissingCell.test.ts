import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { CommandOutcome } from '../../../../shared/office-session/commands/types';
import { isTransferResult } from '../../../../shared/office-session/protocol';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsCellLoc, type YrsSession } from './index';
import { applyDocxCommand, createResidentDocxCommandHooks, preflightDocxCommand } from './residentCommandActions';
import { projectDocxEditState } from './residentCommandEditState';
import { createResidentCommandHost } from './residentCommandHost';
import type { DocxCommandHost, DocxCommandIntent, DocxNavigationIntent } from './residentCommandProtocol';
import { resolveLiveCell } from './residentCommandTableSelection';
import {
  closeCommandEngines, command, context, execute, openCommandEngine, selection, type CommandArm,
} from './residentCommandTestSupport';
import { readResidentCellSelectionRects, readResidentFrameGeometry } from './residentObjectGeometry';

const WASM = fileURLToPath(new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url));
const sessions: YrsSession[] = [];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))), 30_000);
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
  closeCommandEngines();
});

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;
const body = paragraph('Before') +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid><w:tr>${cell('A')}${cell('B')}</w:tr></w:tbl>` +
  paragraph('After');

function docx(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const at = (row: number, column: number): YrsCellLoc => ({ story: 'body', tableIndex: 0, row, column });
const message = (error: unknown) => String((error as Error)?.message ?? error).slice(0, 200);

function sync(a: YrsSession, b: YrsSession): void {
  a.applyUpdate(b.encodeStateAsUpdate(a.encodeStateVector()));
  b.applyUpdate(a.encodeStateAsUpdate(b.encodeStateVector()));
}

function danglingCells(session: YrsSession): string[] {
  const stories = session.storyIds();
  const missing: string[] = [];
  for (const story of stories) {
    for (const segment of session.storySegments(story)) {
      if (segment.kind !== 'embed' || segment.embedKind !== 'table') continue;
      for (const row of (segment.payload.rows as { cells?: { story?: string }[] }[] | undefined) ?? []) {
        for (const named of row.cells ?? []) if (named.story && !stories.includes(named.story)) missing.push(named.story);
      }
    }
  }
  return missing;
}

/**
 * The encoded state of a document whose live table names a missing cell: A splits a cell, B
 * inserts a row (exchanged before the undo, or concurrently with it), A undoes the split.
 */
async function danglingState(base: number, order: 'synced' | 'concurrent'): Promise<{ state: Uint8Array; story: string }> {
  const a = await createYrsSession({ clientId: base + 1 });
  sessions.push(a);
  a.openDocx(docx(), true);
  a.beginUndoCapture();
  const b = await createYrsSession({ clientId: base + 2 });
  sessions.push(b);
  b.openDocx(docx(), false);
  b.loadState(a.encodeState());
  b.beginUndoCapture();
  a.splitCell(at(0, 1), 1, 2);
  sync(a, b);
  b.insertRow(at(0, 0), 'below');
  if (order === 'synced') sync(a, b);
  expect(a.undo()).toBe(true);
  sync(a, b);
  const [story] = danglingCells(b);
  expect(story).toBeDefined();
  return { state: b.encodeState(), story: story! };
}

async function openOn(base: number, order: 'synced' | 'concurrent') {
  const { state, story } = await danglingState(base, order);
  const arm = await openCommandEngine([], { bytes: docx(), checkpoint: state, clientId: base + 5 });
  return { arm, story, paraId: arm.engine.paragraphs(story)[0]!.paraId };
}

function stateBytes(arm: CommandArm): string {
  return `${Array.from(arm.engine.encodeStateVector()).join(',')}|${Array.from(arm.engine.encodeStateAsUpdate()).join(',')}`;
}

function layout(arm: CommandArm) {
  arm.engine.layoutDocumentWithRegionsRetained(JSON.stringify({
    bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } }, renderEnv: {},
  }));
  const frame = applyFrameDelta(null, decodeFrameDelta(arm.engine.buildDisplayListFrame('{}', 0)));
  return { frame, layoutRevision: arm.version };
}

function navigate(arm: CommandArm, payload: DocxNavigationIntent) {
  const intent = command(arm, 'navigate', payload);
  const ctx = context(arm, intent);
  const checked = preflightDocxCommand(arm.engine, arm.state, intent, ctx, { readSelectionLayout: () => layout(arm) });
  if (checked.status === 'refused') return checked;
  return applyDocxCommand(arm.engine, arm.state, checked.prepared, ctx);
}

const NAVIGATION: DocxNavigationIntent[] = [
  { unit: 'character', direction: 'forward', extend: false },
  { unit: 'character', direction: 'backward', extend: false },
  { unit: 'character', direction: 'forward', extend: true },
  { unit: 'word', direction: 'forward', extend: false },
  { unit: 'paragraph', direction: 'backward', extend: false },
  { unit: 'document-boundary', direction: 'forward', extend: false },
  { unit: 'line', direction: 'down', extend: false, goalX: 0 },
  { unit: 'line', direction: 'up', extend: false, goalX: 0 },
];

for (const [index, order] of (['synced', 'concurrent'] as const).entries()) {
  test(`editor reads with the caret in a missing cell (${order}): toolbar, live cell, navigation and object geometry write nothing`, async () => {
    const { arm, story, paraId } = await openOn(11_000 + index * 10, order);
    expect(paraId).toBe(`${story}#p0`);
    const before = stateBytes(arm);
    const caret = selection(paraId, 0, 0, story);
    expect(execute(arm, command(arm, 'select', { selection: caret })).status).toBe('noop');
    expect(arm.state.selection?.focus).toMatchObject({ story, paraId, offset: 0 });
    projectDocxEditState(arm.engine, arm.state);
    arm.engine.selectionContext({ story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } });
    const live = resolveLiveCell(arm.engine, story);
    expect(live).toMatchObject({ story: 'body', tableIndex: 0 });
    const outcomes: string[] = [];
    for (const payload of NAVIGATION) {
      expect(execute(arm, command(arm, 'select', { selection: caret })).status).toBe('noop');
      try { outcomes.push(`${JSON.stringify(payload)} ${navigate(arm, payload).status}`); } catch (error) { outcomes.push(`${JSON.stringify(payload)} throws: ${message(error)}`); }
    }
    console.log(`EDITOR-NAVIGATION ${order} ${JSON.stringify(outcomes)}`);
    expect(outcomes.filter((outcome) => outcome.includes(' throws: '))).toEqual([]);
    expect(execute(arm, command(arm, 'select', { selection: caret })).status).toBe('noop');
    const { frame } = layout(arm);
    await readResidentFrameGeometry(arm.engine, frame, arm.state.selection, 1);
    if (live && live.column > 0) {
      arm.engine.setCellSelection({ anchor: { ...live, column: live.column - 1 }, head: live });
      readResidentCellSelectionRects(arm.engine, frame, arm.engine.cellSelection());
    }
    expect(arm.engine.paragraphs(story).map((entry) => entry.text)).toEqual(['']);
    expect(stateBytes(arm) === before).toBe(true);
  }, 60_000);
}

async function send(host: DocxCommandHost, intent: DocxCommandIntent): Promise<CommandOutcome | undefined> {
  host.policy(intent);
  const reply = await host.handler(undefined, intent);
  const ack = isTransferResult(reply) ? reply.value : reply;
  return ack.outcomes[0];
}

const KEYS: Array<{ name: string; intent: (base: { session: string; incarnation: number }) => DocxCommandIntent; allowed: string[]; reason?: string }> = [
  {
    name: 'typing a character',
    intent: (base) => ({ ...base, seq: 2, basis: { epoch: 1, version: 0 }, kind: 'insertText', payload: { text: 'x' } }),
    allowed: ['refused'],
    reason: 'unsupported',
  },
  {
    name: 'Backspace',
    intent: (base) => ({ ...base, seq: 2, basis: { epoch: 1, version: 0 }, kind: 'deleteBackward', payload: {} }),
    allowed: ['refused', 'noop'],
  },
];

for (const [index, key] of KEYS.entries()) {
  test(`${key.name} with the caret in a missing cell takes the refused-command path: no rejection, no unhandled error, state unchanged`, async () => {
    const { arm, story, paraId } = await openOn(11_100 + index * 10, 'synced');
    const init = {
      session: 'd1-worker', incarnation: 1, epoch: 1, readOnly: false,
      fonts: [], layoutInput: '{}', layoutWithRegions: true, layoutRevision: 1, extras: '{}',
    };
    const hooks = createResidentDocxCommandHooks(arm.engine, init);
    const host = createResidentCommandHost({
      init, registry: { handlers: {}, policies: {} }, hooks: {
        ...hooks,
        frame: ({ coversThroughSeq }) => ({ coversThroughSeq, revision: coversThroughSeq + 1, payload: {
          frame: new ArrayBuffer(0), caret: { frameEpoch: coversThroughSeq + 1, caretRect: null }, layoutRevision: 1,
        } }),
        snapshot: () => arm.engine.encodeState(),
      },
    });
    await host.initial();
    const base = { session: init.session, incarnation: 1 };
    const selected = await send(host, {
      ...base, seq: 1, basis: { epoch: 1, version: 0 }, kind: 'select', payload: { selection: selection(paraId, 0, 0, story) },
    });
    expect(selected?.status).toBe('noop');
    const before = stateBytes(arm);
    const unhandled: unknown[] = [];
    const onUnhandled = (reason: unknown) => { unhandled.push(reason); };
    process.on('unhandledRejection', onUnhandled);
    let outcome: CommandOutcome | undefined;
    let failure: unknown;
    try {
      outcome = await send(host, key.intent(base));
    } catch (error) {
      failure = error;
    } finally {
      await new Promise<void>((resolve) => setImmediate(resolve));
      process.off('unhandledRejection', onUnhandled);
    }
    console.log(`EDITOR-REFUSAL ${key.name} ${JSON.stringify({ outcome, failure: failure === undefined ? null : message(failure) })}`);
    expect(failure === undefined ? 'settled' : `rejected: ${message(failure)}`).toBe('settled');
    expect(key.allowed).toContain(outcome?.status ?? 'missing');
    if (key.reason && outcome?.status === 'refused') expect(outcome.reason).toBe(key.reason);
    expect(unhandled.map(message)).toEqual([]);
    expect(stateBytes(arm) === before).toBe(true);
    expect(arm.engine.paragraphs(story).map((entry) => entry.text)).toEqual(['']);
    const bodyParaId = arm.engine.paragraphs('body')[0]!.paraId;
    const after = await send(host, {
      ...base, seq: 3, basis: { epoch: 1, version: 0 }, kind: 'select', payload: { selection: selection(bodyParaId, 0) },
    });
    expect(after?.status).toBe('noop');
    const typed = await send(host, {
      ...base, seq: 4, basis: { epoch: 1, version: 0 }, kind: 'insertText', payload: { text: 'Z', selection: selection(bodyParaId, 0) },
    });
    expect(typed?.status).toBe('applied');
    expect(arm.engine.paragraphs('body')[0]!.text.startsWith('Z')).toBe(true);
  }, 60_000);
}
