import { afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { applyFrameDeltaOwned, decodeFrameDelta } from '../layout/render/frameDelta';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import { createYrsSession, type YrsCellLoc, type YrsSession } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';

const path = (relative: string) => fileURLToPath(new URL(relative, import.meta.url));
const WASM = path('../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = path('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});

let startWorker: () => InProcessResidentWorker;
const sessions: YrsSession[] = [];
const clients: ResidentEngineWorkerClient[] = [];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  startWorker = await residentWorkerFactory();
}, 30_000);

afterEach(() => {
  for (const client of clients.splice(0)) client.destroy();
  for (const session of sessions.splice(0)) session.destroy();
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

function stateBytes(session: { encodeStateVector(): Uint8Array; encodeStateAsUpdate(vector?: Uint8Array): Uint8Array }): string {
  return `${Array.from(session.encodeStateVector()).join(',')}|${Array.from(session.encodeStateAsUpdate()).join(',')}`;
}

/** The host replica of a document whose live table names a missing cell, laid out for the worker. */
async function danglingHost(base: number, order: 'synced' | 'concurrent'): Promise<{ main: YrsSession; story: string }> {
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
  const main = await createYrsSession({ clientId: base + 3 });
  sessions.push(main);
  main.openDocx(docx(), false);
  main.loadState(b.encodeState());
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  const [story] = danglingCells(main);
  expect(story).toBeDefined();
  return { main, story: story! };
}

for (const [index, order] of (['synced', 'concurrent'] as const).entries()) test(`${order}: typing and Backspace in a missing cell through the resident worker are not applied, fail nothing and leave the worker usable`, async () => {
  const { main, story } = await danglingHost(12_000 + index * 10, order);
  const errors = spyOn(console, 'error');
  try {
    const worker = startWorker();
    const client = new ResidentEngineWorkerClient(worker);
    clients.push(client);
    const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
    let frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
    const caret = { story, paraId: main.paragraphSpans(story)[0]!.paraId, offset: 0 };
    expect(caret.paraId).toBe(`${story}#p0`);
    main.setSelection(caret, caret);
    const hostBefore = stateBytes(main);
    const workerBefore = stateBytes(worker.sessions[0]!);

    const typed = await client.applyInput('x', main.selection()!, frame.frameEpoch);
    expect(typed).toEqual({ applied: false });
    const deleted = await client.applyDelete('backward', main.selection()!, frame.frameEpoch);
    if (deleted.applied) expect(deleted.updates.every((update) => update.byteLength <= 2)).toBe(true);

    expect(client.hasFailed()).toBe(false);
    expect(client.isReady()).toBe(true);
    expect(stateBytes(main) === hostBefore).toBe(true);
    expect(stateBytes(worker.sessions[0]!) === workerBefore).toBe(true);
    expect(danglingCells(main)).toContain(story);

    const after = main.paragraphSpans('body').at(-1)!;
    main.setSelection({ story: 'body', paraId: after.paraId, offset: after.length });
    const elsewhere = await client.applyInput(' typed', main.selection()!, frame.frameEpoch);
    if (!elsewhere.applied) throw new Error('the worker refused input outside the missing cell');
    for (const update of elsewhere.updates) main.applyLocalUpdate(update);
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(elsewhere.frame));
    expect(main.paragraphs('body').some((entry) => entry.text === 'After typed')).toBe(true);
    expect(client.hasFailed()).toBe(false);
    expect(errors).not.toHaveBeenCalled();
  } finally {
    errors.mockRestore();
  }
}, 60_000);
