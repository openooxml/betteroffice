import { afterEach, beforeAll } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import { DirtyProjectionStories, hostSaveMetadata, mergeDocxHostMetadata, saveEditorDocument } from './editorSave';
import { createYrsSession, decodeDocxHostJson, saveYrsDocx, type YrsSession } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import { yrsToDocument } from './yrsToDocument';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
export const SAVE_PATHS = ['agents engine', 'worker-only', 'main-thread (hybrid and in-thread)'] as const;
export type SavePath = typeof SAVE_PATHS[number];
type Engine = Pick<YrsSession, 'readParagraphs' | 'applyEdits' | 'version' | 'deleteRow' | 'insertColumn'>;

export const SOURCE_ROWS = ['first', 'last'].map((label) =>
  `<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>` +
  `<w:p><w:r><w:t xml:space="preserve">${label} row</w:t></w:r></w:p></w:tc></w:tr>`
);
export const SOURCE_PARAGRAPHS = ['before', 'after'].map((label) =>
  `<w:p><w:r><w:t xml:space="preserve">${label} table</w:t></w:r></w:p>`
);
export const TABLE_PREFIX = '<w:tbl><w:tblPr><w:tblW w:w="2400" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>';

export function structuralFixture(root: string): Uint8Array {
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    )],
    ['_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    )],
    ['word/document.xml', toBytes(
      `${root}<w:body>${SOURCE_PARAGRAPHS[0]}${TABLE_PREFIX}${SOURCE_ROWS.join('')}</w:tbl>${SOURCE_PARAGRAPHS[1]}</w:body></w:document>`
    )],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

export function documentXml(bytes: Uint8Array): string {
  return new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
}

export function setupStructuralSaves() {
  const owned: Array<{ destroy(): void }> = [];
  let startWorker: (clientId?: number) => InProcessResidentWorker;
  beforeAll(async () => {
    await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
    await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
    startWorker = await residentWorkerFactory();
  });
  afterEach(() => {
    for (const session of owned.splice(0)) session.destroy();
  });
  return async (path: SavePath, source: Uint8Array, apply: (engine: Engine) => void): Promise<string> => {
    if (path === 'worker-only') {
      const worker = startWorker(411);
      const client = new ResidentEngineWorkerClient(worker);
      owned.push(client);
      const { hostJson } = await client.open(source);
      const host = decodeDocxHostJson(hostJson, source).document;
      const resident = worker.sessions[0]!;
      const peer = await createYrsSession();
      owned.push(peer);
      peer.openDocx(source.slice(), false);
      peer.loadState(resident.encodeStateAsUpdate());
      apply(peer);
      resident.applyUpdate(peer.encodeStateAsUpdate());
      const saved = await client.save({ comments: [], host: hostSaveMetadata(host) });
      return documentXml(new Uint8Array(saved.bytes));
    }
    const session = await createYrsSession();
    owned.push(session);
    if (path === 'agents engine') {
      session.seedFromDocx(source);
      apply(session);
      const draft = await createYrsSession();
      owned.push(draft);
      draft.openDocx(new Uint8Array(session.materializeDocx()!.originalBuffer!), false);
      draft.loadState(session.encodeState());
      return documentXml((await saveYrsDocx(draft)).bytes);
    }
    const host = session.openDocx(source.slice(), true).document;
    apply(session);
    const dirty = new DirtyProjectionStories();
    dirty.add('body');
    const base = session.materializeDocx() as Document;
    const projected = yrsToDocument(session, mergeDocxHostMetadata(base, host), dirty.projectionOptions());
    return documentXml(new Uint8Array(await saveEditorDocument(session, projected, [])));
  };
}

export function insertStructuralParagraph(engine: Engine): void {
  const read = engine.readParagraphs({ story: 'body', view: 'accepted' });
  if (!read.ok) throw new Error(JSON.stringify(read));
  const target = read.paragraphs.find(({ text }) => text === 'after table')!;
  const result = engine.applyEdits({
    expectVersion: engine.version(),
    steps: [{ op: 'insertParagraphs', target: { story: 'body', paraId: target.paraId }, at: 'end', paragraphs: [{ text: 'inserted' }] }],
  });
  if (!result.ok) throw new Error(JSON.stringify(result));
}
