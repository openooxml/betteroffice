import { afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import * as editorSave from './editorSave';
import * as projection from './yrsToDocument';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import type { ResidentSaveRecord } from './residentSave';

const source = new Uint8Array(readFileSync(resolve(import.meta.dir, '__fixtures__/comment-ranges/structure.docx')));
const owned: Array<YrsSession | ResidentEngineSession> = [];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
});

afterEach(() => {
  for (const session of owned.splice(0)) session.destroy();
});

async function open() {
  const session = await createResidentEngineSession();
  owned.push(session);
  const hostJson = session.openDocx(source);
  const host = decodeDocxHostJson(hostJson, source).document;
  const record: ResidentSaveRecord = { full: false };
  const header = session.geometryReader.storyIds().find((story) =>
    story.startsWith('hf:') && host.package.headers?.has(story.slice(3)) &&
    session.geometryReader.paragraphs(story).some((paragraph) => paragraph.text.length > 0)
  )!;
  expect(header).toBeDefined();
  return { session, hostJson, host, record, header };
}

test('resident saves use peer marks even when another story changed or was already projected by the peer', async () => {
  const { session, hostJson, host, record, header } = await open();
  const peer = await createYrsSession();
  owned.push(peer);
  peer.openDocx(source, false);
  peer.loadState(session.encodeState());
  const paragraph = peer.paragraphs('body')[0]!;
  peer.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Remote ');
  session.applyUpdate(peer.encodeStateAsUpdate(session.encodeStateVector()));
  session.markProjectionStories(['body']);
  const project = spyOn(projection, 'yrsToDocument');
  try {
    await session.save(source, hostJson, host, [], record, [header]);
    expect(project.mock.calls[0]![2]?.storyIds).toEqual(new Set([header]));
    expect(record).not.toHaveProperty('revision');
    session.markProjectionStories(['body']);
    project.mockClear();
    await session.save(source, hostJson, host, [], record, []);
    expect(project.mock.calls[0]![2]).toBeUndefined();
  } finally {
    project.mockRestore();
  }
});

test('resident raw comment ops survive a failed save and clear after a successful save', async () => {
  const { session, hostJson, host, record, header } = await open();
  session.applyRawOps(header, [{
    op: 'setComment', id: '999', ranges: [[0, 1]],
    author: 'Worker', date: '2026-10-01T00:00:00Z', body: 'Header comment',
  }]);
  const project = spyOn(projection, 'yrsToDocument');
  const write = spyOn(editorSave, 'saveEditorDocument').mockRejectedValue(new Error('Save failed'));
  try {
    await expect(session.save(source, hostJson, host, [], record)).rejects.toThrow('Save failed');
    expect(project.mock.calls[0]![2]?.storyIds).toEqual(new Set([header]));
    write.mockRestore();
    project.mockClear();
    await session.save(source, hostJson, host, [], record);
    expect(project.mock.calls[0]![2]?.storyIds).toEqual(new Set([header]));
    project.mockClear();
    await session.save(source, hostJson, host, [], record);
    expect(project.mock.calls[0]![2]).toBeUndefined();
    session.applyRawOps(header, [{ op: 'removeComment', id: '999' }]);
    project.mockClear();
    await session.save(source, hostJson, host, [], record);
    expect(project.mock.calls[0]![2]?.storyIds).toEqual(new Set([header]));
  } finally {
    write.mockRestore();
    project.mockRestore();
  }
});
