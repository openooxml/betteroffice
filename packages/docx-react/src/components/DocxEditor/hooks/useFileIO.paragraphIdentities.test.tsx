import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, yrsToDocument, type YrsSession } from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from '../PagedEditor';
import { useFileIO } from './useFileIO';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
const sessions: YrsSession[] = [];
const FIXTURE = resolve(
  import.meta.dir,
  '../../../../../../crates/docx-edit/tests/fixtures/paragraph-identities'
);

function fixture(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const add = (dir: string, prefix: string) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const path = join(dir, entry.name);
      if (entry.isDirectory()) add(path, `${prefix}${entry.name}/`);
      else parts.set(`${prefix}${entry.name}`, toBytes(readFileSync(path, 'utf8')));
    }
  };
  add(FIXTURE, '');
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function session(clientId: number): Promise<YrsSession> {
  const created = await createYrsSession({ clientId });
  sessions.push(created);
  return created;
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
afterEach(() => {
  cleanup();
  for (const created of sessions.splice(0)) created.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** Saves through the editor's Save, projecting the session as the editor does. */
async function editorSave(live: YrsSession): Promise<Uint8Array> {
  const base = live.materializeDocx()!;
  const editor = {
    getYrsSession: () => live,
    flushPendingInput: async () => {},
    getDocument: () => yrsToDocument(live, base),
  } as unknown as PagedEditorRef;
  const saved: ArrayBuffer[] = [];
  const errors: Error[] = [];
  const { result } = renderHook(() =>
    useFileIO({
      pagedEditorRef: { current: editor },
      displayList: null,
      resolveImage: () => null,
      comments: base.package.document.comments ?? [],
      documentName: 'saved',
      onSave: (buffer) => saved.push(buffer),
      downloadOnSave: false,
      onError: (error) => errors.push(error),
      onOpen: undefined,
      onPrint: undefined,
      onDocumentNameChange: undefined,
      loadBuffer: async () => {},
      focusActiveEditor: () => {},
    })
  );
  await act(async () => {
    await result.current.handleSave();
  });
  expect(errors).toEqual([]);
  return new Uint8Array(saved[0]!);
}

function ooxmlParaId(live: YrsSession, key: string): string | null {
  return (
    live.paragraphIdentities().paragraphs.find((identity) => identity.session?.paraId === key)
      ?.ooxmlParaId ?? null
  );
}

test('the editor Save writes the planned paragraph IDs and records them as saved', async () => {
  const live = await session(7);
  live.openDocx(fixture(), true);
  const offline = await session(8);
  offline.loadState(live.encodeState());
  const receipt = live.persistParagraphIds();
  if (receipt.status !== 'applied') throw new Error('refused');
  const separator = receipt.assignments.find(
    ({ paragraph }) =>
      paragraph.kind === 'source' &&
      paragraph.partUri === '/word/footnotes.xml' &&
      paragraph.paragraphOrdinal === 0
  )!;
  const split = live.splitParagraph({ story: 'body', paraId: '0000abcd', offset: 2 });
  const splitId = ooxmlParaId(live, split.secondParaId)!;

  const bytes = await editorSave(live);
  const written = live.writtenParagraphIds(bytes);
  expect(written['/word/footnotes.xml']).toContain(separator.ooxmlParaId.toUpperCase());
  expect(written['/word/document.xml']).toContain(splitId.toUpperCase());

  live.deleteRange({
    story: 'body',
    start: { paraId: '0000abcd', offset: 2 },
    end: { paraId: split.secondParaId, offset: 3 },
  });
  offline.applyRawOps('body', [
    {
      op: 'insertEmbed',
      index: 0,
      kind: 'pilcrow',
      payload: { paraId: 'offline', ooxmlParaId: splitId },
    },
  ]);
  live.applyUpdate(offline.encodeStateAsUpdate(live.encodeStateVector()));
  expect(ooxmlParaId(live, 'offline')).not.toBe(splitId);
});
