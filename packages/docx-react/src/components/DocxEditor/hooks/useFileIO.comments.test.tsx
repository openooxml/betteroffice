import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { unzipContainer } from '@betteroffice/docx/docx/wasm';
import type { Comment } from '@betteroffice/docx/types/content';
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

/** The identity fixture with its "Lower" paragraph holding an internal hyperlink. */
function fixture(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const add = (dir: string, prefix: string) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const path = join(dir, entry.name);
      if (entry.isDirectory()) add(path, `${prefix}${entry.name}/`);
      else {
        const xml = readFileSync(path, 'utf8').replace(
          '<w:r><w:t>Lower</w:t></w:r>',
          '<w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink w:anchor="target"><w:r><w:t>the linked text</w:t></w:r></w:hyperlink>'
        );
        parts.set(`${prefix}${entry.name}`, toBytes(xml));
      }
    }
  };
  add(FIXTURE, '');
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function open(bytes: Uint8Array, clientId: number): Promise<YrsSession> {
  const created = await createYrsSession({ clientId });
  sessions.push(created);
  created.openDocx(bytes, true);
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

/** Saves through the editor's Save, projecting the session with its host's comments as the editor does. */
async function editorSave(live: YrsSession, comments: Comment[]): Promise<Uint8Array> {
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
      resolveImage: () => null,
      comments,
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

function markers(bytes: Uint8Array, id: number): string[] {
  const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
  return [
    ...xml.matchAll(new RegExp(`<w:comment(RangeStart|RangeEnd|Reference) w:id="${id}"/>`, 'g')),
  ].map((match) => match[1]!);
}

function anchored(session: YrsSession, commentId: string): string {
  const [anchor] = session.resolveComment(commentId);
  let offset = 0;
  let text = '';
  for (const segment of session.storySegments(anchor!.story)) {
    if (segment.kind === 'text') {
      text += segment.text.slice(
        Math.max(0, anchor!.start - offset),
        Math.max(0, anchor!.end - offset)
      );
      offset += segment.text.length;
    } else offset += 1;
  }
  return text;
}

test('the editor Save writes the range of a comment added inside a hyperlink', async () => {
  const live = await open(fixture(), 7);
  const comments = live.materializeDocx()!.package.document.comments!;
  const start = live.locateParagraph('body', '0000abcd').start;
  const added: Comment = {
    id: 2,
    author: 'Ada',
    date: '2026-09-28T00:00:00Z',
    content: [{ type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Note' }] }] }],
  };
  live.applyRawOps('body', [
    { op: 'setComment', id: '2', ranges: [[start, start + 7]], author: 'Ada', date: added.date },
  ]);
  const bytes = await editorSave(live, [...comments, added]);
  expect(markers(bytes, 2)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
  expect(anchored(await open(bytes, 8), '2')).toBe('See the');
});

test('the editor Save writes a reanchored comment once, where it now is', async () => {
  const live = await open(fixture(), 9);
  const comments = live.materializeDocx()!.package.document.comments!;
  live.setCommentRanges('1', [
    {
      story: 'body',
      start: { paraId: 'body:p1', offset: 0 },
      end: { paraId: 'body:p1', offset: 4 },
    },
  ]);
  const bytes = await editorSave(live, comments);
  expect(markers(bytes, 1)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
  expect(anchored(await open(bytes, 10), '1')).toBe('Miss');
});
