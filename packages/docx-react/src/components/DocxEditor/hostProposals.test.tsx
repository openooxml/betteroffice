import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { DocxProposalRequest } from '@betteroffice/docx/yrs';
import {
  DocxEditor,
  defineDocxPlugin,
  type DocxEditorProps,
  type DocxEditorRef,
  type DocxPluginContext,
} from '../../index';

const { act, cleanup, render } = await import('@testing-library/react');

const FIXTURE = resolve(import.meta.dir, 'hooks/__fixtures__/probe-linked-header.docx');
const quiet = { error: console.error, warn: console.warn };

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: {
        addEventListener: () => {},
        removeEventListener: () => {},
        ready: Promise.resolve(),
      },
      configurable: true,
    });
  }
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  );
  console.error = () => {};
  console.warn = () => {};
});
afterEach(cleanup);
afterAll(async () => {
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function until(done: () => boolean) {
  for (let attempt = 0; attempt < 300 && !done(); attempt += 1) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
  }
  expect(done()).toBe(true);
}

test('allowHostProposals admits only the proposal methods in a read-only editor', async () => {
  const contexts: DocxPluginContext<null>[] = [];
  const plugin = defineDocxPlugin<null>({
    id: 'acme.review',
    createState: () => null,
    initialize(context) {
      contexts.push(context);
    },
  });
  const sidebar: boolean[] = [];
  const ref = createRef<DocxEditorRef>();
  const bytes = readFileSync(FIXTURE);
  const buffer = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength
  ) as ArrayBuffer;
  const element = (props: Partial<DocxEditorProps>) => (
    <DocxEditor
      ref={ref}
      documentBuffer={buffer}
      readOnly
      plugins={[plugin]}
      pluginGrants={{ 'acme.review': { document: 'write', editBatches: true } }}
      onCommentsSidebarOpenChange={(open) => sidebar.push(open)}
      {...props}
    />
  );
  const view = render(element({}));
  await until(() => ref.current?.getEditorRef()?.getYrsSession() != null && contexts.length > 0);

  const session = ref.current!.getEditorRef()!.getYrsSession()!;
  const paragraph = session.paragraphs('body').find((candidate) => candidate.text.length >= 3)!;
  const word = paragraph.text.slice(0, 3);
  const proposal = (): DocxProposalRequest => ({
    expectVersion: session.version(),
    proposals: [
      {
        id: 'p1',
        paragraph: {
          kind: 'session',
          sessionId: session.paragraphIdentities().sessionId,
          story: 'body',
          paraId: paragraph.paraId,
        },
        suggest: { author: 'Atira', date: '2026-09-29T00:00:00Z' },
        op: 'replaceText',
        search: word,
        replaceWith: 'XYZ',
      },
    ],
  });
  expect(await ref.current!.proposeChanges(proposal())).toMatchObject({
    ok: false,
    failure: { code: 'read-only' },
  });

  view.rerender(element({ allowHostProposals: true }));
  const proposed = await act(() => ref.current!.proposeChanges(proposal()));
  expect(proposed).toMatchObject({
    ok: true,
    snapshot: { proposals: [{ id: 'p1', changed: true }] },
  });
  expect(
    session.paragraphs('body').find((candidate) => candidate.paraId === paragraph.paraId)!.text
  ).toContain('XYZ');

  const append = {
    expectVersion: session.version(),
    steps: [
      {
        op: 'insertText' as const,
        target: { kind: 'paragraph' as const, story: 'body', paraId: paragraph.paraId },
        at: 'end' as const,
        text: '!',
      },
    ],
  };
  expect(await ref.current!.applyEdits(append)).toMatchObject({
    ok: false,
    failure: { code: 'read-only' },
  });
  expect(await contexts.at(-1)!.edits!.applyEdits(append)).toMatchObject({
    ok: false,
    failure: { code: 'read-only' },
  });
  expect(ref.current!.commands.getState('bold').enabled).toBe(false);

  const decided = await act(() =>
    ref.current!.setProposalStates({
      expectVersion: session.version(),
      expectPreviewVersion: 0,
      changes: [{ id: 'p1', state: 'accepted' }],
    })
  );
  expect(decided).toMatchObject({ ok: true, snapshot: { previewVersion: 1 } });
  expect(await ref.current!.getProposals()).toMatchObject({
    version: session.version(),
    previewVersion: 1,
    proposals: [{ id: 'p1', state: 'accepted' }],
  });
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 50));
  });
  expect(sidebar).not.toContain(true);
});

test("proposals' tracked changes do not open the comments sidebar; the user's still do", async () => {
  const sidebar: boolean[] = [];
  const ref = createRef<DocxEditorRef>();
  const bytes = readFileSync(FIXTURE);
  const buffer = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength
  ) as ArrayBuffer;
  render(
    <DocxEditor
      ref={ref}
      documentBuffer={buffer}
      onCommentsSidebarOpenChange={(open) => sidebar.push(open)}
    />
  );
  await until(() => ref.current?.commands.getState('save').enabled === true);
  const session = ref.current!.getEditorRef()!.getYrsSession()!;
  const paragraphs = session.paragraphs('body').filter((candidate) => candidate.text.length >= 3);
  const [first, second] = paragraphs;
  const suggest = { author: 'Atira', date: '2026-09-29T00:00:00Z' };
  const sessionId = session.paragraphIdentities().sessionId;
  const proposed = await act(() =>
    ref.current!.proposeChanges({
      expectVersion: session.version(),
      proposals: [first!, second!].map((paragraph, index) => ({
        id: `p${index}`,
        paragraph: { kind: 'session', sessionId, story: 'body', paraId: paragraph.paraId },
        suggest,
        op: 'insertText',
        at: 'end',
        text: ` proposal ${index}`,
      })),
    })
  );
  if (!proposed.ok) throw new Error(proposed.failure.message);
  const [kept, accepted] = proposed.snapshot.proposals;
  await act(async () => {
    await ref.current!.commands.execute('reviewAccept', { revisionId: accepted!.revisionIds[0]! });
  });
  expect(sidebar).not.toContain(true);

  const typed = await act(() =>
    ref.current!.applyEdits({
      expectVersion: session.version(),
      steps: [
        {
          op: 'insertText',
          target: { kind: 'paragraph', story: 'body', paraId: first!.paraId },
          at: 'start',
          text: 'User ',
          suggest: { author: 'User', date: '2026-09-29T00:00:00Z' },
        },
      ],
    })
  );
  expect(typed).toMatchObject({ ok: true, applied: true });
  await act(async () => {
    await ref.current!.commands.execute('reviewAccept', { revisionId: kept!.revisionIds[0]! });
  });
  expect(sidebar).toContain(true);
});
