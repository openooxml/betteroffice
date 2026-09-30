import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
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

function fixture(text: string, tracked = false): ArrayBuffer {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>'
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>'
    )
  );
  const insertion = tracked
    ? '<w:ins w:id="1" w:author="Reviewer" w:date="2026-09-29T00:00:00Z"><w:r><w:t>Reviewed</w:t></w:r></w:ins>'
    : '';
  parts.set(
    'word/document.xml',
    toBytes(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
        `<w:body><w:p w14:paraId="00000001"><w:r><w:t>${text}</w:t></w:r>${insertion}</w:p><w:sectPr/></w:body></w:document>`
    )
  );
  return rezipPartsToArrayBuffer(parts);
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
  const withdrawn = await act(() =>
    ref.current!.withdrawProposals({ expectVersion: session.version(), ids: ['p1'] })
  );
  expect(withdrawn).toMatchObject({ ok: true, snapshot: { previewVersion: 2, proposals: [] } });
  expect(session.listRevisions().filter((revision) => revision.author === 'Atira')).toEqual([]);
  expect(
    session.paragraphs('body').find((candidate) => candidate.paraId === paragraph.paraId)!.text
  ).toStartWith('XYZ');
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

test('host proposals stay hidden after swapping documentBuffer on the same editor', async () => {
  const sidebar: boolean[] = [];
  const ref = createRef<DocxEditorRef>();
  const element = (documentBuffer: ArrayBuffer) => (
    <DocxEditor
      ref={ref}
      documentBuffer={documentBuffer}
      readOnly
      allowHostProposals
      onCommentsSidebarOpenChange={(open) => sidebar.push(open)}
    />
  );
  const view = render(element(fixture('Document A')));
  await until(() => ref.current?.commands.getState('save').enabled === true);
  const firstSession = ref.current!.getEditorRef()!.getYrsSession()!;
  const propose = async (author: string) => {
    const session = ref.current!.getEditorRef()!.getYrsSession()!;
    const paragraph = session.paragraphs('body')[0]!;
    const result = await act(() =>
      ref.current!.proposeChanges({
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
            suggest: { author, date: '2026-09-29T00:00:00Z' },
            op: 'insertText',
            at: 'end',
            text: ' proposal',
          },
        ],
      })
    );
    expect(result).toMatchObject({ ok: true, snapshot: { proposals: [{ id: 'p1', changed: true }] } });
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 250));
    });
  };

  await propose('Host A');
  expect(firstSession.listRevisions().some((revision) => revision.author === 'Host A')).toBe(true);
  expect(sidebar).not.toContain(true);
  expect(ref.current!.commands.getState('commentsSidebar').active).toBe(false);
  expect(view.container.querySelector('.docx-tracked-change-card')).toBeNull();
  await act(async () => {
    await ref.current!.commands.execute('commentsSidebar', null);
  });
  await until(() => view.container.querySelector('.docx-tracked-change-card') != null);
  expect(view.container.querySelector('.docx-tracked-change-card')!.textContent).toContain('Host A');
  await act(async () => {
    await ref.current!.commands.execute('commentsSidebar', null);
  });
  sidebar.length = 0;

  view.rerender(element(fixture('Document B')));
  await until(
    () =>
      ref.current?.commands.getState('save').enabled === true &&
      ref.current.getEditorRef()?.getYrsSession() != null &&
      ref.current.getEditorRef()?.getYrsSession() !== firstSession
  );
  expect(ref.current!.getEditorRef()!.getYrsSession()!.listRevisions()).toEqual([]);
  expect(ref.current!.commands.getState('commentsSidebar').active).toBe(false);
  expect(view.container.querySelector('.docx-tracked-change-card')).toBeNull();

  await propose('Host B');
  expect(sidebar).not.toContain(true);
  expect(ref.current!.commands.getState('commentsSidebar').active).toBe(false);
  expect(view.container.querySelector('.docx-tracked-change-card')).toBeNull();

  await act(async () => {
    await ref.current!.commands.execute('commentsSidebar', null);
  });
  await until(() => view.container.querySelector('.docx-tracked-change-card') != null);
  expect(ref.current!.commands.getState('commentsSidebar').active).toBe(true);
  const cards = view.container.querySelectorAll('.docx-tracked-change-card');
  expect(cards.length).toBe(1);
  expect(cards[0]!.textContent).toContain('Host B');
  expect(cards[0]!.textContent).not.toContain('Host A');
}, 15_000);

test('swapping to a document with its own revisions auto-opens the sidebar once', async () => {
  const sidebar: boolean[] = [];
  const ref = createRef<DocxEditorRef>();
  const element = (documentBuffer: ArrayBuffer) => (
    <DocxEditor
      ref={ref}
      documentBuffer={documentBuffer}
      onCommentsSidebarOpenChange={(open) => sidebar.push(open)}
    />
  );
  const view = render(element(fixture('Document A', true)));
  await until(
    () =>
      ref.current?.commands.getState('save').enabled === true &&
      view.container.querySelector('.docx-tracked-change-card') != null
  );
  const firstSession = ref.current!.getEditorRef()!.getYrsSession()!;
  expect(sidebar.filter(Boolean)).toHaveLength(1);
  await act(async () => {
    await ref.current!.commands.execute('commentsSidebar', null);
  });

  view.rerender(element(fixture('Document B', true)));
  await until(
    () =>
      ref.current?.commands.getState('save').enabled === true &&
      ref.current.getEditorRef()?.getYrsSession() != null &&
      ref.current.getEditorRef()?.getYrsSession() !== firstSession &&
      ref.current.commands.getState('commentsSidebar').active === true &&
      view.container.querySelector('.docx-tracked-change-card') != null
  );
  expect(sidebar.filter(Boolean)).toHaveLength(2);
  expect(view.container.querySelector('.docx-tracked-change-card')!.textContent).toContain('Reviewer');
  await act(async () => {
    await ref.current!.commands.execute('commentsSidebar', null);
  });

  const session = ref.current!.getEditorRef()!.getYrsSession()!;
  const paragraph = session.paragraphs('body')[0]!;
  const typed = await act(() =>
    ref.current!.applyEdits({
      expectVersion: session.version(),
      steps: [
        {
          op: 'insertText',
          target: { kind: 'paragraph', story: 'body', paraId: paragraph.paraId },
          at: 'end',
          text: ' another revision',
          suggest: { author: 'User', date: '2026-09-29T00:00:00Z' },
        },
      ],
    })
  );
  expect(typed).toMatchObject({ ok: true, applied: true });
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 50));
  });
  expect(sidebar.filter(Boolean)).toHaveLength(2);
  expect(ref.current!.commands.getState('commentsSidebar').active).toBe(false);
  expect(view.container.querySelector('.docx-tracked-change-card')).toBeNull();
}, 15_000);

test('a user revision after a swap can auto-open the sidebar again', async () => {
  const sidebar: boolean[] = [];
  const ref = createRef<DocxEditorRef>();
  const element = (documentBuffer: ArrayBuffer) => (
    <DocxEditor
      ref={ref}
      documentBuffer={documentBuffer}
      onCommentsSidebarOpenChange={(open) => sidebar.push(open)}
    />
  );
  const view = render(element(fixture('Document A', true)));
  await until(
    () =>
      ref.current?.commands.getState('save').enabled === true &&
      view.container.querySelector('.docx-tracked-change-card') != null
  );
  const firstSession = ref.current!.getEditorRef()!.getYrsSession()!;
  view.rerender(element(fixture('Document B')));
  await until(
    () =>
      ref.current?.commands.getState('save').enabled === true &&
      ref.current.getEditorRef()?.getYrsSession() != null &&
      ref.current.getEditorRef()?.getYrsSession() !== firstSession
  );
  expect(ref.current!.commands.getState('commentsSidebar').active).toBe(false);
  expect(view.container.querySelector('.docx-tracked-change-card')).toBeNull();
  expect(sidebar.filter(Boolean)).toHaveLength(1);

  const session = ref.current!.getEditorRef()!.getYrsSession()!;
  const paragraph = session.paragraphs('body')[0]!;
  const typed = await act(() =>
    ref.current!.applyEdits({
      expectVersion: session.version(),
      steps: [
        {
          op: 'insertText',
          target: { kind: 'paragraph', story: 'body', paraId: paragraph.paraId },
          at: 'end',
          text: ' user revision',
          suggest: { author: 'User', date: '2026-09-29T00:00:00Z' },
        },
      ],
    })
  );
  expect(typed).toMatchObject({ ok: true, applied: true });
  await until(() => view.container.querySelector('.docx-tracked-change-card') != null);
  expect(sidebar.filter(Boolean)).toHaveLength(2);
  expect(ref.current!.commands.getState('commentsSidebar').active).toBe(true);
}, 15_000);
