import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import type { ReactNode } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession } from '@betteroffice/docx/yrs';
import {
  DocxEditor,
  EditorToolbar,
  ToolbarCommandButton,
  useDocxCommandState,
  type DocxEditorProps,
  type DocxEditorRef,
} from '../../index';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

const FIXTURE = resolve(import.meta.dir, '../DocxEditor/hooks/__fixtures__/probe-linked-header.docx');
const quiet = { error: console.error, warn: console.warn };

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {}, ready: Promise.resolve() },
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

function documentBytes(): ArrayBuffer {
  const bytes = readFileSync(FIXTURE);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

const COMPACT: ReactNode = (
  <EditorToolbar>
    <EditorToolbar.Toolbar>
      <ToolbarCommandButton id="bold" />
      <ToolbarCommandButton id="undo" />
    </EditorToolbar.Toolbar>
  </EditorToolbar>
);

async function mount(props: Partial<DocxEditorProps> = {}) {
  const ref = createRef<DocxEditorRef>();
  const view = render(<DocxEditor ref={ref} documentBuffer={documentBytes()} {...props} />);
  for (let attempt = 0; attempt < 200; attempt += 1) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
    const save = ref.current?.commands.getState('save');
    if (save?.enabled) break;
  }
  expect(ref.current?.commands.getState('save').enabled).toBe(true);
  return { ref, view };
}

async function selectFirstWord(ref: React.RefObject<DocxEditorRef | null>) {
  const editor = ref.current!.getEditorRef()!;
  const session = editor.getYrsSession()!;
  const paragraph = session.paragraphs('body').find((candidate) => candidate.text.length >= 3)!;
  await act(async () => {
    session.setSelection(
      { story: 'body', paraId: paragraph.paraId, offset: 0 },
      { story: 'body', paraId: paragraph.paraId, offset: 3 }
    );
    editor.syncYrsInputState(false);
  });
}

async function settle() {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 20));
  });
}

describe('DocxEditor toolbar prop', () => {
  test('replaces the default chrome and drives the editor through its commands', async () => {
    const { ref, view } = await mount({ toolbar: COMPACT });
    const body = within(view.container);
    expect(body.queryByTestId('title-bar')).toBeNull();
    const bold = body.getByRole('button', { name: 'Bold' });
    const undo = body.getByRole('button', { name: 'Undo' });
    expect(undo.getAttribute('aria-disabled')).toBe('true');

    await selectFirstWord(ref);
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    fireEvent.click(bold);
    await settle();
    expect(bold.getAttribute('aria-pressed')).toBe('true');
    expect(undo.hasAttribute('aria-disabled')).toBe(false);

    fireEvent.click(undo);
    await settle();
    expect(bold.getAttribute('aria-pressed')).toBe('false');
  });

  test('a command right after scrollToParaId acts on the paragraph it moved to', async () => {
    const { ref } = await mount({ toolbar: COMPACT });
    const editor = ref.current!.getEditorRef()!;
    const session = editor.getYrsSession()!;
    const [from, to] = session.paragraphs('body').filter((paragraph) => paragraph.text.length > 0);
    const caretIn = async (paraId: string) => {
      await act(async () => {
        session.setSelection({ story: 'body', paraId, offset: 0 });
        editor.syncYrsInputState(false);
      });
    };
    await caretIn(from.paraId);
    const before = ref.current!.commands.getState('alignment').value;
    let moved = false;
    await act(async () => {
      moved = ref.current!.scrollToParaId(to.paraId);
      await ref.current!.commands.execute('alignment', { value: 'center' });
    });
    expect(moved).toBe(true);
    expect(ref.current!.commands.getState('alignment').value).toBe('center');
    await caretIn(from.paraId);
    expect(ref.current!.commands.getState('alignment').value).toBe(before);
  });

  test('keeps supplied chrome in read-only mode and applies the restriction', async () => {
    const { view } = await mount({ toolbar: COMPACT, readOnly: true });
    const bold = within(view.container).getByRole('button', { name: 'Bold' });
    expect(bold.getAttribute('aria-disabled')).toBe('true');
    const reason = document.getElementById(bold.getAttribute('aria-describedby')!);
    expect(reason?.textContent).toBe('This document is read-only.');
  });

  test('null hides the chrome, as does showToolbar={false}', async () => {
    const hidden = await mount({ toolbar: null });
    expect(within(hidden.view.container).queryByRole('toolbar')).toBeNull();
    expect(hidden.ref.current?.commands.getState('bold').enabled).toBeDefined();
    cleanup();
    const off = await mount({ toolbar: COMPACT, showToolbar: false });
    expect(within(off.view.container).queryByRole('toolbar')).toBeNull();
  });

  test('omitting the prop keeps the default chrome', async () => {
    const { view } = await mount();
    const body = within(view.container);
    expect(body.getByTestId('title-bar')).toBeDefined();
    expect(body.getByRole('toolbar', { name: 'Formatting toolbar' })).toBeDefined();
    expect(body.getByRole('button', { name: 'Undo' })).toBeDefined();
  });

  test('subscribed controls follow a remote update applied to the live session', async () => {
    function ReviewNext() {
      const state = useDocxCommandState('reviewNext');
      return (
        <output data-testid="review-next">
          {state.enabled ? 'enabled' : state.disabledReason.code}
        </output>
      );
    }
    const toolbar = (
      <EditorToolbar>
        <EditorToolbar.Toolbar>
          <ToolbarCommandButton id="bold" />
          <ReviewNext />
        </EditorToolbar.Toolbar>
      </EditorToolbar>
    );
    const { ref, view } = await mount({ toolbar });
    const body = within(view.container);
    const bold = body.getByRole('button', { name: 'Bold' });
    const reviewNext = body.getByTestId('review-next');
    await selectFirstWord(ref);
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    expect(reviewNext.textContent).toBe('no-revisions');

    const session = ref.current!.getEditorRef()!.getYrsSession()!;
    const replica = await createYrsSession({ clientId: 4242 });
    try {
      replica.applyUpdate(session.encodeStateAsUpdate());
      const selection = session.selection()!;
      replica.toggleMark(
        {
          story: 'body',
          start: { paraId: selection.anchor.paraId, offset: selection.anchor.offset },
          end: { paraId: selection.head.paraId, offset: selection.head.offset },
        },
        { type: 'bold' }
      );
      const paragraph = replica.paragraphs('body').at(-1)!;
      replica.insertText(
        { story: 'body', paraId: paragraph.paraId, offset: paragraph.text.length },
        ' remote',
        { name: 'Remote', date: '2026-01-01T00:00:00Z' }
      );
      const update = replica.encodeStateAsUpdate(session.encodeStateVector());
      await act(async () => {
        session.applyUpdate(update);
      });
      await settle();
      expect(bold.getAttribute('aria-pressed')).toBe('true');
      expect(reviewNext.textContent).toBe('enabled');
    } finally {
      replica.destroy();
    }
  });

  test('an engine refusal fails the command and legacy calls still answer false', async () => {
    const { ref } = await mount({ toolbar: COMPACT });
    await selectFirstWord(ref);
    const editor = ref.current!.getEditorRef()!;
    const session = editor.getYrsSession()!;
    session.toggleMark = () => {
      throw new Error('refused');
    };
    let failed = null as string | null;
    let legacy = true as boolean;
    await act(async () => {
      const result = await ref.current!.commands.execute('bold', null);
      failed = result.ok ? null : result.failure.code;
      legacy = editor.applyYrsFormatting('bold');
    });
    await settle();
    expect(failed).toBe('command-failed');
    expect(legacy).toBe(false);
  });
});

