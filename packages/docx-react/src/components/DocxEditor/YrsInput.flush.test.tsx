import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { YrsInput, type YrsInputProps, type YrsInputRef } from './YrsInput';
import { performYrsHistoryAction } from './yrsCommands';
import { DocxCommandAdmissionError } from '../../commands/createDocxCommandStore';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function seededSession(): Promise<YrsSession> {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Seed');
  session.setSelection({ story: 'body', paraId, offset: 4 });
  session.addUndoBoundary();
  return session;
}

function inputFor(
  session: YrsSession,
  input: React.Ref<YrsInputRef>,
  applyResidentInput?: YrsInputProps['applyResidentInput']
) {
  const map = () =>
    createYrsInputPositionMap(
      'body',
      session.paragraphs('body').map((p) => ({
        paraId: p.paraId,
        length: p.text.length,
      }))
    );
  return (
    <YrsInput
      ref={input}
      enabled
      readOnly={false}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position) => displayPositionToYrsLoc(map(), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(), loc)}
      onStateChange={() => {}}
      onDirectInput={() => {}}
      applyResidentInput={applyResidentInput}
    />
  );
}

async function mount(applyResidentInput?: YrsInputProps['applyResidentInput']) {
  const session = await seededSession();
  const input = createRef<YrsInputRef>();
  const view = render(inputFor(session, input, applyResidentInput));
  return { session, input, view };
}

function text(session: YrsSession): string {
  return session.paragraphs('body')[0].text;
}

function admissionCode(error: unknown): string | null {
  return error instanceof DocxCommandAdmissionError ? error.code : null;
}

test('flush waits for resident input and publishes the latest selection', async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const { session, input } = await mount(async () => {
    await blocked;
    return null;
  });
  act(() => input.current!.insertText(' accepted'));
  let done = false;
  const flush = input.current!.flushPendingInput().then(() => {
    done = true;
  });
  await Promise.resolve();
  expect(done).toBe(false);
  expect(session.paragraphs('body')[0].text).toBe('Seed');
  await act(async () => {
    release();
    await flush;
  });
  expect(session.paragraphs('body')[0].text).toBe('Seed accepted');
  expect(session.selection()?.head.offset).toBe(13);
});

test('flush seals a queued text batch before subsequent input', async () => {
  const calls: string[] = [];
  const { session, input } = await mount(async (text) => {
    calls.push(text);
    return null;
  });
  act(() => input.current!.insertText('A'));
  const first = input.current!.flushPendingInput();
  act(() => input.current!.insertText('B'));
  await act(async () => {
    await first;
    await input.current!.flushPendingInput();
  });
  expect(calls).toEqual(['A', 'B']);
  expect(session.paragraphs('body')[0].text).toBe('SeedAB');
});

test('flush includes a completed IME composition exactly once', async () => {
  const { session, input, view } = await mount();
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  fireEvent.compositionEnd(textarea, { data: '日本' });
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  fireEvent.input(textarea);
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  expect(session.paragraphs('body')[0].text).toBe('Seed日本');
  expect(session.selection()?.head.offset).toBe(6);
});

test('flush waits for an active composition and rejects if the input is removed', async () => {
  const { input, view } = await mount();
  fireEvent.compositionStart(view.getByTestId('yrs-input'));
  let done = false;
  const flush = input.current!.flushPendingInput().finally(() => {
    done = true;
  });
  await Promise.resolve();
  expect(done).toBe(false);
  view.unmount();
  await expect(flush).rejects.toThrow('unavailable');
});

test('flush rejects failed resident input instead of claiming it was committed', async () => {
  const failure = new Error('resident failure');
  const { session, input } = await mount(async () => {
    throw failure;
  });
  act(() => input.current!.insertText('lost'));
  await expect(input.current!.flushPendingInput()).rejects.toBe(failure);
  expect(session.paragraphs('body')[0].text).toBe('Seed');
});

test('undo waits behind pending typing and later typing waits behind undo', async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const { session, input } = await mount(async () => {
    await blocked;
    return null;
  });
  act(() => input.current!.insertText(' A'));
  const undo = input.current!.runAfterPendingInput(
    () => performYrsHistoryAction(session, false).changed
  );
  act(() => input.current!.insertText(' B'));
  expect(text(session)).toBe('Seed');
  await act(async () => {
    release();
    expect(await undo).toBe(true);
    await input.current!.flushPendingInput();
  });
  expect(text(session)).toBe('Seed B');
});

test('a command seals the text batch so input on either side stays separate', async () => {
  const calls: string[] = [];
  const { input } = await mount(async (typed) => {
    calls.push(typed);
    return null;
  });
  act(() => input.current!.insertText('x'));
  const format = input.current!.runAfterPendingInput(() => {
    calls.push('format');
  });
  act(() => input.current!.insertText('y'));
  await act(async () => {
    await format;
    await input.current!.flushPendingInput();
  });
  expect(calls).toEqual(['x', 'format', 'y']);
});

test('repeated undo requests each run', async () => {
  const { session, input } = await mount();
  act(() => input.current!.insertText(' one'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  session.addUndoBoundary();
  act(() => input.current!.insertText(' two'));
  const undo = () =>
    input.current!.runAfterPendingInput(() => performYrsHistoryAction(session, false).changed);
  let results: boolean[] = [];
  await act(async () => {
    results = await Promise.all([undo(), undo()]);
  });
  expect(results).toEqual([true, true]);
  expect(text(session)).toBe('Seed');
});

test('a command issued during composition runs after the committed text', async () => {
  const { session, input, view } = await mount();
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.compositionStart(textarea);
  let seen = null as string | null;
  const command = input.current!.runAfterPendingInput(() => {
    seen = text(session);
  });
  await Promise.resolve();
  expect(seen).toBeNull();
  textarea.value = '日本';
  await act(async () => {
    fireEvent.compositionEnd(textarea, { data: '日本' });
    await command;
  });
  expect(seen).toBe('Seed日本');
});

test('a command after lost input is refused as input-failed', async () => {
  const originalError = console.error;
  console.error = () => {};
  try {
    const { input } = await mount(async () => {
      throw new Error('resident failure');
    });
    act(() => input.current!.insertText('lost'));
    let error: unknown;
    await act(async () => {
      error = await input.current!.runAfterPendingInput(() => 'ran').catch((cause) => cause);
    });
    expect(admissionCode(error)).toBe('input-failed');
  } finally {
    console.error = originalError;
  }
});

test('a command admitted before the document was replaced is refused', async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const first = await seededSession();
  const second = await seededSession();
  const input = createRef<YrsInputRef>();
  const resident = async () => {
    await blocked;
    return null;
  };
  const view = render(inputFor(first, input, resident));
  act(() => input.current!.insertText(' pending'));
  const command = input.current!.runAfterPendingInput(() => 'ran').catch((cause) => cause);
  act(() => {
    view.rerender(inputFor(second, input, resident));
  });
  let error: unknown;
  await act(async () => {
    release();
    error = await command;
  });
  expect(admissionCode(error)).toBe('document-replaced');
});

test('a command waiting on composition is refused when the input unmounts', async () => {
  const { input, view } = await mount();
  fireEvent.compositionStart(view.getByTestId('yrs-input'));
  const command = input.current!.runAfterPendingInput(() => 'ran').catch((cause) => cause);
  view.unmount();
  expect(admissionCode(await command)).toBe('editor-unavailable');
});

test('a stored superscript applies to the next typed text and clears subscript', async () => {
  const { session, input } = await mount();
  act(() =>
    input.current!.applyStoredFormatting({ type: 'toggle', mark: 'superscript', active: false })
  );
  expect(input.current!.storedFormatting()?.delta.other).toEqual({
    superscript: true,
    subscript: null,
  });
  act(() => input.current!.insertText('x'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  const { paraId } = session.paragraphs('body')[0];
  const typed = session.selectionContext({
    story: 'body',
    start: { paraId, offset: 4 },
    end: { paraId, offset: 5 },
  });
  expect(typed.superscript).toBe(true);
  expect(typed.subscript).toBe(false);
});
