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
  applyResidentInput?: YrsInputProps['applyResidentInput'],
  applyResidentDelete?: YrsInputProps['applyResidentDelete'],
  props: Pick<YrsInputProps, 'isSuggesting' | 'author'> = {}
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
      applyResidentDelete={applyResidentDelete}
      {...props}
    />
  );
}

async function mount(
  applyResidentInput?: YrsInputProps['applyResidentInput'],
  applyResidentDelete?: YrsInputProps['applyResidentDelete']
) {
  const session = await seededSession();
  const input = createRef<YrsInputRef>();
  const view = render(inputFor(session, input, applyResidentInput, applyResidentDelete));
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

test('suggesting type-over places the caret and stored formatting on the inserted text', async () => {
  const session = await seededSession();
  const paraId = session.paragraphs('body')[0]!.paraId;
  session.setSelection(
    { story: 'body', paraId, offset: 1 },
    { story: 'body', paraId, offset: 3 }
  );
  const input = createRef<YrsInputRef>();
  render(inputFor(session, input, undefined, undefined, { isSuggesting: true, author: 'Ada' }));
  act(() => {
    input.current!.applyStoredFormatting({ type: 'set', delta: { bold: true } });
    input.current!.insertText('X');
  });
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  expect(text(session)).toBe('SeeXd');
  expect(session.selection()).toEqual({
    anchor: { story: 'body', paraId, offset: 4 },
    head: { story: 'body', paraId, offset: 4 },
  });
  const segments = session.storySegments('body');
  const inserted = segments.find((segment) => segment.kind === 'text' && segment.text === 'X');
  expect(inserted?.attributes.ins).toMatchObject({ author: 'Ada' });
  expect(inserted?.attributes.bold).toBe(true);
  const deleted = segments.find((segment) => segment.kind === 'text' && segment.text === 'ee');
  expect(deleted?.attributes.del).toMatchObject({ author: 'Ada' });
  expect(deleted?.attributes.bold).not.toBe(true);
  act(() => input.current!.insertText('Y'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  expect(text(session)).toBe('SeeXYd');
  expect(session.selection()?.head).toEqual({ story: 'body', paraId, offset: 5 });
});

test('suggesting type-over across paragraphs keeps the stored formatting at the caret', async () => {
  const session = await seededSession();
  const first = session.paragraphs('body')[0]!.paraId;
  const { secondParaId: second } = session.splitParagraph({
    story: 'body',
    paraId: first,
    offset: 2,
  });
  session.setSelection(
    { story: 'body', paraId: second, offset: 1 },
    { story: 'body', paraId: first, offset: 1 }
  );
  const input = createRef<YrsInputRef>();
  render(inputFor(session, input, undefined, undefined, { isSuggesting: true, author: 'Ada' }));
  act(() => {
    input.current!.applyStoredFormatting({ type: 'set', delta: { bold: true } });
    input.current!.insertText('X');
  });
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  act(() => input.current!.insertText('Y'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  const typed = session
    .storySegments('body')
    .filter((segment) => segment.kind === 'text' && /[XY]/.test(segment.text));
  expect(typed.map((segment) => (segment.kind === 'text' ? segment.text : '')).join('')).toBe(
    'XY'
  );
  expect(typed.every((segment) => segment.attributes.bold === true)).toBe(true);
});

test('multiline suggesting type-over keeps the stored formatting on every piece', async () => {
  for (const typed of ['\nX', 'X\nY']) {
    const session = await seededSession();
    const first = session.paragraphs('body')[0]!.paraId;
    const { secondParaId: second } = session.splitParagraph({
      story: 'body',
      paraId: first,
      offset: 2,
    });
    session.setSelection(
      { story: 'body', paraId: first, offset: 1 },
      { story: 'body', paraId: second, offset: 1 }
    );
    const input = createRef<YrsInputRef>();
    render(inputFor(session, input, undefined, undefined, { isSuggesting: true, author: 'Ada' }));
    act(() => {
      input.current!.applyStoredFormatting({ type: 'set', delta: { bold: true } });
      input.current!.insertText(typed);
    });
    await act(async () => {
      await input.current!.flushPendingInput();
    });
    const inserted = session
      .storySegments('body')
      .filter((segment) => segment.kind === 'text' && /[XY]/.test(segment.text));
    expect(inserted.map((segment) => (segment.kind === 'text' ? segment.text : '')).join('')).toBe(
      typed.replace('\n', '')
    );
    expect(inserted.every((segment) => segment.attributes.bold === true)).toBe(true);
    cleanup();
  }
});

test('type-over keeps paragraph stored formatting apart from the carried override', async () => {
  const setup = async (isSuggesting: boolean) => {
    const session = await seededSession();
    const first = session.paragraphs('body')[0]!.paraId;
    const { secondParaId: second } = session.splitParagraph({
      story: 'body',
      paraId: first,
      offset: 2,
    });
    const input = createRef<YrsInputRef>();
    render(inputFor(session, input, undefined, undefined, { isSuggesting, author: 'Ada' }));
    const store = (paraId: string, delta: Record<string, boolean>) => {
      session.setSelection({ story: 'body', paraId, offset: 1 });
      act(() => input.current!.applyStoredFormatting({ type: 'set', delta }));
    };
    store(first, { italic: true });
    store(second, { bold: true });
    const type = async (text: string) => {
      act(() => input.current!.insertText(text));
      await act(async () => {
        await input.current!.flushPendingInput();
      });
    };
    const attributesOf = (text: string) => {
      const segment = session
        .storySegments('body')
        .find((candidate) => candidate.kind === 'text' && candidate.text.includes(text));
      return segment?.attributes ?? {};
    };
    return { session, first, second, type, attributesOf };
  };

  // Plain type-over formats the text with its own paragraph's stored formatting.
  const plain = await setup(false);
  plain.session.setSelection(
    { story: 'body', paraId: plain.first, offset: 1 },
    { story: 'body', paraId: plain.second, offset: 1 }
  );
  await plain.type('X');
  expect(plain.attributesOf('X').italic).toBe(true);
  expect(plain.attributesOf('X').bold).not.toBe(true);
  cleanup();

  // A suggested one carries the head's, without displacing another paragraph's.
  const suggested = await setup(true);
  suggested.session.setSelection(
    { story: 'body', paraId: suggested.second, offset: 1 },
    { story: 'body', paraId: suggested.first, offset: 1 }
  );
  await suggested.type('X\nY');
  expect(suggested.attributesOf('X').italic).toBe(true);
  expect(suggested.attributesOf('Y').italic).toBe(true);
  suggested.session.setSelection({ story: 'body', paraId: suggested.second, offset: 0 });
  await suggested.type('Z');
  expect(suggested.attributesOf('Z').bold).toBe(true);
  cleanup();
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

test('a key queued behind resident typing seals the text batch', async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const { session, input, view } = await mount(async () => {
    await blocked;
    return null;
  });
  const textarea = view.getByTestId('yrs-input');
  act(() => input.current!.insertText('A'));
  act(() => input.current!.insertText('B'));
  fireEvent.keyDown(textarea, { key: 'Backspace' });
  act(() => input.current!.insertText('C'));
  fireEvent.keyDown(textarea, { key: 'ArrowLeft' });
  act(() => input.current!.insertText('D'));
  await act(async () => {
    release();
    await input.current!.flushPendingInput();
  });
  expect(text(session)).toBe('SeedADC');
});

test('deletes queued behind busy input reach the resident engine as one batch', async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const deletes: string[] = [];
  const { session, input, view } = await mount(
    async () => {
      await blocked;
      return null;
    },
    async (direction, count) => {
      deletes.push(`${direction}:${count}`);
      return null;
    }
  );
  const textarea = view.getByTestId('yrs-input');
  act(() => input.current!.insertText('XYZ'));
  for (let i = 0; i < 3; i += 1) fireEvent.keyDown(textarea, { key: 'Backspace' });
  await act(async () => {
    release();
    await input.current!.flushPendingInput();
  });
  expect(deletes[0]).toBe('backward:3');
  expect(text(session)).toBe('Seed');
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

test('input and commands after lost input still apply', async () => {
  const originalError = console.error;
  console.error = () => {};
  try {
    let fail = true;
    const { session, input } = await mount(async () => {
      if (!fail) return null;
      fail = false;
      throw new Error('resident failure');
    });
    act(() => input.current!.insertText(' lost'));
    await expect(input.current!.flushPendingInput()).rejects.toThrow('resident failure');
    act(() => input.current!.insertText(' kept'));
    let ran: unknown;
    await act(async () => {
      await input.current!.flushPendingInput();
      ran = await input.current!.runAfterPendingInput(() => 'ran');
    });
    expect(ran).toBe('ran');
    expect(text(session)).toBe('Seed kept');
  } finally {
    console.error = originalError;
  }
});

test('a flush or command waiting on composition fails when earlier input fails meanwhile', async () => {
  const originalError = console.error;
  let reported!: () => void;
  const failureReported = new Promise<void>((resolve) => {
    reported = resolve;
  });
  console.error = () => reported();
  try {
    let release!: () => void;
    const blocked = new Promise<void>((resolve) => {
      release = resolve;
    });
    let fail = true;
    const { session, input, view } = await mount(async () => {
      if (!fail) return null;
      fail = false;
      await blocked;
      throw new Error('resident failure');
    });
    const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
    act(() => input.current!.insertText(' lost'));
    fireEvent.compositionStart(textarea);
    const flush = input.current!.flushPendingInput().catch((cause) => cause);
    const command = input.current!.runAfterPendingInput(() => 'ran').catch((cause) => cause);
    release();
    await failureReported;
    textarea.value = '日本';
    let outcomes: unknown[] = [];
    await act(async () => {
      fireEvent.compositionEnd(textarea, { data: '日本' });
      outcomes = [await flush, admissionCode(await command)];
    });
    expect(outcomes).toEqual([new Error('resident failure'), 'input-failed']);
    let ran: unknown;
    await act(async () => {
      await input.current!.flushPendingInput();
      ran = await input.current!.runAfterPendingInput(() => 'ran');
    });
    expect(ran).toBe('ran');
    expect(text(session)).toBe('Seed日本');
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

test('typing in a table cell is offered to the resident engine', async () => {
  const session = await seededSession();
  const paraId = session.paragraphs('body')[0]!.paraId;
  session.insertTable({ story: 'body', paraId, offset: 0 }, 1, 1);
  const cell = session.storyIds().find((story) => story.startsWith('body:'))!;
  const cellParagraph = session.paragraphs(cell)[0]!.paraId;
  session.setSelection({ story: cell, paraId: cellParagraph, offset: 0 });
  const map = (story = 'body') =>
    createYrsInputPositionMap(
      story,
      session.paragraphs(story).map((p) => ({ paraId: p.paraId, length: p.text.length }))
    );
  const offered: string[] = [];
  const input = createRef<YrsInputRef>();
  render(
    <YrsInput
      ref={input}
      enabled
      readOnly={false}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position, story) => displayPositionToYrsLoc(map(story), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(loc.story), loc)}
      onStateChange={() => {}}
      onDirectInput={() => {}}
      applyResidentInput={async (text) => {
        offered.push(text);
        return null;
      }}
    />
  );
  act(() => input.current!.insertText('x'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  expect(offered).toEqual(['x']);
  expect(session.paragraphs(cell)[0]!.text).toBe('x');
});
