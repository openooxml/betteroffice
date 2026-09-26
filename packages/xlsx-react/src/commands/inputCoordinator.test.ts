import { describe, expect, test } from 'bun:test';
import { XlsxCommandAdmissionError } from './createXlsxCommandStore';
import { createInputCoordinator, type InputDraft, type InputSeal } from './inputCoordinator';

function harness() {
  const log: string[] = [];
  const state = {
    generation: 1,
    seal: {} as InputSeal,
    accept: true,
    shown: '' as string | null,
    sheet: 0,
  };
  const coordinator = createInputCoordinator({
    generation: () => state.generation,
    seal: () => {
      log.push('seal');
      return state.seal;
    },
    sync: () => {
      const draft = coordinator.draft;
      if (draft && state.shown !== null && state.shown !== draft.value) {
        coordinator.setDraft({ ...draft, value: state.shown });
      }
    },
    write: (draft) => {
      if (!state.accept) {
        log.push(`refused ${draft.value}`);
        return false;
      }
      log.push(`write ${draft.value}`);
      return true;
    },
    close: (draft) => log.push(`close ${draft.value}`),
    restore: (draft) => {
      log.push(`restore ${draft.value}`);
      return draft.sheet === state.sheet;
    },
  });
  const draft = (value: string, row = 0, sheet = 0): InputDraft => {
    state.shown = value;
    return { generation: state.generation, sheet, row, col: 0, value, source: 'cell' };
  };
  const hold = () => {
    let release!: () => void;
    const held = new Promise<void>((resolve) => (release = resolve));
    void coordinator.input(async () => {
      await held;
      log.push('paste');
    });
    return release;
  };
  return { coordinator, log, state, draft, hold };
}

async function codeOf(promise: Promise<unknown>): Promise<string | null> {
  try {
    await promise;
    return null;
  } catch (error) {
    return error instanceof XlsxCommandAdmissionError ? error.code : String(error);
  }
}

describe('input coordinator', () => {
  test('writes the draft once before the command, synchronously when nothing waits', async () => {
    const { coordinator, log, draft } = harness();
    coordinator.setDraft(draft('typed'));
    const result = coordinator.runAfterPendingInput(() => {
      log.push('command');
      return 'done';
    });
    expect(log).toEqual(['seal', 'write typed', 'close typed', 'command']);
    expect(await result).toBe('done');
    expect(coordinator.draft).toBeNull();
  });

  test('runs input writes and commands in the order they were accepted', async () => {
    const { coordinator, log, draft, hold } = harness();
    const release = hold();
    coordinator.setDraft(draft('before'));
    const undo = coordinator.runAfterPendingInput(() => log.push('undo'));
    const secondPaste = coordinator.input(() => {
      log.push('second paste');
    });
    coordinator.setDraft(draft('typed later'));
    expect(coordinator.submit(coordinator.draft!)).toBe(true);
    const save = coordinator.runAfterPendingInput(() => log.push('save'));
    expect(log).toEqual(['seal', 'seal']);
    release();
    await Promise.all([undo, secondPaste, save]);
    expect(log).toEqual([
      'seal',
      'seal',
      'paste',
      'write before',
      'undo',
      'second paste',
      'write typed later',
      'save',
    ]);
  });

  test('fails the commands behind a failed queued write and restores the draft', async () => {
    const { coordinator, log, state, draft, hold } = harness();
    const release = hold();
    const finished = draft('finished');
    coordinator.setDraft(finished);
    expect(coordinator.submit(finished)).toBe(true);
    expect(coordinator.draft).toBeNull();
    const bold = coordinator.runAfterPendingInput(() => log.push('bold'));
    state.accept = false;
    release();
    expect(await codeOf(bold)).toBe('input-failed');
    expect(coordinator.draft).toBe(finished);
    expect(log).toEqual(['seal', 'paste', 'refused finished', 'restore finished']);

    state.accept = true;
    expect(await codeOf(coordinator.runAfterPendingInput(() => log.push('blocked')))).toBe(
      'input-failed'
    );
    const corrected = draft('corrected');
    coordinator.setDraft(corrected);
    expect(coordinator.submit(corrected)).toBe(true);
    expect(coordinator.rejected).toEqual([]);
    await coordinator.runAfterPendingInput(() => log.push('bold again'));
    expect(log.slice(-3)).toEqual(['write corrected', 'seal', 'bold again']);
  });

  test('keeps a draft that fails to write at once, and refuses its command', async () => {
    const { coordinator, state, draft } = harness();
    state.accept = false;
    const kept = draft('kept');
    coordinator.setDraft(kept);
    expect(coordinator.submit(kept)).toBe(false);
    expect(coordinator.draft).toBe(kept);
    expect(await codeOf(coordinator.runAfterPendingInput(() => 'bold'))).toBe('input-failed');
    expect(coordinator.draft).toBe(kept);
    state.accept = true;
    expect(coordinator.submit(kept)).toBe(true);
    expect(await coordinator.runAfterPendingInput(() => 'bold')).toBe('bold');
  });

  test('finishes a composition before the command and takes its final text', async () => {
    const { coordinator, log, state, draft } = harness();
    let end!: (ended: boolean) => void;
    state.seal = { composition: new Promise<boolean>((resolve) => (end = resolve)) };
    coordinator.setDraft(draft('日本'));
    const italic = coordinator.runAfterPendingInput(() => log.push('italic'));
    const after = coordinator.runAfterPendingInput(() => log.push('after'));
    state.shown = '日本語';
    expect(log).toEqual(['seal', 'seal']);
    end(true);
    await Promise.all([italic, after]);
    expect(log).toEqual(['seal', 'seal', 'write 日本語', 'close 日本語', 'italic', 'after']);
  });

  test('fails commands waiting on a composition whose input went away', async () => {
    const { coordinator, log, state, draft } = harness();
    let end!: (ended: boolean) => void;
    state.seal = { composition: new Promise<boolean>((resolve) => (end = resolve)) };
    coordinator.setDraft(draft('日本'));
    const italic = coordinator.runAfterPendingInput(() => log.push('italic'));
    end(false);
    expect(await codeOf(italic)).toBe('input-failed');
    expect(log).toEqual(['seal']);
  });

  test('writes a draft once when several queued commands sealed it', async () => {
    const { coordinator, log, draft, hold } = harness();
    const release = hold();
    const typed = draft('typed');
    coordinator.setDraft(typed);
    const undo = coordinator.runAfterPendingInput(() => log.push('undo'));
    const save = coordinator.runAfterPendingInput(() => log.push('save'));
    expect(coordinator.submit(typed)).toBe(true);
    release();
    await Promise.all([undo, save]);
    expect(log).toEqual(['seal', 'seal', 'paste', 'write typed', 'undo', 'save']);
  });

  test('keeps newer typing when a queued commit fails', async () => {
    const { coordinator, log, state, draft, hold } = harness();
    const release = hold();
    const first = draft('first');
    coordinator.setDraft(first);
    expect(coordinator.submit(first)).toBe(true);
    const second = draft('second', 1);
    coordinator.setDraft(second);
    state.accept = false;
    release();
    await coordinator.input(() => {});
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(coordinator.draft).toBe(second);
    expect(log).toEqual(['paste', 'refused first']);

    expect(coordinator.rejected).toEqual([first]);

    state.accept = true;
    expect(coordinator.submit(second)).toBe(true);
    expect(coordinator.draft).toBeNull();
    expect(coordinator.rejected).toEqual([first]);
    expect(await codeOf(coordinator.runAfterPendingInput(() => log.push('save')))).toBe(
      'input-failed'
    );
    expect(coordinator.settle()).toBe(true);
    coordinator.discard(first);
    await coordinator.runAfterPendingInput(() => log.push('save'));
    expect(log).toEqual(['paste', 'refused first', 'write second', 'seal', 'save']);
  });

  test('never shows a rejected draft outside its sheet, and keeps a live one open until written', async () => {
    const { coordinator, log, state, draft, hold } = harness();
    state.sheet = 1;
    const release = hold();
    const onFirstSheet = draft('oversized', 0, 0);
    coordinator.setDraft(onFirstSheet);
    expect(coordinator.submit(onFirstSheet)).toBe(true);
    state.accept = false;
    release();
    await coordinator.input(() => {});
    expect(coordinator.draft).toBeNull();
    expect(coordinator.rejected).toEqual([onFirstSheet]);
    expect(log).toEqual(['paste', 'refused oversized', 'restore oversized']);

    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(coordinator.pending).toBe(false);
    coordinator.setDraft(onFirstSheet);
    expect(coordinator.settle()).toBe(false);
    expect(coordinator.draft).toBe(onFirstSheet);
    expect(coordinator.rejected).toEqual([onFirstSheet]);
    expect(log).not.toContain('close oversized');

    state.accept = true;
    expect(coordinator.settle()).toBe(true);
    expect(coordinator.draft).toBeNull();
    expect(coordinator.rejected).toEqual([]);
    expect(log.slice(-2)).toEqual(['write oversized', 'close oversized']);
  });

  test('a queued correction supersedes the refusal of its cell', async () => {
    const { coordinator, log, state, draft, hold } = harness();
    state.accept = false;
    const refused = draft('oversized');
    coordinator.setDraft(refused);
    expect(coordinator.submit(refused)).toBe(false);
    const release = hold();
    state.accept = true;
    coordinator.setDraft(draft('corrected'));
    expect(coordinator.settle()).toBe(true);
    expect(coordinator.draft).toBeNull();
    expect(coordinator.rejected).toEqual([]);
    release();
    await coordinator.runAfterPendingInput(() => log.push('save'));
    expect(coordinator.draft).toBeNull();
    expect(log).toEqual([
      'refused oversized',
      'close corrected',
      'seal',
      'paste',
      'write corrected',
      'save',
    ]);
  });

  test('never restores a failed write over a later queued write of its cell', async () => {
    const { coordinator, log, state, draft, hold } = harness();
    const releaseFirst = hold();
    const first = draft('first');
    coordinator.setDraft(first);
    expect(coordinator.submit(first)).toBe(true);
    const releaseSecond = hold();
    coordinator.setDraft(draft('second'));
    expect(coordinator.settle()).toBe(true);
    state.accept = false;
    releaseFirst();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(coordinator.draft).toBeNull();
    expect(coordinator.rejected).toEqual([]);
    state.accept = true;
    releaseSecond();
    await coordinator.input(() => {});
    expect(coordinator.draft).toBeNull();
    expect(coordinator.rejected).toEqual([]);
    expect(log).toEqual(['close second', 'paste', 'refused first', 'paste', 'write second']);
  });

  test('keeps a composed draft ahead of the command even when Enter commits it first', async () => {
    const { coordinator, log, state, draft, hold } = harness();
    const release = hold();
    let end!: (ended: boolean) => void;
    state.seal = { composition: new Promise<boolean>((resolve) => (end = resolve)) };
    coordinator.setDraft(draft('日本'));
    const save = coordinator.runAfterPendingInput(() => log.push('save'));
    state.seal = {};
    state.shown = '日本語';
    end(true);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(coordinator.submit(coordinator.draft!)).toBe(true);
    release();
    await save;
    await coordinator.input(() => {});
    expect(log).toEqual(['seal', 'paste', 'write 日本語', 'save']);
  });

  test('fails the command whose own seal landed input that failed', async () => {
    const log: string[] = [];
    const coordinator = createInputCoordinator({
      generation: () => 1,
      seal: () => {
        void coordinator.input(() => false);
        return {};
      },
      sync: () => {},
      write: () => true,
      close: () => {},
      restore: () => false,
    });
    expect(await codeOf(coordinator.runAfterPendingInput(() => log.push('save')))).toBe(
      'input-failed'
    );
    expect(log).toEqual([]);
  });

  test('refuses on an unfinished gesture or a replaced document', async () => {
    const { coordinator, log, state, hold } = harness();
    state.seal = { refused: 'gesture-active' };
    expect(await codeOf(coordinator.runAfterPendingInput(() => log.push('moved')))).toBe(
      'gesture-active'
    );
    state.seal = {};
    const release = hold();
    const replaced = coordinator.runAfterPendingInput(() => log.push('save'));
    state.generation += 1;
    release();
    expect(await codeOf(replaced)).toBe('document-replaced');
    expect(log).not.toContain('save');
  });

  test('queues a settled draft behind input that is still waiting', async () => {
    const { coordinator, log, draft, hold } = harness();
    const release = hold();
    const first = draft('first');
    coordinator.setDraft(first);
    expect(coordinator.submit(first)).toBe(true);
    coordinator.setDraft(draft('corrected'));
    expect(coordinator.settle()).toBe(true);
    expect(coordinator.draft).toBeNull();
    expect(coordinator.pending).toBe(true);
    expect(log).toEqual(['close corrected']);
    release();
    await coordinator.input(() => {});
    expect(log).toEqual(['close corrected', 'paste', 'write first', 'write corrected']);
  });

  test('settles synchronously for host calls and drops drafts of an older document', () => {
    const { coordinator, log, state, draft } = harness();
    coordinator.setDraft(draft('host'));
    expect(coordinator.settle()).toBe(true);
    expect(log).toEqual(['write host', 'close host']);
    coordinator.setDraft(draft('stale'));
    state.generation += 1;
    expect(coordinator.settle()).toBe(true);
    expect(coordinator.draft).toBeNull();
    expect(log).toEqual(['write host', 'close host']);
  });
});
