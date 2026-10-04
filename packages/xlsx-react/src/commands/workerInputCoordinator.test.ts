import { describe, expect, test } from 'bun:test';
import { XlsxCommandAdmissionError } from './createXlsxCommandStore';
import type { InputDraft, InputSeal } from './inputCoordinator';
import {
  createWorkerInputCoordinator,
  WorkerInputNotReadyError,
  type WorkerInputCoordinator,
} from './workerInputCoordinator';

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, failed) => {
    resolve = done;
    reject = failed;
  });
  return { promise, resolve, reject };
}

function harness() {
  const log: string[] = [];
  const previews: InputDraft[] = [];
  const writes: InputDraft[] = [];
  const errors: unknown[] = [];
  const hydration: string[] = [];
  const focus: string[] = [];
  const state = {
    generation: 1,
    sheet: 0,
    target: 'A1',
    ready: false,
    peer: deferred<void>(),
    paint: null as Promise<void> | null,
    previewEntered: deferred<void>(),
    shown: null as string | null,
    seal: {} as InputSeal,
    accept: true,
    refuse: new Set<string>(),
    writeError: null as Error | null,
    onSeal: null as (() => void) | null,
    onHydration: null as ((reason: string) => void | Promise<void>) | null,
  };
  let coordinator!: WorkerInputCoordinator;
  const hooks = {
    generation: () => state.generation,
    capture: () => ({ sheet: state.sheet, target: state.target }),
    isReady: () => state.ready,
    whenReady: () => (state.ready ? Promise.resolve() : state.peer.promise),
    seal: () => {
      state.onSeal?.();
      return state.seal;
    },
    sync: () => {
      const current = coordinator.draft;
      if (current && state.shown !== null && current.value !== state.shown) {
        coordinator.setDraft({ ...current, value: state.shown });
      }
    },
    preview: (draft: InputDraft) => {
      previews.push(draft);
      state.previewEntered.resolve();
      return state.paint ?? Promise.resolve();
    },
    write: (draft: InputDraft) => {
      if (state.writeError) throw state.writeError;
      if (!state.accept || state.refuse.has(draft.value)) {
        log.push(`refused ${draft.value}`);
        return false;
      }
      writes.push(draft);
      log.push(`write ${draft.value}`);
      return true;
    },
    close: () => focus.push('grid'),
    restore: () => {
      focus.push('cell');
      return true;
    },
    requestHydration: (reason: string) => {
      hydration.push(reason);
      return state.onHydration?.(reason);
    },
    onError: (error: unknown) => errors.push(error),
  };
  coordinator = createWorkerInputCoordinator(hooks);
  const draft = (
    value: string,
    row = 0,
    sheet = state.sheet,
    source: InputDraft['source'] = 'cell'
  ) => {
    state.shown = value;
    return { generation: state.generation, sheet, row, col: 0, value, source } satisfies InputDraft;
  };
  const ready = () => {
    state.ready = true;
    state.peer.resolve();
  };
  return { coordinator, log, state, draft, ready, previews, writes, errors, hydration, focus };
}

async function errorOf(promise: Promise<unknown>): Promise<unknown> {
  try {
    await promise;
    return null;
  } catch (error) {
    return error instanceof XlsxCommandAdmissionError ? error.code : error;
  }
}

describe('worker input coordinator', () => {
  test('replays every typed key accepted before readiness in order', async () => {
    const { coordinator, draft, ready, log, hydration } = harness();
    let text = '';
    const commits: Promise<void>[] = [];
    for (const [row, key] of ['a', 'b', 'c'].entries()) {
      text += key;
      const typed = draft(text, row, 0, row === 2 ? 'formula' : 'cell');
      coordinator.setDraft(typed);
      expect(coordinator.draft?.value).toBe(text);
      commits.push(coordinator.submitAsync(typed));
    }
    expect(log).toEqual([]);
    expect(coordinator.committed.map((entry) => entry.value)).toEqual(['a', 'ab', 'abc']);
    ready();
    await Promise.all(commits);
    expect(log).toEqual(['write a', 'write ab', 'write abc']);
    expect(hydration).toEqual([]);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('keeps queued commits separate from repeated live draft updates', async () => {
    const { coordinator, draft, ready, writes } = harness();
    const first = draft('first');
    coordinator.setDraft(first);
    expect(coordinator.submit(first)).toBe(true);
    const second = draft('second', 1);
    coordinator.setDraft(second);
    const committed = coordinator.submitAsync(second);
    coordinator.setDraft(draft('l', 2));
    coordinator.setDraft(draft('la', 2));
    const live = draft('later', 2);
    coordinator.setDraft(live);
    ready();
    await committed;
    expect(writes).toEqual([first, second]);
    expect(coordinator.draft).toBe(live);
  });

  test('awaits peer readiness and preview paint before the facade write', async () => {
    const { coordinator, draft, ready, state, previews, log } = harness();
    const paint = deferred<void>();
    state.paint = paint.promise;
    const committed = coordinator.submitAsync(draft('typed'));
    expect(previews).toEqual([]);
    ready();
    await state.previewEntered.promise;
    expect(previews.map((entry) => entry.value)).toEqual(['typed']);
    expect(log).toEqual([]);
    paint.resolve();
    await committed;
    expect(log).toEqual(['write typed']);
  });

  test('captures clipboard payload and target during the original gesture', async () => {
    const { coordinator, state, ready, log } = harness();
    const payload = { text: 'pasted', cells: [['one']] };
    let inGesture = true;
    let captures = 0;
    const paste = coordinator.clipboard(
      () => {
        expect(inGesture).toBe(true);
        captures += 1;
        return payload;
      },
      (input, intent) => {
        expect(input).toEqual({ text: 'pasted', cells: [['one']] });
        expect(intent.generation).toBe(1);
        expect(intent.sheet).toBe(0);
        expect(intent.target).toBe('A1');
        log.push(input.text);
      }
    );
    inGesture = false;
    payload.text = 'changed';
    payload.cells[0][0] = 'changed';
    state.sheet = 4;
    state.target = 'Z9';
    expect(captures).toBe(1);
    ready();
    await paste;
    expect(log).toEqual(['pasted']);
    expect(captures).toBe(1);
  });

  test('shares one FIFO across commits, clipboard, charts, commands, hosts and plugins', async () => {
    const { coordinator, draft, ready, state, log } = harness();
    const commit = coordinator.submitAsync(draft('typed'));
    const paste = coordinator.clipboard(() => 'paste', (input) => log.push(input));
    const chartInput = { chartId: 'chart-1', position: { x: 10, y: 20 } };
    const chart = coordinator.runAfterPendingInput(
      (intent) => {
        expect(intent.input).toEqual({ chartId: 'chart-1', position: { x: 10, y: 20 } });
        expect(intent.sheet).toBe(0);
        expect(intent.target).toBe('A1');
        log.push('chart');
      },
      { kind: 'chart', input: chartInput }
    );
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    const host = coordinator.runAfterPendingInput(() => log.push('host'), { kind: 'host' });
    const plugin = coordinator.runAfterPendingInput(() => log.push('plugin'), { kind: 'plugin' });
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual([
      'commit',
      'clipboard',
      'chart',
      'command',
      'host',
      'plugin',
    ]);
    chartInput.position.x = 999;
    state.sheet = 3;
    state.target = 'B4';
    ready();
    await Promise.all([commit, paste, chart, command, host, plugin]);
    expect(log).toEqual(['write typed', 'paste', 'chart', 'command', 'host', 'plugin']);
  });

  test('does not move focus or overwrite live typing when readiness completes', async () => {
    const { coordinator, draft, ready, focus } = harness();
    const committed = coordinator.submitAsync(draft('cell'));
    const live = draft('formula still typing', 1, 0, 'formula');
    coordinator.setDraft(live);
    focus.push('formula');
    ready();
    await committed;
    expect(focus).toEqual(['formula']);
    expect(coordinator.draft).toBe(live);
  });

  test('reports failure once and retains all unapplied intents for recovery', async () => {
    const { coordinator, draft, state, ready, log, errors, hydration } = harness();
    const commit = coordinator.submitAsync(draft('typed'));
    const paste = coordinator.clipboard(() => 'paste', (input) => log.push(input));
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    const accepted = coordinator.unapplied;
    const failed = new Error('worker failed');
    coordinator.fail(failed);
    coordinator.fail(new Error('same session'));
    expect(await errorOf(commit)).toBe(failed);
    expect(await errorOf(paste)).toBe(failed);
    expect(await errorOf(command)).toBe(failed);
    expect(coordinator.unapplied).toEqual(accepted);
    const late = coordinator.submitAsync(draft('accepted after failure', 1));
    expect(await errorOf(late)).toBe(failed);
    expect(errors).toEqual([failed]);
    expect(coordinator.error).toBe(failed);
    state.onHydration = ready;
    await coordinator.recover();
    expect(log).toEqual(['write typed', 'paste', 'command', 'write accepted after failure']);
    expect(hydration).toEqual(['recovery']);
    expect(coordinator.unapplied).toEqual([]);
    expect(coordinator.error).toBeNull();
  });

  test('discards old intents on replacement and ignores late readiness', async () => {
    const { coordinator, draft, state, ready, log, previews } = harness();
    const oldPeer = state.peer;
    const commit = coordinator.submitAsync(draft('old'));
    const command = coordinator.runAfterPendingInput(() => log.push('old command'));
    state.generation += 1;
    coordinator.reset();
    expect(await errorOf(commit)).toBe('document-replaced');
    expect(await errorOf(command)).toBe('document-replaced');
    expect(coordinator.unapplied).toEqual([]);
    expect(coordinator.draft).toBeNull();
    expect(coordinator.pending).toBe(false);
    state.peer = deferred<void>();
    const next = coordinator.submitAsync(draft('new'));
    oldPeer.resolve();
    ready();
    await next;
    expect(log).toEqual(['write new']);
    expect(previews.map((entry) => entry.value)).toEqual(['new']);
  });

  test('prevents synchronous host mutations from overtaking queued work', async () => {
    const { coordinator, draft, ready, state, log } = harness();
    const paint = deferred<void>();
    state.paint = paint.promise;
    const commit = coordinator.submitAsync(draft('typed'));
    expect(() => coordinator.runSync(() => log.push('early host'))).toThrow(WorkerInputNotReadyError);
    ready();
    await state.previewEntered.promise;
    expect(() => coordinator.runSync(() => log.push('early host'))).toThrow(WorkerInputNotReadyError);
    paint.resolve();
    await commit;
    await coordinator.flush();
    expect(coordinator.pending).toBe(false);
    expect(coordinator.runSync(() => log.push('host'))).toBe(2);
    expect(coordinator.runSync(() => log.push('next host'))).toBe(3);
    expect(log).toEqual(['write typed', 'host', 'next host']);
  });

  test('keeps a refused commit and newer typing until a correction is written', async () => {
    const { coordinator, draft, ready, state, log, focus } = harness();
    const refused = draft('refused', 0, 0);
    const commit = coordinator.submitAsync(refused);
    const live = draft('newer', 1, 1);
    coordinator.setDraft(live);
    state.accept = false;
    ready();
    expect(await errorOf(commit)).toBe('input-failed');
    expect(coordinator.draft).toBe(live);
    expect(coordinator.rejected).toEqual([refused]);
    expect(coordinator.unapplied.map((entry) => entry.input)).toEqual([refused]);
    expect(await errorOf(coordinator.runAfterPendingInput(() => log.push('blocked')))).toBe(
      'input-failed'
    );
    state.accept = true;
    const corrected = draft('corrected', 0, 0);
    coordinator.setDraft(corrected);
    await coordinator.submitAsync(corrected);
    expect(coordinator.rejected).toEqual([]);
    expect(coordinator.unapplied).toEqual([]);
    expect(focus).toEqual([]);
    expect(log).toEqual(['refused refused', 'write corrected']);
  });

  test('a queued correction supersedes a refusal without restoring over live input', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    state.refuse.add('bad');
    const refused = coordinator.submitAsync(draft('bad'));
    const corrected = draft('corrected');
    coordinator.setDraft(corrected);
    expect(coordinator.settle()).toBe(true);
    const settled = coordinator.submitAsync(corrected);
    const live = draft('live', 1);
    coordinator.setDraft(live);
    ready();
    expect(await errorOf(refused)).toBe('input-failed');
    await settled;
    expect(coordinator.draft).toBe(live);
    await coordinator.flush();
    expect(coordinator.rejected).toEqual([]);
    expect(coordinator.unapplied).toEqual([]);
    expect(log).toEqual(['refused bad', 'write corrected', 'write live']);
  });

  test('takes final composition text once before queued commands', async () => {
    const { coordinator, draft, state, ready, log, previews } = harness();
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    coordinator.setDraft(draft('日本'));
    const first = coordinator.runAfterPendingInput((intent) => {
      expect(intent.sheet).toBe(0);
      expect(intent.target).toBe('A1');
      log.push('first');
    });
    const second = coordinator.runAfterPendingInput(() => log.push('second'));
    state.shown = '日本語';
    state.sheet = 2;
    state.target = 'C3';
    composition.resolve(true);
    ready();
    await Promise.all([first, second]);
    expect(log).toEqual(['write 日本語', 'first', 'second']);
    expect(previews).toHaveLength(1);
  });

  test('refuses commands when composition ends without its input or a gesture is active', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    state.seal = { refused: 'gesture-active' };
    expect(await errorOf(coordinator.runAfterPendingInput(() => log.push('gesture')))).toBe(
      'gesture-active'
    );
    expect(coordinator.unapplied).toEqual([]);
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    coordinator.setDraft(draft('日本'));
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    composition.resolve(false);
    ready();
    expect(await errorOf(command)).toBe('input-failed');
    expect(log).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['command']);
  });

  test('writes a draft once when commands and Enter accept the same draft', async () => {
    const { coordinator, draft, ready, log, previews } = harness();
    const typed = draft('typed');
    coordinator.setDraft(typed);
    const first = coordinator.runAfterPendingInput(() => log.push('first'));
    const second = coordinator.runAfterPendingInput(() => log.push('second'));
    const enter = coordinator.submitAsync(typed);
    ready();
    await Promise.all([first, second, enter]);
    expect(log).toEqual(['write typed', 'first', 'second']);
    expect(previews).toHaveLength(1);
  });

  test('keeps final composed text ahead of a command when Enter commits first', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    coordinator.setDraft(draft('日本'));
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    coordinator.setDraft(draft('日本語'));
    const enter = coordinator.submitAsync(coordinator.draft!);
    state.seal = {};
    composition.resolve(true);
    ready();
    await Promise.all([command, enter]);
    expect(log).toEqual(['write 日本語', 'command']);
  });

  test('fails a command behind input landed by its own seal', async () => {
    const { coordinator, state, ready, log } = harness();
    state.onSeal = () => {
      void coordinator.input(() => false);
    };
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    ready();
    expect(await errorOf(command)).toBe('input-failed');
    expect(log).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['input', 'command']);
  });

  test('recovers a failed command without writing its already applied draft twice', async () => {
    const { coordinator, draft, ready, log } = harness();
    const failed = new Error('command failed');
    let refuse = true;
    coordinator.setDraft(draft('typed'));
    const command = coordinator.runAfterPendingInput(() => {
      if (refuse) throw failed;
      log.push('command');
    });
    ready();
    expect(await errorOf(command)).toBe(failed);
    expect(log).toEqual(['write typed']);
    expect(coordinator.committed).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['command']);
    refuse = false;
    await coordinator.recover();
    expect(log).toEqual(['write typed', 'command']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('does not replay an applied edit whose asynchronous acknowledgement fails', async () => {
    const { coordinator, ready, log } = harness();
    const acknowledgement = deferred<void>();
    const entered = deferred<void>();
    const edit = coordinator.runAfterPendingInput(
      (_intent, markApplied) => {
        log.push('edit');
        markApplied();
        entered.resolve();
        return acknowledgement.promise;
      },
      { kind: 'host' }
    );
    const after = coordinator.runAfterPendingInput(() => log.push('after'));
    ready();
    await entered.promise;
    expect(() => coordinator.runSync(() => log.push('overtake'))).toThrow(WorkerInputNotReadyError);
    const failed = new Error('acknowledgement failed');
    acknowledgement.reject(failed);
    expect(await errorOf(edit)).toBe(failed);
    expect(await errorOf(after)).toBe(failed);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['command']);
    await coordinator.recover();
    expect(log).toEqual(['edit', 'after']);
  });

  test('retains a draft when preview fails and retries it during recovery', async () => {
    const { coordinator, draft, ready, state, log } = harness();
    const paint = deferred<void>();
    state.paint = paint.promise;
    const commit = coordinator.submitAsync(draft('typed'));
    ready();
    await state.previewEntered.promise;
    const failed = new Error('preview failed');
    paint.reject(failed);
    expect(await errorOf(commit)).toBe(failed);
    expect(log).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['commit']);
    state.paint = null;
    await coordinator.recover();
    expect(log).toEqual(['write typed']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('retains clipboard capture after failure without reading the clipboard again', async () => {
    const { coordinator, state, ready, log } = harness();
    const captured = deferred<{ text: string }>();
    let reads = 0;
    const paste = coordinator.clipboard(
      () => {
        reads += 1;
        return captured.promise;
      },
      (input) => log.push(input.text)
    );
    const failed = new Error('worker failed');
    coordinator.fail(failed);
    expect(await errorOf(paste)).toBe(failed);
    captured.resolve({ text: 'retained clipboard' });
    expect(await coordinator.unapplied[0].input).toEqual({ text: 'retained clipboard' });
    state.onHydration = ready;
    await coordinator.recover();
    expect(reads).toBe(1);
    expect(log).toEqual(['retained clipboard']);
  });

  test('cancels a write when replacement occurs during its preview', async () => {
    const { coordinator, draft, ready, state, log } = harness();
    const paint = deferred<void>();
    state.paint = paint.promise;
    const commit = coordinator.submitAsync(draft('old'));
    ready();
    await state.previewEntered.promise;
    state.generation += 1;
    coordinator.reset();
    expect(await errorOf(commit)).toBe('document-replaced');
    paint.resolve();
    state.paint = null;
    await coordinator.submitAsync(draft('new'));
    expect(log).toEqual(['write new']);
  });

  test('requests hydration only through explicit flush, request and recovery calls', async () => {
    const { coordinator, draft, state, ready, hydration, log } = harness();
    coordinator.setDraft(draft('live'));
    expect(hydration).toEqual([]);
    state.onHydration = ready;
    await coordinator.flush();
    expect(hydration).toEqual(['flush']);
    expect(log).toEqual(['write live']);
    await coordinator.requestHydration('on-demand');
    await coordinator.recover();
    expect(hydration).toEqual(['flush', 'on-demand', 'recovery']);
  });

  test('discards queued work if the generation changes without an explicit reset', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const commit = coordinator.submitAsync(draft('old'));
    state.generation += 1;
    ready();
    expect(await errorOf(commit)).toBe('document-replaced');
    expect(log).toEqual([]);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('retains a failed synchronous host operation with its acceptance snapshot', async () => {
    const { coordinator, state, ready, log } = harness();
    ready();
    const failed = new Error('facade failed');
    const input = { value: 'accepted' };
    let refuse = true;
    expect(() =>
      coordinator.runSync(
        (intent) => {
          if (refuse) throw failed;
          expect(intent.sheet).toBe(0);
          expect(intent.target).toBe('A1');
          expect(intent.input).toEqual({ value: 'accepted' });
          log.push('host');
        },
        { input }
      )
    ).toThrow(failed);
    input.value = 'changed';
    state.sheet = 2;
    state.target = 'B2';
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['host']);
    refuse = false;
    await coordinator.recover();
    expect(log).toEqual(['host']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('drains work accepted as the preceding operation completes', async () => {
    const { coordinator, ready, log } = harness();
    ready();
    await coordinator
      .runAfterPendingInput(() => log.push('first'))
      .then(() => coordinator.runAfterPendingInput(() => log.push('second')));
    expect(log).toEqual(['first', 'second']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('retains a draft when the facade throws and writes it during recovery', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const failed = new Error('facade failed');
    state.writeError = failed;
    const commit = coordinator.submitAsync(draft('typed'));
    ready();
    expect(await errorOf(commit)).toBe(failed);
    expect(log).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['commit']);
    state.writeError = null;
    await coordinator.recover();
    expect(log).toEqual(['write typed']);
  });

  test('recovers final composition text that resolves after session failure', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    coordinator.setDraft(draft('日本'));
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    const failed = new Error('worker failed');
    coordinator.fail(failed);
    expect(await errorOf(command)).toBe(failed);
    state.shown = '日本語';
    composition.resolve(true);
    state.onHydration = ready;
    await coordinator.recover();
    expect(log).toEqual(['write 日本語', 'command']);
  });

  test('does not replay a refused command draft after its correction is written', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    state.refuse.add('bad');
    coordinator.setDraft(draft('bad'));
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    ready();
    expect(await errorOf(command)).toBe('input-failed');
    const corrected = draft('corrected');
    coordinator.setDraft(corrected);
    await coordinator.submitAsync(corrected);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['command']);
    await coordinator.recover();
    expect(log).toEqual(['refused bad', 'write corrected', 'command']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('rejects a flush promptly when replacement interrupts hydration', async () => {
    const { coordinator, draft, state, log } = harness();
    const hydration = deferred<void>();
    state.onHydration = () => hydration.promise;
    coordinator.setDraft(draft('old'));
    const flush = coordinator.flush();
    state.generation += 1;
    coordinator.reset();
    expect(await errorOf(flush)).toBe('document-replaced');
    expect(coordinator.unapplied).toEqual([]);
    expect(log).toEqual([]);
    hydration.resolve();
  });

  test('captures clipboard access before sealing and waits for final composition text', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    state.onSeal = () => log.push('seal');
    coordinator.setDraft(draft('日本'));
    const paste = coordinator.clipboard(
      () => {
        log.push('capture');
        return 'paste';
      },
      (input) => log.push(input)
    );
    expect(log).toEqual(['capture', 'seal']);
    state.shown = '日本語';
    composition.resolve(true);
    ready();
    await paste;
    expect(log).toEqual(['capture', 'seal', 'write 日本語', 'paste']);
  });

  test('retains refused clipboard data and blocks commands behind it until recovery', async () => {
    const { coordinator, ready, log } = harness();
    let accept = false;
    const paste = coordinator.clipboard(
      () => 'paste',
      (input) => {
        if (!accept) return false;
        log.push(input);
        return true;
      }
    );
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    ready();
    expect(await errorOf(paste)).toBe('input-failed');
    expect(await errorOf(command)).toBe('input-failed');
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['clipboard', 'command']);
    expect(coordinator.unapplied[0].input).toBe('paste');
    accept = true;
    await coordinator.recover();
    expect(log).toEqual(['paste', 'command']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('records synchronous operation application before a failure in the next microtask', async () => {
    for (const kind of ['input', 'clipboard', 'command'] as const) {
      const { coordinator, ready, log } = harness();
      const failed = new Error(`${kind} acknowledgement failed`);
      const apply = () => {
        log.push(kind);
        queueMicrotask(() => coordinator.fail(failed));
        return true;
      };
      let result: Promise<unknown>;
      if (kind === 'input') result = coordinator.input(apply);
      else if (kind === 'clipboard') result = coordinator.clipboard(() => 'payload', apply);
      else result = coordinator.runAfterPendingInput(apply);
      const after = coordinator.runAfterPendingInput(() => log.push('after'));
      ready();
      expect(await errorOf(result)).toBe(failed);
      expect(await errorOf(after)).toBe(failed);
      expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['command']);
      await coordinator.recover();
      expect(log).toEqual([kind, 'after']);
    }
  });
});
