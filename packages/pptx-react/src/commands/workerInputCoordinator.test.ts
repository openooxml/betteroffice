import { expect, mock, test } from 'bun:test';
import { PptxPeerNotReadyError, PptxWorkerEditorDisposedError } from '@betteroffice/pptx';
import type { PptxWorkerEditorAccess, PptxWorkerEditorOperation, PptxWorkerEditorSession, PptxWorkerEditorState } from '@betteroffice/pptx';
import { createWorkerInputCoordinator } from './workerInputCoordinator';

function setup(hydrated = true) {
  let sequence = 0;
  let gesture = false;
  let acknowledge!: () => void;
  const ack = new Promise<void>((resolve) => { acknowledge = resolve; });
  const methods: string[] = [];
  const session = {
    get hydrated() { return hydrated; },
    state: { stage: hydrated ? 'ready' : 'hydrating', sequence: 0, acknowledgedSequence: 0 } as PptxWorkerEditorState,
    apply: mock((op: PptxWorkerEditorOperation) => {
      methods.push(op.method);
      return { consumed: true, sequence: ++sequence, revision: sequence, version: `v${sequence}`,
        engineVersion: `v${sequence}`, outcome: { result: true, applied: true, changedTargets: [], canUndo: true, canRedo: false } };
    }),
    flush: mock(() => ack),
  } as unknown as PptxWorkerEditorSession;
  const coordinator = createWorkerInputCoordinator({ session: () => session, gestureActive: () => gesture });
  const peer = { insertText() {}, deleteText() {}, setSlideNotes() {} } as unknown as PptxWorkerEditorAccess;
  const access = hydrated ? coordinator.access(peer) : peer;
  return { coordinator, access, methods, session, acknowledge, gesture: () => { gesture = true; },
    retire: () => Object.defineProperty(session, 'state', { value: { ...session.state, stage: 'disposed' } }) };
}

test('prehydration_input_is_refused_without_queue_or_preview', async () => {
  const { coordinator, methods } = setup(false);
  await expect(coordinator.input(() => { methods.push('queued'); })).rejects.toBeInstanceOf(PptxPeerNotReadyError);
  expect(methods).toEqual([]);
  expect(coordinator.busy()).toBe(false);
  expect(coordinator.keyboardQueued()).toBe(false);
});

test('keyboard_does_not_wait_for_worker_ack', async () => {
  const { coordinator, access, methods, acknowledge, retire } = setup();
  const first = coordinator.input(() => { access.insertText('story', 0, 'a'); }, 'keyboard');
  const flushing = coordinator.flush();
  const second = coordinator.input(() => { access.insertText('story', 1, 'b'); }, 'keyboard');
  expect(methods.filter((method) => method === 'insertText')).toHaveLength(2);
  expect(coordinator.busy()).toBe(false);
  acknowledge();
  await Promise.all([first, second, flushing]);
  const retained = access.insertText;
  retire();
  expect(() => retained('story', 0, 'retired')).toThrow(PptxWorkerEditorDisposedError);
});

test('image_preparation_is_fenced_and_stale_target_is_reported', async () => {
  const { coordinator, methods, acknowledge } = setup();
  let decode!: () => void;
  let current = true;
  const preparing = coordinator.input(async () => {
    await new Promise<void>((resolve) => { decode = resolve; });
    if (!current) throw new Error('The image target slide no longer exists');
    methods.push('addPicture');
  });
  const failure = preparing.catch((error: Error) => error);
  acknowledge();
  await coordinator.flush();
  expect(methods).toEqual([]);
  current = false;
  decode();
  expect((await failure as Error).message).toContain('target slide');
  expect(methods).toEqual([]);
});

test('unfinished_gesture_refuses_flush', async () => {
  const { coordinator, session, gesture } = setup();
  gesture();
  await expect(coordinator.flush()).rejects.toThrow('Finish the pointer gesture');
  expect(session.flush).not.toHaveBeenCalled();
});

test('manual_history_is_independent_of_worker_delay', async () => {
  const { coordinator, access, methods, acknowledge } = setup();
  for (const time of [0, 200, 701]) {
    coordinator.beginTyping('story', time);
    access.deleteText('story', 0, 1);
    access.insertText('story', 0, 'x');
    coordinator.endTyping();
  }
  expect(methods).toEqual(['addUndoBoundary', 'deleteText', 'insertText', 'deleteText', 'insertText',
    'addUndoBoundary', 'deleteText', 'insertText']);
  access.setSlideNotes('slide', 'notes');
  expect(methods.slice(-3)).toEqual(['addUndoBoundary', 'setSlideNotes', 'addUndoBoundary']);
  acknowledge();
  await coordinator.flush();
});
