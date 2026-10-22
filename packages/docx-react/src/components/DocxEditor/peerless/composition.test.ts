import { expect, test } from 'bun:test';
import type {
  CommandTicket,
  SubmitOptions,
  TicketOutcome,
} from '../../../../../../shared/office-session/commands/types';
import type { DocxCommandPayload } from '@betteroffice/docx/peerless';
import { beginComposition, endComposition, updateComposition } from './composition';
import type { PeerlessEchoIntent } from './echoPainter';
import type {
  PeerlessAdmissionTarget,
  PeerlessInputAdmission,
} from './inputAdmission';
import type { PeerlessViewSnapshot } from './viewStore';

function fakeAdmission() {
  let snapshot: PeerlessViewSnapshot = {
    identity: { session: 'composition', incarnation: 1, epoch: 1 },
    version: 0,
    throughSeq: 0,
    presentedThroughSeq: 0,
    workerSelection: {
      anchor: { story: 'body', paraId: 'original', offset: 0 },
      focus: { story: 'body', paraId: 'original', offset: 0 },
    },
    anchors: { anchor: new Uint8Array([1]), focus: new Uint8Array([2]) },
    editState: undefined,
    geometry: null,
  };
  let tail: CommandTicket | undefined;
  let slot: number | null = null;
  let nextSlot = 0;
  let settlements = 0;
  let ignored = false;
  const held: string[] = [];
  const commands: {
    kind: string;
    payload: DocxCommandPayload;
    options: SubmitOptions<PeerlessEchoIntent> | undefined;
    ticket: CommandTicket;
    resolve(outcome: TicketOutcome): void;
    reject(error: Error): void;
  }[] = [];
  const listeners = new Set<() => void>();
  const submit: PeerlessInputAdmission['session']['submit'] = (kind, payload, options) => {
    let resolve!: (outcome: TicketOutcome) => void;
    let reject!: (error: Error) => void;
    const ticket: CommandTicket = {
      id: commands.length + 1,
      seq: commands.length + 1,
      ack: new Promise<TicketOutcome>((accept, decline) => {
        resolve = accept;
        reject = decline;
      }),
      presented: Promise.resolve('presented'),
    };
    commands.push({ kind, payload, options, ticket, resolve, reject });
    tail = ticket;
    return ticket;
  };
  const admission: PeerlessInputAdmission = {
    setReadOnly() { throw new Error('Read-only barriers are unavailable'); },
    subscribeReadOnly: () => () => {},
    session: {
      backend: 'worker',
      get state() {
        return {
          epoch: 1,
          version: snapshot.version,
          throughSeq: snapshot.throughSeq,
          presentedThroughSeq: snapshot.presentedThroughSeq,
          selection: snapshot.workerSelection,
          editState: snapshot.editState,
        };
      },
      submit,
      whenPresented: async () => 'presented',
      subscribe: () => () => {},
      flush: async () => {},
      dispose: async () => {},
    },
    viewStore: {
      subscribe(listener) {
        listeners.add(listener);
        return () => { listeners.delete(listener); };
      },
      getSnapshot: () => snapshot,
      getServerSnapshot: () => snapshot,
      acceptWorkerState(identity, state) {
        snapshot = { ...snapshot, identity, ...state, workerSelection: state.selection };
        listeners.forEach((listener) => listener());
      },
      installFrameGeometry(geometry) {
        snapshot = { ...snapshot, geometry };
        listeners.forEach((listener) => listener());
      },
      setProvisionalCaret() {},
      dispose() {},
    },
    echoPainter: {
      installContext() {},
      restoreBackground: () => false,
      preview: () => undefined,
      paint() {},
      drop() {},
      reconcile() {},
      disable() {},
      dispose() {},
    },
    getSnapshot: () => ({ readOnly: snapshot.editState?.readOnly ?? false,
      tail,
      originalUnackedIntents: commands.length,
      pendingInputs: commands.length,
      echoEnabled: true,
      echoDisabledReason: null,
      compositionSlot: slot,
      heldIntents: held.length,
    }),
    reserveComposition() {
      if (slot !== null) throw new Error('composition already reserved');
      slot = ++nextSlot;
      const target: PeerlessAdmissionTarget = tail
        ? { kind: 'predecessor', ticket: tail }
        : { kind: 'worker-selection', snapshot };
      return { slot, target };
    },
    settleComposition(reserved, commit) {
      expect(reserved).toBe(slot!);
      settlements += 1;
      slot = null;
      const ticket = commit && !ignored
        ? submit('replaceSelection', { text: commit.text }, {
            after: commit.target.kind === 'predecessor' ? commit.target.ticket : undefined,
            fence: true,
          })
        : undefined;
      for (const text of held.splice(0)) submit('insertText', { text }, { after: tail });
      return ticket
        ? { status: 'submitted', ticket, echoEnabled: false }
        : { status: 'ignored', reason: ignored ? 'read-only' : 'empty-composition' };
    },
    dispose() {},
  };
  return {
    admission,
    commands,
    hold: (text: string) => { held.push(text); },
    ignore: () => { ignored = true; },
    settlements: () => settlements,
  };
}

test('d2d_composition_emits_zero_commands_then_one_commit', async () => {
  const fake = fakeAdmission();
  const initial = beginComposition(fake.admission, 7);
  expect(initial).toMatchObject({ id: 7, slot: 1, text: '', phase: 'active' });
  expect(initial.target).toEqual({
    kind: 'worker-selection', snapshot: fake.admission.viewStore.getSnapshot(),
  });
  const first = updateComposition(initial, 'か');
  const second = updateComposition(first, 'かな');
  expect(initial.text).toBe('');
  expect(second.text).toBe('かな');
  expect(fake.commands).toHaveLength(0);
  const result = endComposition(fake.admission, second, '漢字');
  expect(result.status).toBe('committed');
  expect(result.state).toMatchObject({ text: '漢字', phase: 'committed' });
  expect(fake.commands).toHaveLength(1);
  expect(fake.commands[0]).toMatchObject({
    kind: 'replaceSelection', payload: { text: '漢字' }, options: { fence: true },
  });
  fake.commands[0]!.resolve({ status: 'applied', seq: 1 });
  await fake.commands[0]!.ticket.ack;
  expect(result.state.phase).toBe('committed');
});

test('d2d_pending_insert_precedes_atomic_composition_commit', () => {
  const fake = fakeAdmission();
  const insert = fake.admission.session.submit('insertText', { text: 'before' });
  const state = updateComposition(beginComposition(fake.admission, 8), 'かな');
  expect(state.target).toEqual({ kind: 'predecessor', ticket: insert });
  fake.hold('later');
  fake.admission.viewStore.acceptWorkerState(
    fake.admission.viewStore.getSnapshot().identity,
    {
      ...fake.admission.session.state,
      version: 1,
      selection: {
        anchor: { story: 'body', paraId: 'worker-reflowed', offset: 42 },
        focus: { story: 'body', paraId: 'worker-reflowed', offset: 42 },
      },
    }
  );
  expect(fake.commands).toHaveLength(1);
  const result = endComposition(fake.admission, state, '漢字');
  expect(result.status).toBe('committed');
  expect(fake.commands.map((command) => command.kind)).toEqual([
    'insertText', 'replaceSelection', 'insertText',
  ]);
  expect(fake.commands[1]!.options?.after).toBe(insert);
  expect(fake.commands[1]!.payload).toEqual({ text: '漢字' });
  expect(fake.commands[2]!.options?.after).toBe(result.state.ticket);
  expect(fake.admission.getSnapshot().heldIntents).toBe(0);
});

test('d2d_trailing_beforeinput_and_input_do_not_duplicate_commit', () => {
  for (const order of [['beforeinput', 'input'], ['input', 'beforeinput']]) {
    const fake = fakeAdmission();
    const initial = beginComposition(fake.admission, 9);
    const preedit = updateComposition(initial, 'かな');
    const result = endComposition(fake.admission, preedit, '漢字');
    for (const source of order) {
      const trailing = source === 'beforeinput' ? initial : result.state;
      expect(endComposition(fake.admission, trailing, '漢字')).toBe(result);
      expect(updateComposition(trailing, 'duplicate')).toBe(result.state);
    }
    expect(fake.commands).toHaveLength(1);
    expect(fake.settlements()).toBe(1);
    const next = beginComposition(fake.admission, 9);
    endComposition(fake.admission, next, '次');
    expect(fake.commands).toHaveLength(2);
    expect(endComposition(fake.admission, initial, 'old')).toBe(result);
  }
});

test('d2d_empty_composition_cancels_without_command', () => {
  const fake = fakeAdmission();
  const state = updateComposition(beginComposition(fake.admission, 10), 'draft');
  const result = endComposition(fake.admission, state, '');
  expect(result).toMatchObject({
    status: 'cancelled', state: { text: '', phase: 'cancelled' },
  });
  expect(fake.commands).toHaveLength(0);
  expect(fake.settlements()).toBe(1);
  expect(endComposition(fake.admission, state, '')).toBe(result);
  expect(fake.admission.getSnapshot().compositionSlot).toBeNull();
  const pending = fake.admission.session.submit('insertText', { text: 'before' });
  const next = beginComposition(fake.admission, 11);
  fake.hold('later');
  endComposition(fake.admission, next, '');
  expect(fake.commands.map((command) => command.kind)).toEqual(['insertText', 'insertText']);
  expect(fake.commands[1]!.options?.after).toBe(pending);
});

test('d2d_frames_preserve_preedit_and_refusal_retains_text', async () => {
  for (const outcome of ['refused', 'cancelled', 'rejected', 'read-only'] as const) {
    const fake = fakeAdmission();
    const state = updateComposition(beginComposition(fake.admission, 12), 'かな');
    const target = state.target;
    fake.admission.viewStore.installFrameGeometry({
      identity: fake.admission.viewStore.getSnapshot().identity,
      coversThroughSeq: 0,
      frameEpoch: 3,
      layoutRevision: 3,
      pages: [],
      caret: null,
      selectionRects: [],
    });
    expect(state.text).toBe('かな');
    expect(state.phase).toBe('active');
    expect(state.target).toBe(target);
    if (outcome === 'read-only') fake.ignore();
    const result = endComposition(fake.admission, state, '漢字');
    if (outcome !== 'read-only') {
      const command = fake.commands[0]!;
      fake.admission.viewStore.acceptWorkerState(
        fake.admission.viewStore.getSnapshot().identity,
        { ...fake.admission.session.state, selection: null }
      );
      if (outcome === 'rejected') command.reject(new Error('unknown outcome'));
      else command.resolve({ status: outcome, reason: 'locked-target' });
      await command.ticket.ack.catch(() => {});
      await Promise.resolve();
      expect(fake.admission.viewStore.getSnapshot().workerSelection).toBeNull();
    } else {
      expect(result.status).toBe('cancelled');
      expect(fake.commands).toHaveLength(0);
    }
    expect(result.state).toMatchObject({
      phase: 'refused', text: '漢字', recoverableText: '漢字',
    });
    expect(updateComposition(result.state, '')).toBe(result.state);
    expect(endComposition(fake.admission, state, 'duplicate')).toBe(result);
    expect(fake.settlements()).toBe(1);
  }
});
