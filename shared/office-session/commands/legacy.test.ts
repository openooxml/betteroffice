import { describe, expect, it } from 'bun:test';
import { createLegacyEditSession } from './legacy';
import type {
  EditSessionState, LegacyEditSessionOptions, PresentationStatus, SubmitOptions,
} from './types';

type State = EditSessionState<number, { canUndo: boolean }>;

function initial(): State {
  return {
    epoch: 2, version: 0, throughSeq: 0, presentedThroughSeq: 0,
    selection: 0, editState: { canUndo: false },
  };
}

function harness(backend: LegacyEditSessionOptions<string, number, { canUndo: boolean }, string>['backend']) {
  let state = initial();
  let text = '';
  let disposals = 0;
  const applied: Array<{ kind: string; payload: string; options?: SubmitOptions<string> }> = [];
  const renders: Array<{ seq: number; confirm(coversThroughSeq: number): void }> = [];
  const session = createLegacyEditSession<string, number, { canUndo: boolean }, string>({
    backend, initial: state,
    apply(kind, payload, options) {
      applied.push({ kind, payload, options });
      if (kind === 'throw') throw new Error('apply failed');
      if (kind === 'refuse') {
        return { outcome: { status: 'refused', reason: 'locked', committed: false,
          affectedSeqs: [7] }, state };
      }
      if (kind !== 'noop') {
        text += payload;
        state = { ...state, version: state.version + 1,
          selection: text.length, editState: { canUndo: true } };
      }
      return { outcome: { status: kind === 'noop' ? 'noop' : 'applied' }, state };
    },
    render: (seq, confirm) => { renders.push({ seq, confirm }); },
    dispose: () => { disposals += 1; },
  });
  return { session, applied, renders, get text() { return text; },
    get disposals() { return disposals; } };
}

for (const backend of ['legacy-sync', 'legacy-in-thread'] as const) {
  describe(backend, () => {
    it('mutates inside submit, fulfills ack immediately, and waits for render confirmation', async () => {
      const h = harness(backend);
      const options: SubmitOptions<string> = { echo: 'a', coalesceKey: 'typing', lane: 'input' };
      const ticket = h.session.submit('insert', 'a', options);
      expect(h.text).toBe('a');
      expect(h.session.backend).toBe(backend);
      expect(ticket.seq).toBe(1);
      expect(h.applied[0]?.options).toBe(options);
      expect(h.renders.map((render) => render.seq)).toEqual([1]);
      let acknowledged = false;
      let presentation: PresentationStatus | undefined;
      void ticket.ack.then(() => { acknowledged = true; });
      void ticket.presented.then((status) => { presentation = status; });
      await Promise.resolve();
      expect(acknowledged).toBe(true);
      expect(presentation).toBeUndefined();
      expect(await ticket.ack).toEqual({ seq: 1, status: 'applied' });
      expect(h.session.state).toEqual({ ...initial(), version: 1, throughSeq: 1,
        selection: 1, editState: { canUndo: true } });
      h.renders[0]!.confirm(1);
      expect(await ticket.presented).toBe('presented');
      expect(await h.session.whenPresented(1)).toBe('presented');
      expect(h.session.state.presentedThroughSeq).toBe(1);
      await h.session.dispose();
    });

    it('keeps every synchronous apply and monotonic seq without coalescing or dependency changes', async () => {
      const h = harness(backend);
      const first = h.session.submit('refuse', 'x', { coalesceKey: 'typing' });
      const options: SubmitOptions<string> = { after: first, coalesceKey: 'typing', fence: true };
      const second = h.session.submit('insert', 'b', options);
      const third = h.session.submit('noop', '', { coalesceKey: 'typing' });
      expect(h.applied.map((entry) => entry.kind)).toEqual(['refuse', 'insert', 'noop']);
      expect(h.applied[1]?.options).toBe(options);
      expect(h.text).toBe('b');
      expect([first.seq, second.seq, third.seq]).toEqual([1, 2, 3]);
      expect(new Set([first.id, second.id, third.id]).size).toBe(3);
      expect(await first.ack).toEqual({ seq: 1, status: 'refused', reason: 'locked',
        committed: false, affectedSeqs: [7] });
      expect(await first.presented).toBe('dropped');
      expect(await h.session.whenPresented(1)).toBe('dropped');
      expect(await second.ack).toEqual({ seq: 2, status: 'applied' });
      expect(await third.ack).toEqual({ seq: 3, status: 'noop' });
      expect(h.session.state.version).toBe(1);
      expect(h.session.state.throughSeq).toBe(3);
      h.renders[h.renders.length - 1]!.confirm(3);
      expect(await second.presented).toBe('presented');
      expect(await third.presented).toBe('presented');
      expect(await h.session.whenPresented(1)).toBe('dropped');
      await h.session.dispose();
    });

    it('publishes state, unsubscribes, and flushes immediately before rendering', async () => {
      const h = harness(backend);
      const states: State[] = [];
      const off = h.session.subscribe((state) => { states.push(state); });
      const ticket = h.session.submit('insert', 'a');
      expect(states).toEqual([h.session.state]);
      let flushed = false;
      void h.session.flush().then(() => { flushed = true; });
      await Promise.resolve();
      expect(flushed).toBe(true);
      const oldState = states[0]!;
      h.renders[0]!.confirm(1);
      expect(states.map((state) => state.presentedThroughSeq)).toEqual([0, 1]);
      expect(oldState.presentedThroughSeq).toBe(0);
      off();
      h.session.submit('insert', 'b');
      expect(states).toHaveLength(2);
      expect(await ticket.presented).toBe('presented');
      await h.session.dispose();
    });

    it('accepts synchronous render confirmation and ignores stale or future coverage', async () => {
      const h = harness(backend);
      const first = h.session.submit('insert', 'a');
      const second = h.session.submit('insert', 'b');
      h.renders[1]!.confirm(2);
      h.renders[0]!.confirm(1);
      h.renders[0]!.confirm(200);
      expect(h.session.state.presentedThroughSeq).toBe(2);
      expect(await first.presented).toBe('presented');
      expect(await second.presented).toBe('presented');
      await h.session.dispose();
      const session = createLegacyEditSession({
        backend, initial: initial(),
        apply: () => ({ outcome: { status: 'noop' }, state: initial() }),
        render: (seq, confirm) => { confirm(seq); },
      });
      const ticket = session.submit('noop', undefined);
      expect(session.state.presentedThroughSeq).toBe(1);
      expect(await ticket.presented).toBe('presented');
      await session.dispose();
    });

    it('starts after the initial watermark and preserves synchronous apply throws', async () => {
      const h = harness(backend);
      expect(() => h.session.submit('throw', 'x')).toThrow('apply failed');
      expect(h.text).toBe('');
      expect(h.renders).toHaveLength(0);
      const ticket = h.session.submit('insert', 'a');
      expect(ticket.seq).toBe(2);
      await h.session.dispose();
      const session = createLegacyEditSession({
        backend, initial: { ...initial(), throughSeq: 12, presentedThroughSeq: 12 },
        apply: () => ({ outcome: { status: 'noop' }, state: initial() }),
        render: () => {},
      });
      expect(session.submit('noop', undefined).seq).toBe(13);
      await session.dispose();
    });

    it('disposes once, drops waiters, and prevents later apply or render callbacks', async () => {
      const h = harness(backend);
      const ticket = h.session.submit('insert', 'a');
      const waiting = h.session.whenPresented(1);
      const states: State[] = [];
      h.session.subscribe((state) => { states.push(state); });
      const disposal = h.session.dispose();
      expect(h.session.dispose()).toBe(disposal);
      await disposal;
      expect(h.disposals).toBe(1);
      expect(await ticket.ack).toEqual({ seq: 1, status: 'applied' });
      expect(await ticket.presented).toBe('dropped');
      expect(await waiting).toBe('dropped');
      const late = h.session.submit('insert', 'b');
      expect(late.seq).toBeUndefined();
      expect(await late.ack).toEqual({ status: 'cancelled', reason: 'disposed' });
      expect(await late.presented).toBe('dropped');
      h.renders[0]!.confirm(1);
      expect(h.text).toBe('a');
      expect(states).toHaveLength(0);
      expect(await h.session.whenPresented(2)).toBe('dropped');
    });

    it('bounds outstanding presentation promises and retires completed waiters', async () => {
      const h = harness(backend);
      const first = h.session.submit('insert', 'a');
      for (let index = 0; index < 256; index += 1) h.session.submit('noop', '');
      expect(await first.presented).toBe('dropped');
      const last = h.renders[h.renders.length - 1]!;
      last.confirm(last.seq);
      expect(await h.session.whenPresented(last.seq)).toBe('presented');
      expect(h.applied).toHaveLength(257);
      await h.session.dispose();
    });

    it('drops all waiters for an evicted seq and keeps waiter eviction bounded', async () => {
      const h = harness(backend);
      const ticket = h.session.submit('insert', 'a');
      const waiting = Array.from({ length: 256 }, () => h.session.whenPresented(1));
      expect(await ticket.presented).toBe('dropped');
      expect(await Promise.all(waiting)).toEqual(Array.from({ length: 256 }, () => 'dropped'));
      h.renders[0]!.confirm(1);
      expect(await h.session.whenPresented(1)).toBe('dropped');
      await h.session.dispose();
    });

    it('isolates throwing listeners and stops stale notifications after reentrant submit', async () => {
      const h = harness(backend);
      const seen: number[] = [];
      h.session.subscribe(() => { throw new Error('listener'); });
      h.session.subscribe((state) => {
        if (state.throughSeq === 1) h.session.submit('insert', 'b');
      });
      h.session.subscribe((state) => { seen.push(state.throughSeq); });
      h.session.submit('insert', 'a');
      expect(h.text).toBe('ab');
      expect(seen).toEqual([2]);
      expect(h.renders.map((render) => render.seq)).toEqual([2, 1]);
      await h.session.dispose();
    });
  });
}
