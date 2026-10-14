import { describe, expect, it, spyOn } from 'bun:test';
import { SessionFailure } from '../types';
import type { SessionTransport } from '../transport';
import { BASE_REFUSALS, type ArgsOf, type CommandSession, type CommandTable, type Dependency,
  type PayloadOf, type QueryReply, type QueryTable } from '../commands/spec';

/** @internal */
export type ContractPair = () => { client: SessionTransport; host: SessionTransport };
/** @internal */
export type ContractClock = ReturnType<typeof contractClock>;
type Sample<C extends CommandTable> = {
  [K in keyof C & string]: { kind: K; payload: PayloadOf<C[K]> }
}[keyof C & string];
/** @internal */
export interface ContractOpenOptions {
  pair: ContractPair;
  clock: ContractClock;
  recovery?: boolean;
  silenceMs?: number;
}
/** @internal */
export interface SessionContractHandle<C extends CommandTable, Q extends QueryTable> {
  session: CommandSession<C, unknown, unknown, never, Q>;
  crash(): void;
  poison(): void;
  bumpLayout(): void;
  emitAck(): void;
  failRecovery(): void;
  breakQueryPost(): void;
  refuseNext(code: string): void;
  refuseNextCommand(code: string): void;
  malformedReply(): void;
  silence(): void;
  queryReplies(): { reply: Promise<QueryReply>; release(): void };
  park<K extends keyof Q & string>(kind: K, args: ArgsOf<Q[K]>): Promise<QueryReply>;
  readonly queryCount: number;
  readonly transferredBuffers: readonly ArrayBuffer[];
}
/** @internal Exhaustive samples make the same contracts reusable by format adapters. */
export interface SessionContractFixture<C extends CommandTable, Q extends QueryTable> {
  commands: C;
  queries: Q;
  samples: { [K in keyof C]: PayloadOf<C[K]> };
  args: { [K in keyof Q]: ArgsOf<Q[K]> };
  mode: { on: Sample<C>; off: Sample<C> };
  moves: { doc: Sample<C>; selection: Sample<C> };
  open(options: ContractOpenOptions): SessionContractHandle<C, Q> | Promise<SessionContractHandle<C, Q>>;
}

function contractClock() {
  const tasks = new Map<number, { at: number; run(): void }>();
  let nextId = 0;
  let now = 0;
  const date = spyOn(Date, 'now').mockImplementation(() => now);
  const schedule = (...args: Parameters<typeof setTimeout>): ReturnType<typeof setTimeout> => {
    const [callback, delay, ...values] = args;
    if (typeof callback !== 'function') throw new Error('Expected a timer callback');
    const id = ++nextId;
    tasks.set(id, { at: now + (delay ?? 0), run: () => { callback(...values); } });
    return id as unknown as ReturnType<typeof setTimeout>;
  };
  const timeout = spyOn(globalThis, 'setTimeout').mockImplementation(schedule as typeof setTimeout);
  const clear = spyOn(globalThis, 'clearTimeout').mockImplementation((id) => { tasks.delete(id as unknown as number); });
  async function turn() {
    for (const [id, task] of [...tasks]) {
      if (task.at > now || !tasks.has(id)) continue;
      tasks.delete(id); task.run();
    }
    for (let index = 0; index < 32; index += 1) await Promise.resolve();
  }
  async function until(predicate: () => boolean) {
    for (let index = 0; index < 200; index += 1) { await turn(); if (predicate()) return; }
    throw new Error('Controlled event queue did not reach the expected state');
  }
  async function drain() {
    await turn();
    for (let index = 0; index < 200; index += 1) {
      if (![...tasks.values()].some((task) => task.at <= now)) return;
      await turn();
    }
    throw new Error('Controlled event queue did not drain');
  }
  return {
    now: () => now, enqueue: (callback: () => void) => { setTimeout(callback, 0); },
    timer(callback: () => void, ms: number) {
      const id = setTimeout(callback, ms); return () => { clearTimeout(id); };
    },
    drain, until,
    async advance(ms: number) { now += ms; await drain(); },
    async settle<T>(promise: Promise<T>): Promise<T> {
      let settled = false;
      const observed = promise.then((value) => { settled = true; return value; },
        (error: unknown) => { settled = true; throw error; });
      void observed.catch(() => {});
      await until(() => settled);
      return observed;
    },
    restore() { tasks.clear(); clear.mockRestore(); timeout.mockRestore(); date.mockRestore(); },
  };
}

/** @internal Registers the read, admission, transport, and recovery contracts. */
export function describeSessionContract<C extends CommandTable, Q extends QueryTable>(
  fixture: SessionContractFixture<C, Q>, options: { pair: ContractPair }
) {
  describe('session contract', () => {
    type Handle = SessionContractHandle<C, Q>;
    function test(name: string, run: (handle: Handle, clock: ContractClock) => Promise<void>, open: Partial<ContractOpenOptions> = {}) {
      it(name, async () => {
        const clock = contractClock();
        let handle: Handle | undefined;
        try {
          handle = await clock.settle(Promise.resolve(fixture.open({ pair: options.pair, clock, ...open })));
          await run(handle, clock);
        } finally {
          try { if (handle) await clock.settle(handle.session.dispose()); }
          finally { clock.restore(); }
        }
      });
    }
    const kinds = Object.keys(fixture.commands) as Array<keyof C & string>;
    const queries = Object.keys(fixture.queries) as Array<keyof Q & string>;
    const firstQuery = queries[0]!;
    function read(handle: Handle) { return handle.session.read(firstQuery, fixture.args[firstQuery]); }
    function submit(handle: Handle, sample: Sample<C>) { return handle.session.submit(sample.kind, sample.payload); }

    for (const kind of kinds) test(`admits ${kind} according to its read-only policy`, async (h, clock) => {
      await clock.settle(submit(h, fixture.mode.on).ack);
      const outcome = await clock.settle(h.session.submit(kind, fixture.samples[kind]).ack);
      if (fixture.commands[kind]!.admission === 'local-write') expect(outcome).toMatchObject({ status: 'refused', reason: 'read-only' });
      else expect(outcome.reason).not.toBe('read-only');
      await clock.settle(submit(h, fixture.mode.off).ack);
      const allowed = await clock.settle(h.session.submit(kind, fixture.samples[kind]).ack);
      expect(allowed.reason).not.toBe('read-only');
      if (allowed.status === 'refused') expect([...BASE_REFUSALS, ...fixture.commands[kind]!.refusals]).toContain(allowed.reason!);
    });

    for (const kind of kinds) test(`${kind} settles declared refusals and fails on undeclared codes`, async (h, clock) => {
      await clock.settle(submit(h, fixture.mode.off).ack);
      const code = fixture.commands[kind]!.refusals[0] ?? 'busy';
      h.refuseNextCommand(code);
      expect(await clock.settle(h.session.submit(kind, fixture.samples[kind]).ack)).toMatchObject({ status: 'refused', reason: code });
      expect(h.session.stage).toBe('ready');
      h.refuseNextCommand('undeclared');
      expect(await clock.settle(h.session.submit(kind, fixture.samples[kind]).ack)).toMatchObject({ status: 'refused', reason: 'unknown-outcome' });
      expect(h.session.stage).toBe('failed');
      await expect(clock.settle(read(h))).rejects.toMatchObject({ code: 'crash', message: 'Undeclared refusal: undeclared' });
    }, { recovery: false });

    for (const kind of queries) {
      test(`${kind} observes prior submit and holder commit`, async (h, clock) => {
        const ticket = submit(h, fixture.moves.doc);
        const answer = await clock.settle(h.session.read(kind, fixture.args[kind]));
        expect(answer.ok).toBe(true);
        if (!answer.ok) throw new Error('Expected an answer');
        expect(answer.at.throughSeq).toBeGreaterThanOrEqual(ticket.seq!);
        const hold = h.session.hold('draft');
        const before = h.queryCount;
        const held = h.session.read(kind, fixture.args[kind]);
        await clock.drain(); expect(h.queryCount).toBe(before);
        const commit = submit(h, fixture.moves.doc);
        hold.release();
        const committed = await clock.settle(held);
        expect(committed.ok).toBe(true);
        if (!committed.ok) throw new Error('Expected an answer');
        expect(committed.at.throughSeq).toBeGreaterThanOrEqual(commit.seq!);
        expect(committed.at.doc).toBe(h.session.state.version);
      });
      const stale = fixture.queries[kind]!.stale;
      if (stale !== 'answer-at-fence') for (const dependency of ['doc', 'selection', 'layout'] as const) {
        test(`${kind} checks ${dependency} supersession after its stamp`, async (h, clock) => {
          const gate = h.queryReplies();
          const pending = h.session.read(kind, fixture.args[kind]);
          expect((await clock.settle(gate.reply)).ok).toBe(true);
          if (dependency === 'layout') { h.bumpLayout(); h.emitAck(); await clock.drain(); }
          else await clock.settle(submit(h, fixture.moves[dependency]).ack);
          gate.release();
          const answer = await clock.settle(pending);
          if (stale.reject.includes(dependency as Dependency)) expect(answer).toEqual({ ok: false, refusal: { code: 'superseded' } });
          else expect(answer.ok).toBe(true);
        });
      }
      test(`${kind} accepts declared refusals and rejects undeclared codes`, async (h, clock) => {
        const code = fixture.queries[kind]!.refusals[0] ?? 'busy';
        h.refuseNext(code);
        const refused: QueryReply = await clock.settle(h.session.read(kind, fixture.args[kind]));
        expect(refused).toEqual({ ok: false, refusal: { code } });
        h.refuseNext('undeclared');
        await expect(clock.settle(h.session.read(kind, fixture.args[kind]))).rejects.toMatchObject({ code: 'crash', message: 'Undeclared refusal: undeclared' });
      }, { recovery: false });
    }

    test('returns busy at the hold age deadline without posting', async (h, clock) => {
      h.session.hold('composition'); await clock.advance(1500);
      const pending = read(h); await clock.advance(499); expect(h.queryCount).toBe(0);
      await clock.advance(1); expect(await clock.settle(pending)).toEqual({ ok: false, refusal: { code: 'busy' } });
      expect(h.queryCount).toBe(0);
    });
    test('releases a signal owner and supports within', async (h, clock) => {
      const controller = new AbortController();
      const owner = h.session.hold('draft', { signal: controller.signal });
      expect((await clock.settle(h.session.read(firstQuery, fixture.args[firstQuery], { within: owner }))).ok).toBe(true);
      const held = read(h); controller.abort();
      expect(owner.released).toBe(true); expect((await clock.settle(held)).ok).toBe(true);
    });
    test('settles both held and posted reads as disposed', async (h, clock) => {
      const owner = h.session.hold('gesture');
      const held = read(h); h.queryReplies();
      const posted = h.session.read(firstQuery, fixture.args[firstQuery], { within: owner });
      await clock.settle(h.session.dispose());
      expect(await clock.settle(held)).toEqual({ ok: false, refusal: { code: 'disposed' } });
      expect(await clock.settle(posted)).toEqual({ ok: false, refusal: { code: 'disposed' } });
      expect(await read(h)).toEqual({ ok: false, refusal: { code: 'disposed' } });
      expect(owner.released).toBe(true);
    });
    for (const posted of [false, true]) test(`aborts ${posted ? 'posted' : 'held'} reads`, async (h, clock) => {
      const controller = new AbortController();
      const owner = posted ? undefined : h.session.hold('draft');
      const gate = h.queryReplies();
      const pending = h.session.read(firstQuery, fixture.args[firstQuery], { signal: controller.signal });
      if (posted) await clock.settle(gate.reply);
      controller.abort(); owner?.release(); gate.release();
      expect(await clock.settle(pending)).toEqual({ ok: false, refusal: { code: 'superseded', detail: 'aborted' } });
      await clock.drain(); expect(h.session.stage).toBe('ready');
      expect(h.queryCount).toBe(posted ? 1 : 0);
    });
    test('reissues reads after recovery replay with a fresh incarnation', async (h, clock) => {
      await clock.settle(submit(h, fixture.moves.doc).ack);
      const gate = h.queryReplies(); const pending = read(h);
      const old = await clock.settle(gate.reply);
      h.crash(); gate.release();
      const reply = await clock.settle(pending);
      expect(reply.ok).toBe(true);
      if (!reply.ok || !old.ok) throw new Error('Expected answers');
      expect(reply.at.incarnation).toBe(old.at.incarnation + 1);
      expect(reply.at.doc).toBe(old.at.doc); expect(reply.at.throughSeq).toBe(old.at.throughSeq);
      expect(h.queryCount).toBe(2);
    });
    for (const [name, applied, posted] of [
      ['answers reads after recovery when the only sent command was refused', false, false],
      ['answers reads after recovery when the last sent command was refused', true, false],
      ['answers a reissued read after recovery when the last sent command was refused', true, true],
    ] as const) test(name, async (h, clock) => {
      if (applied) expect((await clock.settle(submit(h, fixture.moves.doc).ack)).status).toBe('applied');
      h.refuseNextCommand(fixture.commands[fixture.moves.doc.kind]!.refusals[0] ?? 'busy');
      expect((await clock.settle(submit(h, fixture.moves.doc).ack)).status).toBe('refused');
      const gate = posted ? h.queryReplies() : undefined;
      const early = posted ? read(h) : undefined;
      if (gate) expect((await clock.settle(gate.reply)).ok).toBe(true);
      h.crash(); gate?.release();
      expect(await clock.settle(early ?? read(h))).toMatchObject({ ok: true, at: { incarnation: 1 } });
      expect(h.session.stage).toBe('ready');
      const ticket = submit(h, fixture.moves.doc);
      const later = await clock.settle(read(h));
      expect(later.ok).toBe(true);
      if (later.ok) expect(later.at.throughSeq).toBeGreaterThanOrEqual(ticket.seq!);
    });
    test('keeps captured holds across recovery', async (h, clock) => {
      const owner = h.session.hold('draft'); const pending = read(h);
      h.crash(); await clock.drain(); expect(h.queryCount).toBe(0);
      expect(owner.released).toBe(false);
      submit(h, fixture.moves.doc); owner.release();
      expect((await clock.settle(pending)).ok).toBe(true);
    });
    test('reissues every waiting read after synchronous query transport failure', async (h, clock) => {
      submit(h, fixture.moves.doc); h.breakQueryPost();
      const a = read(h); const b = read(h);
      const replies = await clock.settle(Promise.all([a, b]));
      expect(replies.every((reply) => reply.ok)).toBe(true);
      for (const reply of replies) if (reply.ok) expect(reply.at.incarnation).toBe(1);
    });
    test('rejects a crash without recovery as SessionFailure', async (h, clock) => {
      h.queryReplies(); const pending = read(h); h.crash();
      await expect(clock.settle(pending)).rejects.toBeInstanceOf(SessionFailure);
      expect(h.session.stage).toBe('failed');
    }, { recovery: false });
    test('rejects failed recovery and releases its owners', async (h, clock) => {
      const owner = h.session.hold('mode'); const pending = read(h);
      h.failRecovery(); h.crash();
      await expect(clock.settle(pending)).rejects.toMatchObject({ code: 'message' });
      expect(owner.released).toBe(true);
    });
    test('poisons terminal query throws', async (h, clock) => {
      h.poison(); await expect(clock.settle(read(h))).rejects.toMatchObject({ code: 'trap' });
    }, { recovery: false });
    test('detects a silent query through the RPC watchdog', async (h, clock) => {
      h.silence(); const pending = read(h); await clock.drain(); await clock.advance(50);
      await expect(clock.settle(pending)).rejects.toMatchObject({ code: 'silence' });
    }, { recovery: false, silenceMs: 50 });
    test('fails a malformed query reply with message', async (h, clock) => {
      h.malformedReply(); await expect(clock.settle(read(h))).rejects.toMatchObject({ code: 'message' });
    });
    test('parks queries without blocking later worker edits', async (h, clock) => {
      const parked = h.park(firstQuery, fixture.args[firstQuery]);
      await clock.drain();
      const ticket = submit(h, fixture.moves.doc);
      const reply = await clock.settle(parked);
      expect(reply.ok).toBe(true);
      if (reply.ok) expect(reply.at.throughSeq).toBeGreaterThanOrEqual(ticket.seq!);
      expect((await clock.settle(ticket.ack)).status).toBe('applied');
    });
    test('publishes revise layout clocks through ack subscriptions', async (h, clock) => {
      let updates = 0;
      const off = h.session.subscribe(() => { updates += 1; });
      h.bumpLayout(); h.emitAck(); await clock.drain();
      expect(updates).toBe(1);
      const reply = await clock.settle(read(h));
      expect(reply.ok).toBe(true);
      if (reply.ok) expect(reply.at.layout).toBe(1);
      off();
    });
  });
}
