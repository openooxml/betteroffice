import { createSessionClient } from '../client';
import { createSessionHost } from '../host';
import { createCommandHost, commandSessionMethods } from '../commands/host';
import { createWorkerEditSession } from '../commands/session';
import { command, query, refuse, type CommandPayload, type QueryHandlers, type QueryMethods,
  type QueryReply, type RefusalResult } from '../commands/spec';
import type { CommandAck, CommandConnection, CommandEnvelope, CommandMethods } from '../commands/types';
import { isClientMessage, isHostMessage, transferable, type TransferResult } from '../protocol';
import type { SessionTransport } from '../transport';
import { SessionFailure } from '../types';
import type { ContractOpenOptions, SessionContractFixture } from './contract';

const commands = {
  insert: command<{ text: string }>()({ admission: 'local-write' }),
  select: command<{ position: number }>()({ admission: 'view' }),
  remote: command<{ text: string }>()({ admission: 'remote-update' }),
  mode: command<{ readOnly: boolean }>()({ admission: 'mode' }),
  guarded: command<{ blocked: boolean }>()({ admission: 'local-write', refusals: ['locked'] }),
};
const queries = {
  text: query<[], string>()({ stale: 'answer-at-fence', refusals: ['locked'] }),
  doc: query<[], string>()({ stale: { reject: ['doc'] }, refusals: ['locked'] }),
  selection: query<[], number>()({ stale: { reject: ['selection'] }, refusals: ['locked'] }),
  layout: query<[], number>()({ stale: { reject: ['layout'] }, refusals: ['locked'] }),
  draft: query<[], { text: string; selection: number; layout: number }>()({ stale: { reject: ['doc', 'selection', 'layout'] }, refusals: ['locked'] }),
  bytes: query<[], Uint8Array>()({ stale: 'answer-at-fence', refusals: ['locked'] }),
};
type Payload = CommandPayload<typeof commands>;
type Ack = CommandAck<number, boolean, never>;
type Methods = CommandMethods<Payload, number, boolean, never> & QueryMethods;
type Gate = { reply: Promise<QueryReply>; resolve(reply: QueryReply): void; send?: () => void };

function open(options: ContractOpenOptions) {
  let queryCount = 0;
  let failRecovery = false;
  const transferredBuffers: ArrayBuffer[] = [];
  function worker(incarnation: number, replayAppliedSeqs: readonly number[] = []) {
    const pair = options.pair();
    const faults = new Set<(failure: unknown) => void>();
    const queryIds = new Set<number>();
    let closed = false;
    let text = '';
    let version = 0;
    let selection = 0;
    let layout = 0;
    let readOnly = false;
    let poison = false;
    let breakPost = false;
    let malformed = false;
    let silence = false;
    let refusal: RefusalResult<'locked'> | undefined;
    let commandRefusal: string | undefined;
    let gate: Gate | undefined;
    let admittedSeq: number | undefined;
    pair.host.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'query') {
        queryIds.add(message.id); queryCount += 1;
      }
    });
    function crash() {
      const failure = new SessionFailure('crash', 'Simulated endpoint crash');
      for (const listener of [...faults]) listener(failure);
    }
    function endpoint(transport: SessionTransport, host = false): SessionTransport {
      return {
        post(message, transfer) {
          if (!host && breakPost && isClientMessage(message) && message.kind === 'call' && message.method === 'query') {
            breakPost = false; crash(); throw new SessionFailure('crash', 'Query post failed');
          }
          if (host && isHostMessage(message) && message.kind === 'reply' && message.ok && queryIds.delete(message.id)) {
            if (silence) { silence = false; return; }
            if (malformed) { malformed = false; message = { ...message, value: { ok: true } }; }
            if (gate) {
              const held = gate; gate = undefined;
              const saved = message;
              held.resolve((message as { value: QueryReply }).value);
              held.send = () => { if (!closed) transport.post(saved, transfer); };
              return;
            }
          }
          transport.post(message, transfer);
        },
        listen: (listener) => transport.listen(listener),
        onError(listener) {
          faults.add(listener); const off = transport.onError(listener);
          return () => { faults.delete(listener); off(); };
        },
        close() { closed = true; transport.close(); },
      };
    }
    function answer<R>(value: () => R | TransferResult<R>): R | TransferResult<R> | RefusalResult<'locked'> {
      if (poison) throw new SessionFailure('trap', 'Simulated query trap');
      if (refusal) { const result = refusal; refusal = undefined; return result; }
      return value();
    }
    const handlers = {
      text: () => answer(() => text), doc: () => answer(() => text),
      selection: () => answer(() => selection), layout: () => answer(() => layout),
      draft: () => answer(() => ({ text, selection, layout })),
      bytes: () => answer(() => {
        const bytes = new Uint8Array([text.length]);
        transferredBuffers.push(bytes.buffer);
        return transferable(bytes, [bytes.buffer]);
      }),
    } satisfies QueryHandlers<typeof queries>;
    const commandHost = createCommandHost<Payload, number, boolean, never>({
      session: 'reference', incarnation, epoch: 1, commands, queries: { specs: queries, handlers },
      readOnly: () => readOnly, layoutRevision: () => layout, replayAppliedSeqs, checkRefusals: true,
      getVersion: () => version,
      apply(envelope) {
        if (commandRefusal !== undefined) {
          const reason = commandRefusal; commandRefusal = undefined;
          return { status: 'refused', reason };
        }
        const payload = envelope.payload;
        if ('blocked' in payload && payload.blocked) return { status: 'refused', reason: 'locked' };
        if ('text' in payload) { text += payload.text; version += 1; return { status: 'applied' }; }
        if ('position' in payload) selection = payload.position;
        if ('readOnly' in payload) readOnly = payload.readOnly;
        return { status: 'applied', committed: false, selection, editState: readOnly };
      }, frame: () => undefined, snapshot: () => new Uint8Array(),
    });
    const methods = commandSessionMethods(commandHost);
    const host = createSessionHost<Methods, { ack: Ack }, null>(endpoint(pair.host, true), {
      ...methods, context: null, policies: { ...methods.policies,
        edit(envelope) {
          if (admittedSeq === envelope.seq) {
            admittedSeq = undefined;
            return { lane: 'input', mutates: true, reframes: true, reorderable: false };
          }
          return commandHost.policy(envelope);
        },
      },
    });
    const rpc = createSessionClient<Methods, { ack: Ack }>(endpoint(pair.client), {
      methods: { edit: true, checkpoint: true, query: true },
      now: options.clock.now, timer: options.clock.timer, silenceMs: options.silenceMs,
    });
    const initial = commandHost.revise({ selection, editState: readOnly })!;
    const connection: CommandConnection<Payload, number, boolean, never> = {
      initial, client: { ...rpc, get failure() { return rpc.failure; }, onAck: (listener) => rpc.on('ack', listener) },
    };
    return { connection, rpc, commandHost, crash,
      bumpLayout() { layout += 1; },
      emitAck() { const ack = commandHost.revise(); if (ack) host.emit('ack', ack); },
      poison() { poison = true; }, breakQueryPost() { breakPost = true; },
      refuseNext(code: string) { refusal = refuse(code) as RefusalResult<'locked'>; },
      refuseNextCommand(code: string) { commandRefusal = code; },
      malformedReply() { malformed = true; }, silence() { silence = true; },
      queryReplies() {
        let resolve!: (reply: QueryReply) => void;
        const reply = new Promise<QueryReply>((done) => { resolve = done; });
        const held: Gate = { reply, resolve }; gate = held;
        return { reply, release() { if (gate === held) gate = undefined; held.send?.(); } };
      },
      park(kind: string, args: unknown[]) {
        const seq = initial.throughSeq + 1;
        const envelope: CommandEnvelope<Payload> = { session: 'reference', incarnation, seq,
          basis: { epoch: 1, version }, kind: 'insert', payload: { text: 'parked' } };
        commandHost.policy(envelope); admittedSeq = seq;
        return rpc.call.query({ session: 'reference', incarnation, epoch: 1, afterSeq: seq }, kind, args);
      },
    };
  }
  let active = worker(0);
  const session = createWorkerEditSession<string, Payload, number, boolean, never, never, typeof queries>({
    source: '', connection: active.connection, queries,
    recovery: { async reopen(input) {
      if (failRecovery) throw new SessionFailure('message', 'Reopen failed');
      active = worker(input.incarnation, input.replayAppliedSeqs); return active.connection;
    } }, maxRecoveryAttempts: options.recovery === false ? 0 : 1,
    requestFullFrame: async () => { throw new Error('Unexpected frame request'); },
    checkpointPolicy: () => false, now: options.clock.now, timer: options.clock.timer, enqueue: options.clock.enqueue,
  });
  return { session,
    crash: () => active.crash(), poison: () => active.poison(),
    bumpLayout: () => active.bumpLayout(), emitAck: () => active.emitAck(),
    failRecovery() { failRecovery = true; },
    breakQueryPost: () => active.breakQueryPost(), refuseNext: (code: string) => active.refuseNext(code),
    refuseNextCommand: (code: string) => active.refuseNextCommand(code),
    malformedReply: () => active.malformedReply(), silence: () => active.silence(),
    queryReplies: () => active.queryReplies(), park: (kind: string, args: unknown[]) => active.park(kind, args),
    get queryCount() { return queryCount; }, transferredBuffers,
  };
}

/** @internal Minimal document implementing the reusable session contracts. */
export const referenceFormat = {
  commands, queries,
  samples: { insert: { text: 'a' }, select: { position: 1 }, remote: { text: 'r' }, mode: { readOnly: false }, guarded: { blocked: true } },
  args: { text: [], doc: [], selection: [], layout: [], draft: [], bytes: [] },
  mode: { on: { kind: 'mode', payload: { readOnly: true } }, off: { kind: 'mode', payload: { readOnly: false } } },
  moves: { doc: { kind: 'insert', payload: { text: 'a' } }, selection: { kind: 'select', payload: { position: 1 } } },
  open,
} satisfies SessionContractFixture<typeof commands, typeof queries>;
