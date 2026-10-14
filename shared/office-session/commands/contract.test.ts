import { describe, expect, it } from 'bun:test';
import { createSessionClient } from '../client';
import { createSessionHost } from '../host';
import { createScopeTransport, type SessionScope } from '../transport';
import { createInProcessPair } from '../testing/inProcessTransport';
import { describeSessionContract } from '../testing/contract';
import { referenceFormat } from '../testing/referenceFormat';
import { commandSessionMethods, createCommandHost } from './host';
import { command, query, type QueryMethods } from './spec';
import type { CommandMethods } from './types';

describe('transferring channel', () => {
  describeSessionContract(referenceFormat, { pair: createInProcessPair });
});
describe('copying channel', () => {
  describeSessionContract(referenceFormat, { pair: () => createInProcessPair({ transfer: false }) });
});

it('starts a scope after subscribing and copies transferables when requested', async () => {
  let listener: EventListener | undefined;
  const scope: SessionScope = {
    postMessage() {},
    addEventListener(type, callback) { if (type === 'message') listener = callback; },
    removeEventListener() {},
    start() { expect(listener).toBeDefined(); listener!(new MessageEvent('message', { data: 'started' })); },
  };
  const transport = createScopeTransport(scope);
  const messages: unknown[] = [];
  transport.listen((value) => { messages.push(value); });
  expect(messages).toEqual(['started']);
  transport.close();
  const pair = createInProcessPair({ transfer: false });
  try {
    const received = new Promise<ArrayBuffer>((resolve) => pair.host.listen((value) => resolve(value as ArrayBuffer)));
    const bytes = new Uint8Array([7]);
    pair.client.post(bytes.buffer, [bytes.buffer]);
    expect(bytes.byteLength).toBe(1);
    expect(new Uint8Array(await received)).toEqual(bytes);
  } finally { pair.client.close(); }
});

it('runs an edit and fenced query between MessageChannel endpoints', async () => {
  const channel = new MessageChannel();
  let text = '';
  let version = 0;
  const commands = { insert: command<string>()({ admission: 'local-write' }) };
  const specs = { text: query<[], string>()({ stale: 'answer-at-fence' }) };
  const worker = createCommandHost<string, never, never, never>({
    session: 'port', incarnation: 0, epoch: 1, commands,
    queries: { specs, handlers: { text: () => text } },
    getVersion: () => version,
    apply(envelope) { text += envelope.payload; version += 1; return { status: 'applied' }; },
    frame: () => undefined, snapshot: () => new Uint8Array(),
  });
  type Methods = CommandMethods<string, never, never, never> & QueryMethods;
  createSessionHost<Methods, {}, null>(createScopeTransport(channel.port1), {
    ...commandSessionMethods(worker), context: null,
  });
  const client = createSessionClient<Methods, {}>(createScopeTransport(channel.port2), {
    methods: { edit: true, checkpoint: true, query: true },
  });
  try {
    const edit = client.call.edit({ session: 'port', incarnation: 0, seq: 1,
      basis: { epoch: 1, version: 0 }, kind: 'insert', payload: 'hello' });
    const read = client.call.query({ session: 'port', incarnation: 0, epoch: 1, afterSeq: 1 }, 'text', []);
    expect((await edit).outcomes).toEqual([{ seq: 1, status: 'applied' }]);
    expect(await read).toMatchObject({ ok: true, value: 'hello', at: { throughSeq: 1, doc: 1 } });
  } finally {
    await client.dispose();
    channel.port1.close(); channel.port2.close();
  }
});
