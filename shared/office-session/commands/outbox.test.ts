import { expect, test } from 'bun:test';
import { createOutbox, fenceOutbox } from './outbox';
import type { CommandDraft, CommandTicket, SubmitOptions } from './types';

function ticket(id: number) {
  let seq: number | undefined;
  const value: CommandTicket = {
    id,
    get seq() { return seq; },
    ack: Promise.resolve({ status: 'noop' }),
    presented: Promise.resolve('presented'),
  };
  return { value, send(value: number) { seq = value; } };
}

function draft(payload: string, options: SubmitOptions<string> = {}): CommandDraft<string, string> {
  return { kind: 'insert', payload, options };
}

function coalesce(prev: CommandDraft<string, string>, next: CommandDraft<string, string>) {
  return draft(prev.payload + next.payload, {
    ...next.options,
    echo: (prev.options.echo ?? '') + (next.options.echo ?? ''),
  });
}

test('declines changed coalescer metadata and never dispatches a removed command again', () => {
  const dependency = ticket(90);
  dependency.send(90);
  const changes: SubmitOptions<string>[] = [
    { lane: 'collab' }, { after: dependency.value }, { fence: true },
  ];
  for (const changed of changes) {
    const outbox = createOutbox<string, string>({
      coalesce: (prev, next) => ({
        ...coalesce(prev, next), options: { ...next.options, ...changed },
      }),
    });
    outbox.enqueue(draft('a', { coalesceKey: 'text' }), ticket(1).value);
    outbox.enqueue(draft('b', { coalesceKey: 'text' }), ticket(2).value);
    outbox.enqueue(draft('c'), ticket(3).value);
    const dispatched = Array.from({ length: 4 }, () => outbox.takeNext(() => true));
    expect(dispatched[1]?.id).not.toBe(dispatched[0]?.id);
    expect(dispatched.map((entry) => entry?.draft.payload)).toEqual(['a', 'b', 'c', undefined]);
    expect(outbox.takeNext(() => true)).toBeUndefined();
    expect(outbox.size).toBe(0);
  }
});

test('coalesces adjacent unsent drafts and aliases tickets without consuming extra seqs', () => {
  const outbox = createOutbox({ coalesce });
  const first = ticket(1);
  const second = ticket(2);
  const third = ticket(3);
  outbox.enqueue(draft('a', { coalesceKey: 'text', echo: 'A' }), first.value);
  outbox.enqueue(draft('b', { coalesceKey: 'text', echo: 'B' }), second.value);
  const merged = outbox.enqueue(draft('c', { coalesceKey: 'text', echo: 'C' }), third.value);
  expect(outbox.size).toBe(1);
  expect(merged.id).toBe(1);
  expect(merged.tickets).toEqual([first.value, second.value, third.value]);
  expect(merged.draft.payload).toBe('abc');
  expect(merged.draft.options.echo).toBe('ABC');
  expect(merged.tickets.map((value) => value.seq)).toEqual([undefined, undefined, undefined]);
  expect(outbox.takeNext(() => true)).toBe(merged);
  expect(outbox.size).toBe(0);
  expect(first.value.seq).toBeUndefined();
});

test('blocks merging across fences, incompatible lanes, groups, and dependencies', () => {
  const dependencies = [ticket(90).value, ticket(91).value];
  const pairs: Array<[SubmitOptions<string>, SubmitOptions<string>]> = [
    [{ fence: true }, {}], [{}, { fence: true }],
    [{ lane: 'input' }, { lane: 'collab' }],
    [{ group: 'one' }, { group: 'two' }],
    [{ after: dependencies[0] }, { after: dependencies[1] }],
    [{}, { after: dependencies[0] }],
    [{ expectVersion: true }, { expectVersion: false }],
  ];
  for (const [prev, next] of pairs) {
    let calls = 0;
    const outbox = createOutbox<string, string>({
      coalesce(a, b) { calls += 1; return coalesce(a, b); },
    });
    outbox.enqueue(draft('a', { coalesceKey: 'text', ...prev }), ticket(1).value);
    outbox.enqueue(draft('b', { coalesceKey: 'text', ...next }), ticket(2).value);
    expect(outbox.size).toBe(2);
    expect(calls).toBe(0);
  }
  const outbox = createOutbox({ coalesce });
  outbox.enqueue(draft('a', { coalesceKey: 'text' }), ticket(1).value);
  outbox.enqueue(draft('fence', { fence: true }), ticket(2).value);
  outbox.enqueue(draft('b', { coalesceKey: 'text' }), ticket(3).value);
  expect(outbox.size).toBe(3);
});

function keepPrevious(prev: CommandDraft<string, string>, next: CommandDraft<string, string>) {
  return draft(prev.payload + next.payload, prev.options);
}

test('merges drafts chained on the previous queued entry under that entry\'s predecessor', () => {
  const outbox = createOutbox<string, string>({ coalesce: keepPrevious });
  const sent = ticket(1);
  outbox.enqueue(draft('a', { coalesceKey: 'text' }), sent.value);
  const inFlight = outbox.takeNext(() => true)!;
  sent.send(1);
  const chain = [2, 3, 4].map((id) => ticket(id).value);
  const entries = chain.map((value, index) => outbox.enqueue(
    draft('bcd'[index]!, { coalesceKey: 'text', after: index ? chain[index - 1] : sent.value }), value));
  expect(inFlight.tickets).toEqual([sent.value]);
  expect(inFlight.draft.payload).toBe('a');
  expect(outbox.size).toBe(1);
  const merged = entries[2]!;
  expect(merged.id).toBe(2);
  expect(merged.tickets).toEqual(chain);
  expect(merged.draft.payload).toBe('bcd');
  expect(merged.draft.options.after).toBe(sent.value);
  expect(outbox.takeNext(() => true)).toBe(merged);
  expect(outbox.size).toBe(0);
});

test('keeps chained drafts apart across fences, lanes, groups, versions, and older entries', () => {
  const changes: SubmitOptions<string>[] = [
    { fence: true }, { lane: 'collab' }, { group: 'other' }, { expectVersion: true }, { coalesceKey: 'other' },
  ];
  for (const changed of changes) {
    const outbox = createOutbox<string, string>({ coalesce: keepPrevious });
    const first = ticket(1).value;
    outbox.enqueue(draft('a', { coalesceKey: 'text' }), first);
    outbox.enqueue(draft('b', { coalesceKey: 'text', after: first, ...changed }), ticket(2).value);
    expect(outbox.size).toBe(2);
  }
  const fenced = createOutbox<string, string>({ coalesce: keepPrevious });
  const first = ticket(1).value;
  fenced.enqueue(draft('a', { coalesceKey: 'text' }), first);
  fenceOutbox(fenced);
  fenced.enqueue(draft('b', { coalesceKey: 'text', after: first }), ticket(2).value);
  expect(fenced.size).toBe(2);
  const older = createOutbox<string, string>({ coalesce: keepPrevious });
  const root = ticket(1).value;
  older.enqueue({ ...draft('a', { coalesceKey: 'text' }), kind: 'delete' }, root);
  older.enqueue(draft('b', { coalesceKey: 'text' }), ticket(2).value);
  older.enqueue(draft('c', { coalesceKey: 'text', after: root }), ticket(3).value);
  expect(older.entries.map((entry) => entry.draft.payload)).toEqual(['a', 'b', 'c']);
});

test('merges compatible defaults but respects declined hooks, kinds, keys, and adjacency', () => {
  const outbox = createOutbox({ coalesce });
  const dependency = ticket(90).value;
  outbox.enqueue(draft('a', { coalesceKey: '', after: dependency }), ticket(1).value);
  outbox.enqueue(draft('b', {
    coalesceKey: '', after: dependency, lane: 'input', expectVersion: false,
  }), ticket(2).value);
  expect(outbox.size).toBe(1);
  outbox.enqueue({ ...draft('c', { coalesceKey: '' }), kind: 'delete' }, ticket(3).value);
  outbox.enqueue(draft('d', { coalesceKey: 'other' }), ticket(4).value);
  outbox.enqueue(draft('e'), ticket(5).value);
  outbox.enqueue(draft('f', { coalesceKey: '' }), ticket(6).value);
  expect(outbox.size).toBe(5);
  const declined = createOutbox<string, string>({ coalesce: () => undefined });
  declined.enqueue(draft('a', { coalesceKey: 'text' }), ticket(1).value);
  declined.enqueue(draft('b', { coalesceKey: 'text' }), ticket(2).value);
  expect(declined.size).toBe(2);
});

test('prioritizes input/collab/background while sending dependencies first', () => {
  const outbox = createOutbox<string, string>({});
  const dependency = ticket(1);
  const independent = ticket(2);
  const dependent = ticket(3);
  outbox.enqueue(draft('dependency', { lane: 'background' }), dependency.value);
  outbox.enqueue(draft('remote', { lane: 'collab' }), independent.value);
  outbox.enqueue(draft('dependent', { after: dependency.value }), dependent.value);
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('remote');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('dependency');
  expect(outbox.takeNext(() => true)).toBeUndefined();
  dependency.send(1);
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('dependent');

  outbox.enqueue(draft('background', { lane: 'background' }), ticket(4).value);
  outbox.enqueue(draft('collab', { lane: 'collab' }), ticket(5).value);
  outbox.enqueue(draft('input'), ticket(6).value);
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('input');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('collab');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('background');
});

test('does not dispatch across either side of a fence', () => {
  const outbox = createOutbox<string, string>({});
  outbox.enqueue(draft('before', { lane: 'background' }), ticket(1).value);
  outbox.enqueue(draft('fence', { lane: 'background', fence: true }), ticket(2).value);
  outbox.enqueue(draft('after'), ticket(3).value);
  expect(outbox.takeNext((entry) => entry.id !== 1)).toBeUndefined();
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('before');
  expect(outbox.takeNext((entry) => entry.id !== 2)).toBeUndefined();
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('fence');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('after');
});

test('serves eligible collab after eight inputs under sustained input load', () => {
  const outbox = createOutbox<string, string>({});
  outbox.enqueue(draft('collab', { lane: 'collab' }), ticket(1).value);
  for (let id = 2; id <= 10; id += 1) outbox.enqueue(draft(`input:${id}`), ticket(id).value);
  for (let count = 0; count < 8; count += 1) {
    expect(outbox.takeNext(() => true)?.draft.options.lane ?? 'input').toBe('input');
  }
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('collab');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('input:10');
});

test('fairness counts dispatched inputs, skips ineligible collab, and resets after clear', () => {
  const outbox = createOutbox<string, string>({ inputBurstCap: 2 });
  outbox.enqueue(draft('blocked', { lane: 'collab' }), ticket(1).value);
  outbox.enqueue(draft('a'), ticket(2).value);
  outbox.enqueue(draft('b'), ticket(3).value);
  expect(outbox.takeNext(() => false)).toBeUndefined();
  expect(outbox.takeNext((entry) => entry.id !== 1)?.draft.payload).toBe('a');
  outbox.enqueue(draft('eligible', { lane: 'collab' }), ticket(4).value);
  expect(outbox.takeNext((entry) => entry.id !== 1)?.draft.payload).toBe('b');
  outbox.enqueue(draft('c'), ticket(5).value);
  expect(outbox.takeNext((entry) => entry.id !== 1)?.draft.payload).toBe('eligible');
  expect(outbox.clear().map((entry) => entry.draft.payload)).toEqual(['blocked', 'c']);
  outbox.enqueue(draft('collab', { lane: 'collab' }), ticket(6).value);
  outbox.enqueue(draft('input'), ticket(7).value);
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('input');
});

test('removes unsent descendants through every alias and reports them in submission order', () => {
  const outbox = createOutbox({ coalesce });
  const first = ticket(1).value;
  const alias = ticket(2).value;
  const child = ticket(3).value;
  outbox.enqueue(draft('a', { coalesceKey: 'text' }), first);
  outbox.enqueue(draft('b', { coalesceKey: 'text' }), alias);
  outbox.enqueue(draft('child', { after: alias }), child);
  outbox.enqueue(draft('independent'), ticket(4).value);
  outbox.enqueue(draft('grandchild', { after: child }), ticket(5).value);
  const removed = outbox.remove((entry) => entry.id === 1);
  expect(removed.map((entry) => entry.id)).toEqual([1, 3, 5]);
  expect(removed[0]!.tickets).toEqual([first, alias]);
  expect(outbox.entries.map((entry) => entry.id)).toEqual([4]);
  expect(removed.flatMap((entry) => entry.tickets).every((value) => value.seq === undefined)).toBe(true);
  expect(outbox.remove(() => false)).toEqual([]);
  expect(outbox.clear().map((entry) => entry.id)).toEqual([4]);
  expect(outbox.size).toBe(0);
});

test('keeps alias dependents queued until their merged dependency has left and received a seq', () => {
  const outbox = createOutbox({ coalesce });
  const first = ticket(1);
  const alias = ticket(2);
  outbox.enqueue(draft('a', { lane: 'background', coalesceKey: 'text' }), first.value);
  outbox.enqueue(draft('b', { lane: 'background', coalesceKey: 'text' }), alias.value);
  outbox.enqueue(draft('dependent', { after: alias.value }), ticket(3).value);
  alias.send(1);
  expect(outbox.takeNext((entry) => entry.id !== 1)).toBeUndefined();
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('ab');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('dependent');
});

test('keeps collab fairness inside the current fence segment', () => {
  const outbox = createOutbox<string, string>({ inputBurstCap: 1 });
  outbox.enqueue(draft('first'), ticket(1).value);
  outbox.enqueue(draft('second'), ticket(2).value);
  outbox.enqueue(draft('fence', { lane: 'background', fence: true }), ticket(3).value);
  outbox.enqueue(draft('remote', { lane: 'collab' }), ticket(4).value);
  outbox.enqueue(draft('last'), ticket(5).value);
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('first');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('second');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('fence');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('last');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('remote');
});

test('appends coalesced aliases to the same group', () => {
  const outbox = createOutbox({ coalesce });
  const first = outbox.enqueue(draft('a', { coalesceKey: 'text' }), ticket(1).value);
  const aliases = first.tickets;
  for (let id = 2; id <= 128; id += 1) {
    const merged = outbox.enqueue(draft('a', { coalesceKey: 'text' }), ticket(id).value);
    expect(merged.tickets).toBe(aliases);
  }
  expect(aliases).toHaveLength(128);
});

test('indexes blocked dependents without scanning queued ticket identities', () => {
  const outbox = createOutbox<string, string>({});
  let reads = 0;
  function counted(id: number): CommandTicket {
    const value = ticket(id).value;
    return { ...value, get id() { reads += 1; return id; } };
  }
  for (let id = 1; id <= 128; id += 1) outbox.enqueue(draft('remote', { lane: 'collab' }), counted(id));
  const dependency = counted(129);
  outbox.enqueue(draft('dependency', { lane: 'background' }), dependency);
  for (let id = 130; id <= 257; id += 1) outbox.enqueue(draft('dependent', { after: dependency }), counted(id));
  reads = 0;
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('remote');
  expect(reads).toBeLessThan(16);
});

test('dispatches across empty and repeated submission barriers without consuming tickets', () => {
  const outbox = createOutbox<string, string>({});
  fenceOutbox(outbox);
  outbox.enqueue(draft('before', { lane: 'background' }), ticket(1).value);
  fenceOutbox(outbox);
  fenceOutbox(outbox);
  outbox.enqueue(draft('after'), ticket(2).value);
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('before');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('after');
  expect(outbox.size).toBe(0);
});

test('preserves lane submission order when a queued dependency becomes eligible', () => {
  const outbox = createOutbox<string, string>({});
  const first = ticket(1);
  outbox.enqueue(draft('first'), first.value);
  outbox.enqueue(draft('child', { after: first.value }), ticket(2).value);
  outbox.enqueue(draft('later'), ticket(3).value);
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('first');
  first.send(1);
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('child');
  expect(outbox.takeNext(() => true)?.draft.payload).toBe('later');
});

test('keeps lane order across readiness index blocks and empty fence segments', () => {
  const outbox = createOutbox<string, string>({});
  for (let id = 1; id <= 1100; id += 1) {
    outbox.enqueue(draft(String(id), { lane: 'background', fence: id % 32 === 0 }), ticket(id).value);
  }
  outbox.remove((entry) => entry.id % 2 === 0);
  for (let id = 1; id <= 1100; id += 2) expect(outbox.takeNext(() => true)?.id).toBe(id);
  expect(outbox.takeNext(() => true)).toBeUndefined();
  expect(outbox.size).toBe(0);
});
