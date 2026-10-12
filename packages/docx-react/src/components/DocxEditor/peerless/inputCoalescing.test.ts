import { expect, test } from 'bun:test';
import { createOutbox } from '../../../../../../shared/office-session/commands/outbox';
import type { CommandTicket } from '../../../../../../shared/office-session/commands/types';
import { coalesceInputDrafts, type PeerlessCommandDraft } from './inputCoalescing';

function ticket(id: number): CommandTicket {
  return { id, seq: undefined, ack: Promise.resolve({ status: 'noop' }), presented: Promise.resolve('presented') };
}

function draft(text: string, options: PeerlessCommandDraft['options'] = {}): PeerlessCommandDraft {
  const intent = { kind: 'insertText' as const, payload: { text } };
  return { ...intent, options: {
    coalesceKey: 'typing', ...options,
    echo: { intent, originalIntentCount: 1, contextFrameEpoch: 3 },
  } };
}

test('d2b_coalescing_preserves_aliases_and_fences', () => {
  const after = ticket(90);
  const a = draft('🦊', { after, group: 'text' });
  const b = draft('é', { after, group: 'text' });
  const merged = coalesceInputDrafts(a, b)!;
  expect(merged.payload).toEqual({ text: '🦊é' });
  expect(merged.options.after).toBe(after);
  expect(merged.options.echo?.originalIntentCount).toBe(2);
  const outbox = createOutbox({ coalesce: coalesceInputDrafts });
  const tickets = [ticket(1), ticket(2)];
  outbox.enqueue(a, tickets[0]!);
  const entry = outbox.enqueue(b, tickets[1]!);
  expect(entry.tickets).toEqual(tickets);
  expect(outbox.size).toBe(1);
  for (const options of [
    { after: ticket(91) }, { group: 'other' }, { fence: true },
    { lane: 'background' as const }, { expectVersion: true }, { coalesceKey: 'paste' },
  ]) expect(coalesceInputDrafts(a, draft('b', { after, group: 'text', ...options }))).toBeUndefined();
  for (const next of [
    { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } },
    { kind: 'toggleMark', payload: { mark: { type: 'bold' } } },
    { kind: 'replaceSelection', payload: { text: 'paste' } },
  ] satisfies Array<Pick<PeerlessCommandDraft, 'kind' | 'payload'>>) {
    expect(coalesceInputDrafts(a, { ...next, options: b.options })).toBeUndefined();
  }
  expect(coalesceInputDrafts(draft('a', { fence: true }), draft('b'))).toBeUndefined();
});
