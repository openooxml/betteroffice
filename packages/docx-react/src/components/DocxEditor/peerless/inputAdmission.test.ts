import { expect, spyOn, test } from 'bun:test';
import type { DocxCommandAck, DocxCommandPayload, DocxCommandSelection, DocxWorkerEditSession } from '@betteroffice/docx/peerless';
import type { CommandTicket, SubmitOptions, TicketOutcome } from '../../../../../../shared/office-session/commands/types';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from './echoPainter';
import type { PeerlessViewSnapshot, PeerlessViewStore } from './viewStore';
import { admitInputIntents, admitIntent, createInputAdmission, submitFormattingIntent, type PeerlessFormattingIntent } from './inputAdmission';
import { normalizeInputEvent, type PeerlessInputIntent } from './inputEvents';
import { createPeerlessCommandBinding } from './commandBinding';
import { beginComposition, endComposition, updateComposition } from './composition';

test('widget mutations preserve captured payloads behind pending text and composition', () => {
  for (const kind of ['setContentControlValue', 'addRepeatingItem', 'removeRepeatingItem'] as const) {
    const { admission, submitted } = lane();
    admitIntent(admission, { kind: 'insertText', payload: { text: 'pending' } });
    admission.setSuggesting?.('Writer');
    const payload = { handle: 'physical-control', expectVersion: 'native-token',
      ...(kind === 'setContentControlValue' ? { value: { kind: 'checkbox' as const, checked: true } } : {}) };
    const intent: PeerlessInputIntent = kind === 'setContentControlValue'
      ? { kind, payload: { ...payload, value: { kind: 'checkbox', checked: true } } }
      : { kind, payload };
    expect(admitIntent(admission, intent, { coalesceKey: 'widget' }).status).toBe('submitted');
    expect(submitted[1]!.payload).toEqual(payload);
    expect(submitted[1]!.options).toMatchObject({ after: submitted[0]!.ticket, fence: true });
    expect(submitted[1]!.options.coalesceKey).toBeUndefined();
    expect(submitted[1]!.options.echo).toBeUndefined();
    const reservation = admission.reserveComposition();
    expect(admitIntent(admission, intent).status).toBe('held');
    expect(submitted).toHaveLength(2);
    admission.settleComposition(reservation.slot, { text: 'composed', target: reservation.target });
    expect(submitted.map(entry => entry.kind)).toEqual(['insertText', kind, 'replaceSelection', kind]);
    expect(submitted[3]!.options).toMatchObject({ after: submitted[2]!.ticket, fence: true });
    expect(submitted[3]!.payload).toEqual(payload);
    admission.dispose();
  }
});

test('render settings held by composition report the submitted ticket without author metadata', () => {
  const { admission, submitted } = lane();
  admission.setSuggesting?.('Writer');
  const reservation = admission.reserveComposition();
  const tickets: CommandTicket[] = [];
  const result = admitIntent(admission, { kind: 'setRenderEnvironment', payload: { themeColors: {}, showHiddenText: true } },
    { fence: true, onSubmitted: (ticket) => tickets.push(ticket) });
  expect(result.status).toBe('held');
  expect(tickets).toHaveLength(0);
  admission.settleComposition(reservation.slot, { text: 'text', target: reservation.target });
  expect(tickets).toEqual([submitted[1]!.ticket]);
  expect(submitted[1]!.options.after).toBe(submitted[0]!.ticket);
  expect(submitted[1]!.payload).toEqual({ themeColors: {}, showHiddenText: true });
  expect(submitted[1]!.options).not.toHaveProperty('onSubmitted');
  admission.setReadOnly(true);
  expect(admitIntent(admission, { kind: 'setRenderEnvironment', payload: { themeColors: {}, showHiddenText: false } }).status).toBe('submitted');
  admission.dispose();
});

test('a throwing submission observer cannot strand held intents, trailing input or composition waiters', async () => {
  const { admission, submitted } = lane();
  const failure = new Error('observer failed');
  const report = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const reservation = admission.reserveComposition();
    const first = admitIntent(admission, { kind: 'setRenderEnvironment', payload: { themeColors: {}, showHiddenText: true } },
      { fence: true, onSubmitted: () => { throw failure; } });
    const second = admitIntent(admission, { kind: 'insertText', payload: { text: 'after' } });
    if (first.status !== 'held' || second.status !== 'held') throw new Error('intents were not held');
    const settled = admission.whenCompositionSettled!();
    expect(admission.settleComposition(reservation.slot, null)).toEqual({ status: 'ignored', reason: 'empty-composition' });
    await settled;
    expect(await first.released).toMatchObject({ status: 'submitted', ticket: submitted[0]!.ticket });
    expect(await second.released).toMatchObject({ status: 'submitted', ticket: submitted[1]!.ticket });
    expect(submitted.map(entry => entry.kind)).toEqual(['setRenderEnvironment', 'insertText']);
    expect(report).toHaveBeenCalledWith('[PeerlessInput] submission observer failed', failure);
    for (const [index, entry] of submitted.entries()) entry.settle({ status: 'applied', seq: index + 1 });
    for (let turn = 0; turn < 5; turn++) await Promise.resolve();
    expect(admission.getSnapshot()).toMatchObject({ heldIntents: 0, compositionSlot: null, originalUnackedIntents: 0, tail: undefined });
  } finally { report.mockRestore(); admission.dispose(); }
});

test('a held submission that throws settles the rest as dropped and still releases composition waiters', async () => {
  const { admission, submitted } = lane();
  const failure = new Error('submit failed');
  const reservation = admission.reserveComposition();
  const first = admitIntent(admission, { kind: 'insertText', payload: { text: 'a' } });
  const second = admitIntent(admission, { kind: 'insertText', payload: { text: 'b' } });
  if (first.status !== 'held' || second.status !== 'held') throw new Error('intents were not held');
  const settled = admission.whenCompositionSettled!();
  const submit = spyOn(admission.session, 'submit').mockImplementationOnce(() => { throw failure; });
  try {
    expect(admission.settleComposition(reservation.slot, null)).toEqual({ status: 'ignored', reason: 'empty-composition' });
    await settled;
    await expect(first.released).rejects.toBe(failure);
    await expect(second.released).rejects.toBe(failure);
    expect(submitted).toHaveLength(0);
    expect(admission.getSnapshot()).toMatchObject({ heldIntents: 0, compositionSlot: null, originalUnackedIntents: 0 });
  } finally { submit.mockRestore(); admission.dispose(); }
});

test('a throwing composition clears a refused tail after every held intent drops', async () => {
  const { admission, submitted } = lane();
  admitIntent(admission, { kind: 'insertText', payload: { text: 'refused' } });
  const reservation = admission.reserveComposition();
  const held = ['first', 'second'].map(text => admitIntent(admission, { kind: 'insertText', payload: { text } }));
  submitted[0]!.settle({ status: 'refused', seq: 1, reason: 'stale-version' });
  await submitted[0]!.ticket.ack;
  expect(admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 2, tail: submitted[0]!.ticket });
  const failure = new Error('composition submit failed');
  const submit = spyOn(admission.session, 'submit').mockImplementationOnce(() => { throw failure; });
  try {
    const settled = admission.whenCompositionSettled!();
    expect(() => admission.settleComposition(reservation.slot, { text: '漢字', target: reservation.target })).toThrow(failure);
    await settled;
    for (const entry of held) {
      if (entry.status !== 'held') throw new Error('intent was not held');
      await expect(entry.released).rejects.toBe(failure);
    }
    expect(admission.getSnapshot()).toMatchObject({
      originalUnackedIntents: 0, tail: undefined, echoEnabled: true, echoDisabledReason: null,
      compositionSlot: null, heldIntents: 0,
    });
    const next = submitFormattingIntent(admission, { kind: 'formatRange', payload: { delta: { bold: true } } });
    expect(next.status).toBe('submitted');
    expect(submitted).toHaveLength(2);
    expect(submitted[1]!.options.after).toBeUndefined();
    submitted[1]!.settle({ status: 'applied', seq: 2 });
    if (next.status === 'submitted') expect(await next.ticket.ack).toMatchObject({ status: 'applied' });
    expect(admission.getSnapshot().originalUnackedIntents).toBe(0);
  } finally { submit.mockRestore(); admission.dispose(); }
});

test('a plain submission that throws releases its own intent count and echo fence', () => {
  for (const kind of ['insertText', 'replaceSelection'] as const) {
    const { admission, submitted } = lane();
    const failure = new Error('plain submit failed');
    const submit = spyOn(admission.session, 'submit').mockImplementationOnce(() => { throw failure; });
    try {
      expect(() => admitIntent(admission, { kind, payload: { text: 'failed' } })).toThrow(failure);
      expect(submitted).toHaveLength(0);
      expect(admission.getSnapshot()).toMatchObject({
        originalUnackedIntents: 0, tail: undefined, echoEnabled: true, echoDisabledReason: null,
      });
      expect(admitIntent(admission, { kind: 'insertText', payload: { text: 'next' } })).toMatchObject({
        status: 'submitted', echoEnabled: true,
      });
      expect(submitted[0]!.options.after).toBeUndefined();
    } finally { submit.mockRestore(); admission.dispose(); }
  }
});

test.each(['acceptRevision', 'rejectRevision'] as const)('a refused %s fence without a sequence re-enables echo after acknowledgement', async (kind) => {
  const h = lane();
  try {
    const result = admitIntent(h.admission, { kind, payload: {} });
    if (result.status !== 'submitted') throw new Error('intent was not submitted');
    expect(result.ticket.seq).toBeUndefined();
    expect(h.admission.getSnapshot().echoDisabledReason).toBe('fence');
    expect(await result.ticket.ack).toMatchObject({ status: 'refused' });
    expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, tail: undefined, echoEnabled: true, echoDisabledReason: null });
    expect(admitIntent(h.admission, { kind: 'insertText', payload: { text: 'next' } }))
      .toMatchObject({ status: 'submitted', echoEnabled: true });
    expect(h.submitted[0]!.options.echo?.intent.payload).toEqual({ text: 'next' });
  } finally { h.admission.dispose(); }
});

test.each(['refused', 'cancelled'] as const)('a %s fence with a sequence re-enables echo after acknowledgement', async (status) => {
  const h = lane();
  try {
    admitIntent(h.admission, { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } });
    h.submitted[0]!.settle({ status, seq: 1 });
    await h.submitted[0]!.ticket.ack;
    expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, echoEnabled: true, echoDisabledReason: null });
    expect(admitIntent(h.admission, { kind: 'insertText', payload: { text: 'next' } }))
      .toMatchObject({ status: 'submitted', echoEnabled: true });
    expect(h.submitted[1]!.options.echo?.intent.payload).toEqual({ text: 'next' });
  } finally { h.admission.dispose(); }
});

test('a matching fence ack context attaches queued text before presentation', async () => {
  const h = lane();
  admitIntent(h.admission, { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } });
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'a' } });
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'b' } });
  const install = spyOn(h.echoPainter, 'installAckContext');
  h.ackContext(1);
  expect(install).toHaveBeenCalledTimes(1);
  expect(h.session.state.presentedThroughSeq).toBe(0);
  expect(h.submitted.slice(1).map(({ options }) => options.echo?.intent.payload)).toEqual([{ text: 'a' }, { text: 'b' }]);
  expect(h.admission.getSnapshot().echoDisabledReason).toBeNull();
  h.submitted[0]!.settle({ status: 'noop', seq: 1 });
  await h.submitted[0]!.ticket.ack;
  expect(admitIntent(h.admission, { kind: 'insertText', payload: { text: 'c' } })).toMatchObject({ echoEnabled: true });
  h.admission.dispose();
});

for (const excluded of ['epoch mismatch', 'composition', 'read-only', 'different fence'] as const) {
  test(`fence ack context does not attach text during ${excluded}`, () => {
    const h = lane();
    admitIntent(h.admission, { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } });
    admitIntent(h.admission, { kind: 'insertText', payload: { text: 'a' } });
    if (excluded === 'composition') h.admission.reserveComposition();
    if (excluded === 'read-only') h.admission.setReadOnly(true);
    const install = spyOn(h.echoPainter, 'installAckContext');
    h.ackContext(excluded === 'different fence' ? 2 : 1, excluded === 'epoch mismatch' ? 2 : 1);
    expect(h.submitted[1]!.options.echo).toBeUndefined();
    if (excluded !== 'epoch mismatch') expect(install).not.toHaveBeenCalled();
    if (excluded === 'epoch mismatch') {
      h.present(1);
      expect(h.submitted[1]!.options.echo).toBeDefined();
    }
    h.admission.dispose();
  });
}

test('covered fences re-enable echo with unacked text', async () => {
  const h = lane();
  admitIntent(h.admission, { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } });
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'a' } });
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'b' } });
  h.submitted[0]!.settle({ status: 'applied', seq: 1 });
  await h.submitted[0]!.ticket.ack;
  expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 2, echoEnabled: false, echoDisabledReason: 'fence' });
  h.present(1, false);
  expect(h.admission.getSnapshot().echoDisabledReason).toBe('fence');
  h.present(1);
  expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 2, echoEnabled: true, echoDisabledReason: null });
  expect(h.submitted.slice(1).map(({ options }) => options.echo?.intent.payload)).toEqual([{ text: 'a' }, { text: 'b' }]);
  h.admission.dispose();
});

test('a throwing post-fence submission preserves echo for acknowledged text', async () => {
  const h = lane();
  admitIntent(h.admission, { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } });
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'a' } });
  const failure = new Error('submit failed');
  const submit = spyOn(h.session, 'submit').mockImplementationOnce(() => { throw failure; });
  try {
    expect(() => admitIntent(h.admission, { kind: 'insertText', payload: { text: 'b' } })).toThrow(failure);
    expect(h.submitted).toHaveLength(2);
    for (const entry of h.submitted) entry.settle({ status: 'applied', seq: entry.ticket.seq });
    await Promise.all(h.submitted.map(({ ticket }) => ticket.ack));
    expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, echoDisabledReason: 'fence' });
    h.present(1);
    expect(h.admission.getSnapshot()).toMatchObject({ echoEnabled: true, echoDisabledReason: null });
    expect(h.submitted[1]!.options.echo?.intent.payload).toEqual({ text: 'a' });
  } finally { submit.mockRestore(); h.admission.dispose(); }
});

test('a post-fence ticket cancelled before sequencing does not hold echo detached', async () => {
  const h = lane();
  admitIntent(h.admission, { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } });
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'a' } });
  const cancelled: CommandTicket = {
    id: 99, seq: undefined, ack: Promise.resolve({ status: 'cancelled', reason: 'dependency-refused' }),
    presented: Promise.resolve('dropped'),
  };
  const submit = spyOn(h.session, 'submit').mockImplementationOnce(() => cancelled);
  const attach = h.session.attachEcho!.bind(h.session);
  const attachEcho = spyOn(h.session, 'attachEcho').mockImplementation((tickets, preview) =>
    tickets.includes(cancelled) ? false : attach(tickets, preview));
  try {
    admitIntent(h.admission, { kind: 'insertText', payload: { text: 'b' } });
    for (const entry of h.submitted) entry.settle({ status: 'applied', seq: entry.ticket.seq });
    await Promise.all([...h.submitted.map(({ ticket }) => ticket.ack), cancelled.ack]);
    expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, echoDisabledReason: 'fence' });
    h.present(1);
    expect(h.admission.getSnapshot()).toMatchObject({ echoEnabled: true, echoDisabledReason: null });
    expect(h.submitted[1]!.options.echo?.intent.payload).toEqual({ text: 'a' });
  } finally { submit.mockRestore(); attachEcho.mockRestore(); h.admission.dispose(); }
});

test('failed and cancelled deferred submissions do not block post-fence echo', async () => {
  const h = lane();
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'before' } });
  const fence = submitFormattingIntent(h.admission, { kind: 'formatRange', payload: { delta: { bold: true } } });
  const first = admitIntent(h.admission, { kind: 'insertText', payload: { text: 'a' } });
  const failed = admitIntent(h.admission, { kind: 'insertText', payload: { text: 'b' } });
  const cancelled = admitIntent(h.admission, { kind: 'insertText', payload: { text: 'c' } });
  if (fence.status !== 'submitted' || first.status !== 'submitted' || failed.status !== 'submitted' || cancelled.status !== 'submitted') {
    throw new Error('intents were not submitted');
  }
  const failure = new Error('submit failed');
  const send = h.session.submit.bind(h.session);
  const submit = spyOn(h.session, 'submit').mockImplementation((kind, payload, options) => {
    if ('text' in payload && payload.text === 'b') throw failure;
    return send(kind, payload, options);
  });
  try {
    h.submitted[0]!.settle({ status: 'applied', seq: 1 });
    await h.waitForSubmitted(3);
    await expect(failed.ticket.ack).rejects.toBe(failure);
    expect(await cancelled.ticket.ack).toMatchObject({ status: 'cancelled' });
    for (const entry of h.submitted.slice(1)) entry.settle({ status: 'applied', seq: entry.ticket.seq });
    await Promise.all([fence.ticket.ack, first.ticket.ack]);
    expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, echoDisabledReason: 'fence' });
    h.present(2);
    expect(h.admission.getSnapshot()).toMatchObject({ echoEnabled: true, echoDisabledReason: null });
    expect(h.submitted[2]!.options.echo?.intent.payload).toEqual({ text: 'a' });
  } finally { submit.mockRestore(); h.admission.dispose(); }
});

test.each([1, 2, 3])('acknowledgement gaps preserve post-fence echo through frame %i', async (throughSeq) => {
  const h = lane();
  try {
    admitIntent(h.admission, { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } });
    h.submitted[0]!.settle({ status: 'applied', seq: 1 });
    await h.submitted[0]!.ticket.ack;
    expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, echoDisabledReason: 'fence' });
    for (const text of ['a', 'b']) admitIntent(h.admission, { kind: 'insertText', payload: { text } });
    for (const entry of h.submitted.slice(1)) entry.settle({ status: 'applied', seq: entry.ticket.seq });
    await Promise.all(h.submitted.map(({ ticket }) => ticket.ack));
    expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, echoDisabledReason: 'fence' });
    admitIntent(h.admission, { kind: 'insertText', payload: { text: 'c' } });
    h.present(throughSeq);
    expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 1, echoEnabled: true, echoDisabledReason: null });
    expect(h.submitted.slice(1, throughSeq).every(({ options }) => options.echo === undefined)).toBe(true);
    expect(h.submitted.slice(throughSeq).map(({ options }) => options.echo?.intent.payload))
      .toEqual(['a', 'b', 'c'].slice(throughSeq - 1).map(text => ({ text })));
    expect(admitIntent(h.admission, { kind: 'insertText', payload: { text: 'd' } }))
      .toMatchObject({ status: 'submitted', echoEnabled: true });
    expect(h.submitted[4]!.options.echo?.intent.payload).toEqual({ text: 'd' });
  } finally { h.admission.dispose(); }
});

test('fully covered acknowledged text releases the echo fence without replay', async () => {
  const h = lane();
  const preview = spyOn(h.echoPainter, 'preview');
  try {
    admitIntent(h.admission, { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } });
    admitIntent(h.admission, { kind: 'insertText', payload: { text: 'a' } });
    for (const entry of h.submitted) entry.settle({ status: 'applied', seq: entry.ticket.seq });
    await Promise.all(h.submitted.map(({ ticket }) => ticket.ack));
    expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, echoDisabledReason: 'fence' });
    h.present(2);
    expect(preview).not.toHaveBeenCalled();
    expect(h.submitted[1]!.options.echo).toBeUndefined();
    expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, echoEnabled: true, echoDisabledReason: null });
    expect(admitIntent(h.admission, { kind: 'insertText', payload: { text: 'b' } }))
      .toMatchObject({ status: 'submitted', echoEnabled: true });
    expect(preview).toHaveBeenCalledTimes(1);
  } finally { preview.mockRestore(); h.admission.dispose(); }
});

test('only the latest fence can re-enable echo', () => {
  const h = lane();
  const navigate = { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } } as const;
  admitIntent(h.admission, navigate);
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'before' } });
  admitIntent(h.admission, navigate);
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'after' } });
  h.present(1);
  expect(h.admission.getSnapshot()).toMatchObject({ echoEnabled: false, echoDisabledReason: 'fence' });
  expect(h.submitted[1]!.options.echo).toBeUndefined();
  expect(h.submitted[3]!.options.echo).toBeUndefined();
  h.present(3);
  expect(h.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 4, echoEnabled: true, echoDisabledReason: null });
  expect(h.submitted[1]!.options.echo).toBeUndefined();
  expect(h.submitted[3]!.options.echo?.intent.payload).toEqual({ text: 'after' });
  h.admission.dispose();
});

test('failed post-fence previews leave no partial echo', () => {
  const h = lane();
  admitIntent(h.admission, { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } });
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'a' } });
  admitIntent(h.admission, { kind: 'insertText', payload: { text: 'b' } });
  const preview = spyOn(h.echoPainter, 'preview')
    .mockImplementationOnce((intent) => ({ intent, originalIntentCount: 1, contextFrameEpoch: 1 }))
    .mockReturnValueOnce(undefined);
  try {
    h.present(1);
    expect(preview).toHaveBeenCalledTimes(2);
    expect(h.admission.getSnapshot()).toMatchObject({ echoEnabled: false, echoDisabledReason: 'fence' });
    expect(h.submitted.every(({ options }) => options.echo === undefined)).toBe(true);
    h.present(2);
    expect(h.admission.getSnapshot()).toMatchObject({ echoEnabled: true, echoDisabledReason: null });
    expect(h.submitted[2]!.options.echo?.intent.payload).toEqual({ text: 'b' });
  } finally { preview.mockRestore(); h.admission.dispose(); }
});

test('a held submission failure preserves the committed composition and drops its successors', async () => {
  const { admission, submitted } = lane();
  const selection = admission.viewStore.getSnapshot().workerSelection;
  if (!selection) throw new Error('worker selection is unavailable');
  const state = updateComposition(beginComposition(admission, 1), 'かな');
  const first = admitIntent(admission, { kind: 'insertText', payload: { text: 'first' } });
  const failed = admitIntent(admission, { kind: 'insertText', payload: { text: 'failed' } });
  const last = admitIntent(admission, { kind: 'insertText', payload: { text: 'last' } });
  if (first.status !== 'held' || failed.status !== 'held' || last.status !== 'held') throw new Error('intents were not held');
  const failure = new Error('held submit failed');
  const send = admission.session.submit.bind(admission.session);
  const submit = spyOn(admission.session, 'submit').mockImplementation((kind, payload, options) => {
    if ('text' in payload && payload.text === 'failed') throw failure;
    return send(kind, payload, options);
  });
  try {
    const settled = admission.whenCompositionSettled!();
    const result = endComposition(admission, state, '漢字');
    expect(result).toMatchObject({ status: 'committed', ticket: submitted[0]!.ticket,
      state: { phase: 'committed', text: '漢字', ticket: submitted[0]!.ticket } });
    expect(result.state.recoverableText).toBeUndefined();
    await settled;
    expect(await first.released).toMatchObject({ status: 'submitted', ticket: submitted[1]!.ticket });
    await expect(failed.released).rejects.toBe(failure);
    await expect(last.released).rejects.toBe(failure);
    expect(submitted.map(entry => entry.payload)).toEqual([{ text: '漢字', selection }, { text: 'first' }]);
    expect(submitted[1]!.options.after).toBe(submitted[0]!.ticket);
    expect(admission.getSnapshot()).toMatchObject({ heldIntents: 0, compositionSlot: null, originalUnackedIntents: 2 });
    submitted[1]!.settle({ status: 'applied', seq: 2 });
    await submitted[1]!.ticket.ack;
    expect(admission.getSnapshot().originalUnackedIntents).toBe(1);
    submitted[0]!.settle({ status: 'applied', seq: 1 });
    await submitted[0]!.ticket.ack;
    expect(result.state.phase).toBe('committed');
    expect(result.state.recoverableText).toBeUndefined();
    expect(endComposition(admission, state, '漢字')).toBe(result);
    expect(admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, tail: undefined, echoEnabled: true });
  } finally { submit.mockRestore(); admission.dispose(); }
});

test('a held failure after a read-only held entry releases each count once', async () => {
  const { admission, submitted } = lane();
  const reservation = admission.reserveComposition();
  const ignored = admitIntent(admission, { kind: 'insertText', payload: { text: 'ignored' } });
  const failed = admitIntent(admission, { kind: 'setRenderEnvironment', payload: { themeColors: {}, showHiddenText: true } });
  const last = admitIntent(admission, { kind: 'setRenderEnvironment', payload: { themeColors: {}, showHiddenText: false } });
  if (ignored.status !== 'held' || failed.status !== 'held' || last.status !== 'held') throw new Error('intents were not held');
  const barrier = admission.setReadOnly(true);
  const failure = new Error('environment submit failed');
  const submit = spyOn(admission.session, 'submit').mockImplementationOnce(() => { throw failure; });
  try {
    expect(admission.settleComposition(reservation.slot, null)).toEqual({ status: 'ignored', reason: 'empty-composition' });
    expect(await ignored.released).toEqual({ status: 'ignored', reason: 'read-only' });
    await expect(failed.released).rejects.toBe(failure);
    await expect(last.released).rejects.toBe(failure);
    expect(submitted).toHaveLength(1);
    expect(admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, tail: barrier, heldIntents: 0 });
    submitted[0]!.settle({ status: 'noop', seq: 1 });
    await barrier.ack;
    expect(admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, tail: undefined, echoEnabled: true });
  } finally { submit.mockRestore(); admission.dispose(); }
});

test('read-only admission forwards the host proposal grant and refuses proposals and edits without it', () => {
  const { admission, submitted } = lane();
  admission.setReadOnly(true);
  const propose: PeerlessInputIntent = { kind: 'propose', payload: { expectVersion: 'v1', proposals: [], hostProposal: true } };
  const decide: PeerlessInputIntent = { kind: 'setProposalStates', payload: { expectVersion: 'v1', expectPreviewVersion: 0, changes: [] } };
  const withdraw: PeerlessInputIntent = { kind: 'withdrawProposals', payload: { expectVersion: 'v1', ids: [] } };
  for (const intent of [propose, decide, withdraw]) expect(admitIntent(admission, intent)).toEqual({ status: 'ignored', reason: 'read-only' });
  expect(admitIntent(admission, { kind: 'insertText', payload: { text: 'x' } }, { hostProposal: true })).toEqual({ status: 'ignored', reason: 'read-only' });
  expect(submitted.map(entry => entry.kind)).toEqual(['setReadOnly']);
  for (const intent of [propose, decide, withdraw]) expect(admitIntent(admission, intent, { fence: true, hostProposal: true }).status).toBe('submitted');
  expect(submitted.slice(1).map(({ kind, payload }) => ({ kind, payload }))).toEqual([
    { kind: 'propose', payload: { expectVersion: 'v1', proposals: [], hostProposal: true } },
    { kind: 'setProposalStates', payload: { expectVersion: 'v1', expectPreviewVersion: 0, changes: [], hostProposal: true } },
    { kind: 'withdrawProposals', payload: { expectVersion: 'v1', ids: [], hostProposal: true } },
  ]);
  admission.setReadOnly(false);
  expect(admitIntent(admission, propose).status).toBe('submitted');
  expect(submitted.at(-1)!.payload).toEqual({ expectVersion: 'v1', proposals: [] });
  admission.dispose();
});

test('pending inputs count unacked edits and selections but not environment or mode commands', async () => {
  const { admission, submitted } = lane();
  try {
    admission.setReadOnly(false);
    admitIntent(admission, { kind: 'setRenderEnvironment', payload: { themeColors: {}, showHiddenText: true } }, { fence: true });
    expect(admission.getSnapshot().pendingInputs).toBe(0);
    admitIntent(admission, { kind: 'insertText', payload: { text: 'a' } });
    expect(admission.getSnapshot().pendingInputs).toBe(1);
    for (const [index, entry] of submitted.entries()) entry.settle({ status: 'applied', seq: index + 1 });
    for (let turn = 0; turn < 5; turn++) await Promise.resolve();
    expect(admission.getSnapshot().pendingInputs).toBe(0);
    admitIntent(admission, { kind: 'insertText', payload: { text: 'b' } });
    expect(admission.getSnapshot().pendingInputs).toBe(1);
  } finally { admission.dispose(); }
  expect(admission.getSnapshot().pendingInputs).toBe(0);
  submitted.at(-1)!.settle({ status: 'applied', seq: submitted.length });
  for (let turn = 0; turn < 5; turn++) await Promise.resolve();
  expect(admission.getSnapshot().pendingInputs).toBe(0);
});

function lane() {
  let selectionThroughSeq = 0;
  let selection: DocxCommandSelection = {
    anchor: { story: 'body', paraId: 'old', offset: 0 },
    focus: { story: 'body', paraId: 'old', offset: 0 },
  };
  const snapshot: PeerlessViewSnapshot = {
    identity: { session: 'test', incarnation: 1, epoch: 1 }, version: 0,
    throughSeq: 0, presentedThroughSeq: 0, workerSelection: selection,
    anchors: null, geometry: null, editState: undefined,
  };
  const submitted: Array<{
    kind: string; payload: DocxCommandPayload; options: SubmitOptions<PeerlessEchoIntent>;
    ticket: CommandTicket; settle(outcome: TicketOutcome, selection?: DocxCommandSelection): void;
    reject(error: unknown): void;
  }> = [];
  const waiters: Array<{ count: number; resolve(): void }> = [];
  const waitForSubmitted = (count: number) => submitted.length >= count ? Promise.resolve()
    : new Promise<void>((resolve) => { waiters.push({ count, resolve }); });
  const viewStore: PeerlessViewStore = {
    getSnapshot: () => snapshot, getServerSnapshot: () => snapshot,
    subscribe: () => () => {}, acceptWorkerState() {}, installFrameGeometry() {}, setProvisionalCaret() {}, dispose() {},
  };
  let contextThroughSeq: number | undefined;
  const listeners = new Set<Parameters<DocxWorkerEditSession<PeerlessEchoIntent>['subscribe']>[0]>();
  const ackListeners = new Set<(ack: DocxCommandAck) => void>();
  const echoPainter: PeerlessEchoPainter = {
    installAckContext(context, throughSeq) {
      if (context.frameEpoch !== 1) return false;
      contextThroughSeq = throughSeq;
      return true;
    },
    get contextThroughSeq() { return contextThroughSeq; },
    installContext() {}, restoreBackground: () => true,
    preview: (intent) => ({ intent, originalIntentCount: 1, contextFrameEpoch: 1 }),
    paint() {}, drop() {}, reconcile() {}, disable() {}, dispose() {},
  };
  const session: DocxWorkerEditSession<PeerlessEchoIntent> = {
    backend: 'worker', state: { epoch: 1, version: 0, throughSeq: 0,
      presentedThroughSeq: 0, selection, editState: undefined },
    submit(kind, payload, options = {}) {
      let settle!: (outcome: TicketOutcome) => void;
      let reject!: (error: unknown) => void;
      const ticket: CommandTicket = {
        id: submitted.length + 1, seq: submitted.length + 1,
        ack: new Promise((resolve, fail) => { settle = resolve; reject = fail; }),
        presented: Promise.resolve('presented'),
      };
      submitted.push({ kind, payload, options, ticket, reject, settle(outcome, nextSelection) {
        session.state.throughSeq = snapshot.throughSeq = Math.max(session.state.throughSeq, outcome.seq ?? 0);
        if (nextSelection && (outcome.seq ?? 0) >= selectionThroughSeq) {
          selectionThroughSeq = outcome.seq ?? 0;
          selection = nextSelection;
          session.state.selection = selection;
          snapshot.workerSelection = selection;
        }
        settle(outcome);
      } });
      if (kind === 'formatRange' && (!('selection' in payload) || !payload.selection)) {
        settle({ status: 'refused', seq: ticket.seq, reason: 'invalid-selection' });
      } else if (options.after && !submitted.some((entry) => entry.ticket === options.after)) {
        settle({ status: 'cancelled', reason: 'dependency-refused' });
      }
      for (const waiter of waiters) if (submitted.length >= waiter.count) waiter.resolve();
      return ticket;
    },
    attachEcho(tickets, preview) {
      const attachments: Array<{ options: SubmitOptions<PeerlessEchoIntent>; echo: PeerlessEchoIntent }> = [];
      for (const ticket of tickets) {
        const entry = submitted.find((entry) => entry.ticket === ticket)!;
        const echo = preview({ kind: entry.kind, payload: entry.payload, options: entry.options }, 1);
        if (echo === undefined) return false;
        attachments.push({ options: entry.options, echo });
      }
      for (const { options, echo } of attachments) options.echo = echo;
      return true;
    },
    subscribeAck(listener) { ackListeners.add(listener); return () => { ackListeners.delete(listener); }; },
    subscribe(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
    whenPresented: async () => 'presented',
    flush: async () => {}, dispose: async () => {},
  };
  const admission = createInputAdmission({ session, viewStore, echoPainter });
  let content = 'abcdef';
  const results = new Map<number, DocxCommandSelection>();
  const execute = () => {
    for (const entry of submitted) {
      const afterSeq = entry.options.after?.seq;
      if (afterSeq !== undefined) selection = results.get(afterSeq)!;
      if (entry.kind === 'formatRange') {
        expect('selection' in entry.payload).toBe(true);
      } else expect('selection' in entry.payload).toBe(false);
      let offset = selection.focus.offset;
      if (entry.kind === 'insertText') {
        content = 'abcdef';
        offset = 5;
      } else if (entry.kind === 'navigate') {
        const direction = 'direction' in entry.payload ? entry.payload.direction : '';
        offset += direction === 'backward' ? -1 : 1;
      } else if (entry.kind === 'deleteBackward') {
        content = content.slice(0, offset - 1) + content.slice(offset);
        offset -= 1;
      }
      selection = {
        anchor: { story: 'body', paraId: 'worker-paragraph', offset },
        focus: { story: 'body', paraId: 'worker-paragraph', offset },
      };
      results.set(entry.ticket.seq!, selection);
    }
  };
  return { admission, session, snapshot, submitted, waitForSubmitted, execute, results, echoPainter, content: () => content,
    ackContext(throughSeq: number, frameEpoch = 1) {
      const echoContext: NonNullable<DocxCommandAck['echoContext']> = { page: { pageId: '1', pageIndex: 0 }, frameEpoch,
        layoutRevision: 1, serializedBytes: 100, text: 'ab', styleRuns: [], boundaries: [], damage: [],
        selection: { anchorGrapheme: 1, focusGrapheme: 1 }, transform: [1, 0, 0, 1, 0, 0], shaping: 'simple-ltr' };
      for (const listener of ackListeners) listener({ session: 'test', incarnation: 1, epoch: 1, version: 0,
        throughSeq, outcomes: [{ seq: throughSeq, status: 'noop' }], echoContext });
    },
    present(throughSeq: number, freshContext = true) {
      contextThroughSeq = freshContext ? throughSeq : undefined;
      session.state.presentedThroughSeq = throughSeq;
      for (const listener of listeners) listener(session.state);
    },
  };
}

test('font environments are allowed read-only and ordered behind composition without author metadata', () => {
  const { admission, submitted } = lane();
  admission.setSuggesting?.('Writer');
  const reservation = admission.reserveComposition();
  const payload = { faces: [], chains: {}, families: ['Arial'] };
  const result = admitIntent(admission, { kind: 'setFontEnvironment', payload }, { fence: true });
  expect(result.status).toBe('held');
  admission.settleComposition(reservation.slot, { text: 'text', target: reservation.target });
  expect(submitted[1]!.kind).toBe('setFontEnvironment');
  expect(submitted[1]!.payload).toEqual(payload);
  expect(submitted[1]!.options.after).toBe(submitted[0]!.ticket);
  admission.setReadOnly(true);
  expect(admitIntent(admission, { kind: 'setFontEnvironment', payload }).status).toBe('submitted');
  admission.dispose();
});

test('d2b_intent_behind_unacked_insert_uses_worker_result', async () => {
  const { admission, submitted, execute, results } = lane();
  admitIntent(admission, { kind: 'insertText', payload: { text: '🦊' } });
  admitIntent(admission, { kind: 'deleteBackward', payload: {} });
  expect(submitted[1]!.options.after).toBe(submitted[0]!.ticket);
  expect(admission.getSnapshot().originalUnackedIntents).toBe(2);
  execute();
  expect(results.get(2)?.focus).toEqual({ story: 'body', paraId: 'worker-paragraph', offset: 4 });
  submitted[0]!.settle({ status: 'applied', seq: 1 });
  await submitted[0]!.ticket.ack;
  expect(admission.getSnapshot().originalUnackedIntents).toBe(1);
  expect(admission.getSnapshot().tail).toBe(submitted[1]!.ticket);
  submitted[1]!.settle({ status: 'applied', seq: 2 });
  await submitted[1]!.ticket.ack;
  expect(admission.getSnapshot().originalUnackedIntents).toBe(0);
  expect(admission.getSnapshot().tail).toBeUndefined();
  const composing = lane();
  admitIntent(composing.admission, { kind: 'insertText', payload: { text: 'x' } });
  const reservation = composing.admission.reserveComposition();
  const held = admitIntent(composing.admission, { kind: 'navigate', payload: {
    unit: 'character', direction: 'forward', extend: false,
  } });
  expect(held.status).toBe('held');
  expect(composing.submitted).toHaveLength(1);
  expect(composing.admission.getSnapshot().originalUnackedIntents).toBe(2);
  composing.admission.settleComposition(reservation.slot, { text: '日本語', target: reservation.target });
  expect(composing.submitted.map(({ kind }) => kind)).toEqual(['insertText', 'replaceSelection', 'navigate']);
  expect(composing.submitted.map(({ options }) => options.after?.seq)).toEqual([undefined, 1, 2]);
  expect(composing.admission.getSnapshot().originalUnackedIntents).toBe(3);
  expect(composing.admission.getSnapshot().heldIntents).toBe(0);
});

test('d2b_arrow_backspace_burst_preserves_order', () => {
  const { admission, submitted, execute, results, content } = lane();
  admitIntent(admission, { kind: 'insertText', payload: { text: 'x' } });
  for (let i = 0; i < 2; i += 1) {
    admitIntent(admission, { kind: 'navigate', payload: {
      unit: 'character', direction: 'backward', extend: false,
    } });
    admitIntent(admission, { kind: 'deleteBackward', payload: {} });
  }
  expect(submitted.map(({ kind }) => kind)).toEqual([
    'insertText', 'navigate', 'deleteBackward', 'navigate', 'deleteBackward',
  ]);
  expect(submitted.map(({ options }) => options.after?.seq)).toEqual([undefined, 1, 2, 3, 4]);
  execute();
  expect(content()).toBe('acef');
  expect(results.get(5)?.focus.offset).toBe(1);
});

test('d2b_paste_and_formatting_share_pending_intent_chain', async () => {
  const { admission, submitted, waitForSubmitted } = lane();
  const commands = createPeerlessCommandBinding(admission);
  commands.submit({ kind: 'insertText', payload: { text: 'a' } });
  commands.submit({ kind: 'selectPoint', payload: {
    page: { pageId: '1', pageIndex: 0 }, x: 12, y: 20, layoutRevision: 1,
    frameEpoch: 1, clickCount: 1, modifiers: { shift: false, alt: false, ctrl: false, meta: false },
  } });
  commands.submit({ kind: 'replaceSelection', payload: { text: 'paste' } }, { source: 'paste' });
  commands.format({ kind: 'toggleMark', payload: { mark: { type: 'bold' } } });
  submitFormattingIntent(admission, { kind: 'formatRange', payload: { delta: { italic: true } } });
  expect(submitted).toHaveLength(4);
  for (let i = 0; i < 3; i += 1) submitted[i]!.settle({ status: 'applied', seq: i + 1 });
  submitted[3]!.settle({ status: 'applied', seq: 4 }, {
    anchor: { story: 'body', paraId: 'worker-paragraph', offset: 5 },
    focus: { story: 'body', paraId: 'worker-paragraph', offset: 5 },
  });
  await waitForSubmitted(5);
  expect(submitted.map(({ options }) => options.after?.seq)).toEqual([undefined, 1, 2, 3, 4]);
  expect(submitted.slice(1).every(({ options }) => options.fence)).toBe(true);
  expect(submitted.slice(0, 4).every(({ payload }) => !('selection' in payload))).toBe(true);
  expect('selection' in submitted[4]!.payload).toBe(true);
  expect(submitted.slice(1).every(({ options }) => options.echo === undefined)).toBe(true);
  for (let i = 0; i < 257; i += 1) commands.submit({ kind: 'insertText', payload: { text: 'z' } });
  await waitForSubmitted(262);
  expect(submitted).toHaveLength(262);
  expect(admission.getSnapshot().originalUnackedIntents).toBe(258);
  expect(admission.getSnapshot().echoEnabled).toBe(false);
  expect(admission.getSnapshot().echoDisabledReason).toBe('intent-cap');
  const capped = lane();
  for (let i = 0; i < 257; i += 1) admitIntent(capped.admission, { kind: 'insertText', payload: { text: 'x' } });
  expect(capped.submitted.slice(0, 256).every(({ options }) => options.echo?.originalIntentCount === 1)).toBe(true);
  expect(capped.submitted[256]!.options.echo).toBeUndefined();
  expect(capped.submitted[256]!.options.after).toBe(capped.submitted[255]!.ticket);
});

test('d2b_reversed_acks_keep_the_latest_causal_tail', async () => {
  const { admission, submitted } = lane();
  admitIntent(admission, { kind: 'insertText', payload: { text: 'a' } });
  admitIntent(admission, { kind: 'insertText', payload: { text: 'b' } });
  submitted[1]!.settle({ status: 'applied', seq: 2 });
  await submitted[1]!.ticket.ack;
  expect(admission.getSnapshot().tail).toBe(submitted[1]!.ticket);
  admitIntent(admission, { kind: 'deleteBackward', payload: {} });
  expect(submitted[2]!.options.after).toBe(submitted[1]!.ticket);
  submitted[0]!.settle({ status: 'applied', seq: 1 });
  await submitted[0]!.ticket.ack;
  expect(admission.getSnapshot().tail).toBe(submitted[2]!.ticket);
  submitted[2]!.settle({ status: 'applied', seq: 3 });
  await submitted[2]!.ticket.ack;
  expect(admission.getSnapshot().tail).toBeUndefined();
  expect(admission.getSnapshot().originalUnackedIntents).toBe(0);
});

test('d2b_reversed_suppressed_acks_wait_for_selection_prefix', async () => {
  for (const formatting of [true, false]) {
    for (const selected of [true, false]) {
      for (const settledTail of [true, false]) {
        const input = lane();
        const anchor = { story: 'body', paraId: 'worker-range', offset: 2 };
        const focus = { ...anchor, offset: 5 };
        const selection = { anchor, focus: selected ? focus : anchor };
        input.snapshot.workerSelection = input.session.state.selection = { anchor, focus: selected ? anchor : focus };
        admitIntent(input.admission, { kind: 'navigate', payload: {
          unit: 'character', direction: 'forward', extend: selected,
        } });
        submitFormattingIntent(input.admission, { kind: 'toggleMark', payload: { mark: { type: 'bold' } } });
        if (settledTail) {
          input.submitted[1]!.settle({ status: 'noop', seq: 2 });
          await input.submitted[1]!.ticket.ack;
        }
        if (formatting) submitFormattingIntent(input.admission, {
          kind: 'formatRange', payload: { delta: { italic: true } },
        });
        else admitInputIntents(input.admission, normalizeInputEvent({
          source: 'keydown', eventId: 1, inputTime: 1, key: 'Backspace', code: 'Backspace',
          shift: false, alt: false, ctrl: true, meta: false,
        }, { composing: false, compositionId: null, lastHandledEventId: null, trailingCompositionText: null }).intents, {});
        if (!settledTail) input.submitted[1]!.settle({ status: 'noop', seq: 2 });
        await input.submitted[1]!.ticket.ack;
        for (let i = 0; i < 6; i += 1) await Promise.resolve();
        expect(input.submitted).toHaveLength(2);
        input.submitted[0]!.settle({ status: 'noop', seq: 1 }, selection);
        await input.waitForSubmitted(formatting || selected ? 3 : 4);
        expect(input.submitted.map(({ kind }) => kind)).toEqual(formatting
          ? ['navigate', 'toggleMark', 'formatRange']
          : selected ? ['navigate', 'toggleMark', 'deleteBackward']
            : ['navigate', 'toggleMark', 'navigate', 'deleteBackward']);
        if (formatting) expect(input.submitted[2]!.payload).toEqual({ delta: { italic: true }, selection });
        expect(input.submitted[2]!.options.after).toBe(input.submitted[1]!.ticket);
        if (!formatting && !selected) expect(input.submitted[3]!.options.after).toBe(input.submitted[2]!.ticket);
        input.admission.dispose();
      }
    }
  }
});

test('d2b_format_range_requires_the_acked_worker_selection', async () => {
  const ready = lane();
  const ackedSelection: DocxCommandSelection = {
    anchor: { story: 'body', paraId: 'acked-worker-result', offset: 3 },
    focus: { story: 'body', paraId: 'acked-worker-result', offset: 3 },
  };
  admitIntent(ready.admission, { kind: 'insertText', payload: { text: 'x' } });
  ready.submitted[0]!.settle({ status: 'applied', seq: 1 }, ackedSelection);
  await ready.submitted[0]!.ticket.ack;
  submitFormattingIntent(ready.admission, { kind: 'formatRange', payload: { delta: { bold: true } } });
  expect(ready.submitted[1]!.payload).toEqual({ delta: { bold: true }, selection: ackedSelection });
  expect(ready.submitted[1]!.options.after).toBeUndefined();
  const strict = lane();
  const invalidIntent: PeerlessFormattingIntent = { kind: 'formatRange', payload: { delta: { bold: true } } };
  const invalid = strict.session.submit(invalidIntent.kind, invalidIntent.payload);
  expect(await invalid.ack).toEqual({ status: 'refused', seq: 1, reason: 'invalid-selection' });

  const pending = lane();
  admitIntent(pending.admission, { kind: 'insertText', payload: { text: 'x' } });
  const format = submitFormattingIntent(pending.admission, { kind: 'formatRange', payload: { delta: { italic: true } } });
  admitIntent(pending.admission, { kind: 'insertText', payload: { text: 'y' } });
  expect(format.status).toBe('submitted');
  expect(pending.submitted).toHaveLength(1);
  expect(pending.admission.getSnapshot().originalUnackedIntents).toBe(3);
  const selection: DocxCommandSelection = {
    anchor: { story: 'body', paraId: 'worker-result', offset: 8 },
    focus: { story: 'body', paraId: 'worker-result', offset: 8 },
  };
  pending.submitted[0]!.settle({ status: 'applied', seq: 1 }, selection);
  await pending.waitForSubmitted(3);
  expect(pending.submitted[1]!.payload).toEqual({ delta: { italic: true }, selection });
  expect(pending.submitted[1]!.options.after).toBe(pending.submitted[0]!.ticket);
  expect(pending.submitted[2]!.options.after).toBe(pending.submitted[1]!.ticket);
  expect(pending.submitted[2]!.payload).toEqual({ text: 'y' });
  pending.submitted[1]!.settle({ status: 'noop', seq: 2 });
  if (format.status === 'submitted') expect((await format.ticket.ack).status).toBe('noop');

  const failed = lane();
  admitIntent(failed.admission, { kind: 'insertText', payload: { text: 'x' } });
  const cancelledFormat = submitFormattingIntent(failed.admission, { kind: 'formatRange', payload: { delta: { bold: true } } });
  const cancelledEdit = admitIntent(failed.admission, { kind: 'insertText', payload: { text: 'y' } });
  failed.submitted[0]!.reject(new Error('Terminal session failure'));
  if (cancelledFormat.status === 'submitted' && cancelledEdit.status === 'submitted') {
    await Promise.allSettled([cancelledFormat.ticket.ack, cancelledEdit.ticket.ack]);
  }
  expect(failed.submitted).toHaveLength(1);
  expect(failed.admission.getSnapshot().originalUnackedIntents).toBe(0);
  expect(failed.admission.getSnapshot().tail).toBeUndefined();
});

test('Tab recovers from a cancelled deferred predecessor and restarts the selection prefix', async () => {
  const { admission, submitted } = lane();
  try {
    admitIntent(admission, { kind: 'insertText', payload: { text: 'a' } });
    const cancelled = submitFormattingIntent(admission, { kind: 'formatRange', payload: { delta: { bold: true } } });
    const tab = admitIntent(admission, { kind: 'insertText', payload: { text: '\t' } }, { tabDirection: 'forward' });
    admitIntent(admission, { kind: 'insertText', payload: { text: 'b' } });
    if (cancelled.status !== 'submitted' || tab.status !== 'submitted') throw new Error('input was not submitted');
    submitted[0]!.settle({ status: 'refused', seq: 1, reason: 'invalid-selection' });
    expect(await cancelled.ticket.ack).toMatchObject({ status: 'cancelled' });
    for (let turn = 0; turn < 12; turn++) await Promise.resolve();
    expect(submitted.map(({ kind, payload }) => ({ kind, payload }))).toEqual([
      { kind: 'insertText', payload: { text: 'a' } },
      { kind: 'insertText', payload: { text: '\t' } },
      { kind: 'insertText', payload: { text: 'b' } },
    ]);
    expect(submitted[1]!.options.after).toBeUndefined();
    expect(submitted[2]!.options.after).toBe(submitted[1]!.ticket);
    submitFormattingIntent(admission, { kind: 'formatRange', payload: { delta: { italic: true } } });
    submitted[1]!.settle({ status: 'applied', seq: 2 });
    submitted[2]!.settle({ status: 'applied', seq: 3 });
    for (let turn = 0; turn < 12; turn++) await Promise.resolve();
    expect(submitted.map(({ kind }) => kind)).toEqual(['insertText', 'insertText', 'insertText', 'formatRange']);
    expect(submitted[3]!.options.after).toBe(submitted[2]!.ticket);
  } finally { admission.dispose(); }
});

test('Shift+Tab recovering from a cancelled predecessor inserts nothing', async () => {
  const { admission, submitted } = lane();
  try {
    admitIntent(admission, { kind: 'insertText', payload: { text: 'a' } });
    const cancelled = submitFormattingIntent(admission, { kind: 'formatRange', payload: { delta: { bold: true } } });
    const tab = admitIntent(admission, { kind: 'insertText', payload: { text: '\t' } }, { tabDirection: 'backward' });
    if (cancelled.status !== 'submitted' || tab.status !== 'submitted') throw new Error('input was not submitted');
    submitted[0]!.settle({ status: 'refused', seq: 1, reason: 'invalid-selection' });
    expect(await tab.ticket.ack).toEqual({ status: 'noop' });
    expect(submitted.map(({ kind, payload }) => ({ kind, payload }))).toEqual([
      { kind: 'insertText', payload: { text: 'a' } },
    ]);
  } finally { admission.dispose(); }
});

test('Tab queued behind a click that lands in a table cell navigates cells instead of inserting a tab', async () => {
  for (const [tabDirection, createLastRow] of [['forward', true], ['backward', false]] as const) {
    const { admission, snapshot, submitted, waitForSubmitted } = lane();
    try {
      admitIntent(admission, { kind: 'insertText', payload: { text: 'a' } });
      admitIntent(admission, { kind: 'insertText', payload: { text: '\t' } }, { tabDirection });
      expect(submitted).toHaveLength(1);
      snapshot.editState = { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
        canUndo: false, canRedo: false, hyperlink: null, tableContext: {} as never };
      submitted[0]!.settle({ status: 'applied', seq: 1 });
      await waitForSubmitted(2);
      expect(submitted.map(({ kind, payload }) => ({ kind, payload }))).toEqual([
        { kind: 'insertText', payload: { text: 'a' } },
        { kind: 'navigateCell', payload: { direction: tabDirection, extend: false, createLastRow } },
      ]);
      expect(submitted[1]!.options.after).toBe(submitted[0]!.ticket);
      expect(submitted[1]!.options.fence).toBe(true);
    } finally { admission.dispose(); }
  }
});

test('mode_barriers_restart_dependencies_after_refused_input', async () => {
  const { admission, submitted } = lane();
  admitIntent(admission, { kind: 'insertText', payload: { text: 'refused' } });
  const lock = admission.setReadOnly(true);
  expect(admitIntent(admission, { kind: 'insertText', payload: { text: 'locked' } })).toEqual({ status: 'ignored', reason: 'read-only' });
  const unlock = admission.setReadOnly(false);
  admitIntent(admission, { kind: 'insertText', payload: { text: 'valid' } });
  expect(submitted.map(({ kind }) => kind)).toEqual(['insertText', 'setReadOnly', 'setReadOnly', 'insertText']);
  expect(submitted[1]!.options.after).toBeUndefined();
  expect(submitted[2]!.options.after).toBeUndefined();
  expect(submitted[3]!.options.after).toBe(unlock);
  submitted[0]!.settle({ seq: 1, status: 'refused', reason: 'stale-version' });
  await submitted[0]!.ticket.ack;
  expect(admission.getSnapshot().tail).toBe(submitted[3]!.ticket);
  submitted[1]!.settle({ seq: 2, status: 'noop' });
  await lock.ack;
  expect(admission.getSnapshot().readOnly).toBe(false);
  submitted[2]!.settle({ seq: 3, status: 'noop' });
  await unlock.ack;
  submitted[3]!.settle({ seq: 4, status: 'applied' });
  await submitted[3]!.ticket.ack;
  expect(admission.getSnapshot().tail).toBeUndefined();
});

test('unlock_admits_input_and_deferred_formatting_before_worker_ack', async () => {
  const { admission, snapshot, submitted, waitForSubmitted } = lane();
  snapshot.editState = { readOnly: true, toolbar: null, effectiveFormatting: null, storedMarks: null,
    canUndo: false, canRedo: false, hyperlink: null, tableContext: null };
  const unlock = admission.setReadOnly(false);
  const edit = admitIntent(admission, { kind: 'insertText', payload: { text: 'valid' } });
  expect(edit.status).toBe('submitted');
  submitFormattingIntent(admission, { kind: 'formatRange', payload: { delta: { bold: true } } });
  expect(submitted[1]!.options.after).toBe(unlock);
  submitted[0]!.settle({ seq: 1, status: 'noop' });
  submitted[1]!.settle({ seq: 2, status: 'applied' });
  await waitForSubmitted(3);
  expect(submitted[2]!.kind).toBe('formatRange');
});

test('late_input_ack_cannot_retire_pending_mode_barrier', async () => {
  const { admission, submitted } = lane();
  admitIntent(admission, { kind: 'insertText', payload: { text: 'old' } });
  const barrier = admission.setReadOnly(true);
  submitted[0]!.settle({ seq: 1, status: 'applied' });
  await submitted[0]!.ticket.ack;
  expect(admission.getSnapshot().tail).toBe(barrier);
  expect(admission.getSnapshot().originalUnackedIntents).toBe(0);
  submitted[1]!.settle({ seq: 2, status: 'noop' });
  await barrier.ack;
  expect(admission.getSnapshot().tail).toBeUndefined();
});

test('mode_barrier_waits_for_earlier_deferred_admission_without_inheriting_refusal', async () => {
  for (const refused of [false, true]) {
    const { admission, submitted, waitForSubmitted } = lane();
    admitIntent(admission, { kind: 'insertText', payload: { text: 'earlier' } });
    submitFormattingIntent(admission, { kind: 'formatRange', payload: { delta: { bold: true } } });
    const barrier = admission.setReadOnly(true);
    expect(submitted.map(({ kind }) => kind)).toEqual(['insertText']);
    submitted[0]!.settle({ seq: 1, status: refused ? 'refused' : 'applied' });
    await waitForSubmitted(refused ? 2 : 3);
    const mode = submitted.at(-1)!;
    expect(mode.kind).toBe('setReadOnly');
    expect(mode.options.after).toBeUndefined();
    mode.settle({ seq: mode.ticket.seq, status: 'noop' });
    expect((await barrier.ack).status).toBe('noop');
    admission.dispose();
  }
});

test('cell_navigation_and_typing_share_the_pending_worker_selection_chain', () => {
  const { admission, submitted } = lane();
  admitIntent(admission, { kind: 'insertText', payload: { text: 'old cell' } });
  const cell = { kind: 'navigateCell' as const, payload: { direction: 'forward' as const, extend: false, createLastRow: true } };
  admitIntent(admission, cell);
  admitIntent(admission, { kind: 'insertText', payload: { text: 'next cell' } });
  expect(submitted.map(({ kind }) => kind)).toEqual(['insertText', 'navigateCell', 'insertText']);
  expect(submitted.map(({ options }) => options.after?.seq)).toEqual([undefined, 1, 2]);
  expect(submitted[1]!.options.fence).toBe(true);
});

test('history_intents_share_the_composition_and_typing_admission_tail', () => {
  const { admission, submitted } = lane();
  const binding = createPeerlessCommandBinding(admission);
  binding.submit({ kind: 'insertText', payload: { text: 'a' } });
  const slot = admission.reserveComposition();
  expect(binding.submit({ kind: 'undo', payload: {} })).toMatchObject({ status: 'held' });
  expect(submitted).toHaveLength(1);
  admission.settleComposition(slot.slot, { text: 'b', target: slot.target });
  expect(submitted.map(entry => entry.kind)).toEqual(['insertText', 'replaceSelection', 'undo']);
  expect(submitted[2]!.options).toMatchObject({ after: submitted[1]!.ticket, fence: true });
  expect(submitted[2]!.options.echo).toBeUndefined();
  admission.dispose();
});

test('line breaks and paragraph breaks normalize into distinct fenced commands', () => {
  const state = { composing: false, compositionId: null, lastHandledEventId: null, trailingCompositionText: null };
  for (const [inputType, kind] of [['insertLineBreak', 'lineBreak'], ['insertParagraph', 'paragraphBreak']] as const) {
    for (const source of ['beforeinput', 'input'] as const) {
      const event = { eventId: 1, inputTime: 0, inputType, isComposing: false };
      expect(normalizeInputEvent(source === 'beforeinput'
        ? { ...event, source, data: null } : { ...event, source, value: '' }, state)).toMatchObject({
        intents: [{ kind, payload: {} }], fence: true, preventDefault: true,
      });
    }
  }
  for (const [shift, kind] of [[true, 'lineBreak'], [false, 'paragraphBreak']] as const) {
    expect(normalizeInputEvent({ source: 'keydown', eventId: 1, inputTime: 0,
      key: 'Enter', code: 'Enter', shift, alt: false, ctrl: false, meta: false }, state)).toMatchObject({
      intents: [{ kind, payload: {} }], fence: true, preventDefault: true,
    });
  }
});

test('history_shortcuts_and_beforeinput_normalize_into_fenced_commands', () => {
  const state = { composing: false, compositionId: null, lastHandledEventId: null, trailingCompositionText: null };
  for (const [inputType, kind] of [['historyUndo', 'undo'], ['historyRedo', 'redo']] as const) {
    expect(normalizeInputEvent({ source: 'beforeinput', eventId: 1, inputTime: 0, inputType, data: null, isComposing: false }, state)).toMatchObject({ intents: [{ kind, payload: {} }], fence: true, preventDefault: true });
  }
  for (const [key, shift, kind] of [['z', false, 'undo'], ['z', true, 'redo'], ['y', false, 'redo']] as const) {
    expect(normalizeInputEvent({ source: 'keydown', eventId: 1, inputTime: 0, key, code: '', shift, alt: false, ctrl: true, meta: false }, state)).toMatchObject({ intents: [{ kind, payload: {} }], fence: true, preventDefault: true });
  }
});

test('suggesting mutations capture authors before composition and held input', () => {
  const { admission, submitted } = lane();
  admission.setSuggesting!('Writer');
  const slot = admission.reserveComposition();
  admission.setSuggesting!('Next writer');
  admitIntent(admission, { kind: 'insertText', payload: { text: 'later' } });
  admitIntent(admission, { kind: 'undo', payload: {} });
  admission.setSuggesting!(null);
  admission.settleComposition(slot.slot, { text: 'draft', target: slot.target });
  expect(submitted[0]!.payload).toMatchObject({ text: 'draft', suggesting: { name: 'Writer', date: expect.any(String) } });
  expect(submitted[1]!.payload).toMatchObject({ text: 'later', suggesting: { name: 'Next writer', date: expect.any(String) } });
  expect(submitted[0]!.options.echo).toBeUndefined();
  expect(submitted[1]!.options.fence).toBe(true);
  expect(submitted[2]!.payload).toEqual({});
  admission.dispose();
});

test('editing composition does not acquire a later suggesting author', () => {
  const { admission, submitted } = lane();
  const slot = admission.reserveComposition();
  admission.setSuggesting!('Writer');
  admission.settleComposition(slot.slot, { text: 'plain', target: slot.target });
  expect(submitted[0]!.payload).not.toHaveProperty('suggesting');
  admitIntent(admission, { kind: 'insertText', payload: { text: 'tracked' } });
  expect(submitted[1]!.payload).toMatchObject({ suggesting: { name: 'Writer' } });
  admission.dispose();
});

test('suggesting mode leaves comment and object commands without text suggestion metadata', () => {
  const { admission, submitted } = lane();
  admission.setSuggesting!('Writer');
  const range = { story: 'body', start: { paraId: 'p', offset: 0 }, end: { paraId: 'p', offset: 4 } };
  const intents = [
    { kind: 'addComment', payload: { ranges: [range], author: 'Writer', date: '2026-01-01', body: 'Note', expectVersion: 'native' } },
    { kind: 'resizeObject', payload: { handle: 'image', widthEmu: 9525, heightEmu: 9525, expectVersion: 'native' } },
    { kind: 'moveObject', payload: { handle: 'image', hOffsetEmu: 0, vOffsetEmu: 0, expectVersion: 'native' } },
  ] satisfies PeerlessInputIntent[];
  for (const intent of intents) expect(admitIntent(admission, intent).status).toBe('submitted');
  expect(submitted.map(entry => entry.payload)).toEqual(intents.map(intent => intent.payload));
  admitIntent(admission, { kind: 'insertText', payload: { text: 'tracked' } });
  expect(submitted.at(-1)!.payload).toMatchObject({ suggesting: { name: 'Writer' } });
  admission.dispose();
});
