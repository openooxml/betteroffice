import { expect, test, type Mock } from 'bun:test';
import { inputs, proposalHost, source } from './workerProposalHost.testing';

test('host proposal rounds return four records from the presented review snapshot', async () => {
  const h = proposalHost();
  expect(await h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs })).toEqual({ ok: true, snapshot: h.snapshot() });
  expect(await h.controller.ref.getProposals()).toEqual(h.snapshot());
  expect(h.snapshot().proposals).toHaveLength(4);
  expect(h.submit.mock.calls[0]).toMatchObject(['propose', { expectVersion: 'native-0', proposals: inputs }, { fence: true }]);
  h.controller.dispose();
});

test('a null review after acknowledgement waits for the new preview frame', async () => {
  const h = proposalHost({ automatic: false });
  let settled = false;
  const pending = h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs }).then(value => { settled = true; return value; });
  for (let turn = 0; turn < 20; turn++) await Promise.resolve();
  expect(h.submit).toHaveBeenCalledTimes(1);
  expect(settled).toBe(false);
  expect(h.nullReads()).toBeGreaterThan(0);
  h.present();
  expect(await pending).toEqual({ ok: true, snapshot: h.snapshot() });
  expect(await h.controller.ref.getProposals()).toEqual(h.snapshot());
  h.controller.dispose();
});

test('accept reject and undo cycles keep proposal previews in step with presented frames', async () => {
  const h = proposalHost();
  await h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs });
  for (const state of ['accepted', 'proposed', 'rejected', 'proposed'] as const) {
    const before = await h.controller.ref.getProposals();
    const result = await h.controller.ref.setProposalStates({ expectVersion: before.version, expectPreviewVersion: before.previewVersion,
      changes: [{ id: inputs[0]!.id, state }] });
    expect(result).toEqual({ ok: true, snapshot: h.snapshot() });
    expect(h.snapshot().previewVersion).toBe(before.previewVersion + 1);
    await h.controller.ref.commands.execute('undo', null);
    expect(await h.controller.ref.getProposals()).toEqual(h.snapshot());
  }
  const before = await h.controller.ref.getProposals();
  expect(await h.controller.ref.withdrawProposals({ expectVersion: before.version, ids: [inputs[0]!.id] })).toEqual({ ok: true, snapshot: h.snapshot() });
  expect(h.snapshot().proposals).toHaveLength(3);
  h.controller.dispose();
});

test('proposal refusals return only fields determined exactly by the shell', async () => {
  const h = proposalHost();
  h.refuse('proposal-id-conflict');
  const result: unknown = await h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: [inputs[0]!] });
  expect(result).toEqual({
    ok: false, version: 'native-0', failure: { code: 'proposal-id-conflict', proposalId: inputs[0]!.id },
  });
  h.controller.dispose();
});

test('proposal rounds wait for composition and flush through the shared admission', async () => {
  const h = proposalHost();
  const composition = h.binding.admission.reserveComposition();
  const pending = h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs });
  for (let turn = 0; turn < 20; turn++) await Promise.resolve();
  expect(h.submit).not.toHaveBeenCalled();
  h.binding.admission.settleComposition(composition.slot, null);
  expect(await pending).toEqual({ ok: true, snapshot: h.snapshot() });
  expect(h.binding.session.flush).toHaveBeenCalled();
  h.controller.dispose();
});

const composeDuringFlush = (h: ReturnType<typeof proposalHost>) => {
  let slot: number | undefined;
  (h.binding.session.flush as unknown as Mock<() => Promise<void>>).mockImplementationOnce(async () => {
    slot = h.binding.admission.reserveComposition().slot;
  });
  return async () => {
    for (let turn = 0; turn < 100 && !h.binding.admission.getSnapshot().heldIntents; turn++) await Promise.resolve();
    expect(h.binding.admission.getSnapshot().heldIntents).toBe(1);
    for (let turn = 0; turn < 20; turn++) await Promise.resolve();
    return slot!;
  };
};

test('a composition starting during the proposal flush keeps the round pending until its held command applies', async () => {
  const h = proposalHost();
  const held = composeDuringFlush(h);
  let settled = false;
  const pending = h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs }).then(value => { settled = true; return value; });
  const slot = await held();
  expect(settled).toBe(false);
  h.binding.admission.settleComposition(slot, null);
  expect(await pending).toEqual({ ok: true, snapshot: h.snapshot() });
  expect(h.submit.mock.calls.map(([kind]) => kind)).toEqual(['propose']);
  h.controller.dispose();
});

test('a composition starting during a ref command flush reports the held command once it applies', async () => {
  const h = proposalHost();
  const held = composeDuringFlush(h);
  let settled = false;
  const pending = h.controller.ref.commands.execute('undo', null).then(value => { settled = true; return value; });
  const slot = await held();
  expect(settled).toBe(false);
  h.binding.admission.settleComposition(slot, null);
  expect(await pending).toEqual({ ok: true, status: 'executed' });
  expect(h.submit.mock.calls.map(([kind]) => kind)).toEqual(['undo']);
  h.controller.dispose();
});

test('a superseded review read retries against the newer presented native version', async () => {
  const h = proposalHost();
  h.readAt.mockImplementationOnce(async () => {
    const command = h.binding.commands.submit({ kind: 'undo', payload: {} }, { fence: true });
    if (command.status === 'submitted') await command.ticket.ack;
    return { status: 'superseded' };
  });
  expect(await h.controller.ref.getProposals()).toEqual(h.snapshot());
  expect(h.readAt.mock.calls.at(-1)?.[1]).toBe(h.snapshot().version);
  h.controller.dispose();
});

test('an unavailable presented review rejects within the bounded wait', async () => {
  const h = proposalHost({ automatic: false });
  h.controller.update({ ...h.settings, reviewTimeoutMs: 20 });
  await expect(h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs })).rejects.toThrow('Review snapshot timed out');
  h.controller.dispose();
});

test('read-only and viewing editors apply, decide and withdraw host proposals with allowHostProposals', async () => {
  for (const config of [{ readOnly: true }, { mode: 'viewing' as const }]) {
    const h = proposalHost(config);
    expect(await h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs })).toEqual({ ok: true, snapshot: h.snapshot() });
    expect(h.snapshot().proposals).toHaveLength(4);
    for (const state of ['accepted', 'proposed', 'rejected', 'proposed'] as const) {
      const before = await h.controller.ref.getProposals();
      expect(await h.controller.ref.setProposalStates({ expectVersion: before.version, expectPreviewVersion: before.previewVersion,
        changes: [{ id: inputs[0]!.id, state }] })).toEqual({ ok: true, snapshot: h.snapshot() });
      expect(h.snapshot().proposals[0]!.state).toBe(state);
    }
    const before = await h.controller.ref.getProposals();
    expect(await h.controller.ref.withdrawProposals({ expectVersion: before.version, ids: [inputs[0]!.id] })).toEqual({ ok: true, snapshot: h.snapshot() });
    expect(h.snapshot().proposals).toHaveLength(3);
    expect(h.submit.mock.calls.map(([kind, payload]) => [kind, (payload as { hostProposal?: boolean }).hostProposal]))
      .toEqual([['propose', true], ...Array.from({ length: 4 }, () => ['setProposalStates', true]), ['withdrawProposals', true]]);
    expect(h.binding.viewStore.getSnapshot().editState!.readOnly).toBe(true);
    h.controller.dispose();
  }
});

test('without allowHostProposals read-only and viewing editors refuse proposals, and edits stay refused either way', async () => {
  for (const config of [{ readOnly: true }, { mode: 'viewing' as const }]) {
    for (const allowHostProposals of [false, true]) {
      const h = proposalHost({ ...config, allowHostProposals });
      if (!allowHostProposals) {
        expect(await h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: [inputs[0]!] })).toEqual({
          ok: false, version: 'native-0', failure: { code: 'read-only', message: 'The editor is read-only', proposalId: inputs[0]!.id },
        });
        expect(await h.controller.ref.setProposalStates({ expectVersion: 'native-0', expectPreviewVersion: 0, changes: [] }))
          .toMatchObject({ ok: false, failure: { code: 'read-only' } });
        expect(await h.controller.ref.withdrawProposals({ expectVersion: 'native-0', ids: [] })).toMatchObject({ ok: false, failure: { code: 'read-only' } });
      }
      expect(h.binding.commands.submit({ kind: 'insertText', payload: { text: 'x' } }, { hostProposal: true })).toEqual({ status: 'ignored', reason: 'read-only' });
      expect(await h.controller.ref.applyEdits({ expectVersion: 'native-0', steps: [] })).toMatchObject({ ok: false, failure: { code: 'read-only' } });
      expect(h.submit).not.toHaveBeenCalled();
      h.controller.dispose();
    }
  }
});

test('proposals called before the worker document is ready wait for it and apply in order', async () => {
  const h = proposalHost({ readOnly: true, opening: true });
  const proposed = h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs });
  const read = h.controller.ref.getProposals();
  for (let turn = 0; turn < 20; turn++) await Promise.resolve();
  expect(h.submit).not.toHaveBeenCalled();
  h.controller.opening(source);
  for (let turn = 0; turn < 20; turn++) await Promise.resolve();
  expect(h.submit).not.toHaveBeenCalled();
  h.controller.attach(h.binding);
  expect(await proposed).toEqual({ ok: true, snapshot: h.snapshot() });
  expect(await read).toEqual(h.snapshot());
  expect(h.snapshot().proposals).toHaveLength(4);
  h.controller.dispose();
});

test('a host round queued before the document is ready is submitted on its first presented frame within microtasks', async () => {
  const h = proposalHost({ readOnly: true, opening: true, automatic: false });
  const round = h.controller.ref.getProposals().then(current => h.controller.ref.proposeChanges({ expectVersion: current.version, proposals: inputs }));
  for (let turn = 0; turn < 20; turn++) await Promise.resolve();
  expect(h.submit).not.toHaveBeenCalled();
  const ready = h.binding.viewStore.getSnapshot();
  h.controller.attach(h.binding);
  for (let turn = 0; turn < 200 && !h.submit.mock.calls.length; turn++) await Promise.resolve();
  expect(h.submit.mock.calls.map(([kind]) => kind)).toEqual(['propose']);
  expect(h.binding.viewStore.getSnapshot().geometry).toBe(ready.geometry);
  await (h.submit.mock.results[0]!.value as ReturnType<typeof h.submit>).ack;
  h.present();
  expect(await round).toEqual({ ok: true, snapshot: h.snapshot() });
  h.controller.dispose();
});

test('proposals waiting for the document fail when it is replaced, fails or the editor closes', async () => {
  const request = { expectVersion: 'native-0', proposals: inputs };
  const replaced = proposalHost({ opening: true });
  const pending = replaced.controller.ref.proposeChanges(request);
  replaced.controller.opening(Uint8Array.of(2));
  await expect(pending).rejects.toThrow('Document was replaced');
  replaced.controller.attach(replaced.binding);
  expect(await replaced.controller.ref.proposeChanges(request)).toMatchObject({ ok: true });
  expect(replaced.submit).toHaveBeenCalledTimes(1);
  replaced.controller.dispose();
  const failed = proposalHost({ opening: true });
  const failing = failed.controller.ref.proposeChanges(request);
  failed.controller.fail(new Error('Document is corrupt'));
  await expect(failing).rejects.toThrow('Document is corrupt');
  failed.controller.dispose();
  const closed = proposalHost({ opening: true });
  const closing = closed.controller.ref.getProposals();
  closed.controller.dispose();
  await expect(closing).rejects.toThrow('Editor was closed');
  const empty = proposalHost({ opening: true });
  empty.controller.empty();
  await expect(empty.controller.ref.proposeChanges(request)).rejects.toThrow('The editor is not ready');
  expect(empty.submit).not.toHaveBeenCalled();
  empty.controller.dispose();
});

test('a proposal dropped with a failed composition reports that failure instead of read-only', async () => {
  const h = proposalHost();
  let reservation: ReturnType<typeof h.binding.admission.reserveComposition> | undefined;
  (h.binding.session.flush as unknown as Mock<() => Promise<void>>).mockImplementationOnce(async () => {
    reservation = h.binding.admission.reserveComposition();
  });
  const pending = h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs });
  for (let turn = 0; turn < 100 && !h.binding.admission.getSnapshot().heldIntents; turn++) await Promise.resolve();
  expect(h.binding.admission.getSnapshot().heldIntents).toBe(1);
  const failure = new Error('composition submit failed');
  h.submit.mockImplementationOnce(() => { throw failure; });
  expect(() => h.binding.admission.settleComposition(reservation!.slot, { text: 'composed', target: reservation!.target })).toThrow(failure);
  await expect(pending).rejects.toBe(failure);
  expect(h.snapshot().proposals).toEqual([]);
  h.controller.dispose();
});
test.todo('engine identical proposal retry with the original token returns noop before version checking', async () => {
  const h = proposalHost();
  const request = { expectVersion: 'native-0', proposals: inputs };
  const first = await h.controller.ref.proposeChanges(request);
  expect(await h.controller.ref.proposeChanges(request)).toEqual(first);
  h.controller.dispose();
});
test.todo('engine identity-assigning save presents its advanced native version before proposal reads', async () => {
  const h = proposalHost();
  h.controller.update({ ...h.settings, reviewTimeoutMs: 500, save: async () => { h.assignIdentities(); return new ArrayBuffer(1); } });
  await h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs });
  await h.controller.ref.save();
  expect(await h.controller.ref.getProposals()).toEqual(h.snapshot());
  h.controller.dispose();
});

test('proposals waiting for the document are bounded by time, a stalled first presentation and a queue cap', async () => {
  const request = { expectVersion: 'native-0', proposals: inputs };
  const late = proposalHost({ opening: true });
  late.controller.update({ ...late.settings, openTimeoutMs: 20 });
  await expect(late.controller.ref.proposeChanges(request)).rejects.toThrow('The document did not open in time');
  late.controller.dispose();
  const stalled = proposalHost({ opening: true });
  const pending = stalled.controller.ref.getProposals();
  const stall = new Error('Full frame responses did not satisfy requested coverage and revision');
  stalled.controller.stalled(stall);
  await expect(pending).rejects.toBe(stall);
  stalled.controller.attach(stalled.binding);
  expect(await stalled.controller.ref.getProposals()).toEqual(stalled.snapshot());
  stalled.controller.dispose();
  const capped = proposalHost({ opening: true });
  const queued = Array.from({ length: 64 }, () => capped.controller.ref.getProposals());
  await expect(capped.controller.ref.getProposals()).rejects.toThrow('Too many calls are waiting for the document');
  capped.controller.attach(capped.binding);
  expect(await Promise.all(queued)).toEqual(queued.map(() => capped.snapshot()));
  capped.controller.dispose();
});

test('worker refusals outside the public proposal codes map onto them and keep the worker reason', async () => {
  const h = proposalHost();
  for (const [reason, code] of [['invalid-payload', 'invalid-step'], ['review-unavailable', 'unsupported'], ['stale-version', 'stale-version']] as const) {
    h.refuse(reason);
    expect(await h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: [inputs[0]!] })).toEqual({
      ok: false, version: 'native-0', failure: { code, proposalId: inputs[0]!.id,
        message: code === 'stale-version' ? 'the document changed since the expected version was read' : reason },
    });
  }
  h.controller.dispose();
});
