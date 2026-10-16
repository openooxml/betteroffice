import { expect, test, type Mock } from 'bun:test';
import { inputs, proposalHost } from './workerProposalHost.testing';

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

test('allowed host proposals retain worker read-only state and return its refusal', async () => {
  const h = proposalHost({ readOnly: true });
  expect(await h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: [inputs[0]!] })).toEqual({
    ok: false, version: 'native-0', failure: { code: 'read-only', message: 'The editor is read-only', proposalId: inputs[0]!.id },
  });
  expect(h.binding.viewStore.getSnapshot().editState!.readOnly).toBe(true);
  expect(h.submit.mock.calls.map(([kind]) => kind)).toEqual(['propose']);
  h.controller.dispose();
});

test.todo('engine host proposal admission permits read-only and viewing editors with allowHostProposals', async () => {
  const h = proposalHost({ readOnly: true });
  expect(await h.controller.ref.proposeChanges({ expectVersion: 'native-0', proposals: inputs })).toMatchObject({ ok: true });
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
