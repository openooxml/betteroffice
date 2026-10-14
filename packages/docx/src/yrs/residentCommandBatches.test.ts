import { afterEach, expect, test } from 'bun:test';
import { closeCommandEngines, command, execute, openCommandEngine, snapshot, texts, TIMEOUT } from './residentCommandTestSupport';

afterEach(closeCommandEngines);

test('ordered_batches_return_receipts_and_preserve_atomicity_history_and_readonly', async () => {
  const arm = await openCommandEngine(['Seed']);
  const paraId = arm.state.selection!.focus.paraId;
  const request = () => ({ expectVersion: arm.engine.proposalEngine.version(), steps: [{ op: 'insertText' as const, target: { kind: 'paragraph' as const, story: 'body', paraId }, at: 'end' as const, text: ' tail' }] });
  const applied = execute(arm, command(arm, 'applyEdits', { request: request() }));
  expect(applied).toMatchObject({ status: 'applied', editState: { batch: { result: { ok: true, applied: true, receipts: [{ changed: true }] } } } });
  expect(texts(arm)).toEqual(['Seed tail']);
  expect(execute(arm, command(arm, 'undo', {}))).toMatchObject({ status: 'applied' });
  expect(texts(arm)).toEqual(['Seed']);
  const before = snapshot(arm);
  expect(execute(arm, command(arm, 'applyEdits', { request: { ...request(), expectVersion: 'stale' } }))).toMatchObject({ status: 'refused', reason: 'stale-version' });
  expect(snapshot(arm)).toEqual(before);
  arm.state.readOnly = true;
  const readonly = snapshot(arm);
  expect(execute(arm, command(arm, 'applyEdits', { request: request() }))).toMatchObject({ status: 'refused', reason: 'read-only' });
  expect(snapshot(arm)).toEqual(readonly);
}, TIMEOUT);
