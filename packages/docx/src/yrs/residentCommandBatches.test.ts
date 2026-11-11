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

test('a refused formatText step is a coded batch refusal in the worker and the session keeps working', async () => {
  const arm = await openCommandEngine([], {
    bodyXml: '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p><w:p><w:r><w:t>Plain</w:t></w:r></w:p>',
  });
  const [field, plain] = arm.engine.paragraphs('body').map((paragraph) => paragraph.paraId);
  const paragraph = (paraId: string) => ({ kind: 'paragraph' as const, story: 'body', paraId });
  const batch = (steps: unknown[]) => command(arm, 'applyEdits', { request: { expectVersion: arm.engine.proposalEngine.version(), steps } as never });
  const refused = [
    [{ op: 'formatText', target: paragraph(field!), set: { bold: true }, suggest: { author: 'Ann', date: '2026-11-10T12:00:00Z' } }],
    [{ op: 'formatText', target: paragraph(plain!), set: { link: 'x' } }],
    [{ op: 'formatText', target: { kind: 'range', story: 'body', view: 'accepted', start: { paraId: plain!, offset: 2 }, end: { paraId: plain!, offset: 2 } }, set: { bold: true } }],
  ];
  for (const steps of refused) {
    const before = snapshot(arm);
    const outcome = execute(arm, batch(steps));
    expect(outcome.status, JSON.stringify(steps)).toBe('refused');
    expect(typeof (outcome as { reason?: unknown }).reason).toBe('string');
    expect(snapshot(arm)).toEqual(before);
  }
  expect(execute(arm, batch([{ op: 'formatText', target: paragraph(plain!), set: { bold: true } }])).status).toBe('applied');
  expect(execute(arm, command(arm, 'undo', {})).status).toBe('applied');
}, TIMEOUT);
