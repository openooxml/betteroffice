import { afterEach, expect, test } from 'bun:test';
import { closeCommandEngines, command, execute, openCommandEngine, snapshot, texts, TIMEOUT } from './residentCommandTestSupport';

afterEach(closeCommandEngines);

test('replace_matches_is_atomic_version_checked_and_one_history_step', async () => {
  const arm = await openCommandEngine(['cat cat', 'cat']);
  const found = arm.engine.proposalEngine.findText({ text: 'cat', within: { kind: 'story', story: 'body' }, view: 'accepted' });
  if (!found.ok) throw new Error(found.failure.message);
  const targets = found.matches.map(match => match.range);
  const version = arm.engine.proposalEngine.version();
  const before = snapshot(arm);
  expect(execute(arm, command(arm, 'replaceMatches', { targets, text: 'dog', expectVersion: 'stale' }))).toMatchObject({ status: 'refused', reason: 'stale-version' });
  expect(snapshot(arm)).toEqual(before);
  expect(execute(arm, command(arm, 'replaceMatches', { targets: [targets[0]!, targets[0]!], text: 'dog', expectVersion: version }))).toMatchObject({ status: 'refused', reason: 'overlapping-steps' });
  expect(snapshot(arm)).toEqual(before);
  expect(execute(arm, command(arm, 'replaceMatches', { targets, text: 'wolf', expectVersion: version }))).toMatchObject({ status: 'applied' });
  expect(texts(arm)).toEqual(['wolf wolf', 'wolf']);
  expect(execute(arm, command(arm, 'undo', {}))).toMatchObject({ status: 'applied' });
  expect(texts(arm)).toEqual(['cat cat', 'cat']);
  expect(execute(arm, command(arm, 'redo', {}))).toMatchObject({ status: 'applied' });
  expect(texts(arm)).toEqual(['wolf wolf', 'wolf']);
}, TIMEOUT);

test('replace_matches_validates_all_targets_and_honors_readonly', async () => {
  const arm = await openCommandEngine(['cat']);
  const paraId = arm.state.selection!.focus.paraId;
  const target = { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 3 }, view: 'accepted' as const };
  const version = arm.engine.proposalEngine.version();
  const before = snapshot(arm);
  expect(execute(arm, command(arm, 'replaceMatches', { targets: [target, { ...target, story: 'missing' }], text: 'dog', expectVersion: version }))).toMatchObject({ status: 'refused' });
  expect(snapshot(arm)).toEqual(before);
  arm.state.readOnly = true;
  expect(execute(arm, command(arm, 'replaceMatches', { targets: [target], text: 'dog', expectVersion: version }))).toMatchObject({ status: 'refused', reason: 'read-only' });
}, TIMEOUT);
