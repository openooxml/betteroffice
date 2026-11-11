import { expect, test } from 'bun:test';
import { createDocxSelectionObserver, docxSelectionIdentity } from './selectionIdentity';

test('resolved changes settle locally while remote moves preserve the quiet window', () => {
  const owner = {};
  const observe = createDocxSelectionObserver();
  const identity = (anchor: number, focus: number, story = 'body', paraId = 'p') =>
    docxSelectionIdentity(0, [story, 'sticky-anchor', 'sticky-focus', Math.sign(focus - anchor)],
      [[story, paraId, anchor], [story, paraId, focus]], anchor === focus);
  const a = observe(owner, identity(1, 3));
  const before = observe(owner, identity(3, 5), 'remote');
  const inside = observe(owner, identity(3, 7), 'remote');
  expect(before.settleKey).toBe(a.settleKey);
  expect(inside.settleKey).toBe(a.settleKey);
  expect(inside.resolvedKey).not.toBe(a.resolvedKey);
  const returned = observe(owner, identity(1, 3));
  expect(returned.settleKey).not.toBe(a.settleKey);
  expect(returned.resolvedKey).toBe(a.resolvedKey);
  expect(observe(owner, identity(1, 3)).settleKey).toBe(returned.settleKey);
  const reversed = observe(owner, identity(3, 1));
  expect(reversed.settleKey).not.toBe(returned.settleKey);
  expect(reversed.stickyKey).not.toBe(returned.stickyKey);
  const paragraph = observe(owner, identity(3, 1, 'body', 'q'));
  expect(paragraph.settleKey).not.toBe(reversed.settleKey);
  const story = observe(owner, identity(3, 1, 'hf:header', 'q'));
  expect(story.settleKey).not.toBe(paragraph.settleKey);
  const collapsed = observe(owner, identity(1, 1, 'hf:header', 'q'), 'remote');
  expect(collapsed.settleKey).not.toBe(story.settleKey);
  const removed = observe(owner, docxSelectionIdentity(0, null, null, false), 'remote');
  expect(removed.settleKey).not.toBe(collapsed.settleKey);
  expect(observe(owner, docxSelectionIdentity(0, null, null, false)).settleKey).toBe(removed.settleKey);
});

test('reader seeding advances the baseline without scheduling its own settle', () => {
  const owner = {};
  const observe = createDocxSelectionObserver();
  const range = docxSelectionIdentity(0, ['a', 'b'], [1, 3], false);
  const caret = docxSelectionIdentity(0, ['c', 'c'], [0, 0], true);
  const before = observe(owner, range);
  expect(observe(owner, caret, 'seed').settleKey).toBe(before.settleKey);
  expect(observe(owner, caret).settleKey).toBe(before.settleKey);
  expect(observe(owner, docxSelectionIdentity(0, ['d', 'd'], [1, 1], true)).settleKey).not.toBe(before.settleKey);
});
