import { expect, test } from 'bun:test';
import { DirtyProjectionStories } from './dirtyProjectionStories';

test('dirty projection stories select every story when empty and normalize nested edits', () => {
  const dirty = new DirtyProjectionStories();
  expect(dirty.projectionOptions()).toBeUndefined();
  for (const story of ['body', 'cell:table:0:0', 'hf:rId7:cell:0:0', 'fn:1:cell:0:0', 'en:2']) {
    dirty.add(story);
  }
  expect(dirty.projectionOptions()?.storyIds).toEqual(new Set(['body', 'hf:rId7', 'fn:1', 'en:2']));
  dirty.clear();
  expect(dirty.projectionOptions()).toBeUndefined();
});

test('a save snapshot keeps pending marks and preserves later edits to the same story', () => {
  const dirty = new DirtyProjectionStories();
  dirty.add('body');
  dirty.add('hf:rId7');
  const pending = dirty.capture();
  expect(pending.stories).toEqual(['body', 'hf:rId7']);
  expect(dirty.projectionOptions()?.storyIds).toEqual(new Set(pending.stories));
  dirty.add('body');
  dirty.add('fn:1');
  pending.clear();
  expect(dirty.projectionOptions()?.storyIds).toEqual(new Set(['body', 'fn:1']));
  dirty.capture().clear();
  expect(dirty.projectionOptions()).toBeUndefined();
});
