import { expect, test } from 'bun:test';
import { DirtyProjectionStories, EditorDirtyStories } from './dirtyProjectionStories';

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

test('editor worker saves project every story after a main-thread projection with no later edit', () => {
  const dirty = new EditorDirtyStories();
  dirty.add('body');
  const first = dirty.captureWorkerSave();
  expect(first.stories).toEqual(['body']);
  dirty.add('hf:rId7');
  dirty.projected();
  first.clear();
  expect(dirty.projection.projectionOptions()).toBeUndefined();
  expect(dirty.captureWorkerSave().stories).toEqual([]);
  dirty.add('fn:1');
  expect(dirty.projection.projectionOptions()?.storyIds).toEqual(new Set(['fn:1']));
  const second = dirty.captureWorkerSave();
  expect(second.stories).toEqual(['hf:rId7', 'fn:1']);
  second.clear();
  dirty.add('body');
  expect(dirty.captureWorkerSave().stories).toEqual(['body']);
  dirty.clear();
  expect(dirty.captureWorkerSave().stories).toEqual([]);
  expect(dirty.projection.projectionOptions()).toBeUndefined();
});

test('adopting worker save updates preserves dirty stories and restores marking afterward', () => {
  const dirty = new EditorDirtyStories();
  dirty.add('body');
  dirty.adoptWorkerSaveUpdates(() => {
    dirty.add('hf:rId7');
    dirty.adoptWorkerSaveUpdates(() => dirty.add('fn:1'));
    dirty.add('en:2');
  });
  expect(dirty.captureWorkerSave().stories).toEqual(['body']);
  expect(dirty.projection.projectionOptions()?.storyIds).toEqual(new Set(['body']));
  dirty.projected();
  dirty.adoptWorkerSaveUpdates(() => dirty.add('hf:rId7'));
  expect(dirty.projection.projectionOptions()).toBeUndefined();
  const saved = dirty.captureWorkerSave();
  expect(saved.stories).toEqual([]);
  saved.clear();
  dirty.add('fn:1');
  expect(dirty.captureWorkerSave().stories).toEqual(['fn:1']);
  expect(dirty.projection.projectionOptions()?.storyIds).toEqual(new Set(['fn:1']));
  dirty.clear();
  const failure = new Error('Failed to apply worker save updates');
  expect(() => dirty.adoptWorkerSaveUpdates(() => {
    dirty.add('hf:rId7');
    throw failure;
  })).toThrow(failure);
  expect(dirty.captureWorkerSave().stories).toEqual([]);
  expect(dirty.projection.projectionOptions()).toBeUndefined();
  dirty.add('en:2');
  expect(dirty.captureWorkerSave().stories).toEqual(['en:2']);
  expect(dirty.projection.projectionOptions()?.storyIds).toEqual(new Set(['en:2']));
});
