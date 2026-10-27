import { expect, test } from 'bun:test';
import { toolbarFixture } from './toolbarCommands.testing';

test('worker insertLink is enabled at a collapsed caret', () => {
  const { controller, viewStore, selection, parity, read } = toolbarFixture();
  const view = viewStore.getSnapshot();
  viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: { anchor: selection.anchor, focus: selection.anchor } });
  parity('insertLink');
  expect(controller.commands.store.getState('insertLink').enabled).toBe(true);
  controller.update({ read: read as never, focus() {}, zoom: 1, readOnly: true });
  expect(controller.commands.store.getState('insertLink')).toMatchObject({ enabled: false, disabledReason: { code: 'read-only' } });
  controller.update({ read: read as never, focus() {}, zoom: 1 });
  viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: {
    anchor: { ...selection.anchor, story: 'fn:1' }, focus: { ...selection.anchor, story: 'fn:1' },
  } });
  expect(controller.commands.store.getState('insertLink')).toMatchObject({ enabled: false, disabledReason: { code: 'unsupported-story' } });
});

