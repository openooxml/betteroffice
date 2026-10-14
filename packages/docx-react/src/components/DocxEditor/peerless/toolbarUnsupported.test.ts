import { expect, test } from 'bun:test';
import { toolbarFixture } from './toolbarCommands.testing';

test('worker toolbar preserves unsupported command reasons', () => {
  const { controller } = toolbarFixture();
  for (const id of ['highlightColor', 'insertPageBreak', 'insertSectionBreakNextPage', 'insertSectionBreakContinuous',
    'tableAction', 'insertTOC', 'imageWrap', 'imageTransform',
    'reviewAccept', 'reviewReject', 'reviewPrevious', 'reviewNext', 'commentsSidebar', 'reportIssue'] as const) {
    expect(controller.commands.store.getState(id)).toMatchObject({ enabled: false, disabledReason: { code: 'unsupported-command' } });
  }
});

test('worker image properties need a selected image', () => {
  const { controller } = toolbarFixture();
  expect(controller.commands.store.getState('imageProperties')).toMatchObject({ enabled: false, disabledReason: { code: 'image-required' } });
});
