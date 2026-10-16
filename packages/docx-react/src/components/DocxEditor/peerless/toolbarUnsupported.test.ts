import { expect, test } from 'bun:test';
import { toolbarFixture } from './toolbarCommands.testing';

test('worker toolbar preserves unsupported command reasons', () => {
  const { controller } = toolbarFixture();
  for (const id of ['insertTOC', 'imageWrap', 'imageTransform', 'reportIssue'] as const) {
    expect(controller.commands.store.getState(id)).toMatchObject({ enabled: false, disabledReason: { code: 'unsupported-command' } });
  }
});

test('worker comments sidebar toggles through its host binding and refuses a fixed controlled sidebar', async () => {
  const { controller } = toolbarFixture();
  expect(controller.commands.store.getState('commentsSidebar')).toMatchObject({ enabled: false, disabledReason: { code: 'controlled-sidebar' } });
  for (const [control, status] of [['internal', 'executed'], ['host', 'requested']] as const) {
    let open = false;
    const update = () => controller.update({ read: (async () => null) as never, focus() {}, zoom: 1,
      commentsSidebar: { open, control, toggle: () => { open = !open; update(); } } });
    update();
    expect(controller.commands.store.getState('commentsSidebar')).toMatchObject({ enabled: true, active: false });
    expect(await controller.commands.store.execute('commentsSidebar', null)).toEqual({ ok: true, status });
    expect(controller.commands.store.getState('commentsSidebar')).toMatchObject({ enabled: true, active: true });
  }
  controller.update({ read: (async () => null) as never, focus() {}, zoom: 1, commentsSidebar: { open: true, control: 'fixed', toggle() {} } });
  expect(controller.commands.store.getState('commentsSidebar')).toMatchObject({ enabled: false, active: true, disabledReason: { code: 'controlled-sidebar' } });
});

test('worker highlight and break toolbar commands match hybrid enablement', () => {
  const { controller, parity } = toolbarFixture();
  for (const id of ['highlightColor', 'insertPageBreak', 'insertSectionBreakNextPage', 'insertSectionBreakContinuous'] as const) {
    parity(id);
    expect(controller.commands.store.getState(id).enabled).toBe(true);
  }
});

test('worker table selection follows the shell table allowlist', () => {
  const { controller, viewStore, selection } = toolbarFixture();
  expect(controller.commands.store.getState('tableAction')).toMatchObject({ enabled: false, disabledReason: { code: 'table-required' } });
  const state = viewStore.getSnapshot();
  viewStore.acceptWorkerState(state.identity, { ...state, epoch: 1, selection,
    editState: { ...state.editState!, tableContext: { story: 'body', tableIndex: 0, rowIndex: 0, columnIndex: 0,
      rowCount: 2, columnCount: 2, hasMultiCellSelection: true, canSplitCell: true } } });
  for (const action of ['selectTable', 'selectRow', 'selectColumn', 'addRowAbove', 'addRowBelow', 'addColumnLeft',
    'addColumnRight', 'deleteRow', 'deleteColumn', 'deleteTable', 'mergeCells', 'splitCell'] as const) {
    expect(controller.commands.store.getState('tableAction', action)).toMatchObject({ enabled: true });
  }
  expect(controller.commands.store.getState('tableAction', { type: 'cellFillColor', color: 'FF0000' }))
    .toMatchObject({ enabled: true });
  for (const action of ['borderAll', 'borderOutside', 'borderInside', 'borderNone', 'borderTop', 'borderBottom',
    'borderLeft', 'borderRight'] as const) {
    expect(controller.commands.store.getState('tableAction', action))
      .toMatchObject({ enabled: false, disabledReason: { code: 'unsupported-command' } });
  }
  for (const action of [{ type: 'borderColor', color: 'FF0000' }, { type: 'borderWidth', size: 8 },
    { type: 'cellBorder', side: 'top', style: 'single', size: 4, color: 'FF0000' }] as const) {
    expect(controller.commands.store.getState('tableAction', action as never))
      .toMatchObject({ enabled: false, disabledReason: { code: 'unsupported-command' } });
  }
});

test('worker image properties need a selected image', () => {
  const { controller } = toolbarFixture();
  expect(controller.commands.store.getState('imageProperties')).toMatchObject({ enabled: false, disabledReason: { code: 'image-required' } });
});
