import { expect, mock, test } from 'bun:test';
import { toolbarFixture } from './toolbarCommands.testing';

test('worker toolbar preserves unsupported command reasons', () => {
  const { controller } = toolbarFixture();
  for (const id of ['insertTOC',
    'reviewAccept', 'reviewReject', 'reviewPrevious', 'reviewNext', 'commentsSidebar', 'reportIssue'] as const) {
    expect(controller.commands.store.getState(id)).toMatchObject({ enabled: false, disabledReason: { code: 'unsupported-command' } });
  }
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
    expect(controller.commands.store.getState('tableAction', action)).toMatchObject({ enabled: true });
  }
  for (const action of [{ type: 'borderColor', color: 'FF0000' }, { type: 'borderWidth', size: 8 },
    { type: 'cellBorder', side: 'top', style: 'single', size: 4, color: 'FF0000' },
    { type: 'applyTableStyle', styleId: 'TableGrid' }] as const) {
    expect(controller.commands.store.getState('tableAction', action as never)).toMatchObject({ enabled: true });
  }
});

test('worker table formatting gates match hybrid', () => {
  const { controller, viewStore, selection } = toolbarFixture();
  const state = viewStore.getSnapshot();
  viewStore.acceptWorkerState(state.identity, { ...state, epoch: 1, selection,
    editState: { ...state.editState!, tableContext: { story: 'body', tableIndex: 0, rowIndex: 0, columnIndex: 0,
      rowCount: 2, columnCount: 2, hasMultiCellSelection: true, canSplitCell: true, justification: 'center',
      cellBorderColor: { rgb: '112233' }, cellBackgroundColor: 'AABBCC' } } });
  expect(controller.commands.store.getState('tableAction')).toMatchObject({
    enabled: true, value: { borderColor: '112233', fillColor: 'AABBCC', justification: 'center' },
  });
  for (const action of ['borderAll', 'borderOutside', 'borderInside', 'borderNone', 'borderTop', 'borderBottom',
    'borderLeft', 'borderRight'] as const) {
    expect(controller.commands.store.getState('tableAction', action)).toMatchObject({ enabled: true });
  }
  for (const action of [{ type: 'cellVerticalAlign', align: 'center' }, { type: 'cellMargins', margins: { top: 100 } },
    { type: 'cellTextDirection', direction: 'btLr' }, { type: 'toggleNoWrap' }, { type: 'rowHeight', height: 400 },
    { type: 'toggleHeaderRow' }, { type: 'distributeColumns' }, { type: 'autoFitContents' }] as const) {
    expect(controller.commands.store.getState('tableAction', action as never))
      .toMatchObject({ enabled: false, disabledReason: { code: 'unsupported-command' } });
  }
});

test('worker image properties need a selected image', () => {
  const { controller } = toolbarFixture();
  expect(controller.commands.store.getState('imageProperties')).toMatchObject({ enabled: false, disabledReason: { code: 'image-required' } });
  expect(controller.commands.store.getState('imageWrap')).toMatchObject({ enabled: false, disabledReason: { code: 'image-required' } });
  expect(controller.commands.store.getState('imageTransform')).toMatchObject({ enabled: false, disabledReason: { code: 'image-required' } });
});

test('enables and dispatches imageWrap and imageTransform when an image is selected', async () => {
  const { controller, parity } = toolbarFixture();
  const onImageWrap = mock((_wrap: unknown) => true);
  const onImageTransform = mock((_action: unknown) => true);
  controller.update({
    read: (async () => ({ status: 'ok', version: 'native', value: { version: 'native', paragraphs: [] } })) as never,
    focus() {},
    zoom: 1,
    imagePropertiesAvailable: true,
    image: { wrap: 'inline' },
    onImageWrap,
    onImageTransform,
  });
  expect(controller.commands.store.getState('imageWrap', { wrap: 'inline' })).toMatchObject({ enabled: true, active: true, value: 'inline' });
  expect(controller.commands.store.getState('imageWrap', { wrap: 'squareLeft' })).toMatchObject({ enabled: true, active: false, value: 'inline' });
  expect(controller.commands.store.getState('imageTransform', { action: 'rotateCW' })).toMatchObject({ enabled: true });
  parity('imageWrap', { wrap: 'inline' });
  parity('imageTransform', { action: 'rotateCW' });

  expect(await controller.commands.store.execute('imageWrap', { wrap: 'squareLeft' })).toEqual({ ok: true, status: 'executed' });
  expect(onImageWrap).toHaveBeenCalledWith('squareLeft');
  expect(await controller.commands.store.execute('imageTransform', { action: 'rotateCW' })).toEqual({ ok: true, status: 'executed' });
  expect(onImageTransform).toHaveBeenCalledWith('rotateCW');

  controller.update({
    read: (async () => ({ status: 'ok', version: 'native', value: { version: 'native', paragraphs: [] } })) as never,
    focus() {},
    zoom: 1,
    mode: 'suggesting',
    imagePropertiesAvailable: true,
    image: { wrap: 'inline' },
  });
  expect(controller.commands.store.getState('imageWrap', { wrap: 'squareLeft' })).toMatchObject({ enabled: false, disabledReason: { code: 'suggesting-unsupported' } });
  expect(controller.commands.store.getState('imageTransform', { action: 'rotateCW' })).toMatchObject({ enabled: false, disabledReason: { code: 'suggesting-unsupported' } });

  controller.update({
    read: (async () => ({ status: 'ok', version: 'native', value: { version: 'native', paragraphs: [] } })) as never,
    focus() {},
    zoom: 1,
    readOnly: true,
    imagePropertiesAvailable: true,
    image: { wrap: 'inline' },
  });
  expect(controller.commands.store.getState('imageWrap', { wrap: 'squareLeft' })).toMatchObject({ enabled: false, disabledReason: { code: 'read-only' } });
  expect(controller.commands.store.getState('imageTransform', { action: 'rotateCW' })).toMatchObject({ enabled: false, disabledReason: { code: 'read-only' } });
});
