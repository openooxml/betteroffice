import { describe, expect, test } from 'bun:test';
import { evaluatePptxCommand, nextFontSize } from './evaluate';
import { testEnvironment } from './testing';
import type { PptxCommandArgs, PptxCommandId } from './types';

function gate<K extends PptxCommandId>(
  id: K,
  args?: PptxCommandArgs[K],
  overrides: Parameters<typeof testEnvironment>[0] = {}
): string | null {
  const state = evaluatePptxCommand(id, args, testEnvironment(overrides));
  return state.enabled ? null : state.disabledReason.code;
}

const picture = {
  formattable: false,
  geometry: null,
  fill: null,
  stroke: null,
  strokeWidth: null,
  adjustments: {},
  order: { index: 0, count: 2 },
};

describe('PPTX command gate', () => {
  test('text commands need a writable, reviewed-out text target', () => {
    expect(gate('bold')).toBeNull();
    expect(gate('bold', null, { readOnly: true })).toBe('read-only');
    expect(gate('bold', null, { reviewing: true })).toBe('review-active');
    expect(gate('bold', null, { text: null })).toBe('text-selection-required');
    expect(gate('bold', null, { text: 'unsupported' })).toBe('unsupported-selection');
    expect(gate('alignment', { value: 'ctr' }, { text: null })).toBe('text-selection-required');
  });

  test('mixed selections report mixed marks and no common value', () => {
    const state = evaluatePptxCommand(
      'italic',
      undefined,
      testEnvironment({
        text: {
          kind: 'range',
          bold: false,
          italic: 'mixed',
          underline: true,
          fontFamily: null,
          fontSize: null,
          color: null,
          alignment: null,
        },
      })
    );
    expect(state.active).toBe('mixed');
    const mixedSize = { ...(testEnvironment().text as object), fontSize: null } as never;
    const size = evaluatePptxCommand('fontSize', undefined, testEnvironment({ text: mixedSize }));
    expect(size.value).toBeNull();
    const unread = evaluatePptxCommand(
      'fontSize',
      undefined,
      testEnvironment({ text: 'unsupported' })
    );
    expect(unread.value).toBeUndefined();
    expect(unread.enabled).toBe(false);
  });

  test('alignment options report their own active state', () => {
    const env = testEnvironment();
    expect(evaluatePptxCommand('alignment', { value: 'l' }, env).active).toBe(true);
    expect(evaluatePptxCommand('alignment', { value: 'r' }, env).active).toBe(false);
    const mixed = testEnvironment({ text: { ...(env.text as object), alignment: null } as never });
    expect(evaluatePptxCommand('alignment', undefined, mixed).value).toBeNull();
    expect(evaluatePptxCommand('alignment', { value: 'l' }, mixed).active).toBe(false);
  });

  test('validates arguments', () => {
    expect(gate('fontSize', { points: 0 })).toBe('invalid-arguments');
    expect(gate('fontSize', { points: 401 })).toBe('invalid-arguments');
    expect(gate('fontSize', { points: Number.NaN })).toBe('invalid-arguments');
    expect(gate('fontSize', { points: 400 })).toBeNull();
    expect(gate('textColor', { color: 'red' })).toBe('invalid-arguments');
    expect(gate('zoom', { scale: 5 })).toBe('invalid-arguments');
    expect(gate('zoom', { scale: 'fit' })).toBeNull();
    expect(gate('tool', { value: 'shape:blob' as never })).toBe('invalid-arguments');
    expect(gate('shapeStrokeWidth', { points: 0 })).toBe('invalid-arguments');
    expect(gate('shapeStrokeWidth', { points: null })).toBeNull();
    expect(gate('bold', {} as never)).toBe('invalid-arguments');
  });

  test('shape formatting needs a preset shape; ordering works for pictures', () => {
    expect(gate('shapeFill', { color: '#112233' })).toBeNull();
    expect(gate('shapeFill', { color: null }, { shape: null })).toBe('shape-required');
    expect(gate('shapeFill', { color: null }, { shape: picture })).toBe('unsupported-selection');
    expect(gate('zOrder', { value: 'front' }, { shape: picture })).toBeNull();
    const none = evaluatePptxCommand(
      'zOrder',
      { value: 'front' },
      testEnvironment({ shape: null })
    );
    expect(none.enabled ? null : none.disabledReason).toEqual({
      code: 'shape-required',
      message: 'Select a shape or picture first.',
    });
  });

  test('z-order is disabled at each boundary separately', () => {
    const bottom = { shape: picture };
    expect(gate('zOrder', { value: 'back' }, bottom)).toBe('z-order-boundary');
    expect(gate('zOrder', { value: 'backward' }, bottom)).toBe('z-order-boundary');
    expect(gate('zOrder', { value: 'forward' }, bottom)).toBeNull();
    const top = { shape: { ...picture, order: { index: 1, count: 2 } } };
    expect(gate('zOrder', { value: 'front' }, top)).toBe('z-order-boundary');
    expect(gate('zOrder', { value: 'back' }, top)).toBeNull();
    const alone = { shape: { ...picture, order: { index: 0, count: 1 } } };
    expect(gate('zOrder', undefined, alone)).toBe('z-order-boundary');
  });

  test('adjustments follow the selected shape and its limits', () => {
    expect(gate('shapeAdjustment', { name: 'adj', value: 0.5 })).toBeNull();
    expect(gate('shapeAdjustment', { name: 'adj', value: 0.6 })).toBe('invalid-adjustment');
    expect(gate('shapeAdjustment', { name: 'adj2', value: 0.1 })).toBe('invalid-adjustment');
    const state = evaluatePptxCommand('shapeAdjustment', undefined, testEnvironment());
    expect(state.value).toEqual({ name: 'adj', value: 0.17 });
    const options = state.options ?? [];
    expect(options[options.length - 1].args).toEqual({ name: 'adj', value: 0.5 });
    const shape = testEnvironment().shape as object;
    const arrow = testEnvironment({
      shape: { ...shape, geometry: 'rightArrow', adjustments: { adj1: 0.5, adj2: 0.5 } } as never,
    });
    expect(evaluatePptxCommand('shapeAdjustment', undefined, arrow).value).toEqual({
      name: 'adj1',
      value: 0.5,
    });
    expect(evaluatePptxCommand('shapeAdjustment', { name: 'adj2', value: 1 }, arrow).enabled).toBe(
      true
    );
    const plain = testEnvironment({ shape: { ...shape, adjustments: {} } as never });
    expect(evaluatePptxCommand('shapeAdjustment', undefined, plain).value).toBeNull();
  });

  test('slides insert with the current, no, or an existing layout', () => {
    expect(gate('insertSlide', {})).toBeNull();
    expect(gate('insertSlide', { layoutPartPath: null })).toBeNull();
    expect(gate('insertSlide', { layoutPartPath: 'ppt/slideLayouts/slideLayout2.xml' })).toBeNull();
    expect(gate('insertSlide', { layoutPartPath: 'ppt/slideLayouts/gone.xml' })).toBe(
      'invalid-layout'
    );
    expect(gate('insertSlide', {}, { readOnly: true })).toBe('read-only');
  });

  test('arming placement tools and inserting images need a writable slide', () => {
    expect(gate('tool', { value: 'select' }, { readOnly: true })).toBeNull();
    expect(gate('tool', { value: 'textBox' }, { readOnly: true })).toBe('read-only');
    expect(gate('tool', { value: 'shape:ellipse' }, { slide: null })).toBe('slide-required');
    expect(gate('insertImage', null, { slide: null })).toBe('slide-required');
    expect(gate('insertImage', null, { reviewing: true })).toBe('review-active');
  });

  test('viewing commands stay available while read-only', () => {
    for (const id of ['save', 'exportPng', 'slideshow'] as const) {
      expect(gate(id, null, { readOnly: true })).toBeNull();
    }
    expect(gate('zoom', { scale: 2 }, { readOnly: true })).toBeNull();
    expect(gate('exportPng', null, { slide: null })).toBe('slide-required');
    expect(gate('slideshow', null, { slide: null })).toBe('slide-required');
    expect(gate('save', null, { status: 'empty' })).toBe('no-document');
  });

  test('history waits for pending input before reporting nothing to undo', () => {
    expect(gate('undo', null, { canUndo: false })).toBe('nothing-to-undo');
    expect(gate('undo', null, { canUndo: false, pendingInput: true })).toBeNull();
    expect(gate('redo', null, { canRedo: false })).toBe('nothing-to-redo');
    expect(gate('undo', null, { readOnly: true })).toBe('read-only');
  });

  test('proposal commands check the proposal, staleness and explicit force', () => {
    expect(gate('proposalAccept', { proposalId: 'p1' })).toBeNull();
    expect(gate('proposalAccept', { proposalId: 'p2' })).toBe('proposal-stale');
    expect(gate('proposalAccept', { proposalId: 'p2', force: true })).toBeNull();
    expect(gate('proposalAccept', { proposalId: 'p3' })).toBe('proposal-not-found');
    expect(gate('proposalReject', { proposalId: 'p2' })).toBeNull();
    expect(gate('proposalReject', { proposalId: 'p1' }, { readOnly: true })).toBe('read-only');
    expect(gate('proposalSelect', { proposalId: 'p2' })).toBe('proposal-not-found');
    expect(gate('proposalsPanel', null, { proposals: null })).toBe('proposals-unavailable');
  });

  test('reports every selector choice with its own state', () => {
    const readOnly = testEnvironment({ readOnly: true });
    const tools = evaluatePptxCommand('tool', undefined, readOnly).options ?? [];
    expect(tools[0]).toEqual({
      args: { value: 'select' },
      label: 'Select',
      state: { enabled: true, active: true },
    });
    expect(tools[1].state).toEqual({
      enabled: false,
      active: false,
      disabledReason: { code: 'read-only', message: 'The presentation is read-only.' },
    });
    const sizes = evaluatePptxCommand('fontSize', undefined, testEnvironment()).options ?? [];
    expect(sizes.find((option) => option.args.points === 24)?.state).toEqual({
      enabled: true,
      active: true,
    });
    const option = evaluatePptxCommand('fontSize', { points: 24 }, testEnvironment());
    expect(option.options).toBeUndefined();
    expect(JSON.parse(JSON.stringify(tools))).toEqual(tools);
  });

  test('host-disabled wins over every other reason', () => {
    expect(gate('bold', null, { text: null, hostDisabled: () => true })).toBe('host-disabled');
  });

  test('steps font sizes through the list and by one past its ends', () => {
    expect(nextFontSize(24, [12, 24, 36], 1)).toBe(36);
    expect(nextFontSize(24, [12, 24, 36], -1)).toBe(12);
    expect(nextFontSize(36, [12, 24, 36], 1)).toBe(37);
    expect(nextFontSize(1, [12], -1)).toBe(1);
    expect(nextFontSize(400, [12], 1)).toBe(400);
  });
});
