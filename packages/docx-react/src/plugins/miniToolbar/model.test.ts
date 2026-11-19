import { describe, expect, test } from 'bun:test';
import {
  DOCX_MINI_TOOLBAR_ITEMS,
  INITIAL_MINI_TOOLBAR_STATE,
  afterSelectionChange,
  avoidCommentButton,
  buttonState,
  dismissMiniToolbar,
  miniToolbarEligible,
  miniToolbarSize,
  miniToolbarView,
  resolveMiniToolbarItems,
  rovingIndex,
  type MiniToolbarSnapshot,
  type MiniToolbarState,
} from './model';
import type { DocxPluginSelection } from '../types';

const range = (from: number, to: number, story = 'body') => ({ story, from, to, layoutId: 'layout-1' });
const formatting = (hasSelection: boolean) => ({ hasSelection }) as unknown as DocxPluginSelection['formatting'];

function snapshot(overrides: Partial<MiniToolbarSnapshot> = {}): MiniToolbarSnapshot {
  return {
    inputFocused: true,
    mode: 'editing',
    readOnly: false,
    selectionSettled: true,
    selection: { formatting: formatting(true), displayRange: range(4, 9) },
    ...overrides,
  };
}

const state = (overrides: Partial<MiniToolbarState> = {}): MiniToolbarState => ({ ...INITIAL_MINI_TOOLBAR_STATE, ...overrides });

describe('items', () => {
  test('default to bold, italic, underline, strikethrough, clear formatting and paste as plain text', () => {
    expect(DOCX_MINI_TOOLBAR_ITEMS).toEqual(['bold', 'italic', 'underline', 'strikethrough', 'clearFormatting', 'pasteAsPlainText']);
    expect(resolveMiniToolbarItems()).toEqual(DOCX_MINI_TOOLBAR_ITEMS);
  });

  test('keep the host subset in the host order', () => {
    expect(resolveMiniToolbarItems(['pasteAsPlainText', 'clearFormatting', 'bold'])).toEqual(['pasteAsPlainText', 'clearFormatting', 'bold']);
  });

  test('refuse unknown, repeated or no ids', () => {
    expect(() => resolveMiniToolbarItems(['bold', 'insertLink'])).toThrow(TypeError);
    expect(() => resolveMiniToolbarItems(['bold', 'bold'])).toThrow(TypeError);
    expect(() => resolveMiniToolbarItems([])).toThrow(TypeError);
    expect(() => resolveMiniToolbarItems('bold' as unknown as string[])).toThrow(TypeError);
  });

  test('name an id that is not a string by its type', () => {
    const cyclic: Record<string, unknown> = {};
    cyclic.self = cyclic;
    expect(() => resolveMiniToolbarItems([10n] as unknown as string[])).toThrow('Unknown mini toolbar item of type bigint');
    expect(() => resolveMiniToolbarItems([cyclic] as unknown as string[])).toThrow('Unknown mini toolbar item of type object');
  });
});

describe('eligibility', () => {
  test('a settled, non-collapsed body selection with document input focus while editing or suggesting', () => {
    expect(miniToolbarEligible(snapshot(), state())).toBe(true);
    expect(miniToolbarEligible(snapshot({ mode: 'suggesting' }), state())).toBe(true);
    expect(miniToolbarEligible(snapshot({ selection: { formatting: null, displayRange: range(1, 2, 'body:cell-3') } }), state())).toBe(true);
  });

  test('not while the selection is still changing, as while typing or between clicks', () => {
    expect(miniToolbarEligible(snapshot({ selectionSettled: false }), state())).toBe(false);
  });

  test('not for a caret', () => {
    expect(miniToolbarEligible(snapshot({ selection: { formatting: formatting(false), displayRange: range(4, 4) } }), state())).toBe(false);
    expect(miniToolbarEligible(snapshot({ selection: { formatting: formatting(false), displayRange: null } }), state())).toBe(false);
    expect(miniToolbarEligible(snapshot({ selection: { formatting: null, displayRange: null } }), state())).toBe(false);
  });

  test('a selection waiting for its display range counts by its formatting', () => {
    expect(miniToolbarEligible(snapshot({ selection: { formatting: formatting(true), displayRange: null } }), state())).toBe(true);
  });

  test('not outside the body', () => {
    expect(miniToolbarEligible(snapshot({ selection: { formatting: formatting(true), displayRange: range(1, 5, 'hf:rIdH') } }), state())).toBe(false);
  });

  test('not in viewing or read-only', () => {
    expect(miniToolbarEligible(snapshot({ mode: 'viewing' }), state())).toBe(false);
    expect(miniToolbarEligible(snapshot({ readOnly: true }), state())).toBe(false);
  });

  test('not while the document input has lost focus to a menu, a dialog or a content control', () => {
    expect(miniToolbarEligible(snapshot({ inputFocused: false }), state())).toBe(false);
  });

  test('stays while focus is in the toolbar itself, for a settled selection only', () => {
    expect(miniToolbarEligible(snapshot({ inputFocused: false }), state({ toolbarFocused: true }))).toBe(true);
    expect(miniToolbarEligible(snapshot({ inputFocused: false, selectionSettled: false }), state({ toolbarFocused: true }))).toBe(false);
  });

  test('not for the selection it was dismissed on, until the selection changes', () => {
    const dismissed = dismissMiniToolbar(state({ shown: true }), snapshot());
    expect(dismissed.shown).toBe(false);
    expect(miniToolbarEligible(snapshot(), dismissed)).toBe(false);
    expect(miniToolbarEligible(snapshot({ selection: { formatting: formatting(true), displayRange: null } }), dismissed)).toBe(false);
    expect(afterSelectionChange(dismissed, { formatting: formatting(true), displayRange: { ...range(4, 9), layoutId: 'layout-2' } })).toBe(dismissed);
    expect(afterSelectionChange(dismissed, { formatting: formatting(true), displayRange: null })).toBe(dismissed);
    const moved = afterSelectionChange(dismissed, { formatting: formatting(true), displayRange: range(4, 12) });
    expect(moved.dismissed).toBeNull();
    expect(miniToolbarEligible(snapshot({ selection: { formatting: formatting(true), displayRange: range(4, 12) } }), moved)).toBe(true);
    const untouched = state();
    expect(afterSelectionChange(untouched, { formatting: null, displayRange: range(1, 2) })).toBe(untouched);
  });
});

describe('view', () => {
  const placed = { hidden: false, kept: false };
  test('shown for an eligible selection with a placement', () => {
    expect(miniToolbarView(true, placed)).toBe('shown');
  });
  test('kept, dimmed and inert, while the placement is kept', () => {
    expect(miniToolbarView(true, { hidden: false, kept: true })).toBe('kept');
  });
  test('hidden when the placement is hidden, missing or the selection is not eligible', () => {
    expect(miniToolbarView(true, { hidden: true, kept: false })).toBe('hidden');
    expect(miniToolbarView(true, { hidden: true, kept: true })).toBe('hidden');
    expect(miniToolbarView(true, null)).toBe('hidden');
    expect(miniToolbarView(false, placed)).toBe('hidden');
  });
});

describe('roving focus', () => {
  const enabled = [true, false, true, true];
  test('arrows move to the next or previous enabled button and wrap', () => {
    expect(rovingIndex('ArrowRight', 0, enabled)).toBe(2);
    expect(rovingIndex('ArrowRight', 3, enabled)).toBe(0);
    expect(rovingIndex('ArrowLeft', 2, enabled)).toBe(0);
    expect(rovingIndex('ArrowLeft', 0, enabled)).toBe(3);
  });
  test('Home and End go to the first and last enabled button', () => {
    expect(rovingIndex('Home', 3, [false, true, true, false])).toBe(1);
    expect(rovingIndex('End', 1, [false, true, true, false])).toBe(2);
  });
  test('other keys and an all-disabled toolbar move nothing', () => {
    expect(rovingIndex('Enter', 0, enabled)).toBeNull();
    expect(rovingIndex('ArrowRight', 0, [false, false])).toBeNull();
  });
});

describe('buttons', () => {
  test('toggles report aria-pressed, mixed included', () => {
    expect(buttonState('bold', { enabled: true, active: true })).toEqual({ pressed: 'true', disabled: false });
    expect(buttonState('italic', { enabled: true, active: false })).toEqual({ pressed: 'false', disabled: false });
    expect(buttonState('underline', { enabled: true, active: 'mixed' })).toEqual({ pressed: 'mixed', disabled: false });
    expect(buttonState('strikethrough', { enabled: true })).toEqual({ pressed: 'false', disabled: false });
  });
  test('actions carry no pressed state', () => {
    expect(buttonState('clearFormatting', { enabled: true })).toEqual({ pressed: undefined, disabled: false });
    expect(buttonState('pasteAsPlainText', { enabled: true })).toEqual({ pressed: undefined, disabled: false });
  });
  test('a refused command, such as one without a grant, is disabled and still listed', () => {
    const refused = { enabled: false, disabledReason: { code: 'permission-denied', message: 'not granted' } } as const;
    expect(buttonState('bold', refused)).toEqual({ pressed: 'false', disabled: true });
    expect(buttonState('pasteAsPlainText', refused)).toEqual({ pressed: undefined, disabled: true });
  });
});

describe('size and the comment button', () => {
  test('touch targets are at least 44 pixels under a coarse pointer', () => {
    expect(miniToolbarSize(6, true).button).toBeGreaterThanOrEqual(44);
    expect(miniToolbarSize(6, false).button).toBeLessThan(44);
    expect(miniToolbarSize(3, false).width).toBeLessThan(miniToolbarSize(6, false).width);
  });

  test('the toolbar never reaches past the right edge of the selection page, where the comment button sits', () => {
    const page = { x: 100, y: 0, width: 600, height: 900 };
    expect(avoidCommentButton(620, 200, page)).toBe(500);
    expect(avoidCommentButton(300, 200, page)).toBe(300);
    expect(avoidCommentButton(620, 200, null)).toBe(620);
    expect(avoidCommentButton(150, 800, page)).toBe(100);
  });
});
