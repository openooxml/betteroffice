import type { DocxPluginRect, DocxPluginSelection, DocxPluginSnapshot } from '../types';

/**
 * A button the mini toolbar can show, named by the built-in command it runs.
 *
 * @experimental The plugin API may change in minor releases.
 */
export type DocxMiniToolbarItem =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikethrough'
  | 'clearFormatting'
  | 'pasteAsPlainText';

/** @experimental The plugin API may change in minor releases. */
export const DOCX_MINI_TOOLBAR_ITEMS: readonly DocxMiniToolbarItem[] = Object.freeze([
  'bold', 'italic', 'underline', 'strikethrough', 'clearFormatting', 'pasteAsPlainText',
]);

const TOGGLES: ReadonlySet<DocxMiniToolbarItem> = new Set(['bold', 'italic', 'underline', 'strikethrough']);

export interface MiniToolbarState {
  shown: boolean;
  toolbarFocused: boolean;
  dismissed: string | null;
  focusRequest: number;
}

export const INITIAL_MINI_TOOLBAR_STATE: MiniToolbarState = Object.freeze({
  shown: false, toolbarFocused: false, dismissed: null, focusRequest: 0,
});

export type MiniToolbarSnapshot = Pick<
  DocxPluginSnapshot,
  'inputFocused' | 'mode' | 'readOnly' | 'selection' | 'selectionSettled'
>;

export function resolveMiniToolbarItems(items: readonly string[] = DOCX_MINI_TOOLBAR_ITEMS): readonly DocxMiniToolbarItem[] {
  if (!Array.isArray(items) || items.length === 0) throw new TypeError('items must list at least one mini toolbar item');
  const known = new Set<string>(DOCX_MINI_TOOLBAR_ITEMS);
  const seen = new Set<string>();
  for (const item of items) {
    if (!known.has(item)) {
      const shown = typeof item === 'string' ? JSON.stringify(item) : `of type ${typeof item}`;
      throw new TypeError(`Unknown mini toolbar item ${shown}`);
    }
    if (seen.has(item)) throw new TypeError(`Repeated mini toolbar item "${item}"`);
    seen.add(item);
  }
  return Object.freeze([...items] as DocxMiniToolbarItem[]);
}

/** Identifies a selection by its display range; null while the range is not read. */
export function selectionKey(selection: DocxPluginSelection): string | null {
  const range = selection.displayRange;
  return range ? JSON.stringify([range.story, range.from, range.to]) : null;
}

const inBody = (story: string) => story === 'body' || story.startsWith('body:');

export function miniToolbarEligible(snapshot: MiniToolbarSnapshot, state: MiniToolbarState): boolean {
  if (snapshot.readOnly || (snapshot.mode !== 'editing' && snapshot.mode !== 'suggesting')) return false;
  if (!(snapshot.selectionSettled && (snapshot.inputFocused || state.toolbarFocused))) return false;
  const range = snapshot.selection.displayRange;
  if (range ? range.from === range.to || !inBody(range.story) : snapshot.selection.formatting?.hasSelection !== true) return false;
  const key = selectionKey(snapshot.selection);
  return state.dismissed === null || (key !== null && key !== state.dismissed);
}

export type MiniToolbarView = 'hidden' | 'shown' | 'kept';

export function miniToolbarView(eligible: boolean, placement: { hidden: boolean; kept: boolean } | null): MiniToolbarView {
  if (!eligible || !placement || placement.hidden) return 'hidden';
  return placement.kept ? 'kept' : 'shown';
}

export function dismissMiniToolbar(state: MiniToolbarState, snapshot: Pick<MiniToolbarSnapshot, 'selection'>): MiniToolbarState {
  return { ...state, shown: false, dismissed: selectionKey(snapshot.selection) ?? '' };
}

/** Clears a dismissal once the selection moves to another range. */
export function afterSelectionChange(state: MiniToolbarState, selection: DocxPluginSelection): MiniToolbarState {
  const key = selectionKey(selection);
  return state.dismissed === null || key === null || key === state.dismissed ? state : { ...state, dismissed: null };
}

export function rovingIndex(key: string, current: number, enabled: readonly boolean[]): number | null {
  const indices = enabled.flatMap((on, index) => (on ? [index] : []));
  if (indices.length === 0) return null;
  switch (key) {
    case 'Home':
      return indices[0]!;
    case 'End':
      return indices[indices.length - 1]!;
    case 'ArrowRight':
      return indices.find((index) => index > current) ?? indices[0]!;
    case 'ArrowLeft':
      return [...indices].reverse().find((index) => index < current) ?? indices[indices.length - 1]!;
    default:
      return null;
  }
}

export function buttonState(
  item: DocxMiniToolbarItem,
  state: { enabled: boolean; active?: boolean | 'mixed' }
): { pressed: 'true' | 'false' | 'mixed' | undefined; disabled: boolean } {
  const pressed = !TOGGLES.has(item) ? undefined : state.active === 'mixed' ? 'mixed' : state.active === true ? 'true' : 'false';
  return { pressed, disabled: !state.enabled };
}

const PADDING = 4;
const GAP = 2;

export function miniToolbarSize(count: number, coarse: boolean): { width: number; height: number; button: number } {
  const button = coarse ? 44 : 32;
  return { button, width: count * button + (count - 1) * GAP + 2 * PADDING + 2, height: button + 2 * PADDING + 2 };
}

/** Keeps the toolbar's right edge on its page: the editor's comment button sits past that edge. */
export function avoidCommentButton(x: number, width: number, page: DocxPluginRect | null): number {
  if (!page) return x;
  return Math.min(x, Math.max(page.x, page.x + page.width - width));
}

export const MINI_TOOLBAR_GAP = GAP;
export const MINI_TOOLBAR_PADDING = PADDING;
