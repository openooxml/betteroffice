import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import type {
  XlsxCommandArgs,
  XlsxCommandDescriptor,
  XlsxCommandId,
  XlsxCommandShortcut,
  XlsxPluginCommandId,
} from './types';
import {
  normalizeChord as normalizeHostChord,
  parseChord,
} from '../../../../shared/keyboard/chords';

export { formatChord, isMacPlatform, matchesChord } from '../../../../shared/keyboard/chords';

/** `chord` in canonical form, or null; contributed chords name `Mod`, never a physical `Ctrl`. */
export function normalizeChord(chord: string): string | null {
  const canonical = normalizeHostChord(chord);
  return canonical && !parseChord(canonical).ctrl ? canonical : null;
}

type DescriptorTable = { readonly [K in XlsxCommandId]: XlsxCommandDescriptor<K> };

function descriptor<K extends XlsxCommandId>(
  id: K,
  labelKey: TranslationKey,
  mutatesDocument: boolean,
  shortcuts: readonly XlsxCommandShortcut<K>[] = []
): XlsxCommandDescriptor<K> {
  return Object.freeze({ id, labelKey, mutatesDocument, shortcuts: Object.freeze(shortcuts) });
}

const plain = <K extends XlsxCommandId>(chord: string) =>
  ({ chord, args: null }) as XlsxCommandShortcut<K>;

/** Every XLSX command, with its label and keyboard bindings. */
export const XLSX_COMMAND_DESCRIPTORS: DescriptorTable = Object.freeze({
  bold: descriptor('bold', 'toolbar.bold', true, [plain('Mod+B')]),
  italic: descriptor('italic', 'toolbar.italic', true, [plain('Mod+I')]),
  strikethrough: descriptor('strikethrough', 'toolbar.strikethrough', true),
  paintFormat: descriptor('paintFormat', 'toolbar.paintFormat', true),
  fontFamily: descriptor('fontFamily', 'toolbar.fontFamily', true),
  fontSize: descriptor('fontSize', 'toolbar.fontSize', true),
  fontSizeStep: descriptor('fontSizeStep', 'commands.fontSizeStep', true),
  textColor: descriptor('textColor', 'toolbar.textColor', true),
  fillColor: descriptor('fillColor', 'toolbar.fillColor', true),
  numberFormat: descriptor('numberFormat', 'commands.numberFormat', true),
  decimalPlaces: descriptor('decimalPlaces', 'commands.decimalPlaces', true),
  borderPreset: descriptor('borderPreset', 'toolbar.borders', true),
  borderStyle: descriptor('borderStyle', 'commands.borderStyle', true),
  borderColor: descriptor('borderColor', 'toolbar.borderColor', true),
  merge: descriptor('merge', 'toolbar.mergeCells', true),
  horizontalAlignment: descriptor('horizontalAlignment', 'toolbar.horizontalAlignment', true),
  verticalAlignment: descriptor('verticalAlignment', 'toolbar.verticalAlignment', true),
  textWrapping: descriptor('textWrapping', 'toolbar.textWrapping', true),
  searchMenus: descriptor('searchMenus', 'toolbar.searchMenus', false),
  save: descriptor('save', 'toolbar.save', false, [plain('Mod+S')]),
  exportPng: descriptor('exportPng', 'toolbar.exportPng', false),
  print: descriptor('print', 'toolbar.print', false),
  undo: descriptor('undo', 'toolbar.undo', true, [plain('Mod+Z')]),
  redo: descriptor('redo', 'toolbar.redo', true, [plain('Mod+Shift+Z'), plain('Mod+Y')]),
  zoom: descriptor('zoom', 'toolbar.zoom', false),
  proposalsPanel: descriptor('proposalsPanel', 'proposals.panelLabel', false),
  proposalAccept: descriptor('proposalAccept', 'proposals.accept', true),
  proposalReject: descriptor('proposalReject', 'proposals.reject', false),
  cut: descriptor('cut', 'commands.cut', true, [plain('Mod+X')]),
  copy: descriptor('copy', 'commands.copy', false, [plain('Mod+C')]),
  paste: descriptor('paste', 'commands.paste', true, [plain('Mod+V')]),
  pastePlainText: descriptor('pastePlainText', 'commands.pastePlainText', true, [
    plain('Mod+Shift+V'),
  ]),
  clearContents: descriptor('clearContents', 'commands.clearContents', true),
  insertRows: descriptor('insertRows', 'commands.insertRowsAbove', true),
  insertColumns: descriptor('insertColumns', 'commands.insertColumnsLeft', true),
  deleteRows: descriptor('deleteRows', 'commands.deleteRows', true),
  deleteColumns: descriptor('deleteColumns', 'commands.deleteColumns', true),
  insertSheet: descriptor('insertSheet', 'commands.insertSheet', true),
  deleteSheet: descriptor('deleteSheet', 'commands.deleteSheet', true),
  renameSheet: descriptor('renameSheet', 'commands.renameSheet', true),
  moveSheet: descriptor('moveSheet', 'commands.moveSheetLeft', true),
});

/** Keys the grid handles itself, shown next to their commands in menus but never remappable. */
export const GRID_SHORTCUTS: Partial<Record<XlsxCommandId, string>> = {
  clearContents: 'Delete',
};

export const XLSX_COMMAND_IDS = Object.freeze(
  Object.keys(XLSX_COMMAND_DESCRIPTORS) as XlsxCommandId[]
);

export function isXlsxCommandId(value: unknown): value is XlsxCommandId {
  return (
    typeof value === 'string' && Object.prototype.hasOwnProperty.call(XLSX_COMMAND_DESCRIPTORS, value)
  );
}

const PLUGIN_COMMAND_ID =
  /^plugin:[A-Za-z0-9](?:[A-Za-z0-9._-]*[A-Za-z0-9])?\/[A-Za-z0-9](?:[A-Za-z0-9._-]*[A-Za-z0-9])?$/;

export function isPluginCommandId(value: unknown): value is XlsxPluginCommandId {
  return typeof value === 'string' && PLUGIN_COMMAND_ID.test(value);
}

/** The label of a command bound to `args`, such as "Format as currency". */
export function commandLabelKey<K extends XlsxCommandId>(
  id: K,
  args?: XlsxCommandArgs[K]
): TranslationKey {
  const bound = args as Record<string, unknown> | null | undefined;
  if (bound && typeof bound === 'object') {
    const value = bound.value;
    const direction = bound.direction;
    switch (id) {
      case 'numberFormat':
        if (value === 'currency') return 'toolbar.currency';
        if (value === 'percent') return 'toolbar.percent';
        return `toolbar.numberFormats.${value}` as TranslationKey;
      case 'decimalPlaces':
        return direction === 'increase' ? 'toolbar.increaseDecimal' : 'toolbar.decreaseDecimal';
      case 'fontSizeStep':
        return direction === 'increase' ? 'toolbar.increaseFontSize' : 'toolbar.decreaseFontSize';
      case 'merge':
        return `toolbar.merge.${value}` as TranslationKey;
      case 'borderPreset':
        return `toolbar.borderPresets.${value}` as TranslationKey;
      case 'borderStyle':
        return `toolbar.borderStyles.${value}` as TranslationKey;
      case 'horizontalAlignment':
        return `toolbar.horizontalAlign.${value}` as TranslationKey;
      case 'verticalAlignment':
        return `toolbar.verticalAlign.${value}` as TranslationKey;
      case 'textWrapping':
        return `toolbar.wrapping.${value}` as TranslationKey;
      case 'insertRows':
        return bound.position === 'below' ? 'commands.insertRowsBelow' : 'commands.insertRowsAbove';
      case 'insertColumns':
        return bound.position === 'right' ? 'commands.insertColumnsRight' : 'commands.insertColumnsLeft';
      case 'moveSheet':
        return direction === 'right' ? 'commands.moveSheetRight' : 'commands.moveSheetLeft';
      default:
        break;
    }
  }
  return XLSX_COMMAND_DESCRIPTORS[id].labelKey;
}

/** Every shortcut binding, in descriptor order. */
export const XLSX_COMMAND_BINDINGS: readonly { id: XlsxCommandId; shortcut: XlsxCommandShortcut }[] =
  Object.freeze(
    Object.values(XLSX_COMMAND_DESCRIPTORS).flatMap((entry) =>
      (entry.shortcuts as readonly XlsxCommandShortcut[]).map((shortcut) => ({
        id: entry.id as XlsxCommandId,
        shortcut,
      }))
    )
  );

/** Canonical chords of the built-in commands, which always win over contributed ones. */
export const BUILT_IN_CHORDS: ReadonlySet<string> = new Set(
  XLSX_COMMAND_BINDINGS.map(({ shortcut }) => normalizeChord(shortcut.chord)!)
);

const FUNCTION_KEY = /^f(?:[1-9]|1[0-2])$/;

/** Whether a canonical chord may be a plugin shortcut: it uses Mod or Alt, or is F1 to F12. */
export function isPluginChord(canonical: string): boolean {
  const parsed = parseChord(canonical);
  return parsed.mod || parsed.alt || FUNCTION_KEY.test(parsed.key);
}

const GRID_KEYS: ReadonlySet<string> = new Set([
  'arrowup',
  'arrowdown',
  'arrowleft',
  'arrowright',
  'tab',
  'enter',
  'escape',
  'home',
  'end',
  'pageup',
  'pagedown',
  'f2',
  'delete',
  'backspace',
]);
const GRID_MOD_KEYS: ReadonlySet<string> = new Set(['a', 'c', 'v', 'x']);

/**
 * Whether the grid keeps a canonical chord whatever a host maps: navigation and editing keys
 * with any modifiers, select all, region cycling (F6) and the context menu (Shift+F10). The clipboard chords are commands a host may move.
 */
export function gridKeepsChord(canonical: string): boolean {
  const parsed = parseChord(canonical);
  const command = parsed.mod || parsed.ctrl || parsed.alt;
  return (
    GRID_KEYS.has(parsed.key) ||
    ((parsed.mod || parsed.ctrl) && parsed.key === 'a') ||
    (!command && parsed.key === 'f6') ||
    (!command && parsed.shift && parsed.key === 'f10')
  );
}

/**
 * Whether the grid consumes a canonical chord before shortcuts are dispatched: navigation and
 * editing keys with any modifiers, typing, select all and the clipboard.
 */
export function gridOwnsChord(canonical: string): boolean {
  const parsed = parseChord(canonical);
  if (GRID_KEYS.has(parsed.key)) return true;
  if (parsed.mod) return GRID_MOD_KEYS.has(parsed.key);
  return !parsed.alt && parsed.key.length === 1;
}
