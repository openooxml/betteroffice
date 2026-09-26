import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import type {
  XlsxCommandArgs,
  XlsxCommandDescriptor,
  XlsxCommandId,
  XlsxCommandShortcut,
} from './types';

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
  redo: descriptor('redo', 'toolbar.redo', true, [plain('Mod+Y'), plain('Mod+Shift+Z')]),
  zoom: descriptor('zoom', 'toolbar.zoom', false),
  proposalsPanel: descriptor('proposalsPanel', 'proposals.panelLabel', false),
  proposalAccept: descriptor('proposalAccept', 'proposals.accept', true),
  proposalReject: descriptor('proposalReject', 'proposals.reject', false),
});

export const XLSX_COMMAND_IDS = Object.freeze(
  Object.keys(XLSX_COMMAND_DESCRIPTORS) as XlsxCommandId[]
);

export function isXlsxCommandId(value: unknown): value is XlsxCommandId {
  return (
    typeof value === 'string' && Object.prototype.hasOwnProperty.call(XLSX_COMMAND_DESCRIPTORS, value)
  );
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
      default:
        break;
    }
  }
  return XLSX_COMMAND_DESCRIPTORS[id].labelKey;
}

interface ParsedChord {
  mod: boolean;
  shift: boolean;
  alt: boolean;
  key: string;
}

function parseChord(chord: string): ParsedChord {
  const parts = chord.split('+');
  const key = parts.length > 1 && parts[parts.length - 1] === '' ? '+' : parts[parts.length - 1];
  const modifiers = new Set(parts.slice(0, -1));
  return {
    mod: modifiers.has('Mod'),
    shift: modifiers.has('Shift'),
    alt: modifiers.has('Alt'),
    key: key.toLowerCase(),
  };
}

export function isMacPlatform(): boolean {
  return typeof navigator !== 'undefined' && /Mac|iPod|iPhone|iPad/.test(navigator.platform);
}

type ChordEvent = Pick<KeyboardEvent, 'key' | 'metaKey' | 'ctrlKey' | 'shiftKey' | 'altKey'>;

/** Whether a keydown event presses `chord`. */
export function matchesChord(chord: string, event: ChordEvent, mac = isMacPlatform()): boolean {
  const parsed = parseChord(chord);
  const mod = mac ? event.metaKey : event.ctrlKey;
  const other = mac ? event.ctrlKey : event.metaKey;
  return (
    parsed.mod === mod &&
    !other &&
    parsed.shift === event.shiftKey &&
    parsed.alt === event.altKey &&
    event.key.toLowerCase() === parsed.key
  );
}

/** A chord as people read it on the current platform (`Ctrl+B`, `⌘B`). */
export function formatChord(chord: string, mac = isMacPlatform()): string {
  const parsed = parseChord(chord);
  const key = parsed.key.length === 1 ? parsed.key.toUpperCase() : parsed.key;
  if (mac) {
    return `${parsed.mod ? '⌘' : ''}${parsed.alt ? '⌥' : ''}${parsed.shift ? '⇧' : ''}${key}`;
  }
  return [parsed.mod && 'Ctrl', parsed.alt && 'Alt', parsed.shift && 'Shift', key]
    .filter(Boolean)
    .join('+');
}

/** The display chord for a command invoked with `args`, if it has one. */
export function commandShortcut<K extends XlsxCommandId>(
  id: K,
  args?: XlsxCommandArgs[K]
): string | null {
  const shortcuts = XLSX_COMMAND_DESCRIPTORS[id].shortcuts as readonly XlsxCommandShortcut<K>[];
  const key = JSON.stringify(args ?? null);
  const match = shortcuts.find((shortcut) => JSON.stringify(shortcut.args) === key);
  return match ? formatChord(match.chord) : null;
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

/** The command a keydown event invokes, if any. */
export function commandForEvent(
  event: ChordEvent
): { id: XlsxCommandId; shortcut: XlsxCommandShortcut } | null {
  return XLSX_COMMAND_BINDINGS.find(({ shortcut }) => matchesChord(shortcut.chord, event)) ?? null;
}
