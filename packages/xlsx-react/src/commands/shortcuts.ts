import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import {
  ariaChord,
  formatChord,
  isMacPlatform,
  type ChordLabels,
} from '../../../../shared/keyboard/chords';
import {
  resolveShortcuts,
  type ResolvedShortcuts,
  type ShortcutBinding,
  type ShortcutDefault,
} from '../../../../shared/keyboard/resolve';
import {
  GRID_SHORTCUTS,
  gridKeepsChord,
  isPluginCommandId,
  isXlsxCommandId,
  XLSX_COMMAND_DESCRIPTORS,
} from './descriptors';
import { validArguments } from './evaluate';
import type {
  XlsxCommandArgs,
  XlsxCommandId,
  XlsxCommandShortcut,
  XlsxCommandStore,
  XlsxPluginCommandId,
} from './types';

type XlsxNullArgsCommandId = {
  [K in XlsxCommandId]: null extends XlsxCommandArgs[K] ? K : never;
}[XlsxCommandId];

/**
 * Every id a host can bind: commands that run without arguments and plugin commands.
 * @experimental
 */
export type XlsxShortcutId = XlsxNullArgsCommandId | XlsxPluginCommandId;

/**
 * Host changes to the editor's shortcuts. An omitted id keeps its default; chords replace all of
 * its defaults; `null` disables it and leaves the keystroke to the host or browser. A chord taken
 * from another command's default wins over it. `cut`, `copy`, `paste` and `pastePlainText` read
 * and write the clipboard with the async Clipboard API, which browsers may gate behind a prompt.
 * Navigation and editing keys of the grid, and select all, cannot be bound.
 * @experimental
 */
export type XlsxShortcuts = Readonly<Partial<Record<XlsxShortcutId, ShortcutBinding>>>;

/** Clipboard commands, which the grid runs itself on whatever chord holds them. */
export type XlsxClipboardKind = 'cut' | 'copy' | 'paste' | 'pastePlainText';

const GRID_GESTURES: ReadonlySet<string> = new Set<XlsxClipboardKind>(['cut', 'copy', 'paste', 'pastePlainText']);

export function isClipboardKind(id: string | undefined): id is XlsxClipboardKind {
  return id !== undefined && GRID_GESTURES.has(id);
}

function shortcutDefaults(): ShortcutDefault[] {
  return Object.values(XLSX_COMMAND_DESCRIPTORS).flatMap((descriptor): ShortcutDefault[] => {
    const shortcuts = descriptor.shortcuts as readonly XlsxCommandShortcut[];
    return shortcuts.length === 0
      ? []
      : [{ id: descriptor.id, chords: shortcuts.map(({ chord }) => chord) }];
  });
}

function isBindable(id: string): boolean {
  return isPluginCommandId(id) || (isXlsxCommandId(id) && validArguments(id, null));
}

/**
 * The effective bindings a host map produces and what was wrong with it, for checking a map
 * outside the editor.
 * @experimental
 */
export function resolveXlsxShortcuts(
  shortcuts: XlsxShortcuts | undefined,
  options: { mac?: boolean } = {}
): ResolvedShortcuts {
  return resolveShortcuts(shortcutDefaults(), shortcuts, {
    mac: options.mac ?? isMacPlatform(),
    isBindable,
    isReserved: (chord) => gridKeepsChord(chord),
  });
}

/** Whether the grid runs `id` itself when `chord` is pressed: the clipboard on its own keys. */
export function isGridGesture(id: string, chord: string, defaults: ResolvedShortcuts): boolean {
  return GRID_GESTURES.has(id) && defaults.owners.get(chord) === id;
}

/** The bindings of `id` in `resolved`, in descriptor form. */
export function effectiveCommandShortcuts(
  resolved: ResolvedShortcuts,
  id: XlsxCommandId
): readonly XlsxCommandShortcut[] {
  return (resolved.chords.get(id) ?? []).map((chord) => ({ chord, args: null as never }));
}

type Translate = (key: TranslationKey) => string;

/** Key names in the editor's language. */
export function chordLabels(t: Translate): ChordLabels {
  return {
    ctrl: t('keys.ctrl'),
    alt: t('keys.alt'),
    shift: t('keys.shift'),
    delete: t('keys.del'),
    escape: t('keys.escape'),
    pageup: t('keys.pageUp'),
    pagedown: t('keys.pageDown'),
  };
}

function firstChord<K extends XlsxCommandId>(
  store: XlsxCommandStore,
  id: K,
  args: XlsxCommandArgs[K] | undefined
): string | null {
  const key = JSON.stringify(args ?? null);
  const shortcuts = store.getDescriptor(id).shortcuts as readonly XlsxCommandShortcut[];
  const match = shortcuts.find((shortcut) => JSON.stringify(shortcut.args) === key)?.chord;
  return match ?? (key === 'null' ? (GRID_SHORTCUTS[id] ?? null) : null);
}

/** The chord `store` runs `id` with, as people read it, or null when it has none. */
export function storeShortcutLabel<K extends XlsxCommandId>(
  store: XlsxCommandStore,
  id: K,
  t: Translate,
  args?: XlsxCommandArgs[K],
  mac = isMacPlatform()
): string | null {
  const chord = firstChord(store, id, args);
  return chord ? formatChord(chord, mac, chordLabels(t)) : null;
}

/** The same chord in `aria-keyshortcuts` form. */
export function storeAriaShortcut<K extends XlsxCommandId>(
  store: XlsxCommandStore,
  id: K,
  args?: XlsxCommandArgs[K],
  mac = isMacPlatform()
): string | null {
  const chord = firstChord(store, id, args);
  return chord ? ariaChord(chord, mac) : null;
}
