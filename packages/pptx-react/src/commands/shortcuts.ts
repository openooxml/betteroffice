import type { TranslationKey } from '@betteroffice/pptx-i18n';
import {
  ariaChord,
  formatChord,
  isMacPlatform,
  parseChord,
  type ChordLabels,
} from '../../../../shared/keyboard/chords';
import {
  resolveShortcuts,
  type ResolvedShortcuts,
  type ShortcutBinding,
  type ShortcutDefault,
} from '../../../../shared/keyboard/resolve';
import { defaultArgs, isPluginCommandId, isPptxCommandId, PPTX_COMMAND_DESCRIPTORS } from './descriptors';
import type {
  PptxCommandArgs,
  PptxCommandId,
  PptxCommandShortcut,
  PptxCommandStore,
  PptxPluginCommandId,
} from './types';

type PptxNullArgsCommandId = {
  [K in PptxCommandId]: null extends PptxCommandArgs[K] ? K : never;
}[PptxCommandId];

/**
 * Every id a host can bind: commands that run without arguments and plugin commands.
 * @experimental
 */
export type PptxShortcutId = PptxNullArgsCommandId | PptxPluginCommandId;

/**
 * Host changes to the editor's shortcuts. An omitted id keeps its default; chords replace all of
 * its defaults; `null` disables it and leaves the keystroke to the host or browser. A chord taken
 * from another command's default wins over it. `cut`, `copy`, `paste` and `pasteAsPlainText` on a
 * chord other than their default use the async Clipboard API, which browsers may gate behind a
 * prompt. Text navigation and editing keys, select all and the context menu key cannot be bound.
 * @experimental
 */
export type PptxShortcuts = Readonly<Partial<Record<PptxShortcutId, ShortcutBinding>>>;

/** Keys the text input and slide surface move or edit with, whatever the modifiers. */
const INPUT_KEYS: ReadonlySet<string> = new Set([
  'arrowup',
  'arrowdown',
  'arrowleft',
  'arrowright',
  'home',
  'end',
  'pageup',
  'pagedown',
  'backspace',
  'delete',
  'enter',
  'tab',
  'escape',
]);

/** Whether the editor keeps a canonical chord whatever a host maps. */
export function inputKeepsChord(canonical: string): boolean {
  const parsed = parseChord(canonical);
  const command = parsed.mod || parsed.ctrl || parsed.alt;
  return (
    INPUT_KEYS.has(parsed.key) ||
    ((parsed.mod || parsed.ctrl) && parsed.key === 'a') ||
    (!command && parsed.shift && parsed.key === 'f10')
  );
}

/** Clipboard commands, which run as the browser's own gesture on their default chords. */
export const CLIPBOARD_IDS: ReadonlySet<string> = new Set(['cut', 'copy', 'paste', 'pasteAsPlainText']);

function shortcutDefaults(): ShortcutDefault[] {
  return Object.values(PPTX_COMMAND_DESCRIPTORS).flatMap((descriptor): ShortcutDefault[] => {
    const shortcuts = descriptor.shortcuts as readonly PptxCommandShortcut[];
    return shortcuts.length === 0
      ? []
      : [{ id: descriptor.id, chords: shortcuts.map(({ chord }) => chord) }];
  });
}

const NULL_ARGS: Readonly<Record<PptxNullArgsCommandId, true>> = {
  cut: true, copy: true, paste: true, pasteAsPlainText: true, deleteObject: true, deleteSlide: true,
  duplicateSlide: true, bold: true, italic: true, underline: true, strikethrough: true, insertImage: true,
  save: true, exportPng: true, slideshow: true, undo: true, redo: true, proposalsPanel: true,
};

function isBindable(id: string): boolean {
  return isPluginCommandId(id) || (isPptxCommandId(id) && Object.hasOwn(NULL_ARGS, id));
}

/**
 * The effective bindings a host map produces and what was wrong with it, for checking a map
 * outside the editor.
 * @experimental
 */
export function resolvePptxShortcuts(
  shortcuts: PptxShortcuts | undefined,
  options: { mac?: boolean } = {}
): ResolvedShortcuts {
  return resolveShortcuts(shortcutDefaults(), shortcuts, {
    mac: options.mac ?? isMacPlatform(),
    isBindable,
    isReserved: (chord) => inputKeepsChord(chord),
  });
}

/** The bindings of `id` in `resolved`, in descriptor form. */
export function effectiveCommandShortcuts(
  resolved: ResolvedShortcuts,
  id: PptxCommandId
): readonly PptxCommandShortcut[] {
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

function firstChord<K extends PptxCommandId>(
  store: PptxCommandStore,
  id: K,
  args: PptxCommandArgs[K] | undefined
): string | null {
  const key = JSON.stringify(args ?? defaultArgs(id));
  const shortcuts = store.getDescriptor(id).shortcuts as readonly PptxCommandShortcut[];
  return shortcuts.find((shortcut) => JSON.stringify(shortcut.args) === key)?.chord ?? null;
}

/** The chord `store` runs `id` with, as people read it, or null when it has none. */
export function storeShortcutLabel<K extends PptxCommandId>(
  store: PptxCommandStore,
  id: K,
  t: Translate,
  args?: PptxCommandArgs[K],
  mac = isMacPlatform()
): string | null {
  const chord = firstChord(store, id, args);
  return chord ? formatChord(chord, mac, chordLabels(t)) : null;
}

/** The same chord in `aria-keyshortcuts` form. */
export function storeAriaShortcut<K extends PptxCommandId>(
  store: PptxCommandStore,
  id: K,
  args?: PptxCommandArgs[K],
  mac = isMacPlatform()
): string | null {
  const chord = firstChord(store, id, args);
  return chord ? ariaChord(chord, mac) : null;
}
