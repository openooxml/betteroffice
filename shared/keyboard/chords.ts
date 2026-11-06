/**
 * A chord such as `Mod+Shift+V`. `Mod` is Cmd on macOS and Ctrl elsewhere; `Ctrl` is the physical
 * Control key on macOS and the same as `Mod` elsewhere.
 * @experimental
 */
export type ShortcutChord = string;

export interface ParsedChord {
  ctrl: boolean;
  mod: boolean;
  shift: boolean;
  alt: boolean;
  key: string;
}

/** The subset of a keydown event chords are matched against. */
export type ChordEvent = Pick<
  KeyboardEvent,
  'key' | 'code' | 'ctrlKey' | 'metaKey' | 'shiftKey' | 'altKey'
>;

/** Localized names of keys, for platforms that spell chords out. */
export type ChordLabels = Readonly<Partial<Record<string, string>>>;

const MODIFIERS: ReadonlySet<string> = new Set(['Ctrl', 'Mod', 'Shift', 'Alt']);
const NAMED_KEYS: Readonly<Record<string, string>> = {
  space: 'Space',
  enter: 'Enter',
  tab: 'Tab',
  escape: 'Esc',
  backspace: 'Backspace',
  delete: 'Del',
  home: 'Home',
  end: 'End',
  pageup: 'Page Up',
  pagedown: 'Page Down',
  arrowup: '↑',
  arrowdown: '↓',
  arrowleft: '←',
  arrowright: '→',
};
const FUNCTION_KEY = /^f([1-9]|1[0-2])$/;

export function isMacPlatform(): boolean {
  return typeof navigator !== 'undefined' && /Mac|iPod|iPhone|iPad/.test(navigator.platform);
}

export function parseChord(chord: string): ParsedChord {
  const parts = chord.split('+');
  const key = parts.length > 1 && parts[parts.length - 1] === '' ? '+' : parts[parts.length - 1];
  const modifiers = new Set(parts.slice(0, -1));
  return {
    ctrl: modifiers.has('Ctrl'),
    mod: modifiers.has('Mod'),
    shift: modifiers.has('Shift'),
    alt: modifiers.has('Alt'),
    key: key.toLowerCase(),
  };
}

/**
 * `chord` in canonical form (`Ctrl+Mod+Alt+Shift+key`), or null when it is not a chord. `+` counts
 * as `=`, whose matcher also accepts the `+` key.
 */
export function normalizeChord(chord: string): string | null {
  if (typeof chord !== 'string' || chord.length === 0) return null;
  const parts = chord.split('+');
  const modifiers =
    parts.length > 1 && parts[parts.length - 1] === '' ? parts.slice(0, -2) : parts.slice(0, -1);
  if (!modifiers.every((modifier) => MODIFIERS.has(modifier))) return null;
  const parsed = parseChord(chord);
  if (parsed.key.length === 0) return null;
  const key = parsed.key === '+' ? '=' : parsed.key;
  return [parsed.ctrl && 'Ctrl', parsed.mod && 'Mod', parsed.alt && 'Alt', parsed.shift && 'Shift', key]
    .filter(Boolean)
    .join('+');
}

/** `chord` in canonical form for one platform: outside macOS `Ctrl` and `Mod` are one key. */
export function platformChord(chord: string, mac: boolean): string | null {
  const canonical = normalizeChord(chord);
  if (!canonical || mac) return canonical;
  const parsed = parseChord(canonical);
  return normalizeChord(
    [(parsed.ctrl || parsed.mod) && 'Mod', parsed.alt && 'Alt', parsed.shift && 'Shift', parsed.key]
      .filter(Boolean)
      .join('+')
  );
}

/** Whether `key` is one a host may bind: one printable character or a named key. */
export function isBindableKey(key: string): boolean {
  return (
    [...key].length === 1 || Object.hasOwn(NAMED_KEYS, key) || FUNCTION_KEY.test(key)
  );
}

/** Whether a chord leaves typing alone: it holds a command modifier or is a function key. */
export function isCommandChord(chord: string): boolean {
  const parsed = parseChord(chord);
  return parsed.ctrl || parsed.mod || parsed.alt || FUNCTION_KEY.test(parsed.key);
}

function keyMatches(key: string, event: ChordEvent): boolean {
  if (key === '=') return event.code === 'Equal' || event.key === '=' || event.key === '+';
  if (key === 'space') return event.key === ' ' || event.code === 'Space';
  if (key === '\\') return event.key === '\\' || event.code === 'Backslash';
  if (event.key.toLowerCase() === key) return true;
  if (/^[0-9]$/.test(key)) return event.code === `Digit${key}`;
  if (!/^[a-z]$/.test(key)) return false;
  const latin = /^[\x20-\x7e]$/.test(event.key);
  return !latin && event.code === `Key${key.toUpperCase()}`;
}

/** Whether a keydown event presses `chord`. */
export function matchesChord(chord: string, event: ChordEvent, mac = isMacPlatform()): boolean {
  const parsed = parseChord(chord);
  const mod = mac ? event.metaKey : event.ctrlKey;
  const other = mac ? event.ctrlKey : event.metaKey;
  const wantsMod = parsed.mod || (!mac && parsed.ctrl);
  const wantsOther = mac && parsed.ctrl;
  return (
    wantsMod === mod &&
    wantsOther === other &&
    parsed.shift === event.shiftKey &&
    parsed.alt === event.altKey &&
    keyMatches(parsed.key, event)
  );
}

function keyLabel(key: string, labels: ChordLabels | undefined): string {
  const named = NAMED_KEYS[key];
  if (named) return labels?.[key] ?? named;
  return key.length === 1 || FUNCTION_KEY.test(key) ? key.toUpperCase() : key;
}

/**
 * A chord as people read it: `⌃⌥⇧⌘B` on macOS in Apple's modifier order, otherwise `Ctrl+Shift+B`
 * with the modifier names from `labels`.
 */
export function formatChord(
  chord: string,
  mac = isMacPlatform(),
  labels?: ChordLabels
): string {
  const parsed = parseChord(chord);
  const key = keyLabel(parsed.key, labels);
  if (mac) {
    return `${parsed.ctrl ? '⌃' : ''}${parsed.alt ? '⌥' : ''}${parsed.shift ? '⇧' : ''}${parsed.mod ? '⌘' : ''}${key}`;
  }
  return [
    (parsed.mod || parsed.ctrl) && (labels?.ctrl ?? 'Ctrl'),
    parsed.alt && (labels?.alt ?? 'Alt'),
    parsed.shift && (labels?.shift ?? 'Shift'),
    key,
  ]
    .filter(Boolean)
    .join('+');
}

const ARIA_KEYS: Readonly<Record<string, string>> = {
  space: 'Space',
  enter: 'Enter',
  tab: 'Tab',
  escape: 'Escape',
  backspace: 'Backspace',
  delete: 'Delete',
  home: 'Home',
  end: 'End',
  pageup: 'PageUp',
  pagedown: 'PageDown',
  arrowup: 'ArrowUp',
  arrowdown: 'ArrowDown',
  arrowleft: 'ArrowLeft',
  arrowright: 'ArrowRight',
};

/** A chord in the unlocalized form `aria-keyshortcuts` takes, such as `Control+Shift+V`. */
export function ariaChord(chord: string, mac = isMacPlatform()): string {
  const parsed = parseChord(chord);
  const key = ARIA_KEYS[parsed.key] ?? (parsed.key.length === 1 || FUNCTION_KEY.test(parsed.key) ? parsed.key.toUpperCase() : parsed.key);
  return [
    (mac ? parsed.ctrl : parsed.ctrl || parsed.mod) && 'Control',
    mac && parsed.mod && 'Meta',
    parsed.alt && 'Alt',
    parsed.shift && 'Shift',
    key,
  ]
    .filter(Boolean)
    .join('+');
}
