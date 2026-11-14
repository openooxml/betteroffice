import {
  FUNCTION_KEY,
  isMacPlatform,
  NAMED_KEYS,
  normalizeChord,
  parseChord,
  type ChordEvent,
} from './format';

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
