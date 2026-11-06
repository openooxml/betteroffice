import { parseChord } from '../../../../shared/keyboard/chords';

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
