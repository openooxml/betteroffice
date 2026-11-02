import { describe, expect, test } from 'bun:test';
import {
  BUILT_IN_CHORDS,
  commandShortcut,
  DOCX_COMMAND_DESCRIPTORS,
  formatChord,
  matchesChord,
  normalizeChord,
  shortcutMatchesPlatform,
} from './descriptors';

function key(init: KeyboardEventInit): KeyboardEvent {
  return {
    key: '',
    code: '',
    metaKey: false,
    ctrlKey: false,
    shiftKey: false,
    altKey: false,
    ...init,
  } as KeyboardEvent;
}

function matchesClearFormatting(event: KeyboardEvent, mac: boolean): boolean {
  return DOCX_COMMAND_DESCRIPTORS.clearFormatting.shortcuts.some(
    (shortcut) => shortcutMatchesPlatform(shortcut, mac) && matchesChord(shortcut.chord, event, mac)
  );
}

describe('clear formatting shortcuts', () => {
  test('binds Ctrl+Space only outside macOS and Cmd+backslash only on macOS', () => {
    const ctrlSpace = key({ key: ' ', code: 'Space', ctrlKey: true });
    const cmdSpace = key({ key: ' ', code: 'Space', metaKey: true });
    const cmdBackslash = key({ key: '\\', code: 'Backslash', metaKey: true });
    const ctrlBackslash = key({ key: '\\', code: 'Backslash', ctrlKey: true });
    expect(matchesClearFormatting(ctrlSpace, false)).toBe(true);
    expect(matchesClearFormatting(ctrlSpace, true)).toBe(false);
    expect(matchesClearFormatting(cmdSpace, true)).toBe(false);
    expect(matchesClearFormatting(cmdBackslash, true)).toBe(true);
    expect(matchesClearFormatting(cmdBackslash, false)).toBe(false);
    expect(matchesClearFormatting(ctrlBackslash, false)).toBe(false);
  });

  test('recognizes space and backslash by key or physical code with exact modifiers', () => {
    expect(matchesChord('Mod+Space', key({ key: ' ', ctrlKey: true }), false)).toBe(true);
    expect(matchesChord('Mod+Space', key({ code: 'Space', ctrlKey: true }), false)).toBe(true);
    expect(matchesChord('Mod+\\', key({ key: '\\', metaKey: true }), true)).toBe(true);
    expect(matchesChord('Mod+\\', key({ code: 'Backslash', metaKey: true }), true)).toBe(true);
    expect(matchesChord('Mod+Space', key({ key: ' ', ctrlKey: true, shiftKey: true }), false)).toBe(false);
    expect(matchesChord('Mod+\\', key({ key: '\\', metaKey: true, ctrlKey: true }), true)).toBe(false);
  });

  test('displays the binding for the current platform', () => {
    expect(commandShortcut('clearFormatting', null, false)).toBe('Ctrl+Space');
    expect(commandShortcut('clearFormatting', null, true)).toBe('⌘\\');
    expect(formatChord('Mod+Space', false)).toBe('Ctrl+Space');
    expect(formatChord('Mod+Space', true)).toBe('⌘Space');
    expect(commandShortcut('bold', null, false)).toBe('Ctrl+B');
    expect(commandShortcut('alignment', { value: 'left' }, true)).toBe('⌘L');
    expect(commandShortcut('fontSize', { points: 12 }, false)).toBeNull();
  });

  test('reserves canonical chords from both platforms for built-in commands', () => {
    expect(normalizeChord('Mod+Space')).toBe('Mod+space');
    expect(normalizeChord('Mod+\\')).toBe('Mod+\\');
    expect(BUILT_IN_CHORDS.has('Mod+space')).toBe(true);
    expect(BUILT_IN_CHORDS.has('Mod+\\')).toBe(true);
    expect(shortcutMatchesPlatform({ chord: 'Mod+B' }, true)).toBe(true);
    expect(shortcutMatchesPlatform({ chord: 'Mod+B' }, false)).toBe(true);
  });
});
