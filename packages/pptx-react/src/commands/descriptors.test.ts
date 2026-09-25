import { describe, expect, test } from 'bun:test';
import {
  commandShortcut,
  formatChord,
  matchesChord,
  PPTX_COMMAND_DESCRIPTORS,
} from './descriptors';

function key(init: KeyboardEventInit): KeyboardEvent {
  return {
    key: '',
    metaKey: false,
    ctrlKey: false,
    shiftKey: false,
    altKey: false,
    ...init,
  } as KeyboardEvent;
}

describe('PPTX command descriptors', () => {
  test('bind the editor shortcuts to commands', () => {
    const bindings = Object.values(PPTX_COMMAND_DESCRIPTORS).flatMap((descriptor) =>
      descriptor.shortcuts.map((shortcut) => `${shortcut.chord} ${descriptor.id}`)
    );
    expect(bindings.sort()).toEqual(
      [
        'Mod+B bold',
        'Mod+I italic',
        'Mod+U underline',
        'Mod+S save',
        'Mod+Z undo',
        'Mod+Shift+Z redo',
      ].sort()
    );
  });

  test('label chords for the platform instead of baking in Mac glyphs', () => {
    expect(formatChord('Mod+Shift+Z', true)).toBe('⇧⌘Z');
    expect(formatChord('Mod+Shift+Z', false)).toBe('Ctrl+Shift+Z');
    expect(formatChord('Mod+B', false)).toBe('Ctrl+B');
    expect(commandShortcut('fontSize', { points: 12 })).toBeNull();
  });

  test('match exact modifiers with Mod as Cmd on macOS and Ctrl elsewhere', () => {
    expect(matchesChord('Mod+Z', key({ key: 'z', metaKey: true }), true)).toBe(true);
    expect(matchesChord('Mod+Z', key({ key: 'z', ctrlKey: true }), true)).toBe(false);
    expect(matchesChord('Mod+Z', key({ key: 'z', ctrlKey: true }), false)).toBe(true);
    expect(matchesChord('Mod+Z', key({ key: 'Z', ctrlKey: true, shiftKey: true }), false)).toBe(
      false
    );
    expect(
      matchesChord('Mod+Shift+Z', key({ key: 'Z', ctrlKey: true, shiftKey: true }), false)
    ).toBe(true);
    expect(matchesChord('Mod+B', key({ key: 'b', ctrlKey: true, altKey: true }), false)).toBe(
      false
    );
  });
});
