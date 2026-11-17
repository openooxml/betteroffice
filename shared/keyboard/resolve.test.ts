import { describe, expect, test } from 'bun:test';
import { ariaChord, formatChord, matchesChord, normalizeChord, platformChord } from './chords';
import { resolveShortcuts, type ShortcutDefault } from './resolve';

const DEFAULTS: ShortcutDefault[] = [
  { id: 'bold', chords: ['Mod+B'] },
  { id: 'redo', chords: ['Mod+Shift+Z', 'Mod+Y'] },
  { id: 'pasteAsPlainText', chords: ['Mod+Shift+V'] },
];
const options = { mac: false, isBindable: (id: string) => id === 'strikethrough' };

function key(init: Partial<KeyboardEvent>): KeyboardEvent {
  return { key: '', code: '', metaKey: false, ctrlKey: false, shiftKey: false, altKey: false, ...init } as KeyboardEvent;
}

describe('resolveShortcuts', () => {
  test('keeps defaults for omitted ids and replaces every default of a mapped id', () => {
    const { chords, issues } = resolveShortcuts(DEFAULTS, { redo: 'Mod+Alt+Z' }, options);
    expect(chords.get('bold')).toEqual(['Mod+b']);
    expect(chords.get('redo')).toEqual(['Mod+Alt+z']);
    expect(issues).toEqual([]);
  });

  test('null disables an id and frees its chords', () => {
    const { chords, owners } = resolveShortcuts(DEFAULTS, { pasteAsPlainText: null }, options);
    expect(chords.get('pasteAsPlainText')).toEqual([]);
    expect(owners.has('Mod+Shift+v')).toBe(false);
  });

  test('the per-platform form keeps the default of the absent platform', () => {
    const map = { bold: { mac: 'Mod+Alt+B' } };
    expect(resolveShortcuts(DEFAULTS, map, { ...options, mac: true }).chords.get('bold')).toEqual(['Mod+Alt+b']);
    expect(resolveShortcuts(DEFAULTS, map, options).chords.get('bold')).toEqual(['Mod+b']);
  });

  test('reports unknown ids and invalid chords, keeping the default', () => {
    const { chords, issues } = resolveShortcuts(
      DEFAULTS,
      { nope: 'Mod+K', bold: 'Hyper+B', redo: 'Q' },
      options
    );
    expect(issues.map(({ code, id }) => [code, id])).toEqual([
      ['unknown-id', 'nope'],
      ['invalid-chord', 'bold'],
      ['invalid-chord', 'redo'],
    ]);
    expect(chords.get('bold')).toEqual(['Mod+b']);
    expect(chords.get('redo')).toEqual(['Mod+Shift+z', 'Mod+y']);
  });

  test('accepts function keys alone and refuses reserved chords', () => {
    const { chords, issues } = resolveShortcuts(DEFAULTS, { bold: 'F6', redo: 'Mod+V' }, {
      ...options,
      isReserved: (chord) => chord === 'Mod+v',
    });
    expect(chords.get('bold')).toEqual(['f6']);
    expect(issues).toEqual([{ code: 'reserved-chord', id: 'redo', chord: 'Mod+V' }]);
  });

  test('a host chord takes a default chord and reports the displaced id', () => {
    const { chords, owners, issues } = resolveShortcuts(DEFAULTS, { pasteAsPlainText: 'Mod+Y' }, options);
    expect(owners.get('Mod+y')).toBe('pasteAsPlainText');
    expect(chords.get('redo')).toEqual(['Mod+Shift+z']);
    expect(issues).toEqual([
      { code: 'displaced-default', id: 'pasteAsPlainText', chord: 'Mod+y', other: 'redo' },
    ]);
  });

  test('a chord two host entries claim is dropped from both', () => {
    const { chords, owners, issues } = resolveShortcuts(
      DEFAULTS,
      { bold: 'Mod+K', strikethrough: 'Mod+K' },
      options
    );
    expect(owners.has('Mod+k')).toBe(false);
    expect(chords.get('bold')).toEqual(['Mod+b']);
    expect(issues.map(({ code }) => code)).toEqual(['duplicate-chord', 'duplicate-chord']);
  });

  test('an id without a default can be given a chord', () => {
    expect(resolveShortcuts(DEFAULTS, { strikethrough: 'Mod+Alt+S' }, options).owners.get('Mod+Alt+s'))
      .toBe('strikethrough');
  });
});

describe('chords', () => {
  test('Ctrl is Mod outside macOS and its own key on macOS', () => {
    expect(platformChord('Ctrl+V', false)).toBe('Mod+v');
    expect(platformChord('Ctrl+V', true)).toBe('Ctrl+v');
    expect(matchesChord('Ctrl+Mod+Space', key({ key: ' ', ctrlKey: true, metaKey: true }), true)).toBe(true);
    expect(matchesChord('Mod+B', key({ key: 'b', ctrlKey: true }), true)).toBe(false);
  });

  test('matches by physical key when the layout or Option yields no Latin letter', () => {
    expect(matchesChord('Mod+B', key({ key: 'и', code: 'KeyB', ctrlKey: true }), false)).toBe(true);
    expect(matchesChord('Mod+Alt+V', key({ key: '√', code: 'KeyV', metaKey: true, altKey: true }), true)).toBe(true);
    expect(matchesChord('Mod+Alt+E', key({ key: 'Dead', code: 'KeyE', metaKey: true, altKey: true }), true)).toBe(true);
  });

  test('a Latin layout keeps matching by letter, not position', () => {
    expect(matchesChord('Mod+Z', key({ key: 'y', code: 'KeyZ', ctrlKey: true }), false)).toBe(false);
  });

  test('formats in Apple order on macOS and with localized names elsewhere', () => {
    expect(formatChord('Mod+Alt+Shift+V', true)).toBe('⌥⇧⌘V');
    expect(formatChord('Mod+Shift+V', false, { ctrl: 'Strg', shift: 'Umschalt' })).toBe('Strg+Umschalt+V');
  });

  test('aria form is unlocalized and names the physical modifier', () => {
    expect(ariaChord('Mod+Shift+V', false)).toBe('Control+Shift+V');
    expect(ariaChord('Mod+Shift+V', true)).toBe('Meta+Shift+V');
  });

  test('canonical form orders modifiers', () => {
    expect(normalizeChord('Shift+Alt+Mod+K')).toBe('Mod+Alt+Shift+k');
    expect(normalizeChord('Hyper+K')).toBeNull();
  });
});
