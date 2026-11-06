import { describe, expect, test } from 'bun:test';
import { en } from '@betteroffice/xlsx-i18n';
import { createXlsxCommandController } from './createXlsxCommandStore';
import { resolveXlsxShortcuts, storeAriaShortcut, storeShortcutLabel } from './shortcuts';

function key(init: Partial<KeyboardEvent>): KeyboardEvent {
  return { key: '', code: '', metaKey: false, ctrlKey: false, shiftKey: false, altKey: false, ...init } as KeyboardEvent;
}

function t(path: string): string {
  return path.split('.').reduce<unknown>((node, part) => (node as Record<string, unknown>)[part], en) as string;
}

describe('resolveXlsxShortcuts', () => {
  test('binds the clipboard commands by default', () => {
    const { owners } = resolveXlsxShortcuts(undefined, { mac: false });
    expect(owners.get('Mod+c')).toBe('copy');
    expect(owners.get('Mod+Shift+v')).toBe('pastePlainText');
    expect(owners.get('Mod+b')).toBe('bold');
  });

  test('a command that needs arguments cannot be bound directly', () => {
    const { issues } = resolveXlsxShortcuts({ fontSize: 'Mod+Alt+F' } as never, { mac: false });
    expect(issues).toEqual([{ code: 'unknown-id', id: 'fontSize' }]);
  });

  test('any command without arguments can be given a chord', () => {
    const { owners, issues } = resolveXlsxShortcuts({ strikethrough: 'Mod+5', insertSheet: 'Shift+F11' }, { mac: false });
    expect(issues).toEqual([]);
    expect(owners.get('Mod+5')).toBe('strikethrough');
    expect(owners.get('Shift+f11')).toBe('insertSheet');
  });

  test('keeps the keys the grid navigates and selects with', () => {
    const { issues } = resolveXlsxShortcuts({ bold: ['Mod+ArrowLeft', 'Mod+A', 'Alt+Enter', 'F2'] }, { mac: false });
    expect(issues.map(({ code }) => code)).toEqual(['reserved-chord', 'reserved-chord', 'reserved-chord', 'reserved-chord']);
  });

  test('keeps region cycling and the context menu key', () => {
    const { issues, owners } = resolveXlsxShortcuts({ bold: ['F6', 'Shift+F6', 'Shift+F10', 'F10', 'Mod+F6'] }, { mac: false });
    expect(issues.map(({ code, chord }) => [code, chord])).toEqual([
      ['reserved-chord', 'F6'],
      ['reserved-chord', 'Shift+F6'],
      ['reserved-chord', 'Shift+F10'],
    ]);
    expect(owners.get('f10')).toBe('bold');
    expect(owners.get('Mod+f6')).toBe('bold');
  });

  test('a host chord takes a clipboard default', () => {
    const { owners, chords, issues } = resolveXlsxShortcuts({ bold: 'Mod+C' }, { mac: false });
    expect(owners.get('Mod+c')).toBe('bold');
    expect(chords.get('copy')).toEqual([]);
    expect(issues).toEqual([{ code: 'displaced-default', id: 'bold', chord: 'Mod+c', other: 'copy' }]);
  });

  test('accepts plugin ids before the plugin is active', () => {
    const { owners } = resolveXlsxShortcuts({ 'plugin:acme/run': 'Mod+Alt+R' }, { mac: false });
    expect(owners.get('Mod+Alt+r')).toBe('plugin:acme/run');
  });
});

describe('command controller', () => {
  test('descriptors and dispatch follow the host map', () => {
    const controller = createXlsxCommandController();
    controller.setShortcuts({ paste: 'Mod+Alt+V', bold: null });
    expect(controller.store.getDescriptor('paste').shortcuts.map(({ chord }) => chord)).toEqual(['Mod+Alt+v']);
    expect(controller.store.getDescriptor('bold').shortcuts).toEqual([]);
    expect(controller.shortcutAt(key({ key: 'b', ctrlKey: true }), false)).toBeNull();
    expect(controller.shortcutAt(key({ key: 'v', ctrlKey: true, altKey: true }), false)).toEqual({
      id: 'paste',
      args: null,
      chord: 'Mod+Alt+v',
    });
  });

  test('the grid keeps the clipboard keys until the host moves them', () => {
    const controller = createXlsxCommandController();
    const copy = key({ key: 'c', ctrlKey: true });
    expect(controller.shortcutAt(copy, false)).toBeNull();
    expect(controller.claimsKey(copy, false)).toBe(false);
    expect(controller.claimsKey(key({ key: 'b', ctrlKey: true }), false)).toBe(true);
    controller.setShortcuts({ copy: 'Mod+Alt+C' });
    expect(controller.claimsKey(copy, false)).toBe(true);
    expect(controller.shortcutAt(copy, false)).toBeNull();
    const moved = key({ key: 'c', ctrlKey: true, altKey: true });
    expect(controller.shortcutAt(moved, false)?.id).toBe('copy');
    expect(controller.claimsKey(moved, false)).toBe(false);
    expect(controller.clipboardKey(moved, false)).toBe('copy');
    expect(controller.clipboardKey(copy, false)).toBeNull();
    controller.setShortcuts({ copy: null, bold: 'Mod+X' });
    expect(controller.claimsKey(copy, false)).toBe(true);
    expect(controller.releasedChord('Mod+c')).toBe(true);
    expect(controller.clipboardKey(key({ key: 'x', ctrlKey: true }), false)).toBeNull();
    expect(controller.claimsKey(key({ key: 'x', ctrlKey: true }), false)).toBe(true);
  });

  test('the grid names the clipboard action of a keystroke', () => {
    const controller = createXlsxCommandController();
    expect(controller.clipboardKey(key({ key: 'c', ctrlKey: true }), false)).toBe('copy');
    expect(controller.clipboardKey(key({ key: 'v', ctrlKey: true, shiftKey: true }), false)).toBe('pastePlainText');
    expect(controller.clipboardKey(key({ key: 'v', ctrlKey: true, altKey: true }), false)).toBe('paste');
    expect(controller.clipboardKey(key({ key: 'c', metaKey: true }), true)).toBe('copy');
    expect(controller.clipboardKey(key({ key: 'b', ctrlKey: true }), false)).toBeNull();
    expect(controller.clipboardKey(key({ key: 'c' }), false)).toBeNull();
    controller.setShortcuts({ cut: null, copy: 'Mod+Alt+C' });
    expect(controller.clipboardKey(key({ key: 'x', ctrlKey: true, shiftKey: true }), false)).toBeNull();
    expect(controller.clipboardKey(key({ key: 'c', ctrlKey: true, shiftKey: true }), false)).toBeNull();
    expect(controller.clipboardKey(key({ key: 'v', ctrlKey: true, shiftKey: true, altKey: true }), false)).toBe('paste');
  });

  test('a chord naming the typed character wins over a key-position match', () => {
    const controller = createXlsxCommandController();
    controller.setShortcuts({ bold: 'Mod+Shift+1', italic: 'Mod+Shift+!' });
    const bang = key({ key: '!', code: 'Digit1', ctrlKey: true, shiftKey: true });
    expect(controller.shortcutAt(bang, false)?.id).toBe('italic');
    const one = key({ key: '1', code: 'Digit1', ctrlKey: true, shiftKey: true });
    expect(controller.shortcutAt(one, false)?.id).toBe('bold');
  });

  test('a layout without Latin letters still presses the chord', () => {
    const controller = createXlsxCommandController();
    expect(controller.shortcutAt(key({ key: 'и', code: 'KeyB', ctrlKey: true }), false)?.id).toBe('bold');
  });

  test('plugin chords stay with the plugin dispatcher and bump the revision', () => {
    const controller = createXlsxCommandController();
    const before = controller.shortcutRevision();
    let notified = 0;
    controller.subscribeShortcuts(() => { notified += 1; });
    controller.setShortcuts({ 'plugin:acme/run': 'Mod+Alt+R', italic: null });
    expect(controller.shortcutRevision()).toBeGreaterThan(before);
    expect(notified).toBe(1);
    expect(controller.hostPluginChords('plugin:acme/run')).toEqual(['Mod+Alt+r']);
    expect(controller.hostPluginChords('plugin:acme/other')).toBeUndefined();
    expect(controller.shortcutAt(key({ key: 'r', ctrlKey: true, altKey: true }), false)).toBeNull();
    expect(controller.shortcutOwner('Mod+i')).toBeUndefined();
  });

  test('labels and aria names show the effective chord', () => {
    const controller = createXlsxCommandController();
    controller.setShortcuts({ paste: 'Mod+Alt+V', copy: null });
    expect(storeShortcutLabel(controller.store, 'paste', t, undefined, false)).toBe('Ctrl+Alt+V');
    expect(storeShortcutLabel(controller.store, 'paste', t, undefined, true)).toBe('⌥⌘V');
    expect(storeAriaShortcut(controller.store, 'paste', undefined, false)).toBe('Control+Alt+V');
    expect(storeAriaShortcut(controller.store, 'paste', undefined, true)).toBe('Meta+Alt+V');
    expect(storeShortcutLabel(controller.store, 'copy', t, undefined, false)).toBeNull();
    expect(storeShortcutLabel(controller.store, 'clearContents', t, undefined, false)).toBe('Del');
  });

  test('key names follow the locale', () => {
    const controller = createXlsxCommandController();
    const names: Record<string, string> = { 'keys.ctrl': 'Strg', 'keys.shift': 'Umschalt' };
    const german = (path: string) => names[path] ?? t(path);
    expect(storeShortcutLabel(controller.store, 'redo', german, undefined, false)).toBe('Strg+Umschalt+Z');
  });
});
