import { describe, expect, test } from 'bun:test';
import { en } from '@betteroffice/pptx-i18n';
import { createPptxCommandController } from './createPptxCommandStore';
import { resolvePptxShortcuts } from './shortcuts';
import { storeAriaShortcut, storeShortcutLabel } from './shortcutLabels';
import { hostPluginChords, releasedChord, shortcutOwner, subscribeShortcuts } from './shortcutTable';
import { applyShortcuts, pickChord, pressedChord, shortcutAt } from './useCommandShortcuts';

function key(init: Partial<KeyboardEvent>): KeyboardEvent {
  return { key: '', code: '', metaKey: false, ctrlKey: false, shiftKey: false, altKey: false, ...init } as KeyboardEvent;
}

function t(path: string): string {
  return path.split('.').reduce<unknown>((node, part) => (node as Record<string, unknown>)[part], en) as string;
}

describe('resolvePptxShortcuts', () => {
  test('binds the clipboard commands by default', () => {
    const { owners } = resolvePptxShortcuts(undefined, { mac: false });
    expect(owners.get('Mod+c')).toBe('copy');
    expect(owners.get('Mod+Shift+v')).toBe('pasteAsPlainText');
    expect(owners.get('Mod+u')).toBe('underline');
  });

  test('a command that needs arguments cannot be bound directly', () => {
    const { issues } = resolvePptxShortcuts({ tableInsert: 'Mod+Alt+T' } as never, { mac: false });
    expect(issues).toEqual([{ code: 'unknown-id', id: 'tableInsert' }]);
  });

  test('commands without arguments can be given a chord', () => {
    const { owners, issues } = resolvePptxShortcuts({ duplicateSlide: 'Mod+D', slideshow: 'F5' }, { mac: false });
    expect(issues).toEqual([]);
    expect(owners.get('Mod+d')).toBe('duplicateSlide');
    expect(owners.get('f5')).toBe('slideshow');
  });

  test('keeps the keys the text input moves with, select all and the menu key', () => {
    const { issues } = resolvePptxShortcuts({ bold: ['Mod+ArrowLeft', 'Mod+A', 'Shift+F10', 'Alt+Escape'] }, { mac: false });
    expect(issues.map(({ code }) => code)).toEqual(['reserved-chord', 'reserved-chord', 'reserved-chord', 'reserved-chord']);
  });

  test('refuses ids the editor cannot run without displacing their chord', () => {
    const unsupported = new Set(['copy', 'paste']);
    const { owners, issues } = resolvePptxShortcuts({ copy: 'Mod+S', paste: null, bold: 'Mod+Alt+B' }, { mac: false, unsupported });
    expect(issues).toEqual([{ code: 'unknown-id', id: 'copy' }, { code: 'unknown-id', id: 'paste' }]);
    expect(owners.get('Mod+s')).toBe('save');
    expect(owners.get('Mod+v')).toBe('paste');
    expect(owners.get('Mod+Alt+b')).toBe('bold');
  });

  test('accepts plugin ids before the plugin is active', () => {
    const { owners } = resolvePptxShortcuts({ 'plugin:acme/run': 'Mod+Alt+R' }, { mac: false });
    expect(owners.get('Mod+Alt+r')).toBe('plugin:acme/run');
  });
});

describe('command controller', () => {
  test('descriptors and dispatch follow the host map', () => {
    const controller = createPptxCommandController();
    applyShortcuts(controller, { pasteAsPlainText: 'Mod+Alt+V', bold: null }, { mac: false });
    expect(controller.store.getDescriptor('pasteAsPlainText').shortcuts.map(({ chord }) => chord)).toEqual(['Mod+Alt+v']);
    expect(controller.store.getDescriptor('bold').shortcuts).toEqual([]);
    expect(shortcutAt(controller, key({ key: 'b', ctrlKey: true }), false)).toBeNull();
    expect(shortcutAt(controller, key({ key: 'v', ctrlKey: true, altKey: true }), false)?.id).toBe('pasteAsPlainText');
    applyShortcuts(controller, undefined);
    expect(controller.store.getDescriptor('bold').shortcuts.map(({ chord }) => chord)).toEqual(['Mod+B']);
    expect(shortcutAt(controller, key({ key: 'b', ctrlKey: true }), false)?.id).toBe('bold');
  });

  test('a pressed chord names who holds it now and by default', () => {
    const controller = createPptxCommandController();
    applyShortcuts(controller, { paste: null, italic: 'Mod+B', bold: 'Mod+Alt+B' }, { mac: false });
    expect(pressedChord(controller, key({ key: 'v', ctrlKey: true }), false)).toEqual({
      chord: 'Mod+v', owner: undefined, fallback: 'paste',
    });
    expect(pressedChord(controller, key({ key: 'b', ctrlKey: true }), false)).toEqual({
      chord: 'Mod+b', owner: 'italic', fallback: 'bold',
    });
    expect(pressedChord(controller, key({ key: 'z', ctrlKey: true }), false)).toEqual({
      chord: 'Mod+z', owner: 'undo', fallback: 'undo',
    });
    expect(releasedChord(controller, 'Mod+v')).toBe(true);
    expect(releasedChord(controller, 'Mod+z')).toBe(false);
  });

  test('a chord naming the typed character wins over a key-position match', () => {
    const bang = key({ key: '!', code: 'Digit1', ctrlKey: true, shiftKey: true });
    expect(pickChord(['Mod+Shift+1', 'Mod+Shift+!'], bang, false)).toBe('Mod+Shift+!');
    expect(pickChord(['Mod+Shift+1'], bang, false)).toBe('Mod+Shift+1');
  });

  test('plugin chords stay with the plugin dispatcher and bump the revision', () => {
    const controller = createPptxCommandController();
    const before = controller.shortcutRevision();
    let notified = 0;
    subscribeShortcuts(controller, () => { notified += 1; });
    applyShortcuts(controller, { 'plugin:acme/run': 'Mod+Alt+R', italic: null }, { mac: false });
    expect(controller.shortcutRevision()).toBeGreaterThan(before);
    expect(notified).toBe(1);
    expect(hostPluginChords(controller, 'plugin:acme/run')).toEqual(['Mod+Alt+r']);
    expect(hostPluginChords(controller, 'plugin:acme/other')).toBeUndefined();
    expect(shortcutAt(controller, key({ key: 'r', ctrlKey: true, altKey: true }), false)).toBeNull();
    expect(shortcutOwner(controller, 'Mod+i')).toBeUndefined();
  });

  test('labels and aria names show the effective chord', () => {
    const controller = createPptxCommandController();
    applyShortcuts(controller, { paste: 'Mod+Alt+V', copy: null }, { mac: false });
    expect(storeShortcutLabel(controller.store, 'paste', t, undefined, false)).toBe('Ctrl+Alt+V');
    expect(storeShortcutLabel(controller.store, 'paste', t, undefined, true)).toBe('⌥⌘V');
    expect(storeAriaShortcut(controller.store, 'paste', undefined, false)).toBe('Control+Alt+V');
    expect(storeShortcutLabel(controller.store, 'copy', t, undefined, false)).toBeNull();
    const names: Record<string, string> = { 'keys.ctrl': 'Strg', 'keys.shift': 'Umschalt' };
    expect(storeShortcutLabel(controller.store, 'redo', (path) => names[path] ?? t(path), undefined, false)).toBe('Strg+Umschalt+Z');
  });
});
