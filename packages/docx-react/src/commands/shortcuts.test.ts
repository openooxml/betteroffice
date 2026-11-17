import { describe, expect, test } from 'bun:test';
import { createDocxCommandController } from './createDocxCommandStore';
import { resolveDocxShortcuts } from './shortcuts';

function key(init: Partial<KeyboardEvent>): KeyboardEvent {
  return { key: '', code: '', metaKey: false, ctrlKey: false, shiftKey: false, altKey: false, ...init } as KeyboardEvent;
}

describe('resolveDocxShortcuts', () => {
  test('binds the clipboard commands by default', () => {
    const { owners } = resolveDocxShortcuts(undefined, { mac: false });
    expect(owners.get('Mod+Shift+v')).toBe('pasteAsPlainText');
    expect(owners.get('Mod+a')).toBe('selectAll');
  });

  test('alignment variants move independently', () => {
    const { owners } = resolveDocxShortcuts({ 'alignment.center': 'Mod+Alt+E' }, { mac: false });
    expect(owners.get('Mod+Alt+e')).toBe('alignment.center');
    expect(owners.has('Mod+e')).toBe(false);
    expect(owners.get('Mod+l')).toBe('alignment.left');
  });

  test('a command that needs arguments cannot be bound directly', () => {
    const { issues } = resolveDocxShortcuts({ fontSize: 'Mod+Alt+F' } as never, { mac: false });
    expect(issues).toEqual([{ code: 'unknown-id', id: 'fontSize' }]);
  });

  test('keeps the keys the text input moves with', () => {
    const { issues } = resolveDocxShortcuts({ bold: 'Mod+ArrowLeft' }, { mac: false });
    expect(issues.map(({ code }) => code)).toEqual(['reserved-chord']);
  });

  test('accepts plugin ids before the plugin is active', () => {
    const { owners } = resolveDocxShortcuts({ 'plugin:acme/run': 'Mod+Alt+R' }, { mac: false });
    expect(owners.get('Mod+Alt+r')).toBe('plugin:acme/run');
  });
});

describe('command controller', () => {
  test('descriptors and dispatch follow the host map', () => {
    const controller = createDocxCommandController();
    controller.setShortcuts({ pasteAsPlainText: 'Mod+Alt+V', bold: null });
    expect(controller.store.getDescriptor('pasteAsPlainText').shortcuts.map(({ chord }) => chord)).toEqual(['Mod+Alt+v']);
    expect(controller.store.getDescriptor('bold').shortcuts).toEqual([]);
    expect(controller.shortcutAt(key({ key: 'b', ctrlKey: true }), false)).toBeNull();
    expect(controller.shortcutAt(key({ key: 'v', ctrlKey: true, altKey: true }), false)?.id).toBe('pasteAsPlainText');
  });

  test('a variant dispatches its alignment value', () => {
    const controller = createDocxCommandController();
    expect(controller.shortcutAt(key({ key: 'j', ctrlKey: true }), false)).toEqual({
      id: 'alignment',
      args: { value: 'both' },
      chord: 'Mod+j',
    });
  });

  test('plugin chords stay with the plugin dispatcher and bump the revision', () => {
    const controller = createDocxCommandController();
    const before = controller.shortcutRevision();
    controller.setShortcuts({ 'plugin:acme/run': 'Mod+Alt+R', insertLink: null });
    expect(controller.shortcutRevision()).toBeGreaterThan(before);
    expect(controller.hostPluginChords('plugin:acme/run')).toEqual(['Mod+Alt+r']);
    expect(controller.shortcutAt(key({ key: 'r', ctrlKey: true, altKey: true }), false)).toBeNull();
    expect(controller.shortcutOwner('Mod+k')).toBeUndefined();
  });
});
