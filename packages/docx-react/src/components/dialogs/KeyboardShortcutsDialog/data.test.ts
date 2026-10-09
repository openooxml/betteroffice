import { describe, expect, test } from 'bun:test';
import { DOCX_COMMAND_DESCRIPTORS, formatChord } from '../../../commands/descriptors';
import type { DocxCommandShortcut } from '../../../commands/types';
import { COMMAND_SHORTCUTS, DEFAULT_SHORTCUTS } from './data';

describe('keyboard help', () => {
  test('lists every command binding the descriptors define, and no other', () => {
    const bound = Object.values(DOCX_COMMAND_DESCRIPTORS)
      .flatMap((descriptor) => descriptor.shortcuts as readonly DocxCommandShortcut[])
      .map((shortcut) => formatChord(shortcut.chord, false))
      .sort();
    const listed = COMMAND_SHORTCUTS.flatMap((shortcut) =>
      shortcut.altKeys ? [shortcut.keys, shortcut.altKeys] : [shortcut.keys]
    ).sort();
    expect(listed).toEqual(bound);
  });

  test('omits commands without a binding', () => {
    const ids = DEFAULT_SHORTCUTS.map((shortcut) => shortcut.id);
    for (const unbound of ['print', 'strikethrough', 'indent', 'outdent', 'zoom-in']) {
      expect(ids).not.toContain(unbound);
    }
    expect(DEFAULT_SHORTCUTS.find((shortcut) => shortcut.id === 'redo')).toMatchObject({
      keys: 'Ctrl+Y',
      altKeys: 'Ctrl+Shift+Z',
    });
    expect(DEFAULT_SHORTCUTS.find((shortcut) => shortcut.id === 'superscript')?.keys).toBe(
      'Ctrl+Shift+='
    );
  });
});
