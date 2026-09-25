import { describe, expect, test } from 'bun:test';
import { definitionProblem } from './definePptxPlugin';
import type { PptxPluginDefinition } from './types';

function withShortcuts(shortcuts: string[]): PptxPluginDefinition<unknown> {
  return {
    id: 'acme.keys',
    createState: () => null,
    commands: [
      {
        id: 'run',
        label: 'Run',
        mutatesDocument: false,
        shortcuts,
        execute: () => ({ ok: true, status: 'noop' }),
      },
    ],
  } as unknown as PptxPluginDefinition<unknown>;
}

describe('plugin shortcuts', () => {
  test('need Mod or Alt unless they are a function key', () => {
    for (const chord of ['Delete', 'Enter', 'Escape', 'ArrowUp', 'a', 'Shift+Tab']) {
      expect(definitionProblem(withShortcuts([chord]))?.message).toBe(
        'Command "run" has an invalid shortcut'
      );
    }
    for (const chord of ['Mod+Shift+R', 'Alt+K', 'F7', 'Shift+F12']) {
      expect(definitionProblem(withShortcuts([chord]))).toBeNull();
    }
  });
});
