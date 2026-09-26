import { describe, expect, test } from 'bun:test';
import { definePptxPlugin, definitionProblem, pluginDefinition } from './definePptxPlugin';
import type { PptxPluginCommand, PptxPluginDefinition, PptxPluginPanel } from './types';

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

describe('definePptxPlugin', () => {
  test('keeps the contributions it was given when the host changes them in place', () => {
    const shortcuts = ['Mod+Shift+R'];
    const execute = () => ({ ok: true, status: 'noop' } as const);
    const commands: PptxPluginCommand<null>[] = [
      { id: 'run', label: 'Run', mutatesDocument: false, shortcuts, execute },
    ];
    const toolbar = ['run'];
    const panel: PptxPluginPanel<null> = { title: 'Review', placement: 'left', render: () => null };
    const plugin = definePptxPlugin({
      id: 'acme.keys',
      createState: () => null,
      commands,
      toolbar,
      panel,
    });
    shortcuts.push('Mod+B');
    commands.push({ ...commands[0], id: 'late' });
    toolbar.length = 0;
    panel.title = 'Changed';
    const definition = pluginDefinition(plugin)!;
    expect(definition.commands!.map((command) => [command.id, command.shortcuts])).toEqual([
      ['run', ['Mod+Shift+R']],
    ]);
    expect(definition.toolbar).toEqual(['run']);
    expect(definition.panel!.title).toBe('Review');
    expect(Object.isFrozen(definition.commands![0].shortcuts)).toBe(true);
  });
});
