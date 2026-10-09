import { describe, expect, test } from 'bun:test';
import { defineDocxPlugin, definitionProblem, pluginDefinition } from './defineDocxPlugin';
import type { DocxPluginDefinition } from './types';

function withShortcuts(shortcuts: string[]): DocxPluginDefinition<unknown> {
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
  } as unknown as DocxPluginDefinition<unknown>;
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

describe('plugin definitions', () => {
  test('are copied and frozen down to the panel, commands, shortcuts and toolbar', () => {
    const shortcuts = ['Mod+Shift+R'];
    const command = {
      id: 'run',
      label: 'Run',
      mutatesDocument: false,
      shortcuts,
      execute: () => ({ ok: true, status: 'noop' }),
    };
    const commands = [command];
    const panel = { title: 'Panel', placement: 'left', render: () => null };
    const toolbar = ['run'];
    const plugin = defineDocxPlugin({
      id: 'acme.frozen',
      createState: () => null,
      panel,
      commands,
      toolbar,
    } as unknown as DocxPluginDefinition<null>);
    panel.title = 'Changed';
    command.label = 'Changed';
    shortcuts.push('Delete');
    commands.push({ ...command, id: 'extra' });
    toolbar.push('extra');

    const installed = pluginDefinition(plugin)!;
    expect(installed.panel?.title).toBe('Panel');
    expect(installed.commands?.map(({ id, label, shortcuts }) => [id, label, shortcuts])).toEqual([
      ['run', 'Run', ['Mod+Shift+R']],
    ]);
    expect(installed.toolbar).toEqual(['run']);
    const parts = [installed, installed.panel, installed.commands, installed.toolbar];
    const [frozenCommand] = installed.commands!;
    expect([...parts, frozenCommand, frozenCommand.shortcuts].every(Object.isFrozen)).toBe(true);
    expect(definitionProblem(installed)).toBeNull();
  });
});
