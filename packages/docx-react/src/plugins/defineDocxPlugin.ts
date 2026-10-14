import { normalizeChord } from '../commands/descriptors';
import type { DocxPlugin, DocxPluginCommand, DocxPluginDefinition } from './types';

const definitions = new WeakMap<object, DocxPluginDefinition<unknown>>();

const LOCAL_ID = /^[A-Za-z0-9](?:[A-Za-z0-9._-]{0,126}[A-Za-z0-9])?$/;
const PLACEMENTS: ReadonlySet<string> = new Set(['left', 'right', 'bottom']);

function frozenCopy<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze([...value]) as T;
  return value !== null && typeof value === 'object' ? Object.freeze({ ...value }) : value;
}

function frozenCommand<S>(command: DocxPluginCommand<S>): DocxPluginCommand<S> {
  if (command === null || typeof command !== 'object') return command;
  return Object.freeze({ ...command, shortcuts: frozenCopy(command.shortcuts) });
}

/**
 * Wraps a plugin for `DocxEditor`'s `plugins` prop. The definition is copied and frozen, down to
 * its panel, commands, shortcuts and toolbar, and checked when the editor installs it; a problem
 * is reported through `onPluginError` and disables only this plugin.
 *
 * @experimental The plugin API may change in minor releases.
 */
export function defineDocxPlugin<S>(definition: DocxPluginDefinition<S>): DocxPlugin {
  const commands = definition.commands;
  const frozen = Object.freeze({
    ...definition,
    panel: frozenCopy(definition.panel),
    commands: Array.isArray(commands) ? Object.freeze(commands.map(frozenCommand)) : commands,
    toolbar: frozenCopy(definition.toolbar),
  });
  const plugin = Object.freeze(
    frozen.revision === undefined ? { id: frozen.id } : { id: frozen.id, revision: frozen.revision }
  );
  definitions.set(plugin, frozen as DocxPluginDefinition<unknown>);
  return plugin as unknown as DocxPlugin;
}

/** The definition behind a plugin this package created, or null. */
export function pluginDefinition(plugin: unknown): DocxPluginDefinition<unknown> | null {
  return plugin !== null && typeof plugin === 'object' ? definitions.get(plugin) ?? null : null;
}

function isComponent(value: unknown): boolean {
  return typeof value === 'function' || (value !== null && typeof value === 'object');
}

const FUNCTION_KEY = /^f([1-9]|1[0-2])$/i;

/** Whether a plugin may bind `chord`: it needs Mod or Alt unless it is a function key. */
function pluginChord(chord: string): boolean {
  const normalized = normalizeChord(chord);
  if (normalized === null) return false;
  const parts = normalized.split('+');
  const key = parts[parts.length - 1];
  return parts.includes('Mod') || parts.includes('Alt') || FUNCTION_KEY.test(key);
}

function optionalFunction(value: unknown): boolean {
  return value === undefined || typeof value === 'function';
}

/** What makes a definition unusable, or null. */
export function definitionProblem(definition: DocxPluginDefinition<unknown>): Error | null {
  if (typeof definition.createState !== 'function') {
    return new TypeError('createState must be a function');
  }
  if (!optionalFunction(definition.initialize) || !optionalFunction(definition.onEvent)) {
    return new TypeError('initialize and onEvent must be functions');
  }
  if (!optionalFunction(definition.getSidebarItems)) {
    return new TypeError('getSidebarItems must be a function');
  }
  if (definition.overlay !== undefined && !isComponent(definition.overlay)) {
    return new TypeError('overlay must be a component');
  }
  const panel = definition.panel;
  if (
    panel !== undefined &&
    (typeof panel?.title !== 'string' ||
      !PLACEMENTS.has(panel.placement) ||
      !isComponent(panel.render) ||
      (panel.preferredSize !== undefined &&
        !(Number.isFinite(panel.preferredSize) && panel.preferredSize > 0)))
  ) {
    return new TypeError('panel needs a title, a left, right or bottom placement and a renderer');
  }
  const commands = definition.commands ?? [];
  if (!Array.isArray(commands)) return new TypeError('commands must be an array');
  const ids = new Set<string>();
  for (const command of commands) {
    if (!LOCAL_ID.test(command?.id ?? '')) {
      return new TypeError(`Invalid command id ${JSON.stringify(command?.id)}`);
    }
    if (ids.has(command.id)) return new TypeError(`Duplicate command id "${command.id}"`);
    ids.add(command.id);
    if (
      typeof command.label !== 'string' ||
      typeof command.mutatesDocument !== 'boolean' ||
      typeof command.execute !== 'function' ||
      !optionalFunction(command.getState)
    ) {
      return new TypeError(`Command "${command.id}" needs a label, mutatesDocument and execute`);
    }
    const shortcuts = command.shortcuts ?? [];
    if (!Array.isArray(shortcuts) || shortcuts.some((chord) => !pluginChord(chord))) {
      return new TypeError(`Command "${command.id}" has an invalid shortcut`);
    }
  }
  const toolbar = definition.toolbar ?? [];
  if (!Array.isArray(toolbar) || toolbar.some((id) => !ids.has(id))) {
    return new TypeError('toolbar may only list the plugin’s own command ids');
  }
  return null;
}
