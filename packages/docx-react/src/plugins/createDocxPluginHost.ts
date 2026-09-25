import type { TranslationKey } from '@betteroffice/docx-i18n';
import type { YrsSession } from '@betteroffice/docx/yrs';
import {
  createPluginRuntime,
  type PluginInvocation,
  type PluginRuntime,
  type RuntimeActivation,
  type RuntimePhase,
} from '../../../../shared/plugin-host/runtime';
import {
  UNAVAILABLE_DOCX_COMMANDS,
  type DocxCommandController,
  type DocxCommandScope,
  type DocxPluginCommandBinding,
} from '../commands/createDocxCommandStore';
import { BUILT_IN_CHORDS, EDITING_CHORDS, normalizeChord } from '../commands/descriptors';
import { commandReason } from '../commands/evaluate';
import type { DocxCommandResult, DocxCommandStore, DocxPluginCommandId } from '../commands/types';
import type { EditorMode } from '../components/DocxEditor/internals/editing-modes';
import { readSessionVersion } from '../components/DocxEditor/internals/layoutProvenance';
import {
  createPluginClients,
  pluginCommandScope,
  scopedCommandStore,
  type DocxPluginEditorAccess,
} from './createPluginClients';
import { definitionProblem, pluginDefinition } from './defineDocxPlugin';
import type {
  DocxPlugin,
  DocxPluginCommand,
  DocxPluginContext,
  DocxPluginDefinition,
  DocxPluginError,
  DocxPluginErrorPhase,
  DocxPluginEvent,
  DocxPluginGeometry,
  DocxPluginGrant,
  DocxPluginLayout,
  DocxPluginSelection,
  DocxPluginSnapshot,
} from './types';

type Definition = DocxPluginDefinition<unknown>;
type Context = DocxPluginContext<unknown>;

export type DocxPluginActivation = RuntimeActivation<Definition, Context>;

/** The editor as the plugin host reaches it; members are read at call time. */
export interface DocxPluginHostAccess extends DocxPluginEditorAccess {
  translate(key: TranslationKey): string;
  /** Geometry of the layout that shows the current version, or null. */
  geometry(): DocxPluginGeometry | null;
}

export interface DocxPluginHost {
  setPlugins(plugins: readonly DocxPlugin[] | undefined): void;
  setGrants(grants: Readonly<Record<string, DocxPluginGrant>> | undefined): void;
  setReporter(report: ((error: DocxPluginError) => void) | undefined): void;
  /** Starts a document generation over `session`. */
  open(session: YrsSession): void;
  close(reason: 'document-replaced' | 'unmounted'): void;
  generation(): string | null;
  version(): string;
  selectionChanged(selection: DocxPluginSelection): void;
  modeChanged(mode: EditorMode, readOnly: boolean): void;
  layoutChanged(layout: DocxPluginLayout | null): void;
  /** Geometry moved without a new layout, as on resize or scroll-container changes. */
  geometryChanged(): void;
  activations(): readonly DocxPluginActivation[];
  subscribe(listener: () => void): () => void;
  /** The restricted command store an activation's React contributions see. */
  commandStore(activation: DocxPluginActivation): DocxCommandStore;
  /** Publishes contributed commands, shortcuts and toolbar entries to the editor's commands. */
  syncCommands(): void;
  guard<T>(
    pluginId: string,
    phase: DocxPluginErrorPhase,
    call: (context: Context) => T,
    fallback: T
  ): T;
  fail(pluginId: string, phase: DocxPluginErrorPhase, error: unknown): void;
}

const ENABLED = Object.freeze({ enabled: true as const });
const EMPTY_SELECTION: DocxPluginSelection = Object.freeze({
  formatting: null,
  displayRange: null,
});

function checkedState(value: unknown) {
  const state = value as {
    enabled?: unknown;
    disabledReason?: { code?: unknown; message?: unknown };
  };
  if (!state || typeof state.enabled !== 'boolean') {
    throw new TypeError('A command getState must return a CommandState');
  }
  if (
    !state.enabled &&
    (typeof state.disabledReason?.code !== 'string' ||
      typeof state.disabledReason.message !== 'string')
  ) {
    throw new TypeError('A disabled command state needs a disabledReason');
  }
  return value as ReturnType<NonNullable<DocxPluginCommand<unknown>['getState']>>;
}

function checkedResult(value: unknown): DocxCommandResult {
  const result = value as { ok?: unknown } | null | undefined;
  return result && typeof result.ok === 'boolean'
    ? (value as DocxCommandResult)
    : { ok: true, status: 'executed' };
}

function sameLayout(a: DocxPluginLayout | null, b: DocxPluginLayout | null): boolean {
  return a === b || (!!a && !!b && a.id === b.id && a.version === b.version && a.zoom === b.zoom);
}

export function createDocxPluginHost(access: DocxPluginHostAccess): DocxPluginHost {
  let reporter: ((error: DocxPluginError) => void) | undefined;
  let session: YrsSession | null = null;
  let detachUpdates: (() => void) | null = null;
  let installed = false;
  let generations = 0;
  let commandSignature: string | null = null;
  const state = {
    version: '',
    mode: 'editing' as EditorMode,
    readOnly: false,
    selection: EMPTY_SELECTION,
    layout: null as DocxPluginLayout | null,
  };
  const activationScopes = new WeakMap<object, DocxCommandScope>();
  const identities = new WeakMap<object, number>();
  let nextIdentity = 0;
  const identity = (value: object): number => {
    let id = identities.get(value);
    if (id === undefined) {
      nextIdentity += 1;
      id = nextIdentity;
      identities.set(value, id);
    }
    return id;
  };
  const unknownPlugins = new WeakSet<object>();
  const reportedConflicts = new Set<string>();

  const report = (error: DocxPluginError): void => {
    if (!reporter) {
      console.error(`[DocxEditor] plugin "${error.pluginId}" failed (${error.phase})`, error.error);
      return;
    }
    try {
      reporter(error);
    } catch (reporterError) {
      console.error('[DocxEditor] onPluginError threw', reporterError);
    }
  };

  const controller = (): DocxCommandController | null => access.commands();

  const scopeFor = (invocation: PluginInvocation<DocxPluginSnapshot>): DocxCommandScope => {
    const grant = () => runtime.grant(invocation.pluginId) as DocxPluginGrant;
    if (invocation.signal !== invocation.lifetimeSignal) {
      return pluginCommandScope(invocation, grant, runtime.subscribe);
    }
    let scope = activationScopes.get(invocation.activation);
    if (!scope) {
      scope = pluginCommandScope(invocation, grant, runtime.subscribe);
      activationScopes.set(invocation.activation, scope);
    }
    return scope;
  };

  const createContext = (invocation: PluginInvocation<DocxPluginSnapshot>): Context => {
    const pluginId = invocation.pluginId;
    const clients = createPluginClients(
      invocation,
      access,
      () => runtime.grant(pluginId) as DocxPluginGrant,
      scopedCommandStore(controller(), scopeFor(invocation))
    );
    const geometry = access.geometry();
    return {
      pluginId,
      snapshot: invocation.snapshot,
      get state() {
        return invocation.state() as Readonly<unknown>;
      },
      signal: invocation.signal,
      lifetimeSignal: invocation.lifetimeSignal,
      read: clients.read,
      commands: clients.commands,
      edits: clients.edits,
      geometry: geometry && invocation.snapshot.layout?.id === geometry.layout.id ? geometry : null,
      navigation: clients.navigation,
      setState: (next, atVersion) => invocation.setState(next, atVersion),
      onCleanup: (cleanup) => invocation.onCleanup(cleanup),
      run: (action) => invocation.run((next) => action(createContext(next))),
    };
  };

  const runtime: PluginRuntime<
    Definition,
    Context,
    DocxPluginEvent,
    DocxPluginErrorPhase | RuntimePhase
  > = createPluginRuntime<
    Definition,
    Context,
    DocxPluginEvent,
    DocxPluginSnapshot,
    DocxPluginErrorPhase
  >({
    context: (invocation) => createContext(invocation),
    snapshot: (pluginId, generation) => ({
      generation,
      version: state.version,
      mode: state.mode,
      readOnly: state.readOnly,
      grant: runtime.grant(pluginId) as DocxPluginGrant,
      selection: state.selection,
      layout: state.layout,
    }),
    currentVersion: () =>
      session && runtime.generation() !== null ? readSessionVersion(session) : null,
    loadEvent: (snapshot, reason) => ({
      type: 'load',
      generation: snapshot.generation,
      version: snapshot.version,
      reason,
    }),
    grantsEvent: (snapshot, grant) => ({
      type: 'grants-change',
      generation: snapshot.generation,
      grant: grant as DocxPluginGrant,
    }),
    validate: definitionProblem,
    report,
  });

  const notify = (event: DocxPluginEvent): void => runtime.notify(event);

  const documentChanged = (): void => {
    const generation = runtime.generation();
    const version = session ? readSessionVersion(session) : null;
    if (!generation || !version || version === state.version) return;
    state.version = version;
    notify({ type: 'document-change', generation, version });
    if (state.layout && state.layout.version !== version) {
      state.layout = null;
      notify({ type: 'layout-change', generation, layout: null });
    }
  };

  /** Observes document changes only while plugins are installed. */
  const observe = (): void => {
    if (!installed || !session) {
      detachUpdates?.();
      detachUpdates = null;
      return;
    }
    if (detachUpdates) return;
    state.version = readSessionVersion(session) ?? state.version;
    try {
      detachUpdates = session.onUpdate(documentChanged);
    } catch (error) {
      console.error('[DocxEditor] observing document changes for plugins failed', error);
    }
  };

  const closeSession = (): void => {
    detachUpdates?.();
    detachUpdates = null;
    session = null;
  };

  const commandState = (pluginId: string, command: DocxPluginCommand<unknown>) =>
    runtime.guard(
      pluginId,
      'command-state',
      (context) => checkedState(command.getState ? command.getState(context) : ENABLED),
      {
        enabled: false,
        disabledReason: commandReason('plugin-unavailable', { translate: access.translate }),
      }
    );

  const executeCommand = async (
    pluginId: string,
    command: DocxPluginCommand<unknown>
  ): Promise<DocxCommandResult> => {
    const outcome = await runtime.invoke(pluginId, 'command', (context) =>
      command.execute(context)
    );
    if (outcome.ok) return checkedResult(outcome.value);
    const code = outcome.reason === 'failed' ? 'command-failed' : outcome.reason;
    return { ok: false, failure: commandReason(code, { translate: access.translate }) };
  };

  return {
    setPlugins(plugins) {
      const definitions: Definition[] = [];
      for (const plugin of plugins ?? []) {
        const definition = pluginDefinition(plugin);
        if (definition) {
          definitions.push(definition);
          continue;
        }
        if (plugin !== null && typeof plugin === 'object') {
          if (unknownPlugins.has(plugin)) continue;
          unknownPlugins.add(plugin);
        }
        report({
          pluginId: typeof plugin?.id === 'string' ? plugin.id : '',
          generation: runtime.generation(),
          phase: 'definition',
          error: new TypeError('Plugins must be created with defineDocxPlugin'),
        });
      }
      installed = definitions.length > 0;
      observe();
      runtime.setPlugins(definitions);
    },

    setGrants: (grants) => runtime.setGrants(grants),

    setReporter(next) {
      reporter = next;
    },

    open(next) {
      closeSession();
      session = next;
      state.version = readSessionVersion(next) ?? '';
      state.layout = null;
      state.selection = EMPTY_SELECTION;
      observe();
      generations += 1;
      runtime.open(`document-${generations}`);
    },

    close(reason) {
      closeSession();
      runtime.close(reason);
    },

    generation: () => runtime.generation(),
    version: () => state.version,

    selectionChanged(selection) {
      if (JSON.stringify(selection) === JSON.stringify(state.selection)) return;
      state.selection = selection;
      const generation = runtime.generation();
      if (generation) {
        notify({ type: 'selection-change', generation, version: state.version, selection });
      }
    },

    modeChanged(mode, readOnly) {
      if (mode === state.mode && readOnly === state.readOnly) return;
      state.mode = mode;
      state.readOnly = readOnly;
      const generation = runtime.generation();
      if (generation) notify({ type: 'mode-change', generation, mode, readOnly });
    },

    layoutChanged(next) {
      const layout = next && next.version === state.version ? next : null;
      if (sameLayout(layout, state.layout)) return;
      state.layout = layout;
      const generation = runtime.generation();
      if (generation) notify({ type: 'layout-change', generation, layout });
    },

    geometryChanged: () => runtime.touch(),

    activations: () => runtime.activations(),
    subscribe: (listener) => runtime.subscribe(listener),

    commandStore(activation) {
      const scope = activationScopes.get(activation.activation);
      return scope ? scopedCommandStore(controller(), scope) : UNAVAILABLE_DOCX_COMMANDS;
    },

    syncCommands() {
      const commands = controller();
      if (!commands) return;
      const bindings: DocxPluginCommandBinding[] = [];
      const toolbar: DocxPluginCommandId[] = [];
      const claimed = new Map<string, string>();
      const parts: string[] = [];
      for (const entry of runtime.activations()) {
        const definition = entry.plugin;
        parts.push(`${entry.pluginId}:${identity(definition)}:${identity(entry.activation)}`);
        for (const command of definition.commands ?? []) {
          const id: DocxPluginCommandId = `plugin:${entry.pluginId}/${command.id}`;
          const shortcuts: { chord: string; args: null }[] = [];
          for (const chord of command.shortcuts ?? []) {
            const canonical = normalizeChord(chord)!;
            const owner = BUILT_IN_CHORDS.has(canonical)
              ? 'a built-in command'
              : EDITING_CHORDS.has(canonical)
                ? 'text editing'
                : claimed.get(canonical);
            if (!owner) {
              claimed.set(canonical, id);
              shortcuts.push({ chord, args: null });
              continue;
            }
            const key = `${entry.pluginId}\u0000${canonical}`;
            if (reportedConflicts.has(key)) continue;
            reportedConflicts.add(key);
            report({
              pluginId: entry.pluginId,
              generation: entry.generation,
              phase: 'definition',
              error: new Error(
                `Shortcut ${chord} of ${id} is already used by ${owner}; it was not bound`
              ),
            });
          }
          bindings.push({
            descriptor: Object.freeze({
              id,
              label: command.label,
              mutatesDocument: command.mutatesDocument,
              shortcuts: Object.freeze(shortcuts),
            }),
            state: () => commandState(entry.pluginId, command),
            execute: () => executeCommand(entry.pluginId, command),
          });
        }
        for (const local of definition.toolbar ?? [])
          toolbar.push(`plugin:${entry.pluginId}/${local}`);
      }
      const signature = parts.join('|');
      if (signature === commandSignature) {
        commands.refresh();
        return;
      }
      commandSignature = signature;
      commands.setPluginCommands(bindings, toolbar);
    },

    guard: (pluginId, phase, call, fallback) => runtime.guard(pluginId, phase, call, fallback),
    fail: (pluginId, phase, error) => runtime.fail(pluginId, phase, error),
  };
}
