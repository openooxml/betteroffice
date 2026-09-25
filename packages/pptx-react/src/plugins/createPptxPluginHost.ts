import type { PresentationHandle } from '@betteroffice/pptx';
import type { TranslationKey } from '@betteroffice/pptx-i18n';
import {
  createPluginRuntime,
  type PluginInvocation,
  type PluginRuntime,
  type RuntimeActivation,
  type RuntimePhase,
} from '../../../../shared/plugin-host/runtime';
import {
  UNAVAILABLE_PPTX_COMMANDS,
  type PptxCommandController,
  type PptxCommandScope,
  type PptxPluginCommandBinding,
} from '../commands/createPptxCommandStore';
import { BUILT_IN_CHORDS, EDITING_CHORDS, normalizeChord } from '../commands/descriptors';
import { commandReason } from '../commands/evaluate';
import type { PptxCommandResult, PptxCommandStore, PptxPluginCommandId } from '../commands/types';
import {
  createPluginClients,
  pluginCommandScope,
  scopedCommandStore,
  type PptxPluginEditorAccess,
} from './createPluginClients';
import { definitionProblem, pluginDefinition } from './definePptxPlugin';
import type {
  PptxPlugin,
  PptxPluginCommand,
  PptxPluginContext,
  PptxPluginDefinition,
  PptxPluginError,
  PptxPluginErrorPhase,
  PptxPluginEvent,
  PptxPluginGeometry,
  PptxPluginGrant,
  PptxPluginLayout,
  PptxPluginSelection,
  PptxPluginSnapshot,
} from './types';

type Definition = PptxPluginDefinition<unknown>;
type Context = PptxPluginContext<unknown>;

export type PptxPluginActivation = RuntimeActivation<Definition, Context>;

/** The editor as the plugin host reaches it; members are read at call time. */
export interface PptxPluginHostAccess extends PptxPluginEditorAccess {
  translate(key: TranslationKey): string;
  /** Geometry of the presented frame that shows the current version, or null. */
  geometry(): PptxPluginGeometry | null;
}

export interface PptxPluginHost {
  setPlugins(plugins: readonly PptxPlugin[] | undefined): void;
  setGrants(grants: Readonly<Record<string, PptxPluginGrant>> | undefined): void;
  setReporter(report: ((error: PptxPluginError) => void) | undefined): void;
  /** Starts a document generation over `handle`. */
  open(handle: PresentationHandle): void;
  close(reason: 'document-replaced' | 'unmounted'): void;
  generation(): string | null;
  version(): string;
  selectionChanged(selection: PptxPluginSelection): void;
  modeChanged(readOnly: boolean): void;
  layoutChanged(layout: PptxPluginLayout | null): void;
  /**
   * The id of the layout plugins may use now. It clears in the same step as a committed change,
   * a replaced presentation or a new paint, before React renders.
   */
  layoutId(): string | null;
  /** Geometry moved without a new frame, as on resize. */
  geometryChanged(): void;
  activations(): readonly PptxPluginActivation[];
  subscribe(listener: () => void): () => void;
  /** The restricted command store an activation's React contributions see. */
  commandStore(activation: PptxPluginActivation): PptxCommandStore;
  /** Publishes contributed commands, shortcuts and toolbar entries to the editor's commands. */
  syncCommands(): void;
  guard<T>(
    pluginId: string,
    phase: PptxPluginErrorPhase,
    call: (context: Context) => T,
    fallback: T
  ): T;
  fail(pluginId: string, phase: PptxPluginErrorPhase, error: unknown): void;
}

const ENABLED = Object.freeze({ enabled: true as const });

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
  return value as ReturnType<NonNullable<PptxPluginCommand<unknown>['getState']>>;
}

function checkedResult(value: unknown): PptxCommandResult {
  const result = value as { ok?: unknown } | null | undefined;
  return result && typeof result.ok === 'boolean'
    ? (value as PptxCommandResult)
    : { ok: true, status: 'executed' };
}

function sameLayout(a: PptxPluginLayout | null, b: PptxPluginLayout | null): boolean {
  return a === b || (!!a && !!b && a.id === b.id && a.version === b.version);
}

function readVersion(handle: PresentationHandle): string | null {
  try {
    return handle.version();
  } catch {
    return null;
  }
}

export function createPptxPluginHost(access: PptxPluginHostAccess): PptxPluginHost {
  let reporter: ((error: PptxPluginError) => void) | undefined;
  let handle: PresentationHandle | null = null;
  let detachUpdates: (() => void) | null = null;
  let installed = false;
  let generations = 0;
  let commandSignature: string | null = null;
  const state = {
    version: '',
    readOnly: false,
    selection: null as PptxPluginSelection,
    layout: null as PptxPluginLayout | null,
  };
  const activationScopes = new WeakMap<object, PptxCommandScope>();
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

  const report = (error: PptxPluginError): void => {
    if (!reporter) {
      console.error(`[PptxEditor] plugin "${error.pluginId}" failed (${error.phase})`, error.error);
      return;
    }
    try {
      reporter(error);
    } catch (reporterError) {
      console.error('[PptxEditor] onPluginError threw', reporterError);
    }
  };

  const controller = (): PptxCommandController | null => access.commands();

  const scopeFor = (invocation: PluginInvocation<PptxPluginSnapshot>): PptxCommandScope => {
    const grant = () => runtime.grant(invocation.pluginId) as PptxPluginGrant;
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

  const createContext = (invocation: PluginInvocation<PptxPluginSnapshot>): Context => {
    const pluginId = invocation.pluginId;
    const clients = createPluginClients(
      invocation,
      access,
      () => runtime.grant(pluginId) as PptxPluginGrant,
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
    PptxPluginEvent,
    PptxPluginErrorPhase | RuntimePhase
  > = createPluginRuntime<
    Definition,
    Context,
    PptxPluginEvent,
    PptxPluginSnapshot,
    PptxPluginErrorPhase
  >({
    context: (invocation) => createContext(invocation),
    snapshot: (pluginId, generation) => ({
      generation,
      version: state.version,
      readOnly: state.readOnly,
      grant: runtime.grant(pluginId) as PptxPluginGrant,
      selection: state.selection,
      layout: state.layout,
    }),
    currentVersion: () => (handle && runtime.generation() !== null ? readVersion(handle) : null),
    loadEvent: (snapshot, reason) => ({
      type: 'load',
      generation: snapshot.generation,
      version: snapshot.version,
      reason,
    }),
    grantsEvent: (snapshot, grant) => ({
      type: 'grants-change',
      generation: snapshot.generation,
      grant: grant as PptxPluginGrant,
    }),
    validate: definitionProblem,
    report,
  });

  const notify = (event: PptxPluginEvent): void => runtime.notify(event);

  const documentChanged = (): void => {
    const generation = runtime.generation();
    const version = handle ? readVersion(handle) : null;
    if (!generation || !version || version === state.version) return;
    state.version = version;
    notify({ type: 'document-change', generation, version });
    if (state.layout && state.layout.version !== version) {
      state.layout = null;
      notify({ type: 'layout-change', generation, layout: null });
    }
  };

  /** Observes committed changes only while plugins are installed. */
  const observe = (): void => {
    if (!installed || !handle) {
      detachUpdates?.();
      detachUpdates = null;
      return;
    }
    if (detachUpdates) return;
    state.version = readVersion(handle) ?? state.version;
    try {
      detachUpdates = handle.onUpdate(documentChanged);
    } catch (error) {
      console.error('[PptxEditor] observing presentation changes for plugins failed', error);
    }
  };

  const closeHandle = (): void => {
    detachUpdates?.();
    detachUpdates = null;
    handle = null;
  };

  const commandState = (pluginId: string, command: PptxPluginCommand<unknown>) =>
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
    command: PptxPluginCommand<unknown>
  ): Promise<PptxCommandResult> => {
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
          error: new TypeError('Plugins must be created with definePptxPlugin'),
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
      closeHandle();
      handle = next;
      state.version = readVersion(next) ?? '';
      state.layout = null;
      state.selection = null;
      observe();
      generations += 1;
      runtime.open(`presentation-${generations}`);
    },

    close(reason) {
      closeHandle();
      state.layout = null;
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

    modeChanged(readOnly) {
      if (readOnly === state.readOnly) return;
      state.readOnly = readOnly;
      const generation = runtime.generation();
      if (generation) notify({ type: 'mode-change', generation, readOnly });
    },

    layoutChanged(next) {
      const layout = next && next.version === state.version ? next : null;
      if (sameLayout(layout, state.layout)) return;
      state.layout = layout;
      const generation = runtime.generation();
      if (generation) notify({ type: 'layout-change', generation, layout });
    },

    layoutId: () => (runtime.generation() === null ? null : state.layout?.id ?? null),

    geometryChanged: () => runtime.touch(),

    activations: () => runtime.activations(),
    subscribe: (listener) => runtime.subscribe(listener),

    commandStore(activation) {
      const scope = activationScopes.get(activation.activation);
      return scope ? scopedCommandStore(controller(), scope) : UNAVAILABLE_PPTX_COMMANDS;
    },

    syncCommands() {
      const commands = controller();
      if (!commands) return;
      const bindings: PptxPluginCommandBinding[] = [];
      const toolbar: PptxPluginCommandId[] = [];
      const claimed = new Map<string, string>();
      const parts: string[] = [];
      for (const entry of runtime.activations()) {
        const definition = entry.plugin;
        parts.push(`${entry.pluginId}:${identity(definition)}:${identity(entry.activation)}`);
        for (const command of definition.commands ?? []) {
          const id: PptxPluginCommandId = `plugin:${entry.pluginId}/${command.id}`;
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
