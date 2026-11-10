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
import { EDITING_CHORDS, isPluginCommandId, normalizeChord } from '../commands/descriptors';
import { commandReason } from '../commands/evaluate';
import type {
  DocxCommandStore,
  DocxPluginCommandId,
  DocxPluginCommandResult,
} from '../commands/types';
import type { EditorMode } from '../components/DocxEditor/internals/editing-modes';
import { readSessionVersion } from '../components/DocxEditor/internals/layoutProvenance';
import {
  createPluginClients,
  pluginCommandScope,
  scopedCommandStore,
  type DocxPluginEditorAccess,
} from './createPluginClients';
import { definitionProblem, pluginDefinition } from './defineDocxPlugin';
import { scopePluginGeometry } from './geometry';
import { createDocxSelectionObserver, type DocxSelectionIdentity } from './selectionIdentity';
import { observeProposals, proposalSnapshot, type DocxProposalSnapshot } from './proposalPreview';
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
  DocxPluginRect,
  DocxPluginSelection,
  DocxPluginSnapshot,
} from './types';

type Definition = DocxPluginDefinition<unknown>;
type Context = DocxPluginContext<unknown>;

export type DocxPluginActivation = RuntimeActivation<Definition, Context>;

/** The editor as the plugin host reaches it; members are read at call time. */
export interface DocxPluginHostAccess extends DocxPluginEditorAccess {
  translate(key: TranslationKey): string;
  selectionPublished?(key: string): void;
  /** Geometry of the layout that shows the current version, or null. */
  geometry(): DocxPluginGeometry | null;
  /** Calls `listener` when geometry answers change within one layout, so plugin views read them again. */
  subscribeGeometry?(listener: () => void): () => void;
}

export interface DocxPluginHost {
  setPlugins(plugins: readonly DocxPlugin[] | undefined): void;
  setGrants(grants: Readonly<Record<string, DocxPluginGrant>> | undefined): void;
  setReporter(report: ((error: DocxPluginError) => void) | undefined): void;
  /** Starts a document generation over `session`. */
  open(session: YrsSession): void;
  openWorker(version: string): void;
  workerChanged(version: string): void;
  workerProposalsChanged(snapshot: DocxProposalSnapshot): void;
  close(reason: 'document-replaced' | 'unmounted'): void;
  generation(): string | null;
  version(): string;
  previewVersion(): number;
  selectionChanged(selection: DocxPluginSelection, identity?: string | DocxSelectionIdentity,
    publication?: { owner: object; cause?: 'remote' | 'seed' }): string;
  selectionUnsettled(generation?: string | null): void;
  selectionSettled(source: 'pointer' | 'touch' | 'keyboard', generation?: string | null): void;
  focusChanged(focused: boolean): void;
  localeChanged(locale: string): void;
  modeChanged(mode: EditorMode, readOnly: boolean): void;
  layoutChanged(layout: DocxPluginLayout | null): void;
  /**
   * The id of the layout plugins may use now. It clears in the same step as a committed change or
   * a replaced document, before React renders.
   */
  layoutId(): string | null;
  /** Geometry moved without a new layout, as on resize or scroll-container changes. */
  geometryChanged(): void;
  /** The pages now show `layout`'s pixels; repeats its `layout-change` if it is still current. */
  layoutPresented(layout: DocxPluginLayout): void;
  /** Geometry for the shown `layout` exists now; repeats its `layout-change` if none carried it. */
  geometryPresented(layout: DocxPluginLayout): void;
  /** @internal Resolves after queued plugin notifications finish. */
  notificationsIdle(): Promise<void>;
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

function checkedResult(value: unknown): DocxPluginCommandResult {
  const result = value as { ok?: unknown } | null | undefined;
  return result && typeof result.ok === 'boolean'
    ? (value as DocxPluginCommandResult)
    : { ok: true, status: 'executed' };
}

function sameLayout(a: DocxPluginLayout | null, b: DocxPluginLayout | null): boolean {
  return (
    a === b ||
    (!!a &&
      !!b &&
      a.id === b.id &&
      a.version === b.version &&
      a.previewVersion === b.previewVersion &&
      (a.reservationsVersion ?? null) === (b.reservationsVersion ?? null) &&
      a.zoom === b.zoom)
  );
}

export function createDocxPluginHost(access: DocxPluginHostAccess): DocxPluginHost {
  let observeSelection = createDocxSelectionObserver();
  let reporter: ((error: DocxPluginError) => void) | undefined;
  let session: YrsSession | null = null;
  let workerProposalKey: string | null = null;
  let detachUpdates: (() => void) | null = null;
  let detachProposals: (() => void) | null = null;
  let detachGeometry: (() => void) | null = null;
  let installed = false;
  let generations = 0;
  let selectionIdentity = 'null';
  let resolvedSelection = 'null';
  let selectionSettleKey = 'null';
  const selectionIdentities = new WeakMap<DocxPluginSnapshot, string>();
  const resolvedSelectionKeys = new WeakMap<DocxPluginSnapshot, string>();
  let settledSelection: string | null = null;
  let settleEvent: Extract<DocxPluginEvent, { type: 'selection-settle' }> | null = null;
  let settleDelivered = false;
  const invalidateSettle = () => {
    if (settleEvent && !settleDelivered) settledSelection = null;
    settleEvent = null;
  };
  let viewport: DocxPluginRect | null = null;
  let commandSignature: string | null = null;
  const state = {
    version: '',
    previewVersion: 0,
    mode: 'editing' as EditorMode,
    readOnly: false,
    focused: false,
    locale: 'en',
    selectionSettled: true,
    selection: EMPTY_SELECTION,
    layout: null as DocxPluginLayout | null,
    geometryLayout: null as DocxPluginLayout | null,
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
    const selection = selectionIdentities.get(invocation.snapshot);
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
      geometry:
        geometry && sameLayout(invocation.snapshot.layout, geometry.layout)
          ? scopePluginGeometry(geometry, pluginId, () => !invocation.refusal() &&
              selection === selectionIdentity)
          : null,
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
    snapshot: (pluginId, generation) => {
      const snapshot: DocxPluginSnapshot = {
        generation,
        version: state.version,
        mode: state.mode,
        readOnly: state.readOnly,
        focused: state.focused,
        locale: state.locale,
        grant: runtime.grant(pluginId) as DocxPluginGrant,
        selection: state.selection,
        selectionSettled: state.selectionSettled,
        layout: state.layout,
      };
      selectionIdentities.set(snapshot, selectionIdentity);
      resolvedSelectionKeys.set(snapshot, resolvedSelection);
      return snapshot;
    },
    currentVersion: () => runtime.generation() === null ? null : session ? readSessionVersion(session) : state.version,
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
    eventCurrent: (event, snapshot) => {
      if (event.type !== 'selection-settle') return true;
      const current = event === settleEvent && event.version === snapshot.version &&
        settledSelection === resolvedSelectionKeys.get(snapshot) &&
        JSON.stringify(event.selection) === JSON.stringify(snapshot.selection);
      if (event === settleEvent) {
        if (current) settleDelivered = true;
        else invalidateSettle();
      }
      return current;
    },
    validate: definitionProblem,
    report,
  });

  const observeGeometry = (): void => {
    if (!detachGeometry) detachGeometry = access.subscribeGeometry?.(() => { notifyViewport(); runtime.touch(); }) ?? null;
  };

  const notify = (event: DocxPluginEvent): void => {
    if (event.type === 'selection-change' || event.type === 'document-change') invalidateSettle();
    runtime.notify(event);
  };
  const notifyViewport = (): void => {
    const generation = runtime.generation();
    const geometry = state.layout ? access.geometry() : null;
    const next = geometry && sameLayout(geometry.layout, state.layout) ? geometry.viewport ?? null : null;
    if (viewport === next || (viewport && next && viewport.x === next.x && viewport.y === next.y &&
      viewport.width === next.width && viewport.height === next.height)) return;
    viewport = next ? { ...next } : null;
    if (generation) notify({ type: 'viewport-change', generation, viewport });
  };
  const hasGeometry = (): boolean => {
    const geometry = access.geometry();
    return !!geometry && sameLayout(geometry.layout, state.layout);
  };
  const notifyLayout = (generation: string): void => {
    if (hasGeometry()) state.geometryLayout = state.layout;
    notify({ type: 'layout-change', generation, layout: state.layout });
    notifyViewport();
  };

  const documentChanged = (): void => {
    const generation = runtime.generation();
    const version = session ? readSessionVersion(session) : null;
    if (!generation || !version || version === state.version) return;
    state.version = version;
    notify({ type: 'document-change', generation, version });
    if (state.layout && state.layout.version !== version) {
      state.layout = null;
      notify({ type: 'layout-change', generation, layout: null });
      notifyViewport();
    }
  };

  const proposalsChanged = (snapshot: DocxProposalSnapshot): void => {
    documentChanged();
    state.previewVersion = snapshot.previewVersion;
    const generation = runtime.generation();
    const invalidated = state.layout && state.layout.previewVersion !== snapshot.previewVersion;
    if (invalidated) state.layout = null;
    if (!generation) return;
    notify({
      type: 'proposal-change',
      generation,
      version: snapshot.version,
      previewVersion: snapshot.previewVersion,
    });
    if (invalidated) notify({ type: 'layout-change', generation, layout: null });
    notifyViewport();
  };

  /** Observes document and proposal changes while plugins are installed. */
  const observe = (): void => {
    if (!installed || !session) {
      detachUpdates?.();
      detachUpdates = null;
      detachProposals?.();
      detachProposals = null;
      return;
    }
    if (detachUpdates) return;
    state.version = readSessionVersion(session) ?? state.version;
    state.previewVersion = proposalSnapshot(session)?.previewVersion ?? 0;
    if (state.layout?.previewVersion !== state.previewVersion) state.layout = null;
    try {
      detachUpdates = session.onUpdate(documentChanged);
      detachProposals = observeProposals(session, proposalsChanged);
    } catch (error) {
      console.error('[DocxEditor] observing document changes for plugins failed', error);
    }
  };

  const closeSession = (): void => {
    detachUpdates?.();
    detachUpdates = null;
    detachProposals?.();
    detachProposals = null;
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
  ): Promise<DocxPluginCommandResult> => {
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
      observeSelection = createDocxSelectionObserver();
      session = next;
      state.version = readSessionVersion(next) ?? '';
      state.previewVersion = proposalSnapshot(next)?.previewVersion ?? 0;
      state.layout = null;
      state.selectionSettled = true;
      selectionIdentity = 'null';
      resolvedSelection = 'null';
      selectionSettleKey = 'null';
      settledSelection = null;
      settleEvent = null;
      state.selection = EMPTY_SELECTION;
      observe();
      observeGeometry();
      viewport = null;
      generations += 1;
      runtime.open(`document-${generations}`);
    },

    close(reason) {
      closeSession();
      if (reason === 'unmounted') {
        detachGeometry?.();
        detachGeometry = null;
      }
      runtime.close(reason);
    },

    openWorker(version) {
      closeSession();
      observeSelection = createDocxSelectionObserver();
      workerProposalKey = null;
      state.version = version;
      state.previewVersion = 0;
      state.layout = null;
      state.selectionSettled = true;
      selectionIdentity = 'null';
      resolvedSelection = 'null';
      selectionSettleKey = 'null';
      settledSelection = null;
      settleEvent = null;
      state.selection = EMPTY_SELECTION;
      observeGeometry();
      viewport = null;
      generations += 1;
      runtime.open(`document-${generations}`);
    },

    workerChanged(version) {
      const generation = runtime.generation();
      if (!generation || version === state.version) return;
      state.version = version;
      notify({ type: 'document-change', generation, version });
      if (state.layout) {
        state.layout = null;
        notify({ type: 'layout-change', generation, layout: null });
        notifyViewport();
      }
    },

    workerProposalsChanged(snapshot) {
      const generation = runtime.generation();
      const key = JSON.stringify(snapshot);
      if (!generation || key === workerProposalKey) return;
      workerProposalKey = key;
      if (snapshot.version !== state.version) {
        state.version = snapshot.version;
        notify({ type: 'document-change', generation, version: snapshot.version });
      }
      state.previewVersion = snapshot.previewVersion;
      const invalidated = state.layout && (state.layout.version !== snapshot.version || state.layout.previewVersion !== snapshot.previewVersion);
      if (invalidated) state.layout = null;
      notify({ type: 'proposal-change', generation, version: snapshot.version, previewVersion: snapshot.previewVersion });
      if (invalidated) notify({ type: 'layout-change', generation, layout: null });
      notifyViewport();
    },

    generation: () => runtime.generation(),
    version: () => state.version,
    previewVersion: () => state.previewVersion,

    selectionChanged(selection, identity = JSON.stringify(selection), publication) {
      if (publication && typeof identity !== 'string') {
        identity = observeSelection(publication.owner, identity, publication.cause);
      }
      const sticky = typeof identity === 'string' ? identity : identity.stickyKey;
      const moved = sticky !== selectionIdentity;
      const resolved = typeof identity === 'string' ? identity : identity.resolvedKey;
      const resolvedMoved = resolved !== resolvedSelection;
      const settleKey = typeof identity === 'string' ? identity : identity.settleKey;
      if (settleKey === selectionSettleKey &&
        resolved !== resolvedSelection && state.selectionSettled && settledSelection !== null && settleDelivered) {
        settledSelection = resolved;
      }
      selectionIdentity = sticky;
      resolvedSelection = resolved;
      selectionSettleKey = settleKey;
      const selectionChanged = JSON.stringify(selection) !== JSON.stringify(state.selection);
      state.selection = selection;
      if (!selectionChanged && !resolvedMoved) {
        if (moved) { invalidateSettle(); runtime.touch(); }
      } else {
        const generation = runtime.generation();
        if (generation) {
          notify({ type: 'selection-change', generation, version: state.version, selection });
        }
      }
      access.selectionPublished?.(settleKey);
      return sticky;
    },

    selectionUnsettled(expected = runtime.generation()) {
      if (!expected || runtime.generation() !== expected || !state.selectionSettled) return;
      state.selectionSettled = false;
      invalidateSettle();
      runtime.touch();
    },

    selectionSettled(source, expected = runtime.generation()) {
      if (!expected || runtime.generation() !== expected) return;
      const changed = !state.selectionSettled;
      state.selectionSettled = true;
      const key = resolvedSelection;
      if (key === settledSelection) {
        if (changed) runtime.touch();
        return;
      }
      settledSelection = key;
      settleDelivered = false;
      settleEvent = { type: 'selection-settle', generation: expected, version: state.version,
        selection: structuredClone(state.selection), source };
      notify(settleEvent);
    },

    focusChanged(focused) {
      if (focused === state.focused) return;
      state.focused = focused;
      const generation = runtime.generation();
      if (generation) notify({ type: 'focus-change', generation, focused });
    },

    localeChanged(locale) {
      if (locale === state.locale) return;
      state.locale = locale;
      const generation = runtime.generation();
      if (generation) notify({ type: 'locale-change', generation, locale });
    },

    modeChanged(mode, readOnly) {
      if (mode === state.mode && readOnly === state.readOnly) return;
      state.mode = mode;
      state.readOnly = readOnly;
      const generation = runtime.generation();
      if (generation) notify({ type: 'mode-change', generation, mode, readOnly });
    },

    layoutChanged(next) {
      const layout =
        next && next.version === state.version && next.previewVersion === state.previewVersion
          ? next
          : null;
      if (sameLayout(layout, state.layout)) return;
      state.layout = layout;
      state.geometryLayout = null;
      const generation = runtime.generation();
      if (generation) notifyLayout(generation);
    },

    layoutId: () => (runtime.generation() === null ? null : state.layout?.id ?? null),

    geometryChanged() {
      notifyViewport();
      runtime.touch();
    },

    layoutPresented(layout) {
      const generation = runtime.generation();
      if (!generation || !state.layout || !sameLayout(layout, state.layout)) return;
      notifyLayout(generation);
    },

    geometryPresented(layout) {
      const generation = runtime.generation();
      if (
        !generation ||
        !sameLayout(layout, state.layout) ||
        !hasGeometry() ||
        sameLayout(state.geometryLayout, state.layout)
      ) return;
      notifyLayout(generation);
    },

    notificationsIdle: () => runtime.notificationsIdle(),
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
          const hostChords = commands.hostPluginChords(id);
          for (const chord of hostChords ?? []) {
            claimed.set(chord, id);
            shortcuts.push({ chord, args: null });
          }
          for (const chord of hostChords ? [] : (command.shortcuts ?? [])) {
            const canonical = normalizeChord(chord)!;
            const held = commands.shortcutOwner(canonical);
            const owner =
              held === undefined
                ? claimed.get(canonical)
                : isPluginCommandId(held)
                  ? held
                  : EDITING_CHORDS.has(canonical)
                    ? 'text editing'
                    : 'a built-in command';
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
      parts.push(`shortcuts:${commands.shortcutRevision()}`);
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
