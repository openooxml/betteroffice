/**
 * The format-independent plugin lifecycle: registration, activation and invocation identities,
 * cancellation, coalesced events, guarded state, cleanup and quarantine. Format hosts adapt it
 * with their own contexts, snapshots and events; it never touches React, the DOM or a document.
 */

import type {
  MaybePromise,
  PluginCleanupReason,
  PluginError,
  PluginGrant,
  PluginLoadReason,
} from '../host-contracts/plugins';
import { normalizeGrant, sameGrant } from './grants';

/** The part of a plugin definition the runtime drives. */
export interface RuntimePlugin<Context, Event> {
  readonly id: string;
  readonly revision?: string | number;
  createState(): unknown;
  initialize?(context: Context): MaybePromise<void>;
  onEvent?(context: Context, event: Event): MaybePromise<void>;
}

export interface RuntimeEvent {
  readonly type: string;
  readonly generation: string;
}

export interface RuntimeSnapshot {
  readonly generation: string;
  readonly version: string;
}

/** Why an invocation may no longer act. */
export type InvocationRefusal = 'plugin-unavailable' | 'document-replaced' | 'aborted';

/** Phases the runtime reports itself; hosts add theirs. */
export type RuntimePhase = 'definition' | 'initialize' | 'event' | 'cleanup' | 'action';

/** One call into a plugin, with its snapshot, cancellation and state access. */
export interface PluginInvocation<Snapshot extends RuntimeSnapshot> {
  readonly pluginId: string;
  /** Identifies the activation; readding a plugin creates a new one. */
  readonly activation: object;
  readonly snapshot: Snapshot;
  readonly signal: AbortSignal;
  readonly lifetimeSignal: AbortSignal;
  state(): unknown;
  /** False for a cancelled invocation, an ended activation or a version that is not current. */
  setState(next: unknown, atVersion?: string): boolean;
  onCleanup(cleanup: (reason: PluginCleanupReason) => MaybePromise<void>): void;
  /** Runs `action` with a fresh invocation; a failure quarantines instead of rejecting. */
  run(action: (invocation: PluginInvocation<Snapshot>) => MaybePromise<void>): Promise<void>;
  /**
   * Runs `write`, which commits synchronously through this invocation's clients. The document
   * changes it notifies are this invocation's own: they never abort or repeat its hook.
   */
  commit<T>(write: () => T): T;
  /** Why this invocation may no longer act, or null while it may. */
  refusal(): InvocationRefusal | null;
}

export interface PluginRuntimeOptions<
  Plugin extends RuntimePlugin<Context, Event>,
  Context,
  Event extends RuntimeEvent,
  Snapshot extends RuntimeSnapshot,
  Phase extends string
> {
  context(invocation: PluginInvocation<Snapshot>, plugin: Plugin): Context;
  snapshot(pluginId: string, generation: string): Snapshot;
  /** The document's current version; null without a document. */
  currentVersion(): string | null;
  loadEvent(snapshot: Snapshot, reason: PluginLoadReason): Event;
  grantsEvent(snapshot: Snapshot, grant: PluginGrant<string>): Event;
  /** A problem that disables the plugin, or null. */
  validate?(plugin: Plugin): unknown;
  report(error: PluginError<Phase | RuntimePhase>): void;
}

/** A ready activation, as contributions render it. */
export interface RuntimeActivation<Plugin, Context> {
  readonly pluginId: string;
  readonly plugin: Plugin;
  readonly generation: string;
  readonly activation: object;
  /** Unique per activation, for keying rendered contributions. */
  readonly key: string;
  /** Context for renderers; replaced whenever the snapshot or state changes. */
  readonly context: Context;
}

export type InvokeOutcome<T> =
  | { ok: true; value: T }
  | { ok: false; reason: InvocationRefusal | 'failed' };

export interface PluginRuntime<Plugin, Context, Event, Phase extends string> {
  /** Reconciles registrations by id and revision; call after the host commits. */
  setPlugins(plugins: readonly Plugin[]): void;
  /** Rereads every grant, including one changed in place, and notifies the plugins it changed. */
  setGrants(grants: Readonly<Record<string, PluginGrant<string>>> | null | undefined): void;
  grant(pluginId: string): PluginGrant<string>;
  /** Activates every registration for a ready document. */
  open(generation: string): void;
  /** Ends every activation at once. */
  close(reason: 'document-replaced' | 'unmounted'): void;
  generation(): string | null;
  notify(event: Event): void;
  /** Rebuilds render contexts after snapshot inputs changed without an event. */
  touch(): void;
  /** Ready activations in registration order; stable until listeners are notified. */
  activations(): readonly RuntimeActivation<Plugin, Context>[];
  subscribe(listener: () => void): () => void;
  /**
   * Runs a pure plugin function, as during a host render. A throw returns `fallback` and
   * quarantines the activation right after.
   */
  guard<T>(pluginId: string, phase: Phase, call: (context: Context) => T, fallback: T): T;
  /** Runs a managed plugin effect with a fresh invocation. */
  invoke<T>(
    pluginId: string,
    phase: Phase,
    call: (context: Context) => MaybePromise<T>
  ): Promise<InvokeOutcome<T>>;
  /** Quarantines the current activation of `pluginId`. */
  fail(pluginId: string, phase: Phase, error: unknown): void;
}

const PLUGIN_ID = /^[A-Za-z0-9](?:[A-Za-z0-9._-]{0,126}[A-Za-z0-9])?$/;

/** Plugin ids are letters, digits, `.`, `_` and `-`, starting and ending alphanumerically. */
export function isValidPluginId(id: unknown): id is string {
  return typeof id === 'string' && PLUGIN_ID.test(id);
}

interface Registration<Plugin> {
  readonly id: string;
  plugin: Plugin;
  quarantined: boolean;
  activation: Activation<Plugin> | null;
}

interface Activation<Plugin> {
  readonly token: object;
  readonly key: string;
  readonly registration: Registration<Plugin>;
  readonly generation: string;
  readonly lifetime: AbortController;
  status: 'initializing' | 'ready' | 'disposed';
  ended: PluginCleanupReason | null;
  state: unknown;
  stateRevision: number;
  cleanups: Array<(reason: PluginCleanupReason) => MaybePromise<void>>;
  pending: RuntimeEvent[];
  running: { channel: string; controller: AbortController } | null;
  pumping: boolean;
  rendered: { key: string; entry: RuntimeActivation<Plugin, unknown> } | null;
}

function revisionKey(plugin: RuntimePlugin<unknown, unknown>): string {
  return plugin.revision === undefined ? '' : `${typeof plugin.revision}:${plugin.revision}`;
}

/** How often a load hook runs before changes it did not make count as a failure. */
const MAX_LOAD_RUNS = 10;

/** Load and document changes share a channel: a newer version supersedes both. */
function channelOf(event: RuntimeEvent): string {
  return event.type === 'load' ? 'document-change' : event.type;
}

function isThenable(value: unknown): value is PromiseLike<unknown> {
  return (
    value !== null &&
    (typeof value === 'object' || typeof value === 'function') &&
    typeof (value as PromiseLike<unknown>).then === 'function'
  );
}

/** A controller aborted with `parent`; `release` detaches it once its work settles. */
function linked(parent: AbortSignal): { controller: AbortController; release(): void } {
  const controller = new AbortController();
  if (parent.aborted) {
    controller.abort();
    return { controller, release: () => {} };
  }
  const abort = () => controller.abort();
  parent.addEventListener('abort', abort, { once: true });
  return { controller, release: () => parent.removeEventListener('abort', abort) };
}

export function createPluginRuntime<
  Plugin extends RuntimePlugin<Context, Event>,
  Context,
  Event extends RuntimeEvent,
  Snapshot extends RuntimeSnapshot,
  Phase extends string
>(
  options: PluginRuntimeOptions<Plugin, Context, Event, Snapshot, Phase>
): PluginRuntime<Plugin, Context, Event, Phase | RuntimePhase> {
  type Entry = RuntimeActivation<Plugin, Context>;
  let registrations = new Map<string, Registration<Plugin>>();
  let current: string | null = null;
  let opened = false;
  let grants: Readonly<Record<string, PluginGrant<string>>> = {};
  const normalized = new Map<string, { source: unknown; grant: PluginGrant<string> }>();
  const reportedPlugins = new WeakSet<object>();
  let reportedDuplicates = new Set<string>();
  const listeners = new Set<() => void>();
  let epoch = 0;
  let activations = 0;
  let committing: AbortSignal | null = null;
  let published: readonly Entry[] = [];
  let scheduled = false;

  const report = (error: PluginError<Phase | RuntimePhase>): void => {
    try {
      options.report(error);
    } catch (reporterError) {
      console.error('[plugins] onPluginError threw', reporterError);
    }
  };

  const grantOf = (pluginId: string): PluginGrant<string> => {
    const source = Object.hasOwn(grants, pluginId) ? grants[pluginId] : undefined;
    const cached = normalized.get(pluginId);
    if (cached && cached.source === source) return cached.grant;
    const grant = normalizeGrant(source);
    normalized.set(pluginId, { source, grant });
    return grant;
  };

  const renderEntry = (activation: Activation<Plugin>): Entry => {
    const key = `${epoch}:${activation.stateRevision}`;
    if (activation.rendered?.key === key) return activation.rendered.entry as Entry;
    const invocation = invocationOf(activation, activation.lifetime.signal);
    const entry: Entry = {
      pluginId: activation.registration.id,
      plugin: activation.registration.plugin,
      generation: activation.generation,
      activation: activation.token,
      key: activation.key,
      context: options.context(invocation, activation.registration.plugin),
    };
    activation.rendered = { key, entry: entry as RuntimeActivation<Plugin, unknown> };
    return entry;
  };

  const publish = (): void => {
    scheduled = false;
    const next: Entry[] = [];
    for (const registration of registrations.values()) {
      const activation = registration.activation;
      if (activation?.status === 'ready') next.push(renderEntry(activation));
    }
    const same =
      next.length === published.length && next.every((entry, index) => entry === published[index]);
    if (!same) published = next;
    for (const listener of [...listeners]) listener();
  };

  const changed = (): void => {
    epoch += 1;
    if (scheduled) return;
    scheduled = true;
    queueMicrotask(publish);
  };

  const refusalOf = (
    activation: Activation<Plugin>,
    signal: AbortSignal
  ): InvocationRefusal | null => {
    if (activation.status === 'disposed') {
      return activation.ended === 'document-replaced' ? 'document-replaced' : 'plugin-unavailable';
    }
    return signal.aborted ? 'aborted' : null;
  };

  const runCleanup = (
    activation: Activation<Plugin>,
    cleanup: (reason: PluginCleanupReason) => MaybePromise<void>,
    reason: PluginCleanupReason
  ): void => {
    const failed = (error: unknown) =>
      report({
        pluginId: activation.registration.id,
        generation: activation.generation,
        phase: 'cleanup',
        error,
      });
    try {
      const result = cleanup(reason);
      if (isThenable(result)) result.then(undefined, failed);
    } catch (error) {
      failed(error);
    }
  };

  const dispose = (activation: Activation<Plugin>, reason: PluginCleanupReason): void => {
    if (activation.status === 'disposed') return;
    activation.status = 'disposed';
    activation.ended = reason;
    activation.pending = [];
    if (activation.registration.activation === activation)
      activation.registration.activation = null;
    activation.lifetime.abort();
    for (const cleanup of activation.cleanups.splice(0).reverse()) {
      runCleanup(activation, cleanup, reason);
    }
    changed();
  };

  const quarantine = (
    activation: Activation<Plugin>,
    phase: Phase | RuntimePhase,
    error: unknown
  ): void => {
    if (activation.status === 'disposed') return;
    activation.registration.quarantined = true;
    report({
      pluginId: activation.registration.id,
      generation: activation.generation,
      phase,
      error,
    });
    dispose(activation, 'failed');
  };

  function invocationOf(
    activation: Activation<Plugin>,
    signal: AbortSignal
  ): PluginInvocation<Snapshot> {
    const snapshot = options.snapshot(activation.registration.id, activation.generation);
    const invocation: PluginInvocation<Snapshot> = {
      pluginId: activation.registration.id,
      activation: activation.token,
      snapshot,
      signal,
      lifetimeSignal: activation.lifetime.signal,
      state: () => activation.state,
      setState(next, atVersion) {
        if (refusalOf(activation, signal) !== null) return false;
        const version = options.currentVersion();
        if (version === null || (atVersion ?? snapshot.version) !== version) return false;
        activation.state =
          typeof next === 'function'
            ? (next as (previous: unknown) => unknown)(activation.state)
            : next;
        activation.stateRevision += 1;
        changed();
        return true;
      },
      onCleanup(cleanup) {
        if (typeof cleanup !== 'function') throw new TypeError('A cleanup must be a function');
        if (activation.status === 'disposed') {
          runCleanup(activation, cleanup, activation.ended ?? 'unmounted');
        } else {
          activation.cleanups.push(cleanup);
        }
      },
      run: (action) => runAction(activation, action),
      commit(write) {
        const outer = committing;
        committing = signal;
        try {
          return write();
        } finally {
          committing = outer;
        }
      },
      refusal: () => refusalOf(activation, signal),
    };
    return invocation;
  }

  async function runAction(
    activation: Activation<Plugin>,
    action: (invocation: PluginInvocation<Snapshot>) => MaybePromise<void>
  ): Promise<void> {
    const { controller, release } = linked(activation.lifetime.signal);
    if (controller.signal.aborted) return;
    try {
      await action(invocationOf(activation, controller.signal));
    } catch (error) {
      if (!controller.signal.aborted) quarantine(activation, 'action', error);
    } finally {
      release();
    }
  }

  /** Runs one lifecycle hook single-flight; a newer event on `channel` aborts it. */
  async function hook(
    activation: Activation<Plugin>,
    phase: Phase | RuntimePhase,
    channel: string,
    call: (context: Context, invocation: PluginInvocation<Snapshot>) => MaybePromise<void>
  ): Promise<'ok' | 'aborted' | 'ended'> {
    const { controller, release } = linked(activation.lifetime.signal);
    activation.running = { channel, controller };
    try {
      const invocation = invocationOf(activation, controller.signal);
      await call(options.context(invocation, activation.registration.plugin), invocation);
      if (activation.status === 'disposed') return 'ended';
      return controller.signal.aborted ? 'aborted' : 'ok';
    } catch (error) {
      if (activation.status === 'disposed') return 'ended';
      if (controller.signal.aborted) return 'aborted';
      quarantine(activation, phase, error);
      return 'ended';
    } finally {
      release();
      if (activation.running?.controller === controller) activation.running = null;
    }
  }

  async function pump(activation: Activation<Plugin>): Promise<void> {
    if (activation.pumping) return;
    activation.pumping = true;
    try {
      while (activation.status === 'ready') {
        const event = activation.pending.shift();
        if (!event) break;
        if (!activation.registration.plugin.onEvent) continue;
        const outcome = await hook(activation, 'event', channelOf(event), (context) =>
          activation.registration.plugin.onEvent?.(context, event as Event)
        );
        if (outcome === 'ended') break;
      }
    } finally {
      activation.pumping = false;
    }
  }

  const enqueue = (activation: Activation<Plugin>, event: RuntimeEvent): void => {
    if (activation.status === 'disposed' || event.generation !== activation.generation) return;
    const channel = channelOf(event);
    const index = activation.pending.findIndex((pending) => channelOf(pending) === channel);
    if (index >= 0) {
      if (activation.pending[index].type === 'load') return;
      activation.pending[index] = event;
    } else {
      activation.pending.push(event);
    }
    const running = activation.running;
    if (running?.channel === channel && running.controller.signal !== committing) {
      running.controller.abort();
    }
    if (activation.status === 'ready') queueMicrotask(() => void pump(activation));
  };

  async function activate(
    registration: Registration<Plugin>,
    generation: string,
    reason: PluginLoadReason
  ): Promise<void> {
    activations += 1;
    const activation: Activation<Plugin> = {
      token: Object.freeze({}),
      key: `${registration.id}#${activations}`,
      registration,
      generation,
      lifetime: new AbortController(),
      status: 'initializing',
      ended: null,
      state: undefined,
      stateRevision: 0,
      cleanups: [],
      pending: [],
      running: null,
      pumping: false,
      rendered: null,
    };
    registration.activation = activation;
    try {
      activation.state = registration.plugin.createState();
    } catch (error) {
      quarantine(activation, 'initialize', error);
      return;
    }
    if (registration.plugin.initialize) {
      const initialized = await hook(activation, 'initialize', 'initialize', (context) =>
        activation.registration.plugin.initialize?.(context)
      );
      if (initialized === 'ended') return;
    }
    let loaded: 'ok' | 'aborted' | 'ended' = 'aborted';
    for (let runs = 0; loaded === 'aborted' && activation.status !== 'disposed'; runs += 1) {
      if (runs === MAX_LOAD_RUNS) {
        quarantine(
          activation,
          'event',
          new Error(`Document changes superseded the load hook ${MAX_LOAD_RUNS} times`)
        );
        return;
      }
      activation.pending = [];
      loaded = await hook(activation, 'event', 'document-change', (context, invocation) =>
        activation.registration.plugin.onEvent?.(
          context,
          options.loadEvent(invocation.snapshot, reason)
        )
      );
    }
    if (loaded !== 'ok' || activation.status === 'disposed') return;
    activation.status = 'ready';
    changed();
    void pump(activation);
  }

  const reportDefinition = (pluginId: string, error: unknown): void =>
    report({ pluginId, generation: current, phase: 'definition', error });

  const close = (reason: 'document-replaced' | 'unmounted'): void => {
    current = null;
    if (reason === 'unmounted') opened = false;
    for (const registration of registrations.values()) {
      if (registration.activation) dispose(registration.activation, reason);
    }
    changed();
  };

  return {
    setPlugins(plugins) {
      const byId = new Map<string, Plugin[]>();
      for (const plugin of plugins) {
        const problem = !isValidPluginId(plugin?.id)
          ? new TypeError(`Invalid plugin id ${JSON.stringify(plugin?.id)}`)
          : options.validate?.(plugin);
        if (problem != null) {
          if (plugin && typeof plugin === 'object' && !reportedPlugins.has(plugin)) {
            reportedPlugins.add(plugin);
            reportDefinition(typeof plugin.id === 'string' ? plugin.id : '', problem);
          }
          continue;
        }
        const list = byId.get(plugin.id);
        if (list) list.push(plugin);
        else byId.set(plugin.id, [plugin]);
      }
      const duplicates = new Set<string>();
      const next = new Map<string, Registration<Plugin>>();
      const retired: Array<[Registration<Plugin>, PluginCleanupReason]> = [];
      for (const [id, list] of byId) {
        if (list.length > 1) {
          duplicates.add(id);
          if (!reportedDuplicates.has(id)) {
            reportDefinition(id, new Error(`${list.length} plugins share the id "${id}"`));
          }
          continue;
        }
        const plugin = list[0];
        const existing = registrations.get(id);
        if (existing && revisionKey(existing.plugin) === revisionKey(plugin)) {
          existing.plugin = plugin;
          next.set(id, existing);
          continue;
        }
        if (existing) retired.push([existing, 'definition-replaced']);
        next.set(id, { id, plugin, quarantined: false, activation: null });
      }
      reportedDuplicates = duplicates;
      for (const [id, registration] of registrations) {
        if (!next.has(id)) retired.push([registration, 'removed']);
      }
      registrations = next;
      for (const [registration, reason] of retired) {
        if (registration.activation) dispose(registration.activation, reason);
      }
      if (current !== null) {
        for (const registration of registrations.values()) {
          if (!registration.activation && !registration.quarantined) {
            void activate(registration, current, 'attached');
          }
        }
      }
      changed();
    },

    setGrants(next) {
      const before = new Map(
        [...registrations.keys()].map((pluginId) => [pluginId, grantOf(pluginId)] as const)
      );
      grants = next ?? {};
      normalized.clear();
      let revised = false;
      for (const registration of registrations.values()) {
        const grant = grantOf(registration.id);
        const previous = before.get(registration.id)!;
        if (sameGrant(previous, grant)) {
          normalized.get(registration.id)!.grant = previous;
          continue;
        }
        revised = true;
        const activation = registration.activation;
        if (!activation) continue;
        enqueue(
          activation,
          options.grantsEvent(options.snapshot(registration.id, activation.generation), grant)
        );
      }
      if (revised) changed();
    },

    grant: grantOf,

    open(generation) {
      if (current === generation) return;
      if (current !== null) close('document-replaced');
      current = generation;
      const reason: PluginLoadReason = opened ? 'replaced' : 'loaded';
      opened = true;
      for (const registration of registrations.values()) {
        registration.quarantined = false;
        void activate(registration, generation, reason);
      }
      changed();
    },

    close,

    generation: () => current,

    notify(event) {
      if (event.generation !== current) return;
      for (const registration of registrations.values()) {
        if (registration.activation) enqueue(registration.activation, event);
      }
      changed();
    },

    touch: changed,

    activations: () => published,

    subscribe(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },

    guard(pluginId, phase, call, fallback) {
      const activation = registrations.get(pluginId)?.activation;
      if (activation?.status !== 'ready') return fallback;
      try {
        return call(renderEntry(activation).context);
      } catch (error) {
        queueMicrotask(() => quarantine(activation, phase, error));
        return fallback;
      }
    },

    async invoke(pluginId, phase, call) {
      const activation = registrations.get(pluginId)?.activation;
      if (activation?.status !== 'ready') return { ok: false, reason: 'plugin-unavailable' };
      const { controller, release } = linked(activation.lifetime.signal);
      const invocation = invocationOf(activation, controller.signal);
      try {
        const value = await call(options.context(invocation, activation.registration.plugin));
        return { ok: true, value };
      } catch (error) {
        const refusal = invocation.refusal();
        if (refusal) return { ok: false, reason: refusal };
        quarantine(activation, phase, error);
        return { ok: false, reason: 'failed' };
      } finally {
        release();
      }
    },

    fail(pluginId, phase, error) {
      const activation = registrations.get(pluginId)?.activation;
      if (activation) quarantine(activation, phase, error);
    },
  };
}
