import { describe, expect, test } from 'bun:test';
import type { PluginError } from '../host-contracts/plugins';
import { normalizeGrant, grantsEditBatch } from './grants';
import {
  createPluginRuntime,
  isValidPluginId,
  type PluginInvocation,
  type RuntimePlugin,
} from './runtime';

interface Snapshot {
  generation: string;
  version: string;
  grant: unknown;
}

interface Event {
  type: string;
  generation: string;
  version?: string;
  reason?: string;
  grant?: unknown;
}

interface Context {
  pluginId: string;
  invocation: PluginInvocation<Snapshot>;
}

type Plugin = RuntimePlugin<Context, Event>;

function deferred<T = void>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, fail) => {
    resolve = done;
    reject = fail;
  });
  return { promise, resolve, reject };
}

const tick = async (times = 5) => {
  for (let index = 0; index < times; index += 1) await Promise.resolve();
};

function setup() {
  const state = { version: 'v1' };
  const errors: PluginError<string>[] = [];
  const runtime = createPluginRuntime<Plugin, Context, Event, Snapshot, 'render' | 'command'>({
    context: (invocation) => ({ pluginId: invocation.pluginId, invocation }),
    snapshot: (pluginId, generation) => ({
      generation,
      version: state.version,
      grant: runtime.grant(pluginId),
    }),
    currentVersion: () => (runtime.generation() === null ? null : state.version),
    loadEvent: (snapshot, reason) => ({
      type: 'load',
      generation: snapshot.generation,
      version: snapshot.version,
      reason,
    }),
    grantsEvent: (snapshot, grant) => ({
      type: 'grants-change',
      generation: snapshot.generation,
      grant,
    }),
    report: (error) => {
      errors.push(error);
    },
  });
  const change = (version: string) => {
    state.version = version;
    runtime.notify({ type: 'document-change', generation: runtime.generation()!, version });
  };
  return { runtime, errors, state, change };
}

function recorder(id: string, extra: Partial<Plugin> = {}) {
  const log: string[] = [];
  const contexts: Context[] = [];
  const plugin: Plugin = {
    id,
    createState: () => ({ count: 0 }),
    initialize(context) {
      contexts.push(context);
      log.push('initialize');
      context.invocation.onCleanup((reason) => {
        log.push(`cleanup:${reason}`);
      });
    },
    onEvent(context, event) {
      contexts.push(context);
      log.push(`${event.type}:${event.reason ?? event.version ?? ''}`);
    },
    ...extra,
  };
  return { plugin, log, contexts };
}

describe('plugin runtime lifecycle', () => {
  test('initializes, loads, then delivers changes and cleans up once on removal', async () => {
    const { runtime, change } = setup();
    const { plugin, log } = recorder('acme.review');
    runtime.setPlugins([plugin]);
    await tick();
    expect(log).toEqual([]);
    runtime.open('g1');
    await tick();
    expect(log).toEqual(['initialize', 'load:loaded']);
    expect(runtime.activations().map((entry) => entry.pluginId)).toEqual(['acme.review']);
    change('v2');
    await tick();
    expect(log.at(-1)).toBe('document-change:v2');
    runtime.setPlugins([]);
    runtime.setPlugins([]);
    await tick();
    expect(log.filter((entry) => entry.startsWith('cleanup'))).toEqual(['cleanup:removed']);
    expect(runtime.activations()).toEqual([]);
  });

  test('a plugin added to an open document loads with reason attached', async () => {
    const { runtime } = setup();
    runtime.open('g1');
    const { plugin, log } = recorder('late');
    runtime.setPlugins([plugin]);
    await tick();
    expect(log).toEqual(['initialize', 'load:attached']);
  });

  test('replacement during initialization aborts it and starts a fresh activation', async () => {
    const { runtime, state } = setup();
    const gate = deferred();
    const seen: Context[] = [];
    const { plugin, log } = recorder('slow', {
      async initialize(context) {
        seen.push(context);
        context.invocation.onCleanup((reason) => {
          log.push(`cleanup:${reason}`);
        });
        await gate.promise;
        log.push(`initialized:${context.invocation.setState({ count: 1 })}`);
      },
    });
    runtime.setPlugins([plugin]);
    runtime.open('g1');
    await tick();
    runtime.close('document-replaced');
    expect(seen[0].invocation.signal.aborted).toBe(true);
    expect(seen[0].invocation.refusal()).toBe('document-replaced');
    expect(log).toEqual(['cleanup:document-replaced']);
    state.version = 'w1';
    runtime.open('g2');
    gate.resolve();
    await tick(10);
    expect(log).toEqual([
      'cleanup:document-replaced',
      'initialized:false',
      'initialized:true',
      'load:replaced',
    ]);
    expect(seen[1].invocation.activation).not.toBe(seen[0].invocation.activation);
    expect(seen[1].invocation.snapshot.generation).toBe('g2');
  });

  test('removal and readdition use distinct activations', async () => {
    const { runtime } = setup();
    const first = recorder('same');
    runtime.setPlugins([first.plugin]);
    runtime.open('g1');
    await tick();
    const old = first.contexts[0].invocation;
    runtime.setPlugins([]);
    runtime.setPlugins([first.plugin]);
    await tick();
    const fresh = first.contexts.at(-1)!.invocation;
    expect(fresh.activation).not.toBe(old.activation);
    expect(old.refusal()).toBe('plugin-unavailable');
    expect(old.setState({ count: 9 })).toBe(false);
    expect(first.log).toEqual([
      'initialize',
      'load:loaded',
      'cleanup:removed',
      'initialize',
      'load:attached',
    ]);
  });

  test('setup, cleanup, setup (StrictMode) creates distinct activations', async () => {
    const { runtime } = setup();
    const { plugin, log, contexts } = recorder('strict');
    runtime.setPlugins([plugin]);
    runtime.open('g1');
    runtime.close('unmounted');
    runtime.open('g2');
    await tick();
    expect(log).toEqual(['initialize', 'cleanup:unmounted', 'initialize', 'load:loaded']);
    expect(contexts[0].invocation.activation).not.toBe(contexts[1].invocation.activation);
    expect(runtime.activations()).toHaveLength(1);
  });

  test('duplicate ids disable every conflicting registration until resolved', async () => {
    const { runtime, errors } = setup();
    const a = recorder('dup');
    const b = recorder('dup');
    const other = recorder('other');
    runtime.open('g1');
    runtime.setPlugins([a.plugin, other.plugin, b.plugin]);
    runtime.setPlugins([a.plugin, other.plugin, b.plugin]);
    await tick();
    expect(errors.map((error) => [error.pluginId, error.phase])).toEqual([['dup', 'definition']]);
    expect(runtime.activations().map((entry) => entry.pluginId)).toEqual(['other']);
    runtime.setPlugins([other.plugin, b.plugin]);
    await tick();
    expect(runtime.activations().map((entry) => entry.pluginId)).toEqual(['other', 'dup']);
    expect(isValidPluginId('acme/x')).toBe(false);
    runtime.setPlugins([{ ...other.plugin, id: 'bad id' }]);
    expect(errors.at(-1)?.phase).toBe('definition');
  });

  test('new array identities keep state; reordering reorders; a new revision restarts', async () => {
    const { runtime, change } = setup();
    const a = recorder('a');
    const b = recorder('b');
    runtime.setPlugins([a.plugin, b.plugin]);
    runtime.open('g1');
    await tick();
    expect(a.contexts[0].invocation.setState({ count: 5 })).toBe(true);
    const events: string[] = [];
    runtime.setPlugins([
      b.plugin,
      { ...a.plugin, onEvent: (_context, event) => void events.push(event.type) },
    ]);
    await tick();
    expect(runtime.activations().map((entry) => entry.pluginId)).toEqual(['b', 'a']);
    expect(a.log.filter((entry) => entry === 'initialize')).toHaveLength(1);
    change('v2');
    await tick();
    expect(events).toEqual(['document-change']);
    expect(runtime.activations()[1].context.invocation.state()).toEqual({ count: 5 });
    runtime.setPlugins([b.plugin, { ...a.plugin, revision: 2 }]);
    await tick();
    expect(a.log).toContain('cleanup:definition-replaced');
    expect(runtime.activations()[1].context.invocation.state()).toEqual({ count: 0 });
  });

  test('a cleanup registered after the activation ended runs at once with its reason', async () => {
    const { runtime } = setup();
    const { plugin, contexts } = recorder('late-cleanup');
    runtime.setPlugins([plugin]);
    runtime.open('g1');
    await tick();
    runtime.close('document-replaced');
    const reasons: string[] = [];
    contexts[0].invocation.onCleanup((reason) => {
      reasons.push(reason);
    });
    expect(reasons).toEqual(['document-replaced']);
  });

  test('changes coalesce by kind and a superseded hook cannot publish state', async () => {
    const { runtime, change, state } = setup();
    const gates: Array<ReturnType<typeof deferred>> = [];
    const seen: string[] = [];
    const invocations: PluginInvocation<Snapshot>[] = [];
    const plugin: Plugin = {
      id: 'coalesce',
      createState: () => 0,
      async onEvent(context, event) {
        seen.push(`${event.type}:${event.version}`);
        invocations.push(context.invocation);
        if (event.type !== 'document-change') return;
        const gate = deferred();
        gates.push(gate);
        await gate.promise;
        context.invocation.setState(event.version);
      },
    };
    runtime.setPlugins([plugin]);
    runtime.open('g1');
    await tick();
    change('v2');
    await tick();
    change('v3');
    change('v4');
    runtime.notify({ type: 'selection-change', generation: 'g1', version: 'v4' });
    expect(invocations[1].signal.aborted).toBe(true);
    gates[0].resolve();
    await tick(10);
    expect(seen).toEqual(['load:v1', 'document-change:v2', 'document-change:v4']);
    gates[1].resolve();
    await tick(10);
    expect(seen.at(-1)).toBe('selection-change:v4');
    expect(invocations[1].setState('late')).toBe(false);
    expect(runtime.activations()[0].context.invocation.state()).toBe('v4');
    expect(state.version).toBe('v4');
  });

  test('state updates need the current version', async () => {
    const { runtime, change } = setup();
    const { plugin, contexts } = recorder('versioned');
    runtime.setPlugins([plugin]);
    runtime.open('g1');
    await tick();
    const loadContext = contexts[1];
    change('v2');
    await tick();
    expect(loadContext.invocation.setState({ count: 1 })).toBe(false);
    expect(loadContext.invocation.setState({ count: 1 }, 'v2')).toBe(true);
  });

  test('a failing plugin is quarantined without affecting another or the reporter', async () => {
    const { runtime, errors, change } = setup();
    const healthy = recorder('healthy');
    const cleanups: string[] = [];
    const broken: Plugin = {
      id: 'broken',
      createState: () => null,
      initialize(context) {
        context.invocation.onCleanup(() => {
          cleanups.push('first');
        });
        context.invocation.onCleanup(() => {
          throw new Error('cleanup failed');
        });
      },
      onEvent(_context, event) {
        if (event.type === 'document-change') throw new Error('boom');
      },
    };
    runtime.setPlugins([broken, healthy.plugin]);
    runtime.open('g1');
    await tick();
    change('v2');
    await tick(10);
    expect(errors.map((error) => `${error.pluginId}:${error.phase}`)).toEqual([
      'broken:event',
      'broken:cleanup',
    ]);
    expect(cleanups).toEqual(['first']);
    expect(runtime.activations().map((entry) => entry.pluginId)).toEqual(['healthy']);
    expect(healthy.log.at(-1)).toBe('document-change:v2');
    runtime.setPlugins([broken, healthy.plugin]);
    await tick();
    expect(runtime.activations().map((entry) => entry.pluginId)).toEqual(['healthy']);
    runtime.open('g2');
    await tick();
    expect(runtime.activations().map((entry) => entry.pluginId)).toEqual(['broken', 'healthy']);
  });

  test('a throwing reporter and failing guarded calls stay contained', async () => {
    const original = console.error;
    console.error = () => {};
    try {
      const runtime = createPluginRuntime<Plugin, Context, Event, Snapshot, 'command'>({
        context: (invocation) => ({ pluginId: invocation.pluginId, invocation }),
        snapshot: (_pluginId, generation) => ({ generation, version: 'v', grant: null }),
        currentVersion: () => 'v',
        loadEvent: (snapshot, reason) => ({
          type: 'load',
          generation: snapshot.generation,
          reason,
        }),
        grantsEvent: (snapshot) => ({ type: 'grants-change', generation: snapshot.generation }),
        report: () => {
          throw new Error('reporter');
        },
      });
      runtime.setPlugins([
        { id: 'p', createState: () => null },
        { id: 'q', createState: () => null },
      ]);
      runtime.open('g');
      await tick();
      expect(
        runtime.guard(
          'p',
          'command',
          () => {
            throw new Error('state');
          },
          'fallback'
        )
      ).toBe('fallback');
      const outcome = await runtime.invoke('q', 'command', async () => {
        throw new Error('run');
      });
      expect(outcome).toEqual({ ok: false, reason: 'failed' });
      await tick();
      expect(runtime.activations()).toEqual([]);
    } finally {
      console.error = original;
    }
  });

  test('grant changes reach only the plugins whose grant changed', async () => {
    const { runtime } = setup();
    const a = recorder('a');
    const b = recorder('b');
    runtime.setPlugins([a.plugin, b.plugin]);
    runtime.setGrants({ a: { document: 'write', editBatches: true } });
    runtime.open('g1');
    await tick();
    runtime.setGrants({ a: { document: 'write', editBatches: true }, b: { commands: ['bold'] } });
    await tick();
    expect(a.log.some((entry) => entry.startsWith('grants-change'))).toBe(false);
    expect(b.log.some((entry) => entry.startsWith('grants-change'))).toBe(true);
    expect(runtime.grant('b')).toEqual({ commands: ['bold'] });
    expect(runtime.grant('constructor')).toEqual({ commands: [] });
  });

  test('rendered activations keep their identity until something changes', async () => {
    const { runtime } = setup();
    const { plugin } = recorder('stable');
    runtime.setPlugins([plugin]);
    runtime.open('g1');
    await tick();
    const first = runtime.activations();
    expect(runtime.activations()).toBe(first);
    runtime.touch();
    await tick();
    expect(runtime.activations()).not.toBe(first);
    expect(runtime.activations()[0].context).not.toBe(first[0].context);
  });
});

describe('plugin grants', () => {
  test('normalizes to the vocabulary and gates batches', () => {
    const read = normalizeGrant({ commands: ['bold', 'bold'], editBatches: true } as never);
    expect(read).toEqual({ commands: ['bold'] });
    expect(grantsEditBatch(read, undefined)).toBe(false);
    const write = normalizeGrant({ document: 'write', editBatches: true });
    expect(grantsEditBatch(write, 'separate')).toBe(true);
    expect(grantsEditBatch(write, 'none')).toBe(false);
    expect(
      grantsEditBatch(
        normalizeGrant({ document: 'write', editBatches: true, untrackedHistory: true }),
        'none'
      )
    ).toBe(true);
  });
});
