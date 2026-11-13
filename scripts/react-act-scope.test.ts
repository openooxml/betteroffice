import { expect, test } from 'bun:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { type ActScopeLog, logTestStart, openActQueueLength, openAtStartLine, openSummaryLine } from './react-act-scope';

const INTERNALS = '__CLIENT_INTERNALS_DO_NOT_USE_OR_WARN_USERS_THEY_CANNOT_UPGRADE';
const REACT_PATH = '/repo/node_modules/react/index.js';
const reactWithQueue = (actQueue: unknown) => ({ [INTERNALS]: { actQueue } });
const cacheWith = (exports: unknown, path = REACT_PATH) => ({ [path]: { exports } });
const recorder = () => {
  const lines: string[] = [];
  return { lines, write: (line: string) => { lines.push(line); } };
};
const freshLog = (): ActScopeLog => ({ started: 0, open: 0 });

test('an open act queue at test start prints one line naming the test and the queue length', () => {
  const out = recorder();
  const log = freshLog();
  logTestStart(log, cacheWith(reactWithQueue(null)), out.write);
  logTestStart(log, cacheWith(reactWithQueue([() => null, () => null])), out.write);
  expect(out.lines).toEqual([
    'ACT-SCOPE-OPEN at the start of test #2 in this file: queue length 2 (a test before this one left a React act scope open, or caught a rejected act)',
  ]);
  expect(out.lines[0]).toBe(openAtStartLine(2, 2));
});

test('a null act queue prints nothing', () => {
  const out = recorder();
  const log = freshLog();
  logTestStart(log, cacheWith(reactWithQueue(null)), out.write);
  expect(out.lines).toEqual([]);
  expect(log.started).toBe(1);
  expect(log.open).toBe(0);
});

test('a nested React development or production entry in the cache is read', () => {
  const nested = '/repo/node_modules/.bun/react@19.2.7/node_modules/react/cjs/react.development.js';
  expect(openActQueueLength({ [REACT_PATH]: { exports: reactWithQueue(null) }, [nested]: { exports: reactWithQueue([1]) } })).toBe(1);
  expect(openActQueueLength(cacheWith({ [INTERNALS]: {} }, '/repo/node_modules/react/cjs/react.production.js'))).toBeUndefined();
});

test('no input shape throws or prints', () => {
  const throwing = new Proxy({}, { get: () => { throw new Error('getter'); } });
  const throwingEntries = new Proxy({}, { ownKeys: () => { throw new Error('ownKeys'); } });
  const throwingLength = new Proxy([], { get: (target, key) => { if (key === 'length') throw new Error('length'); return Reflect.get(target, key); } });
  const inputs: unknown[] = [
    undefined,
    null,
    0,
    'cache',
    {},
    throwing,
    throwingEntries,
    { '/repo/node_modules/lodash/index.js': { exports: reactWithQueue([]) } },
    { [REACT_PATH]: undefined },
    { [REACT_PATH]: null },
    { [REACT_PATH]: {} },
    cacheWith(undefined),
    cacheWith({}),
    cacheWith({ [INTERNALS]: null }),
    cacheWith({ [INTERNALS]: {} }),
    cacheWith(throwing),
    cacheWith({ [INTERNALS]: throwing }),
    cacheWith(reactWithQueue(undefined)),
    cacheWith(reactWithQueue(0)),
    cacheWith(reactWithQueue('queue')),
    cacheWith(reactWithQueue({ length: 1 })),
    cacheWith(reactWithQueue(throwingLength)),
  ];
  const out = recorder();
  const log = freshLog();
  for (const input of inputs) {
    expect(() => openActQueueLength(input)).not.toThrow();
    expect(openActQueueLength(input)).toBeUndefined();
    expect(() => logTestStart(log, input, out.write)).not.toThrow();
  }
  expect(out.lines).toEqual([]);
  expect(log.started).toBe(inputs.length);
});

test('a throwing getter anywhere in the read path prints nothing and throws nothing', () => {
  const throwingAt = (target: object, key: PropertyKey) =>
    Object.defineProperty(target, key, { enumerable: true, get: () => { throw new Error(String(key)); } });
  const queue = new Proxy([], { get: (target, key) => { if (key === 'length') throw new Error('length'); return Reflect.get(target, key); } });
  const caches: unknown[] = [
    throwingAt({}, REACT_PATH),
    { [REACT_PATH]: throwingAt({}, 'exports') },
    cacheWith(throwingAt({}, INTERNALS)),
    cacheWith({ [INTERNALS]: throwingAt({}, 'actQueue') }),
    cacheWith(reactWithQueue(queue)),
  ];
  const out = recorder();
  for (const cache of caches) expect(() => logTestStart(freshLog(), cache, out.write)).not.toThrow();
  expect(out.lines).toEqual([]);
});

test('reading the act queue leaves it untouched', () => {
  const queue = [() => null];
  const react = reactWithQueue(queue);
  const internals = react[INTERNALS];
  logTestStart(freshLog(), cacheWith(react), recorder().write);
  expect(internals.actQueue).toBe(queue);
  expect(queue).toHaveLength(1);
  expect(react[INTERNALS]).toBe(internals);
});

test('the summary line gives the count and the first test number, and is absent when nothing was open', () => {
  const log = freshLog();
  const open = cacheWith(reactWithQueue([]));
  const closed = cacheWith(reactWithQueue(null));
  expect(openSummaryLine(log)).toBeUndefined();
  for (const cache of [closed, closed, open, closed, open, open]) logTestStart(log, cache, recorder().write);
  expect(openSummaryLine(log)).toBe('ACT-SCOPE-OPEN summary: 3 tests started with the queue open; first at test #3 in this file');
});

test('the test preload loads in a plain bun process without the test runner', () => {
  const root = resolve(import.meta.dir, '..');
  const dir = mkdtempSync(join(tmpdir(), 'test-preload-plain-'));
  try {
    const script = join(dir, 'empty.ts');
    writeFileSync(script, 'export {};\n');
    const env = { ...process.env };
    delete env.BUN_OPTIONS;
    const child = Bun.spawnSync([process.execPath, '--preload', resolve(root, 'scripts/test-preload.ts'), script], { cwd: root, env, stderr: 'pipe' });
    expect(child.stderr.toString()).toBe('');
    expect(child.exitCode).toBe(0);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('under the test runner the preload counts each test start', () => {
  const log: unknown = Reflect.get(process, Symbol.for('betteroffice.test.actScopeLog'));
  expect(typeof log).toBe('object');
  expect(log).not.toBeNull();
  expect(typeof (log as ActScopeLog).started).toBe('number');
  expect((log as ActScopeLog).started).toBeGreaterThanOrEqual(1);
});
