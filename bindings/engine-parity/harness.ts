import { expect } from 'bun:test';
import type { Surface } from '../../scripts/check-engine-surface';
import { expectSavedPartsEqual } from '../native-parity';

export interface EngineHandle<T> {
  target: T;
  /** Values the engine generated itself, such as nonces, read before and after each call. */
  generated(): readonly string[];
  save(): Promise<Uint8Array>;
  free(): void;
}

export type MakeEngine<T> = () => Promise<EngineHandle<T>>;

export interface RowExpectation {
  method: string;
  /** Distinguishes a second row of the same method in test names. */
  name?: string;
  mutates?: boolean;
  throws?: { name: string; reason: string };
  voidReason?: string;
}

export interface Row<T> extends RowExpectation {
  setup?: (target: T) => unknown;
  call: (target: T, context: never) => unknown;
}

export function rowName(row: RowExpectation): string {
  return row.name ?? row.method;
}

export interface Exclusion {
  method: string;
  reason: string;
}

export interface KnownDifference {
  method: string;
  reason: string;
}

export type Outcome = { value: unknown } | { error: { name: string; code: unknown; message: string } };

export type Normalization = (value: unknown, generated: readonly string[]) => unknown;

export function byteContainer(value: unknown): unknown {
  return Buffer.isBuffer(value) ? Uint8Array.from(value) : value;
}

export function generatedTokens(value: unknown, generated: readonly string[]): unknown {
  if (typeof value !== 'string') return value;
  return [...new Set(generated)]
    .filter((token) => token !== '')
    .reduce((text, token, index) => text.replaceAll(token, `<generated ${index}>`), value);
}

export const normalize: readonly Normalization[] = [byteContainer, generatedTokens];

export function defineRow<T, C>(row: RowExpectation & {
  setup: (target: T) => C | Promise<C>;
  call: (target: T, context: Awaited<C>) => unknown;
}): Row<T> {
  return row;
}

function mapDeep(value: unknown, leaf: (value: unknown) => unknown): unknown {
  if (Array.isArray(value)) return value.map((item) => mapDeep(item, leaf));
  if (value !== null && typeof value === 'object' && [Object.prototype, null].includes(Object.getPrototypeOf(value))) {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, mapDeep(item, leaf)]));
  }
  return leaf(value);
}

function normalized(value: unknown, generated: readonly string[]): unknown {
  return normalize.reduce((current, step) => mapDeep(current, (leaf) => step(leaf, generated)), value);
}

async function capture(run: () => unknown): Promise<Outcome> {
  try {
    return { value: await run() };
  } catch (error) {
    if (error instanceof Error) {
      return { error: { name: error.name, code: (error as Error & { code?: unknown }).code, message: error.message } };
    }
    return { error: { name: typeof error, code: undefined, message: String(error) } };
  }
}

async function observe<T>(row: Row<T>, engine: EngineHandle<T>): Promise<{ outcome: unknown; saved?: Outcome }> {
  try {
    const context = await row.setup?.(engine.target);
    const before = engine.generated();
    const outcome = await capture(() => row.call(engine.target, context as never));
    const generated = [...before, ...engine.generated()];
    if (!row.mutates) return { outcome: normalized(outcome, generated) };
    const saved = await capture(() => engine.save());
    return { outcome: normalized(outcome, generated), saved: normalized(saved, generated) as Outcome };
  } finally {
    engine.free();
  }
}

/** Runs one row on a fresh wasm and a fresh native session and requires equal outcomes and saves. */
export async function runRow<T>(row: Row<T>, makeWasm: MakeEngine<T>, makeNative: MakeEngine<T>): Promise<void> {
  const wasm = await observe(row, await makeWasm());
  const native = await observe(row, await makeNative());
  const name = rowName(row);
  expect(native.outcome, name).toStrictEqual(wasm.outcome);
  expectDeclaredOutcome(row, wasm.outcome as Outcome);
  if (!row.mutates) return;
  const [left, right] = [native.saved!, wasm.saved!];
  if ('value' in left && 'value' in right) {
    await expectSavedPartsEqual(left.value as Uint8Array, right.value as Uint8Array);
  } else {
    expect(left, `${name} save`).toStrictEqual(right);
  }
}

/** An agreed outcome still has to be the declared one: an error needs `throws`, `undefined` needs `mutates` or `voidReason`. */
export function expectDeclaredOutcome(row: RowExpectation, outcome: Outcome): void {
  const name = rowName(row);
  if ('error' in outcome) {
    const { error } = outcome;
    if (!row.throws) throw new Error(`${name}: both engines threw ${error.name}: ${error.message}`);
    expect(error.name, `${name} error`).toBe(row.throws.name);
    return;
  }
  if (row.throws) throw new Error(`${name}: declares throws ${row.throws.name} but both engines returned a value`);
  if (outcome.value === undefined && !row.mutates && !row.voidReason) {
    throw new Error(`${name}: both engines returned undefined; mark the row mutates or give a voidReason`);
  }
}

/** Members declared on both engines: `Class.member` and bare function names. */
export function surfaceMethods(wasm: readonly Surface[], native: Surface): string[] {
  const names = (surface: Surface) => [
    ...[...surface.classes].flatMap(([owner, members]) =>
      [...members.values()].filter((member) => member.kind !== 'constructor').map((member) => `${owner}.${member.name}`)),
    ...surface.functions.keys(),
  ];
  const nativeNames = new Set(names(native));
  return [...new Set(wasm.flatMap(names))].filter((name) => nativeNames.has(name)).sort();
}

export function coverageProblems(
  surface: readonly string[],
  plan: {
    rows: readonly RowExpectation[];
    exclusions: readonly Exclusion[];
    notYetCovered: readonly string[];
    knownDifferences: readonly KnownDifference[];
  },
): string[] {
  const problems: string[] = [];
  const declared = new Set(surface);
  const rowMethods = [...new Set(plan.rows.map((row) => row.method))];
  const owner = new Map<string, string>();
  for (const [list, names] of [
    ['rows', rowMethods],
    ['exclusions', plan.exclusions.map((entry) => entry.method)],
    ['notYetCovered', plan.notYetCovered],
  ] as const) {
    for (const name of names) {
      if (!declared.has(name)) problems.push(`${name}: in ${list} but not on the surface`);
      const previous = owner.get(name);
      if (previous === undefined) owner.set(name, list);
      else problems.push(previous === list ? `${name}: listed twice in ${list}` : `${name}: in both ${previous} and ${list}`);
    }
  }
  for (const name of surface) {
    if (!owner.has(name)) problems.push(`${name}: on the surface but not a row, an exclusion or not yet covered`);
  }
  for (const entry of [...plan.exclusions, ...plan.knownDifferences]) {
    if (entry.reason.trim() === '') problems.push(`${entry.method}: needs a reason`);
  }
  for (const entry of plan.knownDifferences) {
    if (!rowMethods.includes(entry.method)) problems.push(`${entry.method}: a known difference without a row`);
  }
  const names = new Set<string>();
  for (const row of plan.rows) {
    const name = rowName(row);
    if (names.has(name)) problems.push(`${name}: two rows share this name; give a variant its own name`);
    names.add(name);
    if (row.throws && row.throws.reason.trim() === '') problems.push(`${name}: throws needs a reason`);
    if (row.voidReason !== undefined && row.voidReason.trim() === '') problems.push(`${name}: voidReason is empty`);
  }
  return problems;
}
