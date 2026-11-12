import { expect } from 'bun:test';
import JSZip from 'jszip';
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

/** A row return that is a zip package. */
export class ZipReturn {
  constructor(readonly bytes: Uint8Array) {}
}

/** A row return framed as a u64 little-endian warnings JSON length, that JSON, then a zip package. */
export class FramedZipReturn {
  constructor(readonly bytes: Uint8Array) {}
}

export async function zipEntries(bytes: Uint8Array): Promise<{ entryNames: string[]; parts: [string, Uint8Array][] }> {
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const entryNames = Object.keys(zip.files);
  const files = entryNames.filter((name) => !zip.files[name]!.dir).sort();
  const parts = await Promise.all(files.map(async (name): Promise<[string, Uint8Array]> => [name, await zip.files[name]!.async('uint8array')]));
  return { entryNames, parts };
}

export async function zipParts(value: unknown): Promise<unknown> {
  if (value instanceof ZipReturn) return zipEntries(value.bytes);
  if (!(value instanceof FramedZipReturn)) return value;
  const { bytes } = value;
  const length = Number(new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getBigUint64(0, true));
  const warningsJson = new TextDecoder().decode(bytes.subarray(8, 8 + length));
  return { warningsJson, warnings: JSON.parse(warningsJson) as unknown, ...await zipEntries(bytes.subarray(8 + length)) };
}

export function byteContainer(value: unknown): unknown {
  return Buffer.isBuffer(value) ? Uint8Array.from(value) : value;
}

export function generatedTokens(value: unknown, generated: readonly string[]): unknown {
  if (typeof value !== 'string') return value;
  return [...new Set(generated)]
    .filter((token) => token !== '')
    .reduce((text, token, index) => text.replaceAll(token, `<generated ${index}>`), value);
}

export const normalize: readonly Normalization[] = [zipParts, byteContainer, generatedTokens];

export function defineRow<T, C>(row: RowExpectation & {
  setup: (target: T) => C | Promise<C>;
  call: (target: T, context: Awaited<C>) => unknown;
}): Row<T> {
  return row;
}

async function mapDeep(value: unknown, leaf: (value: unknown) => unknown): Promise<unknown> {
  if (Array.isArray(value)) return Promise.all(value.map((item) => mapDeep(item, leaf)));
  if (value !== null && typeof value === 'object' && [Object.prototype, null].includes(Object.getPrototypeOf(value))) {
    return Object.fromEntries(await Promise.all(Object.entries(value).map(async ([key, item]) => [key, await mapDeep(item, leaf)])));
  }
  return leaf(value);
}

async function normalized(value: unknown, generated: readonly string[]): Promise<unknown> {
  let current = value;
  for (const step of normalize) current = await mapDeep(current, (leaf) => step(leaf, generated));
  return current;
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

interface Observation {
  outcome: unknown;
  saved?: Outcome;
}

async function observe<T>(row: Row<T>, engine: EngineHandle<T>, side: string): Promise<Observation> {
  try {
    const setup = await capture(() => row.setup?.(engine.target));
    if ('error' in setup) throw new Error(`${rowName(row)}: ${side} setup threw ${setup.error.name}: ${setup.error.message}`);
    const before = engine.generated();
    const outcome = await capture(() => row.call(engine.target, setup.value as never));
    const generated = [...before, ...engine.generated()];
    if (!row.mutates) return { outcome: await normalized(outcome, generated) };
    const saved = await capture(() => engine.save());
    return { outcome: await normalized(outcome, generated), saved: await normalized(saved, generated) as Outcome };
  } finally {
    engine.free();
  }
}

async function observeBoth<T>(row: Row<T>, makeWasm: MakeEngine<T>, makeNative: MakeEngine<T>) {
  const wasm = await observe(row, await makeWasm(), 'wasm');
  const native = await observe(row, await makeNative(), 'native');
  return { wasm, native };
}

async function expectPackage(name: string, side: string, saved: Outcome, mainPart: string): Promise<Uint8Array> {
  if ('error' in saved) throw new Error(`${name}: ${side} save threw ${saved.error.name}: ${saved.error.message}`);
  const bytes = saved.value;
  if (!(bytes instanceof Uint8Array) || bytes.byteLength === 0) throw new Error(`${name}: ${side} save returned no bytes`);
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true }).catch((error: unknown) => {
    throw new Error(`${name}: ${side} save is not a valid zip: ${String(error)}`);
  });
  if (!zip.file(mainPart)) throw new Error(`${name}: ${side} save has no ${mainPart}`);
  return bytes;
}

/** Runs one row on a fresh wasm and a fresh native session and requires equal outcomes and saves. */
export async function runRow<T>(
  row: Row<T>,
  makeWasm: MakeEngine<T>,
  makeNative: MakeEngine<T>,
  mainPart = 'word/document.xml',
): Promise<void> {
  const { wasm, native } = await observeBoth(row, makeWasm, makeNative);
  const name = rowName(row);
  expect(native.outcome, name).toStrictEqual(wasm.outcome);
  expectDeclaredOutcome(row, wasm.outcome as Outcome);
  if (!row.mutates) return;
  const nativeBytes = await expectPackage(name, 'native', native.saved!, mainPart);
  const wasmBytes = await expectPackage(name, 'wasm', wasm.saved!, mainPart);
  await expectSavedPartsEqual(nativeBytes, wasmBytes);
}

async function differs(check: () => unknown): Promise<boolean> {
  try {
    await check();
    return false;
  } catch {
    return true;
  }
}

/** A known difference holds only while both engines run the row and their outcomes or saves still differ. */
export async function runKnownDifference<T>(row: Row<T>, makeWasm: MakeEngine<T>, makeNative: MakeEngine<T>): Promise<void> {
  const name = rowName(row);
  const observed = await observeBoth(row, makeWasm, makeNative).catch((error: unknown) => {
    throw new Error(`${error instanceof Error ? error.message : String(error)}; a known difference needs both engines to run the row, so fix the setup or remove the knownDifferences entry`);
  });
  const { wasm, native } = observed;
  if (await differs(() => expect(native.outcome).toStrictEqual(wasm.outcome))) return;
  if (row.mutates) {
    const [left, right] = [native.saved!, wasm.saved!];
    const savedDiffer = 'value' in left && 'value' in right
      ? await differs(() => expectSavedPartsEqual(left.value as Uint8Array, right.value as Uint8Array))
      : await differs(() => expect(left).toStrictEqual(right));
    if (savedDiffer) return;
  }
  throw new Error(`${name}: both engines now agree; remove the knownDifferences entry`);
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

export type EngineSide = 'native' | 'wasm';

export interface OneEngineExport {
  name: string;
  side: EngineSide;
}

export interface OneEngineOnly extends OneEngineExport {
  reason: 'BY_DESIGN' | 'PARKED' | 'GAP' | 'BUILD_METADATA';
  note: string;
}

function exportNames(surfaces: readonly Surface[]) {
  const classes = new Map<string, Set<string>>();
  for (const surface of surfaces) {
    for (const [owner, members] of surface.classes) {
      const names = classes.get(owner) ?? new Set<string>();
      for (const member of members.keys()) names.add(member);
      classes.set(owner, names);
    }
  }
  return { classes, bare: new Set(surfaces.flatMap((surface) => [...surface.functions.keys(), ...surface.consts])) };
}

/** Exports declared on one engine only: a whole class by its name, otherwise `Class.member`, functions and consts. */
export function oneEngineExports(wasm: readonly Surface[], native: Surface): OneEngineExport[] {
  const sides = { wasm: exportNames(wasm), native: exportNames([native]) };
  const only = (side: EngineSide, other: EngineSide): OneEngineExport[] => [
    ...[...sides[side].classes].flatMap(([owner, members]) => {
      const shared = sides[other].classes.get(owner);
      return shared ? [...members].filter((member) => !shared.has(member)).map((member) => `${owner}.${member}`) : [owner];
    }),
    ...[...sides[side].bare].filter((name) => !sides[other].bare.has(name)),
  ].map((name) => ({ name, side }));
  return [...only('wasm', 'native'), ...only('native', 'wasm')].sort((a, b) => a.name.localeCompare(b.name));
}

export function oneEngineProblems(actual: readonly OneEngineExport[], listed: readonly OneEngineOnly[]): string[] {
  const problems: string[] = [];
  const key = (entry: OneEngineExport) => `${entry.side}:${entry.name}`;
  const actualKeys = new Set(actual.map(key));
  const listedKeys = new Set<string>();
  for (const entry of listed) {
    if (listedKeys.has(key(entry))) problems.push(`${entry.name}: listed twice in oneEngineOnly`);
    listedKeys.add(key(entry));
    if (!actualKeys.has(key(entry))) problems.push(`${entry.name}: listed as ${entry.side} only but is not one-engine-only`);
    if (entry.note.trim() === '') problems.push(`${entry.name}: oneEngineOnly needs a note`);
  }
  for (const entry of actual) {
    if (!listedKeys.has(key(entry))) problems.push(`${entry.name}: only on ${entry.side} but not in oneEngineOnly`);
  }
  return problems;
}

export interface CoveragePlan {
  rows: readonly RowExpectation[];
  exclusions: readonly Exclusion[];
  notYetCovered: readonly string[];
  knownDifferences: readonly KnownDifference[];
  oneEngineOnly: readonly OneEngineOnly[];
}

export function printParityCounts(format: string, surface: readonly string[], plan: CoveragePlan): void {
  console.log([
    `PARITY-COUNTS ${format}`,
    `surface=${surface.length}`,
    `rows=${plan.rows.length}`,
    `methodsCovered=${new Set(plan.rows.map((row) => row.method)).size}`,
    `excluded=${plan.exclusions.length}`,
    `notYetCovered=${plan.notYetCovered.length}`,
    `errorRows=${plan.rows.filter((row) => row.throws).length}`,
    `knownDifferences=${plan.knownDifferences.length}`,
    `oneEngineOnly=${plan.oneEngineOnly.length}`,
  ].join(' '));
}

export function coverageProblems(surface: readonly string[], plan: CoveragePlan): string[] {
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
