import { normalizeType, type MemberInfo, type Surface } from './check-engine-surface';

export type MemberKind = 'method' | 'getter' | 'setter' | 'accessor' | 'value';
export type ExportKind = 'class' | 'function' | 'value';

/** One engine module's runtime surface: export paths (`Name`) and member paths
 * (`Name.prototype.member`, `Name.member` for statics, symbols as `[Symbol.dispose]`). */
export type RuntimeSurface = Map<string, ExportKind | MemberKind>;

/** A native-only (or wasm-only) member that is meant to exist. Intended extras name the
 * public API they back; pending ones name the gap-list item that removes them. */
export type Extra = { reason: string } & ({ backs: string } | { pending: `item ${number}` });

/** Engine glue members that are not API: wasm-bindgen and napi internals. */
export type Internal = { reason: string };

export interface Allowances {
  nativeOnly: Record<string, Extra>;
  wasmOnly: Record<string, Extra>;
  internals: { wasm: Record<string, Internal>; native: Record<string, Internal> };
}

export interface SurfaceStats {
  classes: number;
  compared: number;
  equal: number;
  extras: number;
  pending: number;
}

export interface SurfaceReport {
  problems: string[];
  stats: SurfaceStats;
}

const FUNCTION_OWN = new Set(['length', 'name', 'prototype', 'arguments', 'caller']);

function keyName(key: string | symbol): string {
  return typeof key === 'symbol' ? `[${key.description ?? String(key)}]` : key;
}

function memberKind(descriptor: PropertyDescriptor): MemberKind {
  if (descriptor.get && descriptor.set) return 'accessor';
  if (descriptor.get) return 'getter';
  if (descriptor.set) return 'setter';
  return typeof descriptor.value === 'function' ? 'method' : 'value';
}

/** Runtime surface of a module namespace or a CommonJS exports object, classes found by discovery. */
export function discover(module: object): RuntimeSurface {
  const surface: RuntimeSurface = new Map();
  for (const key of Reflect.ownKeys(module)) {
    if (key === Symbol.toStringTag) continue;
    const name = keyName(key);
    const value: unknown = Reflect.get(module, key);
    if (typeof value !== 'function') {
      surface.set(name, 'value');
      continue;
    }
    const members: [string, MemberKind][] = [];
    const prototype: unknown = (value as { prototype?: unknown }).prototype;
    if (prototype !== null && typeof prototype === 'object') {
      for (const member of Reflect.ownKeys(prototype)) {
        if (member === 'constructor') continue;
        members.push([`${name}.prototype.${keyName(member)}`, memberKind(Object.getOwnPropertyDescriptor(prototype, member)!)]);
      }
    }
    for (const member of Reflect.ownKeys(value)) {
      if (typeof member === 'string' && FUNCTION_OWN.has(member)) continue;
      members.push([`${name}.${keyName(member)}`, memberKind(Object.getOwnPropertyDescriptor(value, member)!)]);
    }
    surface.set(name, members.length > 0 ? 'class' : 'function');
    for (const [path, kind] of members) surface.set(path, kind);
  }
  return surface;
}

/** Merges the surfaces of several modules that one binding replaces; a name exported twice must agree. */
export function mergeSurfaces(parts: readonly [string, RuntimeSurface][]): { surface: RuntimeSurface; problems: string[] } {
  const surface: RuntimeSurface = new Map();
  const origin = new Map<string, string>();
  const problems: string[] = [];
  for (const [module, part] of parts) {
    for (const [path, kind] of part) {
      const seen = surface.get(path);
      if (seen !== undefined && seen !== kind) {
        problems.push(`${path}: ${seen} in ${origin.get(path)}, ${kind} in ${module}`);
        continue;
      }
      surface.set(path, kind);
      if (!origin.has(path)) origin.set(path, module);
    }
  }
  return { surface, problems };
}

function owner(path: string): string {
  return path.split('.', 1)[0];
}

function isExport(path: string): boolean {
  return !path.includes('.');
}

function stripInternals(surface: RuntimeSurface, internals: Record<string, Internal>, side: string, problems: string[]): RuntimeSurface {
  const kept = new Map(surface);
  for (const path of Object.keys(internals)) {
    if (!kept.has(path)) problems.push(`${path}: listed as a ${side} engine internal but not present`);
    kept.delete(path);
  }
  return kept;
}

/** Compares two runtime surfaces strictly, in both directions. */
export function compareRuntime(wasmSurface: RuntimeSurface, nativeSurface: RuntimeSurface, allow: Allowances): SurfaceReport {
  const problems: string[] = [];
  const wasm = stripInternals(wasmSurface, allow.internals.wasm, 'wasm', problems);
  const native = stripInternals(nativeSurface, allow.internals.native, 'native', problems);
  const stats: SurfaceStats = { classes: 0, compared: 0, equal: 0, extras: 0, pending: 0 };
  const used = new Set<string>();
  const oneSided = (path: string, kind: string, present: RuntimeSurface, absent: RuntimeSurface, extras: Record<string, Extra>, only: string, missing: string) => {
    if (!isExport(path) && !absent.has(owner(path))) {
      if (present.has(owner(path)) && extras[owner(path)]) used.add(owner(path));
      return;
    }
    const extra = extras[path];
    if (extra) {
      used.add(path);
      return;
    }
    problems.push(`${path}: ${only} (${kind}); ${missing}`);
  };
  for (const path of new Set([...wasm.keys(), ...native.keys()])) {
    const wasmKind = wasm.get(path);
    const nativeKind = native.get(path);
    if (wasmKind === undefined) {
      oneSided(path, nativeKind!, native, wasm, allow.nativeOnly, 'only on native', 'declare it in the native extras or add it to wasm');
      continue;
    }
    if (nativeKind === undefined) {
      oneSided(path, wasmKind, wasm, native, allow.wasmOnly, 'missing on native', 'add it to the native binding or declare it wasm-only');
      continue;
    }
    if (wasmKind === 'class') stats.classes += 1;
    if (!isExport(path) || wasmKind !== 'class') stats.compared += 1;
    if (wasmKind !== nativeKind) {
      problems.push(`${path}: ${wasmKind} on wasm, ${nativeKind} on native`);
      continue;
    }
    if (!isExport(path) || wasmKind !== 'class') stats.equal += 1;
  }
  for (const [side, extras, present, other] of [
    ['native-only', allow.nativeOnly, native, wasm],
    ['wasm-only', allow.wasmOnly, wasm, native],
  ] as const) {
    for (const [path, extra] of Object.entries(extras)) {
      if (!present.has(path)) problems.push(`${path}: declared ${side} but not present; remove the stale entry`);
      else if (other.has(path)) problems.push(`${path}: declared ${side} but present on both engines; remove the stale entry`);
      else if (!used.has(path)) problems.push(`${path}: declared ${side} but not reached`);
      if ('pending' in extra) stats.pending += 1;
      else stats.extras += 1;
    }
  }
  return { problems, stats };
}

/** snake_case and camelCase parameter names compare equal. */
export function normalizeParamName(name: string): string {
  return name.replace(/_+([a-z0-9])/g, (_, next: string) => next.toUpperCase());
}

function declaration(surface: Surface, path: string, kind: ExportKind | MemberKind): MemberInfo | 'class' | undefined {
  if (isExport(path)) {
    if (kind === 'class') return surface.classes.has(path) ? 'class' : undefined;
    return surface.functions.get(path);
  }
  const [className, ...rest] = path.split('.');
  const members = surface.classes.get(className);
  if (!members) return undefined;
  const isInstance = rest[0] === 'prototype';
  const name = isInstance ? rest.slice(1).join('.') : rest.join('.');
  const member = members.get(name);
  if (!member) return undefined;
  if (isInstance && (member.kind === 'method' || member.kind === 'getter')) return member;
  if (!isInstance && member.kind === 'static') return member;
  return undefined;
}

function describe(member: MemberInfo): string {
  const params = member.params.map((param) => `${param.name ?? '_'}${param.optional ? '?' : ''}: ${param.type}`).join(', ');
  return `${member.kind} (${params}): ${member.returnType}`;
}

/** Joins each runtime member to its declaration and compares the declared signatures.
 * Symbol-keyed members are compared at runtime only: the declaration parser skips them. */
export function compareDeclarations(
  wasmRuntime: RuntimeSurface,
  nativeRuntime: RuntimeSurface,
  wasmDeclared: Surface,
  nativeDeclared: Surface,
  allow: Allowances,
): string[] {
  const problems: string[] = [];
  const sides = [
    ['wasm', wasmRuntime, wasmDeclared, allow.internals.wasm],
    ['native', nativeRuntime, nativeDeclared, allow.internals.native],
  ] as const;
  for (const [side, runtime, declared, internals] of sides) {
    for (const [path, kind] of runtime) {
      if (internals[path] || internals[owner(path)] || path.includes('[') || kind === 'value') continue;
      if (declaration(declared, path, kind) === undefined) {
        problems.push(`${path}: ${kind} on ${side} at runtime but not in its declarations`);
      }
    }
    for (const [className, members] of declared.classes) {
      if (internals[className]) continue;
      for (const [name, member] of members) {
        if (member.kind === 'constructor') continue;
        const path = member.kind === 'static' ? `${className}.${name}` : `${className}.prototype.${name}`;
        if (!runtime.has(path) && !internals[path]) problems.push(`${path}: declared on ${side} but absent at runtime`);
      }
    }
    for (const name of declared.functions.keys()) {
      if (!runtime.has(name) && !internals[name]) problems.push(`${name}: declared on ${side} but absent at runtime`);
    }
  }
  const signature = (path: string, wasm: MemberInfo, native: MemberInfo) => {
    const detail = `wasm ${describe(wasm)}, native ${describe(native)}`;
    if (wasm.params.length !== native.params.length) return `${path}: arity ${wasm.params.length} on wasm, ${native.params.length} on native (${detail})`;
    for (const [index, param] of wasm.params.entries()) {
      const other = native.params[index];
      if (param.optional !== other.optional) return `${path}: parameter ${index + 1} optionality differs (${detail})`;
      if (normalizeParamName(param.name ?? '') !== normalizeParamName(other.name ?? '')) return `${path}: parameter ${index + 1} name differs (${detail})`;
      if (normalizeType(param.type) !== normalizeType(other.type)) return `${path}: parameter ${index + 1} type differs (${detail})`;
    }
    if (normalizeType(wasm.returnType) !== normalizeType(native.returnType)) return `${path}: return type differs (${detail})`;
    return undefined;
  };
  for (const [path, kind] of wasmRuntime) {
    if (!nativeRuntime.has(path) || path.includes('[') || kind === 'value') continue;
    const wasm = declaration(wasmDeclared, path, kind);
    const native = declaration(nativeDeclared, path, kind);
    if (wasm === undefined || native === undefined) continue;
    if (wasm === 'class' || native === 'class') {
      const wasmCtor = wasmDeclared.classes.get(path)?.get('constructor');
      const nativeCtor = nativeDeclared.classes.get(path)?.get('constructor');
      if (wasmCtor && !nativeCtor) problems.push(`${path}: public constructor only on wasm`);
      else if (!wasmCtor && nativeCtor) problems.push(`${path}: public constructor only on native`);
      else if (wasmCtor && nativeCtor) {
        const problem = signature(`${path}.constructor`, wasmCtor, nativeCtor);
        if (problem) problems.push(problem);
      }
      continue;
    }
    if (wasm.kind !== native.kind) {
      problems.push(`${path}: declared ${wasm.kind} on wasm, ${native.kind} on native`);
      continue;
    }
    const problem = signature(path, wasm, native);
    if (problem) problems.push(problem);
  }
  return problems;
}

/** One line per format: the figures a regression is read against. */
export function summaryLine(format: string, stats: SurfaceStats): string {
  return `${format}: classes ${stats.classes}, members compared ${stats.compared}, equal ${stats.equal}, declared extras ${stats.extras}, pending ${stats.pending}`;
}

function isNativeCode(fn: Function): boolean {
  return /\{\s*\[native code\]\s*\}\s*$/.test(Function.prototype.toString.call(fn));
}

/** Declared parameter counts of a module's top-level JavaScript functions; native functions report none. */
export function jsArity(module: object): Map<string, number> {
  const arity = new Map<string, number>();
  for (const key of Reflect.ownKeys(module)) {
    if (typeof key === 'symbol') continue;
    const value: unknown = Reflect.get(module, key);
    if (typeof value === 'function' && !isNativeCode(value)) arity.set(key, value.length);
  }
  return arity;
}

/** Compares the arity of functions that are JavaScript on both sides; `declared` lists accepted differences. */
export function compareArity(wasm: Map<string, number>, native: Map<string, number>, declared: Record<string, Extra>): string[] {
  const problems: string[] = [];
  const differing = new Set<string>();
  for (const [name, length] of wasm) {
    const other = native.get(name);
    if (other === undefined || other === length) continue;
    differing.add(name);
    if (!declared[name]) problems.push(`${name}: takes ${length} parameters on wasm, ${other} on native`);
  }
  for (const name of Object.keys(declared)) {
    if (!differing.has(name)) problems.push(`${name}: declared arity difference no longer exists; remove the stale entry`);
  }
  return problems;
}

/** Merges the declarations of several modules; a name declared twice must be declared alike. */
export function mergeDeclarations(parts: readonly [string, Surface][]): { surface: Surface; problems: string[] } {
  const surface: Surface = { classes: new Map(), functions: new Map() };
  const problems: string[] = [];
  const origin = new Map<string, string>();
  const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
  for (const [module, part] of parts) {
    for (const [name, member] of part.functions) {
      const seen = surface.functions.get(name);
      if (seen && !same(seen, member)) problems.push(`${name}: declared differently in ${origin.get(name)} and ${module}`);
      if (!seen) {
        surface.functions.set(name, member);
        origin.set(name, module);
      }
    }
    for (const [name, members] of part.classes) {
      const seen = surface.classes.get(name);
      if (seen && !same([...seen].sort(), [...members].sort())) problems.push(`${name}: class declared differently in ${origin.get(name)} and ${module}`);
      if (!seen) {
        surface.classes.set(name, members);
        origin.set(name, module);
      }
    }
  }
  return { surface, problems };
}

/** The allowances that concern one module pair: entries whose export is on either side. */
export function scopeAllowances(allow: Allowances, wasm: RuntimeSurface, native: RuntimeSurface, reached: Set<string>): Allowances {
  const keep = <T>(entries: Record<string, T>, ...surfaces: RuntimeSurface[]) => Object.fromEntries(
    Object.entries(entries).filter(([path]) => {
      const inScope = surfaces.some((surface) => surface.has(owner(path)));
      if (inScope) reached.add(path);
      return inScope;
    }),
  );
  return {
    nativeOnly: keep(allow.nativeOnly, wasm, native),
    wasmOnly: keep(allow.wasmOnly, wasm, native),
    internals: { wasm: keep(allow.internals.wasm, wasm), native: keep(allow.internals.native, native) },
  };
}

/** Every allowance must concern at least one module pair, or it is stale. */
export function unreached(allow: Allowances, reached: Set<string>): string[] {
  const all = [
    ...Object.keys(allow.nativeOnly), ...Object.keys(allow.wasmOnly),
    ...Object.keys(allow.internals.wasm), ...Object.keys(allow.internals.native),
  ];
  return all.filter((path) => !reached.has(path)).map((path) => `${path}: allowed but not exported by any engine module; remove the stale entry`);
}
