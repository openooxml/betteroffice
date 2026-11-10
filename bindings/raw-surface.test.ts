import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { join } from 'node:path';
import { describe, expect, test } from 'bun:test';
import { parseSurface, type Surface } from '../scripts/check-engine-surface';
import { NATIVE_ENGINES } from '../scripts/native-engine';
import {
  compareArity,
  compareDeclarations,
  compareRuntime,
  discover,
  jsArity,
  mergeDeclarations,
  mergeSurfaces,
  scopeAllowances,
  summaryLine,
  unreached,
  type RuntimeSurface,
  type SurfaceStats,
} from '../scripts/raw-engine-surface';
import { arityAllowances, engineAllowances, rawAllowances, typeAllowances } from './native-extras';

const root = join(import.meta.dir, '..');
const require = createRequire(import.meta.url);

/** The wasm-bindgen modules each native binding replaces. */
const GLUE = {
  xlsx: ['packages/xlsx/src/wasm/generated/xlsx_wasm'],
  pptx: ['packages/pptx/src/wasm/generated/pptx_wasm', 'packages/pptx/src/wasm/generated/metafile_svg_wasm'],
  docx: ['edit/docx_edit', 'parse/docx_parse', 'query/docx_query', 'opc/ooxml_opc', 'preflight/docx_preflight']
    .map((module) => `packages/docx/src/wasm/generated/${module}`),
} as const;

type Format = keyof typeof GLUE;

/** Each binding as loaded, taken before any engine module runs: engine modules install members on the binding classes. */
const RAW_NATIVE = Object.fromEntries((Object.keys(GLUE) as Format[])
  .map((format) => [format, discover(require(join(root, `bindings/${format}-native/index.js`)))])) as Record<Format, RuntimeSurface>;

function declared(paths: readonly string[]): { surface: Surface; problems: string[] } {
  return mergeDeclarations(paths.map((path) => [path, parseSurface(readFileSync(join(root, path), 'utf8'), { paramNames: true })]));
}

async function surfaceOf(module: string): Promise<RuntimeSurface> {
  return discover(await import(join(root, module)));
}

describe('native and wasm engine surfaces', () => {
  for (const format of Object.keys(GLUE) as Format[]) {
    test(`${format}: raw binding matches the wasm-bindgen glue`, async () => {
      const parts: [string, RuntimeSurface][] = [];
      for (const module of GLUE[format]) parts.push([module, await surfaceOf(`${module}.js`)]);
      const { surface: wasm, problems: runtimeMerge } = mergeSurfaces(parts);
      const native = RAW_NATIVE[format];
      const allow = rawAllowances(format);
      const report = compareRuntime(wasm, native, allow);
      const wasmDeclared = declared(GLUE[format].map((module) => `${module}.d.ts`));
      const nativeDeclared = declared([`bindings/${format}-native/index.d.ts`]);
      const declarations = compareDeclarations(wasm, native, wasmDeclared.surface, nativeDeclared.surface, allow);
      console.log(`raw ${summaryLine(format, report.stats)}`);
      expect([...runtimeMerge, ...wasmDeclared.problems, ...report.problems, ...declarations]).toEqual([]);
    });

    test(`${format}: each native engine module matches its wasm engine module`, async () => {
      const allow = engineAllowances(format);
      const arityAllow = arityAllowances(format);
      const reached = new Set<string>();
      const totals: SurfaceStats = { classes: 0, compared: 0, equal: 0, extras: 0, pending: 0 };
      const problems: string[] = [];
      for (const { engine, native: shim } of NATIVE_ENGINES[format] as readonly { engine: string; native: string }[]) {
        const wasm = await surfaceOf(engine);
        const native = await surfaceOf(shim);
        const report = compareRuntime(wasm, native, scopeAllowances(allow, wasm, native, reached));
        for (const key of Object.keys(totals) as (keyof SurfaceStats)[]) totals[key] += report.stats[key];
        const arity = compareArity(jsArity(await import(join(root, engine))), jsArity(await import(join(root, shim))), arityAllow[engine] ?? {});
        totals.pending += Object.keys(arityAllow[engine] ?? {}).length;
        totals.pending += Object.keys(typeAllowances[engine] ?? {}).length;
        problems.push(...[...report.problems, ...arity].map((problem) => `${engine}: ${problem}`));
      }
      const pairs = new Set((NATIVE_ENGINES[format] as readonly { engine: string }[]).map(({ engine }) => engine));
      for (const engine of Object.keys(arityAllow)) {
        if (!pairs.has(engine)) problems.push(`${engine}: arity allowance for a module that is not an engine module`);
      }
      console.log(`engine ${summaryLine(format, totals)}`);
      expect([...problems, ...unreached(allow, reached)]).toEqual([]);
    });
  }

  test('each native engine module declares the types of its wasm engine module', () => {
    const typecheck = readFileSync(join(import.meta.dir, 'engine-surface.typecheck.ts'), 'utf8');
    const checked = [...typecheck.matchAll(/from '\.\.\/(packages\/[^']+)';/g)].map((match) => `${match[1]}.ts`).sort();
    const engines = Object.values(NATIVE_ENGINES).flatMap((pairs) => (pairs as readonly { engine: string }[]).map(({ engine }) => engine)).sort();
    expect(checked).toEqual(engines);
    const pending = Object.fromEntries([...typecheck.matchAll(/type (\w+)Pending = ([^;]+);/g)]
      .map((match) => [match[1], match[2].split('|').map((name) => name.trim().replace(/'/g, '')).sort()]));
    const expected = Object.fromEntries(Object.entries(typeAllowances).map(([engine, names]) => {
      const [, format, module] = /^packages\/(\w+)\/src\/wasm\/(\w+)\.ts$/.exec(engine)!;
      return [`${format[0].toUpperCase()}${format.slice(1)}${module[0].toUpperCase()}${module.slice(1)}`, Object.keys(names).sort()];
    }));
    expect(pending).toEqual(expected);
    const result = spawnSync('bunx', ['tsc', '-p', join(import.meta.dir, 'engine-surface.tsconfig.json')], { encoding: 'utf8', cwd: root });
    expect(`${result.stdout}${result.stderr}`.trim()).toBe('');
    expect(result.status).toBe(0);
  }, 180_000);
});
