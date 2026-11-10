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
  mergeSurfaces,
  summaryLine,
  type RuntimeSurface,
} from '../scripts/raw-engine-surface';
import { arityAllowances, engineAllowances, rawAllowances } from './native-extras';

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

function declared(paths: readonly string[]): Surface {
  const merged: Surface = { classes: new Map(), functions: new Map() };
  for (const path of paths) {
    const surface = parseSurface(readFileSync(join(root, path), 'utf8'), { paramNames: true });
    for (const [name, members] of surface.classes) merged.classes.set(name, members);
    for (const [name, member] of surface.functions) merged.functions.set(name, member);
  }
  return merged;
}

async function merged(modules: readonly string[]): Promise<RuntimeSurface> {
  const parts: [string, RuntimeSurface][] = [];
  for (const module of modules) parts.push([module, discover(await import(join(root, module)))]);
  const { surface, problems } = mergeSurfaces(parts);
  expect(problems).toEqual([]);
  return surface;
}

describe('native and wasm engine surfaces', () => {
  for (const format of Object.keys(GLUE) as Format[]) {
    test(`${format}: raw binding matches the wasm-bindgen glue`, async () => {
      const wasm = await merged(GLUE[format].map((module) => `${module}.js`));
      const native = discover(require(join(root, `bindings/${format}-native/index.js`)));
      const allow = rawAllowances(format);
      const report = compareRuntime(wasm, native, allow);
      const declarations = compareDeclarations(
        wasm,
        native,
        declared(GLUE[format].map((module) => `${module}.d.ts`)),
        declared([`bindings/${format}-native/index.d.ts`]),
        allow,
      );
      console.log(`raw ${summaryLine(format, report.stats)}`);
      expect([...report.problems, ...declarations]).toEqual([]);
    });

    test(`${format}: native engine modules match the wasm engine modules`, async () => {
      const pairs = NATIVE_ENGINES[format] as readonly { engine: string; native: string }[];
      const wasm = await merged(pairs.map(({ engine }) => engine));
      const native = await merged(pairs.map(({ native }) => native));
      const report = compareRuntime(wasm, native, engineAllowances(format));
      const arity = new Map<string, number>();
      const nativeArity = new Map<string, number>();
      for (const { engine, native: shim } of pairs) {
        for (const [name, length] of jsArity(await import(join(root, engine)))) arity.set(name, length);
        for (const [name, length] of jsArity(await import(join(root, shim)))) nativeArity.set(name, length);
      }
      const arityPending = Object.keys(arityAllowances(format)).length;
      console.log(`engine ${summaryLine(format, { ...report.stats, pending: report.stats.pending + arityPending })}`);
      expect([...report.problems, ...compareArity(arity, nativeArity, arityAllowances(format))]).toEqual([]);
    });
  }
});
