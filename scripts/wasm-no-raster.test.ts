import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ROOT = resolve(import.meta.dir, '..');
const WASM = 'wasm32-unknown-unknown';
const RASTER = /^(tiny-skia(-path)?|\S*raster\S*) v/;

const INITIAL_RASTER_ARTIFACTS = ['vsdx-wasm', 'xlsx-wasm'];
const STILL_RASTER: string[] = [];

const BUILD_SCRIPTS = ['build-docx-wasm', 'build-pptx-wasm', 'build-vsdx-wasm', 'build-xlsx-wasm'];

function moduleArtifacts(script: string): string[][] {
  const source = readFileSync(resolve(ROOT, 'scripts', `${script}.ts`), 'utf8');
  return [...source.matchAll(/\{\s*crate: '([^']+)',[^}]*?\}/gs)].map(([block, crate]) => {
    const args = block.match(/cargoArgs: \[([^\]]*)\]/)?.[1] ?? '';
    return [crate, ...[...args.matchAll(/'([^']+)'/g)].map((arg) => arg[1]).filter((arg) => arg !== '--locked')];
  });
}

const keyOf = ([crate, ...args]: string[]) => [crate, ...args].join(' ');
const ARTIFACTS = [...new Map([
  ...BUILD_SCRIPTS.flatMap(moduleArtifacts),
  ['betteroffice-fonts', '--no-default-features', '--features', 'wasm'],
  ['ooxml-opc', '--features', 'wasm'],
].map((artifact) => [keyOf(artifact), artifact])).values()];

function rasterPackages([crate, ...args]: string[]): string[] {
  const run = Bun.spawnSync([
    'cargo', 'tree', '--locked', '--manifest-path', `crates/${crate}/Cargo.toml`, '--target', WASM, '-e', 'normal',
    '--prefix', 'none', '-f', '{p}', ...args,
  ], { cwd: ROOT });
  expect(run.exitCode, run.stderr.toString()).toBe(0);
  return [...new Set(run.stdout.toString().split('\n').filter((line) => RASTER.test(line)))];
}

describe('wasm artifacts carry no rasterizer', () => {
  test('the allow-list only shrinks', () => {
    for (const key of STILL_RASTER) expect(INITIAL_RASTER_ARTIFACTS).toContain(key);
  });

  test('every build script yields artifacts', () => {
    for (const script of BUILD_SCRIPTS) expect(moduleArtifacts(script).length).toBeGreaterThan(0);
  });

  test('every allow-listed artifact is a built artifact', () => {
    const built = ARTIFACTS.map(keyOf);
    for (const key of STILL_RASTER) expect(built).toContain(key);
  });

  for (const artifact of ARTIFACTS) {
    const key = keyOf(artifact);
    test(key, () => {
      const found = rasterPackages(artifact);
      if (STILL_RASTER.includes(key)) expect(found.length, `${key} no longer has a rasterizer: delete it from STILL_RASTER`).toBeGreaterThan(0);
      else expect(found).toEqual([]);
    }, 120_000);
  }
});
