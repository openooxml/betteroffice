import { expect, test } from 'bun:test';
import { resolve } from 'node:path';

const ROOT = resolve(import.meta.dir, '..');
const WASM = 'wasm32-unknown-unknown';

const ARTIFACTS: Array<[name: string, dir: string, target: string | undefined, args: string[]]> = [
  ['docx edit', 'crates/docx-edit', WASM, ['--features', 'wasm,tiff,metafile,heap-stats']],
  ['docx parse', 'crates/docx-parse', WASM, ['--features', 'wasm-layout,tiff,metafile']],
  ['docx preflight', 'crates/ooxml-font-preflight', WASM, ['--no-default-features', '--features', 'wasm,docx']],
  ['pptx preflight', 'crates/ooxml-font-preflight', WASM, ['--no-default-features', '--features', 'wasm,pptx']],
  ['xlsx preflight', 'crates/ooxml-font-preflight', WASM, ['--no-default-features', '--features', 'wasm,xlsx']],
  ['docx query', 'crates/docx-layout', WASM, ['--features', 'wasm-query']],
  ['pptx', 'crates/pptx-wasm', WASM, []],
  ['pptx lean', 'crates/pptx-wasm', WASM, ['--no-default-features']],
  ['xlsx', 'crates/xlsx-wasm', WASM, []],
  ['xlsx lean', 'crates/xlsx-wasm', WASM, ['--no-default-features']],
  ['vsdx', 'crates/vsdx-wasm', WASM, []],
  ['docx native', 'bindings/docx-native', undefined, []],
  ['pptx native', 'bindings/pptx-native', undefined, []],
  ['xlsx native', 'bindings/xlsx-native', undefined, []],
  ['docx python', 'bindings/python-docx', undefined, []],
  ['pptx python', 'bindings/python-pptx', undefined, []],
  ['vsdx python', 'bindings/python-vsdx', undefined, []],
  ['xlsx python', 'bindings/python-xlsx', undefined, []],
];

for (const [name, dir, target, args] of ARTIFACTS) {
  test(`${name} reads JSON numbers back exactly`, () => {
    const run = Bun.spawnSync([
      'cargo', 'tree', '--manifest-path', `${dir}/Cargo.toml`, ...(target ? ['--target', target] : []),
      '-e', 'normal', '--prefix', 'none', '-f', '{p} [{f}]', ...args,
    ], { cwd: ROOT });
    expect(run.exitCode, run.stderr.toString()).toBe(0);
    const lines = run.stdout.toString().split('\n').filter((line) => line.startsWith('serde_json v'));
    expect(lines.length).toBeGreaterThan(0);
    for (const line of lines) expect(line).toContain('float_roundtrip');
  }, 120_000);
}
