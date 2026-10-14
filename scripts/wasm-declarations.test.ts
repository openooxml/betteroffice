import { describe, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { sortWasmExportDeclarations } from './wasm.ts';

const ROOT = resolve(import.meta.dir, '..');

const DECLARATIONS = [
  'export function open(bytes: Uint8Array): number;',
  '',
  'export interface InitOutput {',
  '    readonly memory: WebAssembly.Memory;',
  '    readonly zeta: (a: number) => void;',
  '    readonly alpha: () => number;',
  '    readonly __wbindgen_malloc: (a: number, b: number) => number;',
  '}',
  '',
].join('\n');

const BG = [
  '/* tslint:disable */',
  '/* eslint-disable */',
  'export const memory: WebAssembly.Memory;',
  'export const zeta: (a: number) => void;',
  'export const alpha: () => number;',
  '',
].join('\n');

function generatedDeclarations(dir: string): string[] {
  return readdirSync(dir, { recursive: true, encoding: 'utf8' })
    .filter((name) => name.endsWith('.d.ts'))
    .map((name) => join(dir, name));
}

describe('sortWasmExportDeclarations', () => {
  test('sorts InitOutput members by name and leaves the rest', () => {
    expect(sortWasmExportDeclarations(DECLARATIONS)).toBe(
      [
        'export function open(bytes: Uint8Array): number;',
        '',
        'export interface InitOutput {',
        '    readonly __wbindgen_malloc: (a: number, b: number) => number;',
        '    readonly alpha: () => number;',
        '    readonly memory: WebAssembly.Memory;',
        '    readonly zeta: (a: number) => void;',
        '}',
        '',
      ].join('\n')
    );
  });

  test('sorts the consts of a _bg.wasm.d.ts by name', () => {
    expect(sortWasmExportDeclarations(BG)).toBe(
      [
        '/* tslint:disable */',
        '/* eslint-disable */',
        'export const alpha: () => number;',
        'export const memory: WebAssembly.Memory;',
        'export const zeta: (a: number) => void;',
        '',
      ].join('\n')
    );
  });

  test('gives the same file for any export order', () => {
    const reordered = BG.replace(
      'export const zeta: (a: number) => void;\nexport const alpha: () => number;',
      'export const alpha: () => number;\nexport const zeta: (a: number) => void;'
    );
    expect(sortWasmExportDeclarations(reordered)).toBe(sortWasmExportDeclarations(BG));
  });

  test('rejects a layout it does not recognize', () => {
    expect(() => sortWasmExportDeclarations('export interface InitOutput {\n}\n')).toThrow();
    expect(() => sortWasmExportDeclarations(`${BG}// trailing\nexport const late: () => void;\n`)).toThrow();
  });

  test('the vendored declarations are in sorted order', () => {
    const files = ['docx', 'xlsx', 'pptx', 'vsdx'].flatMap((pkg) =>
      generatedDeclarations(resolve(ROOT, 'packages', pkg, 'src/wasm/generated'))
    );
    expect(files.length).toBe(14);
    for (const file of files) {
      const text = readFileSync(file, 'utf8');
      expect({ file, text: sortWasmExportDeclarations(text) }).toEqual({ file, text });
    }
  });
});
