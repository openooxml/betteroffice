import { expect, test } from 'bun:test';
import { existsSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { LICENSE, packedPaths, tarballFolders } from './check-packed-license.mjs';
import { NATIVE_PACKAGES } from './native-bindings.mjs';

const root = fileURLToPath(new URL('..', import.meta.url));
const folders = NATIVE_PACKAGES.flatMap((path) => tarballFolders(join(root, path)));

test('a native package is its root folder and one folder per platform', () => {
  const docx = tarballFolders(join(root, 'bindings/docx-native')).map((folder) => relative(root, folder));
  expect(docx[0]).toBe('bindings/docx-native');
  expect(docx.slice(1)).toEqual(
    [
      'darwin-arm64',
      'darwin-x64',
      'linux-arm64-gnu',
      'linux-arm64-musl',
      'linux-x64-gnu',
      'linux-x64-musl',
      'win32-arm64-msvc',
      'win32-x64-msvc',
    ].map((platform) => `bindings/docx-native/npm/${platform}`)
  );
  expect(folders).toHaveLength(27);
  expect(tarballFolders(join(root, 'packages/docx'))).toEqual([join(root, 'packages/docx')]);
});

test('the packed paths are read from either shape npm prints', () => {
  const files = [{ path: 'LICENSE' }, { path: 'package.json' }];
  expect(packedPaths(JSON.stringify([{ name: 'a', files }]))).toEqual(['LICENSE', 'package.json']);
  expect(packedPaths(JSON.stringify({ name: 'a', files }))).toEqual(['LICENSE', 'package.json']);
  expect(() => packedPaths('[]')).toThrow('npm pack printed no file list');
});

test('every native package folder holds the project license, as the wasm packages do', () => {
  const license = readFileSync(join(root, LICENSE), 'utf8');
  for (const format of ['docx', 'xlsx', 'pptx']) {
    expect(readFileSync(join(root, 'packages', format, LICENSE), 'utf8')).toBe(license);
  }
  const without = folders.filter((folder) => !existsSync(join(folder, LICENSE))).map((folder) => relative(root, folder));
  expect(without).toEqual([]);
  for (const folder of folders) {
    expect(readFileSync(join(folder, LICENSE), 'utf8')).toBe(license);
  }
});

test('the workflows that pack the native packages check every tarball for the license', () => {
  const read = (name: string) => readFileSync(`${root}.github/workflows/${name}`, 'utf8');
  expect(read('ci.yml')).toContain('node scripts/check-packed-license.mjs $(node scripts/native-bindings.mjs --paths)\n');
  expect(read('native-packages.yml')).toContain('node scripts/check-packed-license.mjs bindings/${{ matrix.binding }}-native\n');
  expect(read('publish-native.yml')).toContain('node scripts/check-packed-license.mjs bindings/${{ inputs.binding }}-native\n');
});
