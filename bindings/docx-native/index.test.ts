import { readFileSync, readdirSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { describe, expect, spyOn, test } from 'bun:test';
import { documentFixture, exercisePackage } from './package-script';

const manifest = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'));
const wasmManifest = JSON.parse(readFileSync(new URL('../../packages/docx/package.json', import.meta.url), 'utf8'));

function distFiles(directory = new URL('./dist/', import.meta.url)): URL[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = new URL(entry.name + (entry.isDirectory() ? '/' : ''), directory);
    return entry.isDirectory() ? distFiles(path) : [path];
  });
}

describe('@betteroffice/docx-native', () => {
  test('exposes exactly the wasm package subpaths with a browser guard', () => {
    expect(Object.keys(manifest.exports)).toEqual(Object.keys(wasmManifest.exports));
    expect(Object.keys(manifest.exports)).toHaveLength(28);
    for (const [entry, target] of Object.entries(wasmManifest.exports) as [string, { types: string; import: string }][]) {
      expect(manifest.exports[entry]).toEqual({
        types: target.types,
        browser: './dist/browser.js',
        import: target.import,
      });
      expect(Object.keys(manifest.exports[entry])).toEqual(['types', 'browser', 'import']);
    }
    expect(manifest.main).toBe(wasmManifest.main);
    expect(manifest.module).toBe(wasmManifest.module);
    expect(manifest.types).toBe(wasmManifest.types);
    expect(manifest.files).toEqual(['dist', 'index.js', 'index.d.ts']);
    expect(manifest.peerDependencies).toEqual(wasmManifest.peerDependencies);
    expect(manifest.peerDependenciesMeta).toEqual(wasmManifest.peerDependenciesMeta);
    expect(readFileSync(new URL('./dist/core.mjs', import.meta.url), 'utf8').length).toBeGreaterThan(0);
    expect(readFileSync(new URL('./dist/core.d.mts', import.meta.url), 'utf8').length).toBeGreaterThan(0);
  });

  test('browser conditions reject every entry', () => {
    for (const entry of Object.keys(manifest.exports)) {
      const specifier = '@betteroffice/docx-native' + (entry === '.' ? '' : entry.slice(1));
      const result = spawnSync(process.execPath, ['--conditions=browser', '--eval', `import(${JSON.stringify(specifier)}).catch(error => { console.error(error.message); process.exit(1); })`], {
        cwd: import.meta.dir,
        encoding: 'utf8',
      });
      expect(result.status).not.toBe(0);
      expect(result.stderr).toContain('use @betteroffice/docx in browsers');
    }
  });

  test('ships no wasm assets or wasm-bindgen runtime', () => {
    for (const file of distFiles()) {
      expect(file.pathname).not.toEndWith('.wasm');
      if (!/\.(?:m?js|json)$/.test(file.pathname)) continue;
      const contents = readFileSync(file, 'utf8');
      expect(contents).not.toContain('__wbindgen');
      expect(contents).not.toContain('_bg.wasm');
      expect(contents).not.toMatch(/\.wasm\b/);
    }
  });

  test('opens, edits, saves and reopens in-thread without compiling or instantiating wasm', async () => {
    const reject = () => { throw new Error('Unexpected WebAssembly call'); };
    const spies = [
      spyOn(WebAssembly, 'compile').mockImplementation(reject),
      spyOn(WebAssembly, 'instantiate').mockImplementation(reject),
      spyOn(WebAssembly, 'compileStreaming').mockImplementation(reject),
      spyOn(WebAssembly, 'instantiateStreaming').mockImplementation(reject),
    ];
    try {
      const core = await import('@betteroffice/docx-native');
      const yrs = await import('@betteroffice/docx-native/yrs');
      const fixture = await documentFixture();
      const { results, bytes } = await exercisePackage(yrs, core, fixture);
      expect(results.before).toBe('Package text');
      expect(results.texts).toEqual(['Native Ω Package text']);
      expect(results.reopenedTexts).toEqual(results.texts);
      expect(results.blocks).toBeGreaterThan(0);
      expect(bytes.subarray(0, 2)).toEqual(new Uint8Array([80, 75]));
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
  });
});
