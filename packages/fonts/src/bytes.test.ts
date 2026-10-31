import { expect, spyOn, test } from 'bun:test';
import { readFontBytes } from './bytes';

test('package assets on a Node.js without getBuiltinModule fail loudly instead of fetching', async () => {
  const host = globalThis.process as unknown as { getBuiltinModule?: unknown };
  const original = host.getBuiltinModule;
  const fetch = spyOn(globalThis, 'fetch');
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  host.getBuiltinModule = undefined;
  try {
    const url = new URL('../assets/manifest.json', import.meta.url);
    await expect(readFontBytes('manifest.json', url)).rejects.toThrow('needs Node.js 20.16 or later');
    await expect(readFontBytes('manifest.json', url)).rejects.toThrow('needs Node.js 20.16 or later');
    expect(fetch).not.toHaveBeenCalled();
    expect(warn).toHaveBeenCalledTimes(1);
  } finally {
    host.getBuiltinModule = original;
    fetch.mockRestore();
    warn.mockRestore();
  }
});

test('package assets read from disk when getBuiltinModule exists', async () => {
  const fetch = spyOn(globalThis, 'fetch');
  try {
    const bytes = await readFontBytes('index.ts', new URL('./index.ts', import.meta.url));
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(fetch).not.toHaveBeenCalled();
  } finally {
    fetch.mockRestore();
  }
});
