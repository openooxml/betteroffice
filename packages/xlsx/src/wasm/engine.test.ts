import { describe, expect, test } from 'bun:test';
import { readEnvOverride, selectBackend, type BackendRuntime } from './engine';

const server: BackendRuntime = {
  hasProcess: true,
  hasGetBuiltinModule: true,
  hasNodeVersion: true,
  hasDom: false,
};
const pkg = '0.5.1';

function nativeModule(version: string = pkg) {
  return { XlsxDocument: class {}, engineVersion: () => version };
}

describe('selectBackend', () => {
  test('wasm forced never loads the native engine', () => {
    let loads = 0;
    const backend = selectBackend({
      env: 'wasm',
      runtime: server,
      loadNative: () => {
        loads += 1;
        return nativeModule();
      },
      packageVersion: pkg,
    });
    expect(backend).toBe('wasm');
    expect(loads).toBe(0);
  });

  test('native forced throws with the load error when the addon is missing', () => {
    expect(() =>
      selectBackend({
        env: 'native',
        runtime: server,
        loadNative: () => {
          throw new Error('Cannot find module @betteroffice/xlsx-native');
        },
        packageVersion: pkg,
      })
    ).toThrow('BETTEROFFICE_ENGINE=native but @betteroffice/xlsx-native failed to load: Cannot find module');
  });

  test('native forced throws off server runtimes', () => {
    expect(() =>
      selectBackend({
        env: 'native',
        runtime: { ...server, hasDom: true },
        loadNative: () => nativeModule(),
        packageVersion: pkg,
      })
    ).toThrow('BETTEROFFICE_ENGINE=native requires a server-side JS runtime');
  });

  test('version mismatch falls back to wasm when automatic', () => {
    expect(
      selectBackend({
        env: undefined,
        runtime: server,
        loadNative: () => nativeModule('0.0.1'),
        packageVersion: pkg,
      })
    ).toBe('wasm');
  });

  test('version mismatch throws with both versions when native is forced', () => {
    expect(() =>
      selectBackend({
        env: 'native',
        runtime: server,
        loadNative: () => nativeModule('0.0.1'),
        packageVersion: pkg,
      })
    ).toThrow(
      'BETTEROFFICE_ENGINE=native but @betteroffice/xlsx-native version 0.0.1 does not match @betteroffice/xlsx version 0.5.1'
    );
  });

  test('a DOM keeps the wasm engine', () => {
    let loads = 0;
    expect(
      selectBackend({
        env: undefined,
        runtime: { ...server, hasDom: true },
        loadNative: () => {
          loads += 1;
          return nativeModule();
        },
        packageVersion: pkg,
      })
    ).toBe('wasm');
    expect(loads).toBe(0);
  });

  test('no getBuiltinModule keeps the wasm engine', () => {
    expect(
      selectBackend({
        env: undefined,
        runtime: { ...server, hasGetBuiltinModule: false },
        loadNative: () => nativeModule(),
        packageVersion: pkg,
      })
    ).toBe('wasm');
  });

  test('a missing XlsxDocument export keeps the wasm engine', () => {
    expect(
      selectBackend({
        env: undefined,
        runtime: server,
        loadNative: () => ({ engineVersion: () => pkg }),
        packageVersion: pkg,
      })
    ).toBe('wasm');
  });

  test('a load failure keeps the wasm engine when automatic', () => {
    expect(
      selectBackend({
        env: undefined,
        runtime: server,
        loadNative: () => {
          throw new Error('not installed');
        },
        packageVersion: pkg,
      })
    ).toBe('wasm');
  });

  test('an unknown override value behaves like automatic', () => {
    expect(
      selectBackend({
        env: 'bogus',
        runtime: server,
        loadNative: () => nativeModule(),
        packageVersion: pkg,
      })
    ).toBe('native');
  });

  test('happy path selects the native engine', () => {
    expect(
      selectBackend({
        env: undefined,
        runtime: server,
        loadNative: () => nativeModule(),
        packageVersion: pkg,
      })
    ).toBe('native');
    expect(
      selectBackend({
        env: 'native',
        runtime: server,
        loadNative: () => nativeModule(),
        packageVersion: pkg,
      })
    ).toBe('native');
  });
});

describe('readEnvOverride', () => {
  test('reads BETTEROFFICE_ENGINE when the environment is readable', () => {
    expect(readEnvOverride({ env: { BETTEROFFICE_ENGINE: 'native' } })).toBe('native');
    expect(readEnvOverride({ env: {} })).toBeUndefined();
    expect(readEnvOverride({})).toBeUndefined();
  });

  test('treats a denied environment as no override', () => {
    const denied = {
      get env(): Record<string, string | undefined> {
        throw new Error('NotCapable: Requires env access to "BETTEROFFICE_ENGINE"');
      },
    };
    expect(readEnvOverride(denied)).toBeUndefined();
    const deniedKey = { env: new Proxy({}, { get() { throw new Error('NotCapable'); } }) };
    expect(readEnvOverride(deniedKey)).toBeUndefined();
  });
});
