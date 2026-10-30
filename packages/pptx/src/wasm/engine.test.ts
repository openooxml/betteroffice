import { describe, expect, test } from 'bun:test';
import { createRequire } from 'node:module';
import { readEnvOverride, selectBackend, type BackendRuntime } from './engine';

const server: BackendRuntime = {
  hasProcess: true,
  hasGetBuiltinModule: true,
  hasNodeVersion: true,
  hasDom: false,
};
const pkg = '0.5.1';

function nativeModule(version: string = pkg) {
  return { PptxDocument: class {}, PptxRenderer: class {}, engineVersion: () => version };
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
          throw Object.assign(new Error('Cannot find module @betteroffice/pptx-native'), {
            code: 'MODULE_NOT_FOUND',
          });
        },
        packageVersion: pkg,
      })
    ).toThrow('BETTEROFFICE_ENGINE=native but @betteroffice/pptx-native failed to load: Cannot find module');
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
      'BETTEROFFICE_ENGINE=native but @betteroffice/pptx-native version 0.0.1 does not match @betteroffice/pptx version 0.5.1'
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

  test('a missing PptxDocument export keeps the wasm engine', () => {
    expect(
      selectBackend({
        env: undefined,
        runtime: server,
        loadNative: () => ({ engineVersion: () => pkg }),
        packageVersion: pkg,
      })
    ).toBe('wasm');
  });

  test('fallback works when the native package is missing (MODULE_NOT_FOUND)', () => {
    expect(
      selectBackend({
        env: undefined,
        runtime: server,
        loadNative: () => {
          throw Object.assign(new Error('Cannot find module @betteroffice/pptx-native'), {
            code: 'MODULE_NOT_FOUND',
          });
        },
        packageVersion: pkg,
      })
    ).toBe('wasm');
  });

  test('the real selector falls back to wasm when the native package cannot resolve', () => {
    const requireFromHere = createRequire(import.meta.url);
    let loads = 0;
    let failure: unknown;
    expect(
      selectBackend({
        env: undefined,
        runtime: server,
        loadNative: () => {
          loads += 1;
          try {
            return requireFromHere('@betteroffice/pptx-native-missing-for-test');
          } catch (error) {
            failure = error;
            throw error;
          }
        },
        packageVersion: pkg,
      })
    ).toBe('wasm');
    expect(loads).toBe(1);
    expect(failure).toMatchObject({ code: 'MODULE_NOT_FOUND' });
  });

  test('native forced throws when PptxDocument is missing', () => {
    expect(() =>
      selectBackend({
        env: 'native',
        runtime: server,
        loadNative: () => ({ engineVersion: () => pkg }),
        packageVersion: pkg,
      })
    ).toThrow('BETTEROFFICE_ENGINE=native but @betteroffice/pptx-native does not export PptxDocument');
  });

  test('a missing PptxRenderer export keeps the wasm engine', () => {
    expect(
      selectBackend({
        env: undefined,
        runtime: server,
        loadNative: () => ({ PptxDocument: class {}, engineVersion: () => pkg }),
        packageVersion: pkg,
      })
    ).toBe('wasm');
  });

  test('native forced throws when PptxRenderer is missing', () => {
    expect(() =>
      selectBackend({
        env: 'native',
        runtime: server,
        loadNative: () => ({ PptxDocument: class {}, engineVersion: () => pkg }),
        packageVersion: pkg,
      })
    ).toThrow('BETTEROFFICE_ENGINE=native but @betteroffice/pptx-native does not export PptxRenderer;');
  });

  for (const [name, loadNative] of [
    ['missing engineVersion', () => ({ PptxDocument: class {}, PptxRenderer: class {} })],
    ['throwing engineVersion', () => ({
      PptxDocument: class {},
      PptxRenderer: class {},
      engineVersion: () => { throw new Error('version failed'); },
    })],
  ] as const) {
    test(`${name} falls back to wasm when automatic`, () => {
      expect(selectBackend({ env: undefined, runtime: server, loadNative, packageVersion: pkg }))
        .toBe('wasm');
    });

    test(`${name} throws when native is forced`, () => {
      expect(() => selectBackend({ env: 'native', runtime: server, loadNative, packageVersion: pkg }))
        .toThrow('BETTEROFFICE_ENGINE=native but @betteroffice/pptx-native version unknown does not match @betteroffice/pptx version 0.5.1');
    });
  }

  for (const [name, runtime] of [
    ['no process', { ...server, hasProcess: false }],
    ['no getBuiltinModule', { ...server, hasGetBuiltinModule: false }],
    ['no node version', { ...server, hasNodeVersion: false }],
  ] as const) {
    test(`${name} falls back without loading native`, () => {
      let loads = 0;
      expect(
        selectBackend({
          env: undefined,
          runtime,
          loadNative: () => {
            loads += 1;
            return nativeModule();
          },
          packageVersion: pkg,
        })
      ).toBe('wasm');
      expect(loads).toBe(0);
    });

    test(`${name} throws when native is forced`, () => {
      expect(() =>
        selectBackend({
          env: 'native',
          runtime,
          loadNative: () => nativeModule(),
          packageVersion: pkg,
        })
      ).toThrow('BETTEROFFICE_ENGINE=native requires a server-side JS runtime');
    });
  }

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
