import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, spyOn, test } from 'bun:test';
import { StrictMode, useEffect } from 'react';

import {
  configureDefaultFonts,
  type ResidentFontRequirement,
  type ResidentMeasurementConfig,
  type RustTextEngine,
} from '@betteroffice/docx/layout';
import { useRustMeasurement } from './useRustMeasurement';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook, waitFor } = await import('@testing-library/react');

function bytesOf(text: string): ArrayBuffer {
  return new TextEncoder().encode(text).buffer as ArrayBuffer;
}

afterEach(() => {
  cleanup();
  configureDefaultFonts({});
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

describe('useRustMeasurement warm requirements', () => {
  const regular: ResidentFontRequirement = {
    key: 'regular',
    family: 'Calibri',
    bold: false,
    italic: false,
  };
  const warm: ResidentFontRequirement = {
    key: 'warm',
    family: 'Warm',
    bold: false,
    italic: false,
  };

  const later: ResidentFontRequirement = {
    key: 'later',
    family: 'Later',
    bold: false,
    italic: false,
  };

  async function prepared(
    loadWarm: () => Promise<ArrayBuffer>,
    loadLater: () => Promise<ArrayBuffer> = () => new Promise<ArrayBuffer>(() => {})
  ) {
    const registered: string[][] = [];
    const calls = { warmLoads: 0, warmPasses: 0, requiredPasses: 0 };
    const engineWith = (): RustTextEngine => {
      const fonts: string[] = [];
      registered.push(fonts);
      return {
        registerFont: (bytes) => fonts.push(new TextDecoder().decode(bytes)),
        clearFonts() {},
      };
    };
    const fontProvider = {
      resolve: (family: string) => () => {
        if (family === 'Later') return loadLater();
        if (family !== 'Warm') return Promise.resolve(bytesOf(family));
        calls.warmLoads++;
        return loadWarm();
      },
    };
    const hook = renderHook(
      ({ engine }: { engine: RustTextEngine }) =>
        useRustMeasurement({ document: null, textEngine: engine, fontProvider }),
      { initialProps: { engine: engineWith() } }
    );
    await waitFor(() =>
      expect(hook.result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
        regular: [1],
      })
    );
    hook.result.current.scheduleWarmLayoutRef.current = () => {
      calls.warmPasses++;
    };
    hook.result.current.runLayoutPipelineRef.current = () => {
      calls.requiredPasses++;
    };
    return { ...hook, calls, registered, engineWith };
  }

  function deferred() {
    let finishLoad!: (bytes: ArrayBuffer) => void;
    const pending = new Promise<ArrayBuffer>((resolve) => {
      finishLoad = resolve;
    });
    return { pending, finishLoad };
  }

  test('pending warm fonts leave the current config ready and request one pass when ready', async () => {
    let finishLoad!: (bytes: ArrayBuffer) => void;
    const pending = new Promise<ArrayBuffer>((resolve) => {
      finishLoad = resolve;
    });
    const { result, calls } = await prepared(() => pending);
    const plain = result.current.residentMeasurementConfig([regular]);

    result.current.warmFontRequirements([warm]);
    result.current.warmFontRequirements([warm]);
    expect(result.current.residentMeasurementConfig([regular])).toEqual(plain);
    await waitFor(() => expect(calls.warmLoads).toBe(1));
    expect(calls.warmPasses).toBe(0);

    finishLoad(bytesOf('warm'));
    await waitFor(() => expect(calls.warmPasses).toBe(1));
    expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
      regular: [1],
      warm: [2],
    });
    result.current.warmFontRequirements([warm]);
    expect(calls).toEqual({ warmLoads: 1, warmPasses: 1, requiredPasses: 0 });
  });

  test('a warm font that never settles does not hold back another that loads', async () => {
    const { result, calls } = await prepared(() => Promise.resolve(bytesOf('warm')));
    result.current.warmFontRequirements([later, warm]);
    await waitFor(() => expect(calls.warmPasses).toBe(1));
    expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
      regular: [1],
      warm: [2],
    });
  });

  test('a warm font that never settles never blocks plain requirements', async () => {
    const { result, calls } = await prepared(() => new Promise<ArrayBuffer>(() => {}));
    result.current.warmFontRequirements([warm]);
    expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
      regular: [1],
    });
    await waitFor(() => expect(calls.warmLoads).toBe(1));
    expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
      regular: [1],
    });
    expect(result.current.deferLayoutPass()).toBe(false);
    expect(calls.warmPasses).toBe(0);
  });

  test('a rejected warm font requests one pass and is never retried or kept', async () => {
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const { result, calls } = await prepared(() => Promise.reject(new Error('warm load failed')));
      result.current.warmFontRequirements([warm]);
      expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
        regular: [1],
      });
      await waitFor(() => expect(calls.warmPasses).toBe(1));
      for (let pass = 0; pass < 3; pass++) {
        result.current.warmFontRequirements([warm]);
        expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
          regular: [1],
        });
      }
      await new Promise((resolve) => setTimeout(resolve, 10));
      expect(calls).toEqual({ warmLoads: 1, warmPasses: 1, requiredPasses: 0 });
    } finally {
      warn.mockRestore();
    }
  });

  test('a warm font that loads after unmount registers nothing and requests no pass', async () => {
    const { pending, finishLoad } = deferred();
    const { result, unmount, calls, registered } = await prepared(() => pending);
    result.current.warmFontRequirements([warm]);
    await waitFor(() => expect(calls.warmLoads).toBe(1));

    unmount();
    finishLoad(bytesOf('Warm'));
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(registered).toEqual([['Calibri']]);
    expect(calls).toEqual({ warmLoads: 1, warmPasses: 0, requiredPasses: 0 });
  });

  test('a warm font the replaced source was loading registers nothing and requests no pass', async () => {
    const { pending, finishLoad } = deferred();
    const { result, rerender, calls, registered, engineWith } = await prepared(() => pending);
    result.current.warmFontRequirements([warm]);
    await waitFor(() => expect(calls.warmLoads).toBe(1));

    rerender({ engine: engineWith() });
    await waitFor(() =>
      expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
        regular: [1],
      })
    );
    // The new source's first load asks for one pass of its own.
    expect(calls.requiredPasses).toBe(1);
    finishLoad(bytesOf('Warm'));
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(registered).toEqual([['Calibri'], ['Calibri']]);
    expect(calls).toEqual({ warmLoads: 1, warmPasses: 0, requiredPasses: 1 });
    expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
      regular: [1],
    });
  });

  test('already-ready warm requirements add no keys and request no pass', async () => {
    const { result, calls } = await prepared(() => Promise.resolve(bytesOf('warm')));
    const readyWarm = { ...warm, family: regular.family };
    const plain = result.current.residentMeasurementConfig([regular]);
    result.current.warmFontRequirements([readyWarm]);
    expect(result.current.residentMeasurementConfig([regular])).toEqual(plain);
    expect(result.current.residentMeasurementConfig([regular])).toEqual(plain);
    expect(calls).toEqual({ warmLoads: 0, warmPasses: 0, requiredPasses: 0 });
  });
});

describe('useRustMeasurement default fonts', () => {
  test('retries the same chain after a transient default-provider failure', async () => {
    let attempts = 0;
    configureDefaultFonts({
      load: () => {
        attempts++;
        if (attempts === 1) return Promise.reject(new Error('transient chunk failure'));
        return Promise.resolve({
          createFontProvider: () => ({
            resolve: () => () => Promise.resolve(bytesOf('recovered-provider')),
          }),
        });
      },
    });
    const registered: Uint8Array[] = [];
    const engine: RustTextEngine = {
      registerFont(bytes) {
        registered.push(bytes);
        return registered.length;
      },
      clearFonts() {},
    };
    const regular: ResidentFontRequirement = {
      key: 'regular',
      family: 'Calibri',
      bold: false,
      italic: false,
    };
    const warn = spyOn(console, 'warn').mockImplementation(() => {});

    try {
      const { result } = renderHook(() =>
        useRustMeasurement({ document: null, textEngine: engine })
      );
      await waitFor(() => expect(result.current.deferLayoutPass()).toBe(false));

      await waitFor(() => {
        const failed: ResidentMeasurementConfig | null =
          result.current.residentMeasurementConfig([regular]);
        expect(failed).not.toBeNull();
        expect(failed?.fontChains).toEqual({});
      });
      expect(attempts).toBe(1);

      await waitFor(() => {
        const recovered: ResidentMeasurementConfig | null =
          result.current.residentMeasurementConfig([regular]);
        expect(recovered?.fontChains).toEqual({ regular: [1] });
      });
      expect(new TextDecoder().decode(registered[0])).toBe('recovered-provider');
      expect(attempts).toBe(2);
    } finally {
      warn.mockRestore();
    }
  });

  test('a requirement a later pass no longer asks for stays in the measurement config', async () => {
    configureDefaultFonts({
      load: () =>
        Promise.resolve({
          createFontProvider: () => ({
            resolve: (family: string) => () => Promise.resolve(bytesOf(family)),
          }),
        }),
    });
    let registered = 0;
    const engine: RustTextEngine = {
      registerFont() {
        registered += 1;
        return registered;
      },
      clearFonts() {},
    };
    const regular: ResidentFontRequirement = {
      key: 'regular',
      family: 'Calibri',
      bold: false,
      italic: false,
    };
    const symbol: ResidentFontRequirement = {
      key: 'symbol',
      family: 'Symbol',
      bold: false,
      italic: false,
    };
    const { result } = renderHook(() =>
      useRustMeasurement({ document: null, textEngine: engine })
    );
    await waitFor(() => expect(result.current.deferLayoutPass()).toBe(false));
    let both: ResidentMeasurementConfig | null = null;
    await waitFor(() => {
      both = result.current.residentMeasurementConfig([regular, symbol]);
      expect(Object.keys(both?.fontChains ?? {})).toEqual(['regular', 'symbol']);
    });
    expect(result.current.residentMeasurementConfig([regular])).toEqual(both);
  });

  test('a replaced text engine forgets the fonts earlier passes asked for', async () => {
    configureDefaultFonts({
      load: () =>
        Promise.resolve({
          createFontProvider: () => ({
            resolve: (family: string) => () => Promise.resolve(bytesOf(family)),
          }),
        }),
    });
    const engineWith = (): RustTextEngine => {
      let registered = 0;
      return {
        registerFont() {
          registered += 1;
          return registered;
        },
        clearFonts() {},
      };
    };
    const regular: ResidentFontRequirement = {
      key: 'regular',
      family: 'Calibri',
      bold: false,
      italic: false,
    };
    const symbol: ResidentFontRequirement = {
      key: 'symbol',
      family: 'Symbol',
      bold: false,
      italic: false,
    };
    const { result, rerender } = renderHook(
      ({ engine }: { engine: RustTextEngine }) =>
        useRustMeasurement({ document: null, textEngine: engine }),
      { initialProps: { engine: engineWith() } }
    );
    await waitFor(() =>
      expect(
        Object.keys(result.current.residentMeasurementConfig([regular, symbol])?.fontChains ?? {})
      ).toEqual(['regular', 'symbol'])
    );
    rerender({ engine: engineWith() });
    await waitFor(() =>
      expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
        regular: [1],
      })
    );
  });

  test('a font an earlier pass is still loading holds back no later pass', async () => {
    configureDefaultFonts({
      load: () =>
        Promise.resolve({
          createFontProvider: () => ({
            resolve: (family: string) => () =>
              family === 'Pending'
                ? new Promise<ArrayBuffer>(() => {})
                : Promise.resolve(bytesOf(family)),
          }),
        }),
    });
    let registered = 0;
    const engine: RustTextEngine = {
      registerFont() {
        registered += 1;
        return registered;
      },
      clearFonts() {},
    };
    const regular: ResidentFontRequirement = {
      key: 'regular',
      family: 'Calibri',
      bold: false,
      italic: false,
    };
    const pending: ResidentFontRequirement = {
      key: 'pending',
      family: 'Pending',
      bold: false,
      italic: false,
    };
    const { result } = renderHook(() =>
      useRustMeasurement({ document: null, textEngine: engine })
    );
    await waitFor(() => expect(result.current.deferLayoutPass()).toBe(false));
    await waitFor(() => expect(result.current.residentMeasurementConfig([regular])).not.toBeNull());
    expect(result.current.residentMeasurementConfig([regular, pending])).toBeNull();
    expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
      regular: [1],
    });
  });

  test('a font that loads after the editor let go of its engine registers nothing on it', async () => {
    let finishLoad: (() => void) | undefined;
    configureDefaultFonts({
      load: () =>
        Promise.resolve({
          createFontProvider: () => ({
            resolve: () => () =>
              new Promise<ArrayBuffer>((resolve) => {
                finishLoad = () => resolve(bytesOf('late'));
              }),
          }),
        }),
    });
    const registered: Uint8Array[] = [];
    const engine: RustTextEngine = {
      registerFont(bytes) {
        registered.push(bytes);
        return registered.length;
      },
      clearFonts() {},
    };
    const regular: ResidentFontRequirement = {
      key: 'regular',
      family: 'Calibri',
      bold: false,
      italic: false,
    };
    const { result, unmount } = renderHook(() =>
      useRustMeasurement({ document: null, textEngine: engine })
    );
    await waitFor(() => expect(result.current.deferLayoutPass()).toBe(false));
    let passes = 0;
    result.current.runLayoutPipelineRef.current = () => {
      passes++;
    };
    expect(result.current.residentMeasurementConfig([regular])).toBeNull();
    await waitFor(() => expect(finishLoad).toBeDefined());

    unmount();
    finishLoad!();
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(registered).toEqual([]);
    expect(passes).toBe(0);
  });

  test('a source replaced on the same engine loads the fonts its predecessor was loading', async () => {
    configureDefaultFonts({
      load: () =>
        Promise.resolve({
          createFontProvider: () => ({
            resolve: () => () =>
              new Promise<ArrayBuffer>((resolve) => setTimeout(() => resolve(bytesOf('font')), 5)),
          }),
        }),
    });
    let fonts = 0;
    const engine: RustTextEngine = {
      registerFont: () => ++fonts,
      clearFonts() {},
    };
    const regular: ResidentFontRequirement = {
      key: 'regular',
      family: 'Calibri',
      bold: false,
      italic: false,
    };
    // StrictMode replays the mount: the first source's warmup starts, then the source is released.
    const { result } = renderHook(
      () => {
        const measurement = useRustMeasurement({ document: null, textEngine: engine });
        const { residentMeasurementConfig } = measurement;
        useEffect(() => {
          residentMeasurementConfig([regular]);
        }, [residentMeasurementConfig]);
        return measurement;
      },
      { wrapper: StrictMode }
    );
    await waitFor(() =>
      expect(result.current.residentMeasurementConfig([regular])?.fontChains).toEqual({
        regular: [fonts],
      })
    );
  });
});
