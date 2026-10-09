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
