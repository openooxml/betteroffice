import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, spyOn, test } from 'bun:test';

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
});
