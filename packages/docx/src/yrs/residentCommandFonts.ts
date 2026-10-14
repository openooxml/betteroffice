import type { RustTextEngine } from '../layout/measure/rustMeasureSource';
import type { YrsResidentFontRegistration } from './index';

export function collectCommandFonts(signal?: AbortSignal) {
  const fonts: YrsResidentFontRegistration[] = [];
  let disposed = false;
  const dispose = () => {
    disposed = true;
    fonts.length = 0;
    signal?.removeEventListener('abort', dispose);
  };
  const append = (font: YrsResidentFontRegistration) => {
    signal?.throwIfAborted();
    if (disposed) throw new Error('DOCX command fonts are disposed');
    return fonts.push(font instanceof Uint8Array ? font.slice() : font) - 1;
  };
  const engine: RustTextEngine = {
    registerFont: (bytes) => append(bytes),
    registerSubstituteFont: (substituteOf, family) => append({ substituteOf, family }),
    clearFonts: dispose,
  };
  signal?.throwIfAborted();
  signal?.addEventListener('abort', dispose, { once: true });
  return { fonts, engine, dispose };
}

export function registerCommandFonts(
  engine: Pick<RustTextEngine, 'registerFont' | 'registerSubstituteFont'>,
  fonts: readonly YrsResidentFontRegistration[],
) {
  const ids: (number | null)[] = [];
  const resolve = (slot: number): number | null => {
    if (!Number.isSafeInteger(slot) || slot < 0 || slot >= ids.length) {
      throw new Error('Invalid DOCX command font slot');
    }
    return ids[slot];
  };
  for (const font of fonts) {
    let id: number | null;
    if (font instanceof Uint8Array) {
      try { id = engine.registerFont(font); }
      catch (error) {
        if (error instanceof WebAssembly.RuntimeError) throw error;
        id = null;
      }
    } else {
      const base = resolve(font.substituteOf);
      try { id = base === null ? null : engine.registerSubstituteFont?.(base, font.family) ?? base; }
      catch (error) {
        if (error instanceof WebAssembly.RuntimeError) throw error;
        id = base;
      }
    }
    if (id !== null && (!Number.isSafeInteger(id) || id < 0)) throw new Error('Invalid DOCX command font ID');
    ids.push(id);
  }
  const remap = (json: string, fontChains?: Record<string, number[]>): string => {
    const visit = (value: unknown): void => {
      if (!value || typeof value !== 'object') return;
      for (const [key, child] of Object.entries(value)) {
        if (key === 'fontChains' && child && typeof child === 'object') {
          const chains: Record<string, number[]> = {};
          for (const [family, chain] of Object.entries(fontChains ?? child)) {
            if (!Array.isArray(chain)) throw new Error('Invalid DOCX command font chain');
            const mapped = chain.map(resolve).filter((id): id is number => id !== null);
            if (mapped.length) chains[family] = [...new Set(mapped)];
          }
          (value as Record<string, unknown>)[key] = chains;
        } else visit(child);
      }
    };
    const value: unknown = JSON.parse(json || '{}');
    visit(value);
    return JSON.stringify(value);
  };
  return { ids, remap };
}
