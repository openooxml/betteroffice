import { installBaseLatinFace } from '../../wasm/baseLatinFonts';
import type { FontCoverage } from './fontCoverage';
import type { FontDecoder } from '../../../../../shared/office-session/fontDecoder';
import type { BundledFontProvider } from './fontRegistry';

type Loader = () => Promise<ArrayBuffer>;

/** Registers the backstop faces covering `missing`, pushing each useful id onto `ids`. */
export async function loadBaseLatinFaces(options: {
  bold: boolean;
  italic: boolean;
  missing: string[];
  ids: number[];
  provider: () => Promise<BundledFontProvider | undefined>;
  wrapped: WeakMap<Loader, Loader>;
  register: (loader: Loader, face: { family: string; bold: boolean; italic: boolean }) => Promise<number>;
  coverage: (id: number) => FontCoverage | undefined;
  candidate: (character: string, styled: boolean) => boolean;
}): Promise<void> {
  const { bold, italic, ids, wrapped, candidate } = options;
  let missing = options.missing;
  const provider = await options.provider();
  const styles: Array<[boolean, boolean]> = bold || italic ? [[bold, italic], [false, false]] : [[false, false]];
  for (const [faceBold, faceItalic] of styles) {
    if (!missing.some((character) => candidate(character, faceBold || faceItalic))) continue;
    const loader = provider?.resolveBaseLatin?.(faceBold, faceItalic);
    if (!loader) continue;
    let registered = wrapped.get(loader);
    if (!registered) {
      registered = async () => {
        const bytes = await loader();
        await installBaseLatinFace(bytes, faceBold, faceItalic);
        return bytes;
      };
      wrapped.set(loader, registered);
    }
    const id = await options.register(registered, { family: 'BetterOffice Base Latin', bold: faceBold, italic: faceItalic });
    const covers = options.coverage(id);
    if (missing.some((character) => covers?.(character))) {
      ids.push(id);
      missing = missing.filter((character) => !covers?.(character));
    }
    if (missing.length === 0) break;
  }
}

export function resolveBaseLatinProvider(
  state: {
    baseUrl: string | URL | undefined;
    cache: WeakMap<object, Promise<BundledFontProvider>>;
    runtime: { decoder?: FontDecoder };
  },
  runtime?: { decoder?: FontDecoder },
): Promise<BundledFontProvider> {
  const key = runtime ?? state.runtime;
  const cache = state.cache;
  let pending = cache.get(key);
  if (!pending) {
    const baseUrl = state.baseUrl;
    pending = import('@betteroffice/fonts').then(manager =>
      manager.createFontProvider(baseUrl === undefined ? undefined : { baseUrl }, key));
    const current = pending;
    pending.catch(() => { if (cache.get(key) === current) cache.delete(key); });
    cache.set(key, pending);
  }
  return pending;
}
