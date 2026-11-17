import type { BundledFontLoader, BundledFontScript } from '@betteroffice/fonts';
import type { FontCdnConfiguration } from '../../../../shared/office-session/fontCdn';
import { scriptFontFace } from './commandScriptFonts';
import type { DefaultFontFace, DefaultFontLanding, RequestedFontFace } from '../../../../shared/pptxDefaultFonts';
import { decodeFontFace } from '../wasm/loader';
import { sessionFontDecoder } from './sessionFontDecoder';
import type { createSlicedFontFaces } from './commandSlicedFonts';

type FontRequest =
  | { kind: 'primary'; family: string; bold: boolean; italic: boolean }
  | { kind: 'coverage'; script: string; bold: boolean; italic: boolean };

/** @internal */
const DEFAULT_FONT_MAX_ATTEMPTS = 3;
/** @internal */
const DEFAULT_FONT_RETRY_DELAYS_MS = [2000, 8000];

let decoderReady: Promise<unknown> = Promise.resolve();

const sessionDecoder = sessionFontDecoder(async (bytes, options) => {
  await decoderReady.catch(() => {});
  return decodeFontFace(bytes, options);
});

export const commandFontProvider = {
  /** @internal */
  async load(fontCdn?: FontCdnConfiguration) {
    const fonts = await import('@betteroffice/fonts');
    if (fontCdn) fonts.configureFontCdn(fontCdn);
    return fonts.createFontProvider(undefined, { decoder: sessionDecoder });
  },
};

/**
 * Starts loading the regular coverage face of each script before the deck is parsed. `provider` returns the
 * session's one font provider, the one later given to `createCommandDefaultFonts`, so a face started here is
 * fetched and decoded once. Decoding waits for `wasmReady`, where the font decoder lives.
 */
export function prefetchCoverageFaces(
  scripts: readonly string[], wasmReady: Promise<unknown>, provider: () => ReturnType<typeof commandFontProvider.load>,
  sliced = false,
): void {
  decoderReady = wasmReady;
  if (!scripts.length) return;
  void provider().then(async (source) => {
    for (const script of scripts) {
      const load = source.resolveScriptFallback(script as BundledFontScript, false, false);
      if (load?.slices && sliced && (await import('./commandSlicedFonts')).slicedFacesSupported()) {
        const { openSlicedFace } = await import('@betteroffice/fonts/progressive');
        void openSlicedFace(load.slices).catch(() => {});
      } else void load?.().catch(() => {});
    }
  }).catch(() => {});
}

export function createCommandDefaultFonts(
  supplied: readonly { family: string }[],
  collect: (slideId?: string) => RequestedFontFace[],
  register: (face: DefaultFontFace) => void,
  settled: (faces: DefaultFontLanding[]) => void,
  install?: (faces: readonly DefaultFontFace[]) => Promise<readonly DefaultFontFace[]>,
  beforeRegister?: () => void,
  shared?: () => ReturnType<typeof commandFontProvider.load>,
  sliced?: ReturnType<typeof createSlicedFontFaces>,
) {
  const families = new Set(supplied.map((face) => face.family.trim().toLowerCase()));
  const attempted = new Set<string>();
  const attempts = new Map<string, number>();
  const retryAt = new Map<string, number>();
  const pending = new Map<string, Promise<DefaultFontFace | undefined>>();
  const installations = new Map<string, Promise<boolean>>();
  const registered = new Set<string>();
  let provider: ReturnType<typeof commandFontProvider.load> | undefined;
  let disposed = false;
  let prepared = false;
  const cancellations = new Set<() => void>();
  const key = (face: FontRequest) => JSON.stringify([face.kind,
    face.kind === 'coverage' ? face.script : face.family.trim().toLowerCase(), face.bold, face.italic]);
  const suppliedFamily = (face: FontRequest) => face.kind === 'primary' && families.has(face.family.trim().toLowerCase());

  const retryable = (id: string) => !attempted.has(id) ||
    ((attempts.get(id) ?? 0) < DEFAULT_FONT_MAX_ATTEMPTS && Date.now() >= (retryAt.get(id) ?? Infinity));
  const failed = (id: string) => retryAt.set(id,
    Date.now() + (DEFAULT_FONT_RETRY_DELAYS_MS[(attempts.get(id) ?? 1) - 1] ?? Infinity));

  function publish(wanted: FontRequest[], faces: (DefaultFontFace | undefined)[]): DefaultFontLanding[] {
    if (disposed) return [];
    const additions: DefaultFontLanding[] = [];
    for (const [index, face] of faces.entries()) {
      if (!face || suppliedFamily(wanted[index])) continue;
      const registeredKey = JSON.stringify([face.kind ?? 'primary', face.family.trim().toLowerCase(), face.bold, face.italic]);
      if (registered.has(registeredKey)) continue;
      try {
        if (!prepared) {
          beforeRegister?.();
          prepared = true;
        }
        register(face);
        registered.add(registeredKey);
        additions.push({ family: face.family, bold: face.bold, italic: face.italic,
          ...(face.kind ? { kind: face.kind } : {}), ...(face.ownWeight ? { ownWeight: true } : {}) });
      } catch { failed(key(wanted[index])); }
    }
    if (additions.length) settled(additions);
    return additions;
  }

  function request(face: FontRequest): Promise<DefaultFontFace | undefined> {
    const id = key(face);
    const existing = pending.get(id);
    if (existing) return existing;
    if (!retryable(id) || suppliedFamily(face)) return Promise.resolve(undefined);
    attempted.add(id);
    attempts.set(id, (attempts.get(id) ?? 0) + 1);
    retryAt.delete(id);
    const task = new Promise<DefaultFontFace | undefined>((resolve) => {
      let finished = false;
      const finish = (value?: DefaultFontFace) => {
        if (finished) return;
        finished = true;
        clearTimeout(timer);
        cancellations.delete(cancel);
        resolve(value);
      };
      const cancel = () => finish();
      const timer = setTimeout(cancel, 5000);
      cancellations.add(cancel);
      let available = false;
      provider ??= Promise.resolve().then(() => shared ? shared() : commandFontProvider.load()).catch((error) => {
        provider = undefined;
        throw error;
      });
      void provider.then(async (source) => {
        if (disposed) return undefined;
        let bold = face.bold;
        let italic = face.italic;
        let load: BundledFontLoader | undefined;
        let light: (() => Promise<ArrayBuffer>) | undefined;
        if (face.kind === 'coverage') {
          load = source.resolveScriptFallback(face.script as BundledFontScript, bold, italic);
          if (!load && (bold || italic)) {
            bold = false;
            italic = false;
            load = source.resolveScriptFallback(face.script as BundledFontScript, bold, italic);
          }
        } else {
          const named = source.resolve(face.family, bold, italic) ?? source.resolveFamily(face.family, bold, italic);
          light = named ? undefined : source.resolveLightVariant?.(face.family, bold, italic);
          load = named ?? light ?? source.resolveLastResort(face.family, bold, italic, 'powerpoint');
        }
        if (!load) return undefined;
        available = true;
        if (load.slices && sliced?.supported()) {
          try {
            const { openSlicedFace } = await import('@betteroffice/fonts/progressive');
            const opened = await openSlicedFace(load.slices);
            if (disposed || suppliedFamily(face)) return undefined;
            const skeleton = face.kind === 'coverage' ? await scriptFontFace(face.script, bold, italic, opened.bytes())
              : { family: face.family, bold: face.bold, italic: face.italic, bytes: opened.bytes() };
            if (skeleton && !families.has(skeleton.family.trim().toLowerCase())) {
              sliced.add(skeleton.family, { get revision() { return opened.revision; }, bytes: () => opened.bytes(),
                ensureText: (text) => opened.ensureText(text), prefetch: (text) => opened.prefetch(text),
                chunksOf: (text) => opened.chunksOf(text), loadedChunks: () => opened.loadedChunks() }, load);
              return disposed ? undefined : skeleton;
            }
          } catch {}
        }
        const bytes = await load();
        if (disposed || suppliedFamily(face)) return undefined;
        const loaded = face.kind === 'coverage'
          ? await scriptFontFace(face.script, bold, italic, bytes)
          : { family: face.family, bold: face.bold, italic: face.italic, bytes, ...(light ? { ownWeight: true } : {}) };
        if (!loaded || disposed || suppliedFamily(face)) return undefined;
        sliced?.plain(loaded.family, loaded.bold, loaded.italic);
        const identity = JSON.stringify([loaded.family.trim().toLowerCase(), loaded.bold, loaded.italic]);
        let installation = installations.get(identity);
        if (!installation) {
          installation = Promise.resolve().then(async () => !disposed && (!install || (await install([loaded])).length > 0));
          installations.set(identity, installation);
          const current = installation;
          const forget = () => { if (installations.get(identity) === current) installations.delete(identity); };
          void installation.then((success) => { if (!success) forget(); }, forget);
        }
        return await installation && !disposed ? loaded : undefined;
      }).then((loaded) => {
        pending.delete(id);
        if (!loaded && available && !disposed && !suppliedFamily(face)) failed(id);
        if (finished && loaded) publish([face], [loaded]);
        else finish(loaded);
      }, () => {
        pending.delete(id);
        if (!disposed && !suppliedFamily(face)) failed(id);
        finish();
      });
    });
    pending.set(id, task);
    return task;
  }

  function requests(slideId?: string): FontRequest[] {
    if (disposed) return [];
    try {
      const wanted = new Map<string, FontRequest>();
      for (const face of collect(slideId)) {
        const primary: FontRequest = { kind: 'primary', family: face.family, bold: face.bold, italic: face.italic };
        wanted.set(key(primary), primary);
        for (const script of face.scripts ?? []) {
          const coverage: FontRequest = { kind: 'coverage', script, bold: face.bold, italic: face.italic };
          wanted.set(key(coverage), coverage);
        }
      }
      return [...wanted.values()].filter((face) => !suppliedFamily(face) &&
        (retryable(key(face)) || pending.has(key(face))));
    } catch { return []; }
  }

  return {
    isLoading(): boolean { return pending.size > 0; },
    hasPending(slideId?: string): boolean { return requests(slideId).length > 0; },
    load(slideId?: string): DefaultFontLanding[] | Promise<DefaultFontLanding[]> {
      const wanted = requests(slideId);
      if (!wanted.length) return [];
      return Promise.all(wanted.map(request)).then((faces) => publish(wanted, faces));
    },
    supply(family: string): void { families.add(family.trim().toLowerCase()); },
    dispose(): void {
      disposed = true;
      for (const cancel of cancellations) cancel();
      pending.clear();
      installations.clear();
    },
  };
}
