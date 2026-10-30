import type { BundledFontScript } from '@betteroffice/fonts';
import { scriptFontFace } from './commandScriptFonts';
import type { DefaultFontFace, RequestedFontFace } from '../../../../shared/pptxDefaultFonts';

type FontRequest =
  | { kind: 'primary'; family: string; bold: boolean; italic: boolean }
  | { kind: 'coverage'; script: string; bold: boolean; italic: boolean };

export const commandFontProvider = {
  async load() { return (await import('@betteroffice/fonts')).createFontProvider(); },
};

export function createCommandDefaultFonts(
  supplied: readonly { family: string }[],
  collect: (slideId?: string) => RequestedFontFace[],
  register: (face: DefaultFontFace) => void,
  settled: (faces: DefaultFontFace[]) => void,
  install?: (faces: readonly DefaultFontFace[]) => Promise<readonly DefaultFontFace[]>,
  beforeRegister?: () => void,
) {
  const families = new Set(supplied.map((face) => face.family.trim().toLowerCase()));
  const attempted = new Set<string>();
  const pending = new Map<string, Promise<DefaultFontFace | undefined>>();
  const coverageInstallations = new Map<string, Promise<DefaultFontFace | undefined>>();
  const registered = new Set<string>();
  let provider: ReturnType<typeof commandFontProvider.load> | undefined;
  let disposed = false;
  let prepared = false;
  const cancellations = new Set<() => void>();
  const key = (face: FontRequest) => JSON.stringify([face.kind,
    face.kind === 'coverage' ? face.script : face.family.trim().toLowerCase(), face.bold, face.italic]);
  const suppliedFamily = (face: FontRequest) => face.kind === 'primary' && families.has(face.family.trim().toLowerCase());

  function request(face: FontRequest): Promise<DefaultFontFace | undefined> {
    const id = key(face);
    const existing = pending.get(id);
    if (existing) return existing;
    if (attempted.has(id) || suppliedFamily(face)) return Promise.resolve(undefined);
    attempted.add(id);
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
      provider ??= commandFontProvider.load();
      void provider.then(async (source) => {
        if (disposed || finished) return undefined;
        let bold = face.bold;
        let italic = face.italic;
        let load: (() => Promise<ArrayBuffer>) | undefined;
        if (face.kind === 'coverage') {
          load = source.resolveScriptFallback(face.script as BundledFontScript, bold, italic);
          if (!load && (bold || italic)) {
            bold = false;
            italic = false;
            load = source.resolveScriptFallback(face.script as BundledFontScript, bold, italic);
          }
        } else {
          load = source.resolve(face.family, bold, italic)
            ?? source.resolveFamily(face.family, bold, italic)
            ?? source.resolveLightVariant?.(face.family, bold, italic);
        }
        if (!load) return undefined;
        const bytes = await load();
        if (disposed || finished || suppliedFamily(face)) return undefined;
        const loaded = face.kind === 'coverage'
          ? await scriptFontFace(face.script, bold, italic, bytes)
          : { family: face.family, bold: face.bold, italic: face.italic, bytes };
        if (!loaded || disposed || finished) return undefined;
        if (face.kind === 'coverage') {
          const identity = JSON.stringify([loaded.family.trim().toLowerCase(), loaded.bold, loaded.italic]);
          let installation = coverageInstallations.get(identity);
          if (!installation) {
            installation = Promise.resolve().then(async () => {
              if (disposed) return undefined;
              if (install && !(await install([loaded])).length) return undefined;
              return loaded;
            });
            coverageInstallations.set(identity, installation);
          }
          return installation;
        }
        if (install && !(await install([loaded])).length) return undefined;
        return loaded;
      }).then(finish, () => finish());
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
        (!attempted.has(key(face)) || pending.has(key(face))));
    } catch { return []; }
  }

  return {
    isLoading(): boolean { return pending.size > 0; },
    hasPending(slideId?: string): boolean { return requests(slideId).length > 0; },
    load(slideId?: string): DefaultFontFace[] | Promise<DefaultFontFace[]> {
      const wanted = requests(slideId);
      if (!wanted.length) return [];
      return Promise.all(wanted.map(request)).then((faces) => {
        if (disposed) return [];
        const additions: DefaultFontFace[] = [];
        for (const [index, face] of faces.entries()) {
          pending.delete(key(wanted[index]));
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
            additions.push({ ...face, bytes: face.bytes.slice(0) });
          } catch {}
        }
        if (additions.length) settled(additions);
        return additions;
      });
    },
    supply(family: string): void { families.add(family.trim().toLowerCase()); },
    dispose(): void {
      disposed = true;
      for (const cancel of cancellations) cancel();
      pending.clear();
      coverageInstallations.clear();
    },
  };
}
