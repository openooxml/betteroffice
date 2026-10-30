import type { DefaultFontFace } from '../../../../shared/pptxDefaultFonts';

type RequestedFace = Omit<DefaultFontFace, 'bytes'>;

export const commandFontProvider = {
  async load() { return (await import('@betteroffice/fonts')).createFontProvider(); },
};

export function createCommandDefaultFonts(
  supplied: readonly { family: string }[],
  collect: (slideId?: string) => RequestedFace[],
  register: (face: DefaultFontFace) => void,
  settled: (faces: DefaultFontFace[]) => void,
  install?: (faces: readonly DefaultFontFace[]) => Promise<readonly DefaultFontFace[]>,
  beforeRegister?: () => void,
) {
  const families = new Set(supplied.map((face) => face.family.trim().toLowerCase()));
  const attempted = new Set<string>();
  const pending = new Map<string, Promise<DefaultFontFace | undefined>>();
  const registered = new Set<string>();
  let provider: ReturnType<typeof commandFontProvider.load> | undefined;
  let disposed = false;
  let prepared = false;
  const cancellations = new Set<() => void>();
  const key = (face: RequestedFace) => JSON.stringify([face.family.trim().toLowerCase(), face.bold, face.italic]);

  function request(face: RequestedFace): Promise<DefaultFontFace | undefined> {
    const id = key(face);
    const existing = pending.get(id);
    if (existing) return existing;
    if (attempted.has(id) || families.has(face.family.trim().toLowerCase())) return Promise.resolve(undefined);
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
        const load = source.resolve(face.family, face.bold, face.italic)
          ?? source.resolveFamily(face.family, face.bold, face.italic);
        if (!load) return undefined;
        const bytes = await load();
        if (disposed || finished || families.has(face.family.trim().toLowerCase())) return undefined;
        const loaded = { ...face, bytes };
        if (install && !(await install([loaded])).length) return undefined;
        return loaded;
      }).then(finish, () => finish());
    });
    pending.set(id, task);
    return task;
  }

  function requests(slideId?: string): RequestedFace[] {
    if (disposed) return [];
    try {
      return collect(slideId).filter((face) => !families.has(face.family.trim().toLowerCase()) &&
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
          if (!face || registered.has(key(face)) || families.has(face.family.trim().toLowerCase())) continue;
          try {
            if (!prepared) {
              beforeRegister?.();
              prepared = true;
            }
            register(face);
            registered.add(key(face));
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
    },
  };
}
