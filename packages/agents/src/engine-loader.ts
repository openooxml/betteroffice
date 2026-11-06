export type EngineFormat = 'docx' | 'xlsx' | 'pptx';
export interface EngineModules {
  docx: {
    yrs: typeof import('@betteroffice/docx/yrs');
    editor: typeof import('@betteroffice/docx/editor');
    layout: typeof import('@betteroffice/docx/layout');
    utils: typeof import('@betteroffice/docx/utils');
    render: typeof import('@betteroffice/docx/layout/render');
  };
  xlsx: typeof import('@betteroffice/xlsx');
  pptx: typeof import('@betteroffice/pptx');
}

type Importers = { [F in EngineFormat]: () => Promise<EngineModules[F]> };
interface Engine<T> { module: T; native: boolean }

/** @internal */
export function createEngineLoader(native: Importers | undefined, wasm: Importers, warn: (warning: Error) => void) {
  const loaded: Partial<{ [F in EngineFormat]: Engine<EngineModules[F]> }> = {};
  const pending: Partial<{ [F in EngineFormat]: Promise<Engine<EngineModules[F]>> }> = {};
  const failures = new Set<EngineFormat>();
  let warning: (Error & { code: string }) | undefined;
  const unavailable = (format: EngineFormat, error: unknown) => {
    failures.add(format);
    if (warning) return;
    const reason = error instanceof Error ? error.message : String(error);
    warning = Object.assign(new Error(), { code: 'BETTEROFFICE_AGENTS_NATIVE_UNAVAILABLE' });
    Object.defineProperty(warning, 'message', { get: () => `Native engines unavailable for ${[...failures].join(', ')}: ${reason}. DOCX, PPTX and XLSX fall back to WebAssembly when their native engine cannot load. Previews on this path need @napi-rs/canvas installed.` });
    warn(warning);
  };
  const remember = <F extends EngineFormat>(format: F, engine: Engine<EngineModules[F]>) => {
    const ready = loaded[format] as Engine<EngineModules[F]> | undefined;
    if (ready) return ready;
    (loaded as Record<F, Engine<EngineModules[F]>>)[format] = engine;
    return engine;
  };
  return {
    load<F extends EngineFormat>(format: F): Promise<Engine<EngineModules[F]>> {
      const ready = loaded[format] as Engine<EngineModules[F]> | undefined;
      if (ready) return Promise.resolve(ready);
      let task = pending[format] as Promise<Engine<EngineModules[F]>> | undefined;
      if (!task) {
        task = (async () => {
          if (native) {
            try { return remember(format, { module: await native[format]() as EngineModules[F], native: true }); }
            catch (error) { unavailable(format, error); }
          }
          return remember(format, { module: await wasm[format]() as EngineModules[F], native: false });
        })();
        (pending as Record<F, Promise<Engine<EngineModules[F]>>>)[format] = task;
      }
      return task;
    },
    get<F extends EngineFormat>(format: F): Engine<EngineModules[F]> {
      const ready = loaded[format] as Engine<EngineModules[F]> | undefined;
      if (ready) return ready;
      throw new Error(`The ${format} engine has not loaded. Await open${format[0].toUpperCase()}${format.slice(1)} first.`);
    },
  };
}
