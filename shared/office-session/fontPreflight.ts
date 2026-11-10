import { readWasmSync } from './readWasmAsset';

export type PreflightWasmInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

/** @experimental */
export interface FontPreflightFace {
  family: string;
  bold: boolean;
  italic: boolean;
  scripts: string[];
}

/** @experimental */
export interface FontPreflight {
  faces: FontPreflightFace[];
  /** False when some font rules could not be resolved. */
  complete: boolean;
}

/** @internal */
export function createPreflightState(options: {
  assetUrl: () => URL;
  engineKind: 'wasm' | 'native';
  initAsync: (input: { module_or_path: PreflightWasmInput }) => Promise<unknown>;
}) {
  let ready = false;
  let pending: Promise<void> | undefined;
  let memory: WebAssembly.Memory | undefined;
  return {
    status(): 'idle' | 'loading' | 'ready' {
      return ready ? 'ready' : pending ? 'loading' : 'idle';
    },
    memory(): WebAssembly.Memory | undefined {
      return memory;
    },
    preload(input?: PreflightWasmInput): Promise<void> {
      if (ready) return Promise.resolve();
      if (options.engineKind === 'native') {
        ready = true;
        return Promise.resolve();
      }
      if (pending) return pending;
      const url = options.assetUrl();
      pending = Promise.resolve().then(() => options.initAsync({
        module_or_path: input ?? readWasmSync(url) ?? url,
      })).then((output) => {
        memory = (output as { memory?: WebAssembly.Memory }).memory;
        ready = true;
      }).catch((error: unknown) => {
        pending = undefined;
        throw error instanceof Error ? error : new Error(String(error));
      });
      return pending;
    },
  };
}
