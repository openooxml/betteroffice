interface NodeFsLike {
  readFileSync(path: string): Uint8Array;
}

interface NodeUrlLike {
  fileURLToPath(url: string): string;
}

function builtinModule<T>(name: string): T | undefined {
  const proc = (globalThis as { process?: { getBuiltinModule?: (id: string) => unknown } }).process;
  if (typeof proc?.getBuiltinModule !== 'function') return undefined;
  try {
    return proc.getBuiltinModule(name) as T;
  } catch {
    return undefined;
  }
}

export function readWasmSync(url: URL, throwOnReadError = false): Uint8Array | undefined {
  if (url.protocol !== 'file:') return undefined;
  const fs = builtinModule<NodeFsLike>('node:fs');
  const nodeUrl = builtinModule<NodeUrlLike>('node:url');
  if (!fs || !nodeUrl) return undefined;
  try {
    return fs.readFileSync(nodeUrl.fileURLToPath(url.href));
  } catch (error) {
    if (throwOnReadError) throw error;
    return undefined;
  }
}

export function packagedWasmInput(url: URL): Uint8Array | URL {
  return readWasmSync(url) ?? url;
}

export function packagedWasmError(url: URL, error: unknown): Error {
  const reason = error instanceof Error ? error.message : String(error);
  return Object.assign(
    new Error(`Could not load the packaged WASM asset ${url.href} (${reason}); pass its bytes or a compiled WebAssembly.Module to initWasm(input)`),
    { cause: error }
  );
}
