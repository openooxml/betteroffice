import { absoluteHref } from './assetHref';

interface NodeFsLike {
  readFileSync(path: string): Uint8Array;
}

interface NodeUrlLike {
  fileURLToPath(url: string): string;
}

export async function readWasm(url: URL, builtinModule: <T>(name: string) => T | undefined): Promise<Uint8Array> {
  let fs = builtinModule<NodeFsLike>('node:fs');
  let nodeUrl = builtinModule<NodeUrlLike>('node:url');
  if (!fs || !nodeUrl) {
    const fsName = 'node:fs';
    const urlName = 'node:url';
    [fs, nodeUrl] = await Promise.all([import(fsName), import(urlName)]);
  }
  return fs!.readFileSync(nodeUrl!.fileURLToPath(url.href));
}

export function bundledDecoderUrl(): URL | undefined {
  try {
    const url = new URL('./decoder/fontpack.wasm', import.meta.url);
    return new URL(absoluteHref(url.href));
  } catch {
    return undefined;
  }
}
