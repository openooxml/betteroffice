import type { BundledFontFace } from './manifest';

interface NodeFsLike {
  readFileSync(path: string): Uint8Array;
}
interface NodeUrlLike {
  fileURLToPath(url: string): string;
}

function builtinModule<T>(name: string): T | undefined {
  const proc = (
    globalThis as { process?: { getBuiltinModule?: (id: string) => unknown } }
  ).process;
  if (typeof proc?.getBuiltinModule !== 'function') return undefined;
  try {
    return proc.getBuiltinModule(name) as T;
  } catch {
    return undefined;
  }
}

/** Node fetch cannot read file: package assets, so server imports use built-in I/O. */
function readFileAsset(url: URL): ArrayBuffer | undefined {
  if (url.protocol !== 'file:') return undefined;
  const fs = builtinModule<NodeFsLike>('node:fs');
  const nodeUrl = builtinModule<NodeUrlLike>('node:url');
  if (!fs || !nodeUrl) return undefined;
  try {
    // `.href` and not the URL object: fileURLToPath brand-checks its argument.
    const bytes = fs.readFileSync(nodeUrl.fileURLToPath(url.href));
    return bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;
  } catch {
    return undefined;
  }
}

function validateByteLength(
  face: BundledFontFace,
  bytes: ArrayBuffer,
): ArrayBuffer {
  if (bytes.byteLength !== face.byteLength) {
    throw new Error(
      `Bundled font ${face.file} has ${bytes.byteLength} bytes; expected ${face.byteLength}`,
    );
  }
  return bytes;
}

const bytesCache = new Map<string, Promise<ArrayBuffer>>();

export function loadFontBytes(
  face: BundledFontFace,
  url: URL,
  timeoutMs?: number,
): Promise<ArrayBuffer> {
  const key = `${url.href}|${face.byteLength}|${timeoutMs ?? ''}`;
  const cached = bytesCache.get(key);
  if (cached) return cached;
  const promise = (async () => {
    const onDisk = readFileAsset(url);
    if (onDisk) return validateByteLength(face, onDisk);
    const response = await fetch(
      url,
      timeoutMs === undefined
        ? undefined
        : {
            signal: AbortSignal.timeout(timeoutMs),
            credentials: 'omit',
            referrerPolicy: 'no-referrer',
          },
    );
    if (!response.ok) {
      throw new Error(
        `Failed to fetch bundled font ${face.file}: HTTP ${response.status}`,
      );
    }
    return validateByteLength(face, await response.arrayBuffer());
  })();
  promise.catch(() => {
    if (bytesCache.get(key) === promise) bytesCache.delete(key);
  });
  bytesCache.set(key, promise);
  return promise;
}
