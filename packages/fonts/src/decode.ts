interface NodeZlibLike {
  zstdDecompressSync?: (
    bytes: Uint8Array,
    options: { dictionary?: Uint8Array; maxOutputLength: number },
  ) => Uint8Array;
}

interface NodeFsLike {
  readFileSync(path: string): Uint8Array;
}

interface NodeUrlLike {
  fileURLToPath(url: string): string;
}

interface NodeBufferLike {
  Buffer: { from(bytes: Uint8Array): Uint8Array };
}

interface Decoder {
  memory: WebAssembly.Memory;
  fp_alloc(len: number): number;
  fp_free(ptr: number, len: number): void;
  fp_decode(
    srcPtr: number,
    srcLen: number,
    dictPtr: number,
    dictLen: number,
    maxOutput: number,
  ): number;
  fp_out_ptr(): number;
  fp_out_len(): number;
  fp_out_free(): void;
}

const DEFAULT_MAX_OUTPUT = 64 * 1024 * 1024;
const instances = new Map<string, Promise<Decoder>>();

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

function view(bytes: ArrayBuffer | Uint8Array): Uint8Array {
  return bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
}

function copy(bytes: Uint8Array): ArrayBuffer {
  return new Uint8Array(bytes).buffer;
}

function corrupt(): Error {
  return new Error('Font face is corrupt or truncated, or uses a wrong dictionary');
}

function contentSize(bytes: Uint8Array, maxOutput: number): number {
  if (bytes.length < 5) throw corrupt();
  const descriptor = bytes[4];
  if (descriptor & 8) throw corrupt();
  const singleSegment = (descriptor & 32) !== 0;
  const sizeFlag = descriptor >>> 6;
  const sizeLength = sizeFlag === 0 ? (singleSegment ? 1 : 0) : 1 << sizeFlag;
  let offset = 5 + (singleSegment ? 0 : 1);
  const dictLength = [0, 1, 2, 4][descriptor & 3];
  offset += dictLength;
  if (bytes.length < offset + sizeLength) throw corrupt();
  if (sizeLength === 0) throw new Error('Font face has no decoded content size');
  let size = 0n;
  for (let index = 0; index < sizeLength; index++) {
    size |= BigInt(bytes[offset + index]) << BigInt(8 * index);
  }
  if (sizeFlag === 1) size += 256n;
  if (size > BigInt(maxOutput)) {
    throw new Error(`Font face is too large: ${size} decoded bytes exceed ${maxOutput}`);
  }
  if ((descriptor & 4) === 0) throw new Error('Font face has no content checksum');
  offset += sizeLength;
  let last = false;
  while (!last) {
    if (offset + 3 > bytes.length) throw corrupt();
    const block = bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16);
    last = (block & 1) !== 0;
    const type = (block >>> 1) & 3;
    if (type === 3) throw corrupt();
    offset += 3 + (type === 1 ? 1 : block >>> 3);
    if (offset > bytes.length) throw corrupt();
  }
  if (offset + 4 !== bytes.length) {
    throw new Error('Font face is corrupt or must contain exactly one zstd frame');
  }
  return Number(size);
}

async function readWasm(url: URL): Promise<Uint8Array> {
  let fs = builtinModule<NodeFsLike>('node:fs');
  let nodeUrl = builtinModule<NodeUrlLike>('node:url');
  if (!fs || !nodeUrl) {
    const fsName = 'node:fs';
    const urlName = 'node:url';
    [fs, nodeUrl] = await Promise.all([import(fsName), import(urlName)]);
  }
  return fs!.readFileSync(nodeUrl!.fileURLToPath(url.href));
}

async function instantiate(url: URL): Promise<Decoder> {
  let instance: WebAssembly.Instance;
  if (url.protocol === 'file:') {
    const result = await WebAssembly.instantiate(copy(await readWasm(url)), {});
    instance = result.instance;
  } else {
    if (typeof fetch !== 'function') throw new Error('Cannot fetch font decoder wasm');
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`Cannot load font decoder wasm: HTTP ${response.status}`);
    }
    if (
      /^https?:$/.test(url.protocol) &&
      typeof Response !== 'undefined' &&
      typeof WebAssembly.instantiateStreaming === 'function'
    ) {
      try {
        instance = (await WebAssembly.instantiateStreaming(response.clone(), {})).instance;
      } catch {
        instance = (await WebAssembly.instantiate(await response.arrayBuffer(), {})).instance;
      }
    } else {
      instance = (await WebAssembly.instantiate(await response.arrayBuffer(), {})).instance;
    }
  }
  const exports = instance.exports;
  if (
    !(exports.memory instanceof WebAssembly.Memory) ||
    ['fp_alloc', 'fp_free', 'fp_decode', 'fp_out_ptr', 'fp_out_len', 'fp_out_free'].some(
      (name) => typeof exports[name] !== 'function',
    )
  ) {
    throw new Error('Font decoder wasm has incompatible exports');
  }
  return exports as unknown as Decoder;
}

function decoder(url: URL): Promise<Decoder> {
  const key = url.href;
  const cached = instances.get(key);
  if (cached) return cached;
  const promise = instantiate(url);
  instances.set(key, promise);
  promise.catch(() => {
    if (instances.get(key) === promise) instances.delete(key);
  });
  return promise;
}

function wasmError(code: number): Error {
  switch (code) {
    case -2:
      return new Error('Font face has no decoded content size');
    case -3:
      return new Error('Font face is too large for maxOutput');
    case -4:
      return new Error('Font face checksum failed: corrupt face or wrong dictionary');
    case -5:
      return new Error('Font face dictionary is missing or wrong');
    case -7:
      return new Error('Cannot allocate memory for font face');
    case -9:
      return new Error('Font face has no content checksum');
    default:
      return corrupt();
  }
}

function decodeWasm(
  wasm: Decoder,
  bytes: Uint8Array,
  dict: Uint8Array | undefined,
  maxOutput: number,
  expectedSize: number,
): ArrayBuffer {
  let srcPtr = 0;
  let dictPtr = 0;
  try {
    srcPtr = wasm.fp_alloc(bytes.length) >>> 0;
    if (!srcPtr) throw wasmError(-7);
    new Uint8Array(wasm.memory.buffer, srcPtr, bytes.length).set(bytes);
    if (dict?.length) {
      dictPtr = wasm.fp_alloc(dict.length) >>> 0;
      if (!dictPtr) throw wasmError(-7);
      new Uint8Array(wasm.memory.buffer, dictPtr, dict.length).set(dict);
    }
    const code = wasm.fp_decode(srcPtr, bytes.length, dictPtr, dict?.length ?? 0, maxOutput);
    if (code !== 0) throw wasmError(code);
    const len = wasm.fp_out_len() >>> 0;
    if (len !== expectedSize) throw corrupt();
    return copy(new Uint8Array(wasm.memory.buffer, wasm.fp_out_ptr() >>> 0, len));
  } finally {
    wasm.fp_out_free();
    if (dictPtr) wasm.fp_free(dictPtr, dict?.length ?? 0);
    if (srcPtr) wasm.fp_free(srcPtr, bytes.length);
  }
}

export async function decodeFontBytes(
  bytes: ArrayBuffer | Uint8Array,
  dict?: ArrayBuffer | Uint8Array,
  options?: { maxOutput?: number; wasmUrl?: URL; codec?: string },
): Promise<ArrayBuffer> {
  if (options?.codec !== undefined && options.codec !== 'zstd') {
    throw new Error(`Unsupported font face codec: ${options.codec}`);
  }
  const input = view(bytes);
  if (
    input.length < 4 ||
    input[0] !== 0x28 || input[1] !== 0xb5 || input[2] !== 0x2f || input[3] !== 0xfd
  ) {
    return copy(input);
  }
  const maxOutput = options?.maxOutput ?? DEFAULT_MAX_OUTPUT;
  if (!Number.isSafeInteger(maxOutput) || maxOutput < 0 || maxOutput > 0x7fffffff) {
    throw new Error('Font face maxOutput must be an integer from 0 to 2147483647');
  }
  const expectedSize = contentSize(input, maxOutput);
  const dictionary = dict === undefined ? undefined : view(dict);
  const zlib = options?.wasmUrl ? undefined : builtinModule<NodeZlibLike>('node:zlib');
  if (typeof zlib?.zstdDecompressSync === 'function') {
    try {
      const buffer = builtinModule<NodeBufferLike>('node:buffer');
      const output = zlib.zstdDecompressSync(input, {
        dictionary: dictionary === undefined ? undefined : buffer?.Buffer.from(dictionary) ?? dictionary,
        maxOutputLength: Math.max(1, maxOutput),
      });
      if (output.byteLength !== expectedSize) throw corrupt();
      return copy(output);
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      throw new Error(`Font face is corrupt or uses a wrong dictionary: ${detail}`);
    }
  }
  const wasm = await decoder(options?.wasmUrl ?? new URL('../dist/decoder/fontpack.wasm', import.meta.url));
  return decodeWasm(wasm, input, dictionary, maxOutput, expectedSize);
}
