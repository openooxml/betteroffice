export type FontDecoder = (
  bytes: ArrayBuffer | Uint8Array,
  options: { codec: string; dict?: ArrayBuffer | Uint8Array; maxOutput?: number },
) => Promise<{ font: ArrayBuffer; payload?: ArrayBuffer }>;

export type FontDecodeResult = Awaited<ReturnType<FontDecoder>>;

export type FontDecoderMethods = {
  decodeFontFace(bytes: ArrayBuffer, options: Parameters<FontDecoder>[1], wasm?: ArrayBuffer | WebAssembly.Module): FontDecodeResult;
};

export function rpcFontDecoder(call: (bytes: ArrayBuffer, options: Parameters<FontDecoder>[1], transfer: Transferable[]) => Promise<FontDecodeResult>): FontDecoder {
  return (bytes, options) => {
    const encoded = fontBytes(bytes).buffer;
    const dict = options.dict === undefined ? undefined : fontBytes(options.dict).buffer;
    return call(encoded, { ...options, dict }, dict === undefined ? [encoded] : [encoded, dict]);
  };
}

export function fontBytes(bytes: ArrayBuffer | Uint8Array): Uint8Array<ArrayBuffer> {
  return (bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)).slice();
}

export function copyDecodedFace(face: { font: Uint8Array; payload?: Uint8Array }): FontDecodeResult {
  return { font: fontBytes(face.font).buffer,
    ...(face.payload === undefined ? {} : { payload: fontBytes(face.payload).buffer }) };
}
