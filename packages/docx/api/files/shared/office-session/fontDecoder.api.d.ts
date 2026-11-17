// @betteroffice/docx declaration shared/office-session/fontDecoder.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
export type FontDecoder = (bytes: ArrayBuffer | Uint8Array, options: {
    codec: string;
    dict?: ArrayBuffer | Uint8Array;
    maxOutput?: number;
}) => Promise<{
    font: ArrayBuffer;
    payload?: ArrayBuffer;
}>;
export type FontDecodeResult = Awaited<ReturnType<FontDecoder>>;
export type FontDecoderMethods = {
    decodeFontFace(bytes: ArrayBuffer, options: Parameters<FontDecoder>[1], wasm?: ArrayBuffer | WebAssembly.Module): FontDecodeResult;
};
export declare function rpcFontDecoder(call: (bytes: ArrayBuffer, options: Parameters<FontDecoder>[1], transfer: Transferable[]) => Promise<FontDecodeResult>): FontDecoder;
export declare function fontBytes(bytes: ArrayBuffer | Uint8Array): Uint8Array<ArrayBuffer>;
export declare function copyDecodedFace(face: {
    font: Uint8Array;
    payload?: Uint8Array;
}): FontDecodeResult;
