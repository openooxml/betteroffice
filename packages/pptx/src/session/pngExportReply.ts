import { WorkerFontModuleUnavailableError } from '../../../../shared/office-session/optionalFontPacks';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import { deferReply, transferable } from '../../../../shared/office-session/protocol';

export function pngExportReply(rendered: Promise<Uint8Array>) {
  return deferReply<ArrayBuffer>(rendered.then((bytes) => {
    const buffer = bytes.buffer instanceof ArrayBuffer && bytes.byteOffset === 0 &&
      bytes.byteLength === bytes.buffer.byteLength ? bytes.buffer : new Uint8Array(bytes).buffer;
    return transferable(buffer, [buffer]);
  }, (error: unknown) => {
    if (error instanceof WorkerFontModuleUnavailableError) {
      throw new PngExportError('font-unavailable', error.message);
    }
    throw error;
  }));
}
