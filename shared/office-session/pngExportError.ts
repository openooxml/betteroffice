export type PngExportCode = 'canvas-unavailable' | 'canvas-too-large' | 'font-unavailable';

export class PngExportError extends Error {
  readonly code: PngExportCode;
  constructor(code: PngExportCode, message: string) {
    super(message);
    this.name = 'PngExportError';
    this.code = code;
  }
}
