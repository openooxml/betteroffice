export type PngExportCode = 'canvas-unavailable' | 'canvas-too-large' | 'font-unavailable' | 'picture-unavailable';

export class PngExportError extends Error {
  readonly code: PngExportCode;
  constructor(code: PngExportCode, message: string) {
    super(message);
    this.name = 'PngExportError';
    this.code = code;
  }
}

/** Why a picture was left as a placeholder. */
export type PngPictureUnavailableReason = 'missing' | 'unsupported' | 'decode-failed' | 'metafile' | 'too-large' | 'budget';

/** A picture a PNG export drew as a placeholder. `limit` and `actual` say by how much a size or budget limit was passed. */
export interface PngUnavailablePicture {
  id: string;
  reason: PngPictureUnavailableReason;
  limit?: number;
  actual?: number;
  unit?: 'pixels' | 'bytes' | 'work';
}

/** A problem a PNG export worked around: the page still exports. */
export interface PngDiagnostic {
  code: 'picture-unavailable';
  pictures: PngUnavailablePicture[];
}

/**
 * Whether a PNG export rejects instead of drawing a placeholder when a picture
 * cannot be drawn, when the caller does not say: every format's
 * `strictPictures` defaults to this.
 */
export const DEFAULT_STRICT_PICTURES = false;

/** The rejection `strictPictures` asks for, naming each picture that could not be drawn. */
export function strictPicturesError(pictures: readonly PngUnavailablePicture[]): PngExportError {
  const named = pictures.map((picture) => `${picture.id} (${picture.reason})`).join(', ');
  return new PngExportError('picture-unavailable', `${pictures.length} picture(s) on this page could not be drawn: ${named}`);
}
