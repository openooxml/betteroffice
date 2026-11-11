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
  /** The engine's reference for the picture; empty for one whose relationship has no part, which has nothing to name it by. */
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
