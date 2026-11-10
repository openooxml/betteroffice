import { PngExportError, type PngUnavailablePicture } from './pngExportError';

/**
 * Whether a PNG export rejects instead of drawing a placeholder when a picture
 * cannot be drawn, when the caller does not say: every format's
 * `strictPictures` defaults to this.
 */
export const DEFAULT_STRICT_PICTURES = false;

/** How the `strictPictures` rejection names a picture whose `id` is empty. */
export const UNNAMED_PICTURE = 'a picture with no part';

/** The rejection `strictPictures` asks for, naming each picture that could not be drawn. */
export function strictPicturesError(pictures: readonly PngUnavailablePicture[]): PngExportError {
  const named = pictures.map((picture) => `${picture.id || UNNAMED_PICTURE} (${picture.reason})`).join(', ');
  return new PngExportError('picture-unavailable', `${pictures.length === 1 ? '1 picture' : `${pictures.length} pictures`} on this page could not be drawn: ${named}`);
}
