import { missingWorkerRenderCapabilities } from './workerCapabilities';

/** Chromium and Firefox support slices. WebKit and unknown engines use whole faces. */
export function slicedFacesSupported(
  userAgent = (globalThis as { navigator?: { userAgent?: string } }).navigator?.userAgent ?? '',
): boolean {
  return /(?:Chrome|Chromium|Firefox)\//.test(userAgent) && !/\b(?:iPhone|iPad|iPod|CriOS|FxiOS)\b/.test(userAgent);
}

export function shouldSlicePresentationFonts(sliced?: boolean): boolean {
  return sliced !== false && slicedFacesSupported() && missingWorkerRenderCapabilities().length === 0;
}
