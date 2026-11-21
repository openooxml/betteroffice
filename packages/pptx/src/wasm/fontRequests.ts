import type { RequestedFontFace } from '../../../../shared/pptxDefaultFonts';
import type { PresentationHandle } from './loader';
import { presentationContexts } from './presentationContext';

export function presentationEnableBrowserCoverageFallback(handle: PresentationHandle): void {
  const context = presentationContexts.get(handle);
  if (!context) throw new Error('Presentation browser coverage fallback is unavailable');
  context.wasmCall(() => context.renderer.setBrowserCoverageFallback(true));
}

export function presentationRequestedFontFaces(
  handle: PresentationHandle, slideId?: string,
): RequestedFontFace[] {
  const context = presentationContexts.get(handle);
  if (!context) throw new Error('Presentation font requests are unavailable');
  return context.jsonWasmCall<{ faces: RequestedFontFace[] }>(
    () => context.renderer.requestedFontFacesJson(context.doc, slideId)).faces;
}
