import type { DefaultFontFace } from '../../../../shared/pptxDefaultFonts';
import type { PresentationHandle } from './loader';
import { presentationContexts } from './presentationContext';

export function presentationRequestedFontFaces(
  handle: PresentationHandle, slideId?: string,
): Omit<DefaultFontFace, 'bytes'>[] {
  const context = presentationContexts.get(handle);
  if (!context) throw new Error('Presentation font requests are unavailable');
  return context.jsonWasmCall<{ faces: Omit<DefaultFontFace, 'bytes'>[] }>(
    () => context.renderer.requestedFontFacesJson(context.doc, slideId)).faces;
}
