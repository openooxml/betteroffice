import type { FontDecoder } from '../../../../shared/office-session/fontDecoder';
import type { PptxFontFace } from '../types';

export type PresentationFontPreparation = (
  runtime: { decoder: FontDecoder }, families: () => Promise<string[]>,
) => Promise<readonly PptxFontFace[] | undefined>;

export const presentationFontPreparations = new WeakMap<object, PresentationFontPreparation>();

export function bindPresentationFontPreparation<T extends object>(options: T, prepare: PresentationFontPreparation): T {
  presentationFontPreparations.set(options, prepare);
  return options;
}
