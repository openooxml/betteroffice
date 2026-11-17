import type { FontPackDiagnostic } from '../../../../shared/fontPackDiagnostics';
export {
  default,
  decodeFontFace,
  decodeTiffPng,
  exportPptxMarkdownJson,
  exportPptxStructuredJson,
  parsePptxJson,
  parsePptxJsonWithoutMedia,
  PptxDocument,
  PptxRenderer,
  renderPptxMarkdownJson,
  rendererVersion,
} from './generated/pptx_wasm.js';
export type { InitInput } from './generated/pptx_wasm.js';

/** Which engine this build links; the native package builds these sources with its own engine module. */
export const engineKind: 'wasm' | 'native' = 'wasm';

/** The metafile SVG converter's bindings, imported only when a session needs them. */
export function metafileSvgBindings(): Promise<typeof import('./generated/metafile_svg_wasm')> {
  return import('./generated/metafile_svg_wasm');
}

/** The installed font add-on packages that cannot be used, and why; their faces are substituted. */
export async function fontPackDiagnostics(): Promise<FontPackDiagnostic[]> {
  return (await import('@betteroffice/fonts')).checkFontPacks();
}
