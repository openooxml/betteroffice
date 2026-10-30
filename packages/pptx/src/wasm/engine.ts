export {
  default,
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

/** Which engine this build links; @betteroffice/pptx-native builds these sources with its own engine module. */
export const engineKind: 'wasm' | 'native' = 'wasm';
