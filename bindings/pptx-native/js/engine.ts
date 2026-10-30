import { createRequire } from 'node:module';
import type * as Wasm from '../../../packages/pptx/src/wasm/generated/pptx_wasm.js';

export type { InitInput } from '../../../packages/pptx/src/wasm/generated/pptx_wasm.js';

type Engine = Pick<
  typeof Wasm,
  | 'decodeFontFace'
  | 'decodeTiffPng'
  | 'exportPptxMarkdownJson'
  | 'exportPptxStructuredJson'
  | 'parsePptxJson'
  | 'parsePptxJsonWithoutMedia'
  | 'PptxDocument'
  | 'PptxRenderer'
  | 'renderPptxMarkdownJson'
  | 'rendererVersion'
>;

const binding = createRequire(import.meta.url)('../index.js') as Engine;

export const decodeFontFace = binding.decodeFontFace;
export const decodeTiffPng = binding.decodeTiffPng;
export const exportPptxMarkdownJson = binding.exportPptxMarkdownJson;
export const exportPptxStructuredJson = binding.exportPptxStructuredJson;
export const parsePptxJson = binding.parsePptxJson;
export const parsePptxJsonWithoutMedia = binding.parsePptxJsonWithoutMedia;
export const PptxDocument = binding.PptxDocument;
export type PptxDocument = Wasm.PptxDocument;
export const PptxRenderer = binding.PptxRenderer;
export type PptxRenderer = Wasm.PptxRenderer;
export const renderPptxMarkdownJson = binding.renderPptxMarkdownJson;
export const rendererVersion = binding.rendererVersion;

export const engineKind: 'wasm' | 'native' = 'native';

async function initNativeEngine(): Promise<{ memory: undefined }> {
  return { memory: undefined };
}

export default initNativeEngine as unknown as typeof Wasm.default;
