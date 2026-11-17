import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { warnFontPackDiagnostics, type FontPackDiagnostic } from '../../../shared/fontPackDiagnostics';
import { installDispose } from '../../../shared/nativeDispose';
import type * as Wasm from '../../../packages/pptx/src/wasm/generated/pptx_wasm.js';
import type { PptxNativeExtras } from '../../native-extras-types';

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

type NativeFontLoader = Pick<PptxNativeExtras['functions'], 'configureFontPacks' | 'fontPackDiagnosticsJson'>;

const require = createRequire(import.meta.url);
const binding = require('../index.js') as Engine & NativeFontLoader;
installDispose(binding.PptxDocument, binding.PptxRenderer);

function packDirectory(pack: 'latin' | 'intl' | 'cjk'): string | undefined {
  try {
    return join(dirname(require.resolve(`@betteroffice/fonts-${pack}/package.json`)), 'assets');
  } catch {
    return undefined;
  }
}

binding.configureFontPacks(packDirectory('latin'), packDirectory('intl'), packDirectory('cjk'));
const packDiagnostics = JSON.parse(binding.fontPackDiagnosticsJson()) as FontPackDiagnostic[];
warnFontPackDiagnostics(packDiagnostics);

/** The installed font add-on packages that cannot be used, and why; their faces are substituted. */
export async function fontPackDiagnostics(): Promise<FontPackDiagnostic[]> {
  return packDiagnostics.map((diagnostic) => ({ ...diagnostic }));
}

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

export function metafileSvgBindings(): Promise<never> {
  return Promise.reject(new Error('The native engine converts metafiles itself'));
}

async function initNativeEngine(): Promise<{ memory: undefined }> {
  return { memory: undefined };
}

export default initNativeEngine as unknown as typeof Wasm.default;
