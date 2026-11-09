import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { warnFontPackDiagnostics, type FontPackDiagnostic } from '../../../shared/fontPackDiagnostics';
import type * as Wasm from '../../../packages/xlsx/src/wasm/generated/xlsx_wasm.js';

export type { InitInput } from '../../../packages/xlsx/src/wasm/generated/xlsx_wasm.js';

type Engine = Pick<
  typeof Wasm,
  | 'XlsxDocument'
  | 'XlsxSnapshotBuilder'
  | 'allocationFailure'
  | 'decodeFontFace'
  | 'displayListCellLimit'
  | 'exportXlsxMarkdownJson'
  | 'exportXlsxStructuredJson'
  | 'renderXlsxMarkdownJson'
>;

type NativeFontLoader = {
  configureFontPacks(latin?: string, intl?: string, cjk?: string): boolean;
  fontPackDiagnosticsJson(): string;
};

const require = createRequire(import.meta.url);
const binding = require('../index.js') as Engine & NativeFontLoader;

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

export const XlsxDocument = binding.XlsxDocument;
export type XlsxDocument = Wasm.XlsxDocument;
export const XlsxSnapshotBuilder = binding.XlsxSnapshotBuilder;
export type XlsxSnapshotBuilder = Wasm.XlsxSnapshotBuilder;
export const allocationFailure = binding.allocationFailure;
export const decodeFontFace = binding.decodeFontFace;
export const displayListCellLimit = binding.displayListCellLimit;
export const exportXlsxMarkdownJson = binding.exportXlsxMarkdownJson;
export const exportXlsxStructuredJson = binding.exportXlsxStructuredJson;
export const renderXlsxMarkdownJson = binding.renderXlsxMarkdownJson;

export const engineKind: 'wasm' | 'native' = 'native';

async function initNativeEngine(): Promise<{ memory: undefined }> {
  return { memory: undefined };
}

export default initNativeEngine as unknown as typeof Wasm.default;
