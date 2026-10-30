import { createRequire } from 'node:module';
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

const binding = createRequire(import.meta.url)('../index.js') as Engine;

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
