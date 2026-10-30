export {
  default,
  XlsxDocument,
  XlsxSnapshotBuilder,
  allocationFailure,
  displayListCellLimit,
  exportXlsxMarkdownJson,
  exportXlsxStructuredJson,
  renderXlsxMarkdownJson,
} from './generated/xlsx_wasm.js';
export type { InitInput } from './generated/xlsx_wasm.js';

/** Which engine this build links; the native package builds these sources with its own engine module. */
export const engineKind: 'wasm' | 'native' = 'wasm';
