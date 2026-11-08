import type { FontPackDiagnostic } from '../../../../shared/fontPackDiagnostics';
export {
  default,
  XlsxDocument,
  XlsxSnapshotBuilder,
  allocationFailure,
  decodeFontFace,
  displayListCellLimit,
  exportXlsxMarkdownJson,
  exportXlsxStructuredJson,
  renderXlsxMarkdownJson,
} from './generated/xlsx_wasm.js';
export type { InitInput } from './generated/xlsx_wasm.js';

/** Which engine this build links; the native package builds these sources with its own engine module. */
export const engineKind: 'wasm' | 'native' = 'wasm';

/** The installed font add-on packages that cannot be used, and why; their faces are substituted. */
export async function fontPackDiagnostics(): Promise<FontPackDiagnostic[]> {
  return (await import('@betteroffice/fonts')).checkFontPacks();
}
