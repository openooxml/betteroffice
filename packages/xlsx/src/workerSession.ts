/** @experimental */
export { openWorkbookSession } from './session/client';
/** @experimental */
export type { OpenWorkbookSessionOptions, WorkbookSession } from './session/client';
/** @experimental */
export { workbookSessionRecovery, WorkbookRecoveryReplayError } from './session/clientRecovery';
/** @experimental */
export type { WorkbookSessionRecovery } from './session/clientRecovery';
export { createWorkbookEditSession } from './session/editSession';
export { WorkbookPngUnavailableError } from './session/methods';
export { PngExportError } from '../../../shared/office-session/pngExportError';
export type { PngExportCode } from '../../../shared/office-session/pngExportError';
export type { SaveResult, SaveWarning, SaveWarningCode } from './saveWarnings';
export type { XlsxFontFace, XlsxPngFont, XlsxPngOptions, XlsxPngTarget, XlsxRangeAddress, XlsxSheetRef } from './wasm/pngTarget';
export type { WorkbookEditSession } from './session/editSession';
/** @experimental */
export type {
  WorkbookCellGeometry,
  WorkbookCellInputs,
  WorkbookFrame,
  WorkbookFrameOptions,
  WorkbookInitialPaint,
  WorkbookInitialPaintRequest,
  WorkbookSessionOpenReply,
  WorkbookSessionEvents,
  WorkbookSessionMethods,
  WorkbookSessionOpenOptions,
  WorkbookSessionState,
  WorkbookSheetSummary,
  WorkbookSheetView,
  WorkbookPngOptions,
} from './session/methods';
