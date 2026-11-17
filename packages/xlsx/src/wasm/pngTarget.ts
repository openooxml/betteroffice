import type { Viewport } from './loader';

/** A sheet by zero-based index or by tab name. */
export type XlsxSheetRef = number | string;

/** A cell range as an A1 string, an A1 object, or zero-based row and column corners. */
export type XlsxRangeAddress =
  | string
  | { kind: 'a1'; a1: string }
  | { kind: 'rowCol'; start: { row: number; col: number }; end: { row: number; col: number } };

/**
 * What `renderPng` draws: a viewport of a sheet, or a cell range of it.
 * `range: 'used'` draws the sheet's used range as the engine finds it at call time.
 */
export type XlsxPngTarget =
  | { sheet: number | string; viewport: Viewport }
  | { sheet: number | string; range: XlsxRangeAddress | 'used' };

/** A font face an export paints with, the same shape as `VsdxFontFace`. */
export interface XlsxFontFace {
  family: string;
  bold?: boolean;
  italic?: boolean;
  bytes: Uint8Array;
}

/**
 * A family and style a PNG export paints, for callers that supply faces.
 * @internal
 */
export interface XlsxPngFont {
  family: string;
  bold: boolean;
  italic: boolean;
}

export interface XlsxPngOptions {
  /** Device pixels per sheet pixel. Defaults to 1. */
  scale?: number;
  /** Largest PNG width in pixels; a wider target is cut at its right edge. */
  maxWidth?: number;
  /** Largest PNG height in pixels; a taller target is cut at its bottom edge. */
  maxHeight?: number;
  signal?: AbortSignal;
  /**
   * Faces to paint the sheet's fonts with, matched by family and style. Each
   * export installs them under private names and drops them afterwards; one
   * that fails to load rejects with `PngExportError` `font-unavailable`. The
   * editors pass the faces they paint with. Without this option the export
   * loads the bundled faces `@betteroffice/xlsx-native` draws the sheet's fonts
   * with, and rejects if one fails to load. A list, even an empty one, replaces
   * them: a family it does not cover is painted with the engine's built-in
   * substitute.
   */
  fonts?: ReadonlyArray<XlsxFontFace>;
}
