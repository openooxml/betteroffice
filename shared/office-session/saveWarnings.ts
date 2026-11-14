/**
 * `unwritable-reference`: the document holds a reference with a character XML 1.0 cannot carry, which the
 * part cannot be written with, so the save kept the part as the file wrote it. `unrewritable-formula`: the part
 * holds a formula, kept as the file wrote it, that a sheet rename could not rewrite, so it was kept unrenamed.
 * New codes may be added; handle a code you do not know as a generic warning.
 */
export type SaveWarningCode = 'unwritable-reference' | 'unrewritable-formula';

/** A part a save wrote differently from what the document holds, and why. */
export interface SaveWarning {
  /** Stable, to match on. */
  code: SaveWarningCode;
  /** The package part, such as `xl/charts/chart1.xml`. */
  part: string;
  /** What inside the part, such as `chart`. */
  element?: string;
  /** The sheet the part belongs to, in a workbook. */
  sheet?: string;
  /** A sentence for people; not stable, match on `code`. */
  message: string;
}

/** The saved bytes and the warnings the save carries; `warnings` is empty on an ordinary save. */
export interface SaveResult {
  bytes: Uint8Array;
  warnings: SaveWarning[];
}

/**
 * Emits each warning as a Node process warning with code `BETTEROFFICE_SAVE_WARNING`, so server hosts can
 * filter it. Does nothing outside Node.
 */
export function warnSaveWarnings(warnings: readonly SaveWarning[]): void {
  if (warnings.length === 0) return;
  const proc = (globalThis as {
    process?: { emitWarning?: (warning: string, options: { type: string; code: string }) => void };
  }).process;
  if (typeof proc?.emitWarning !== 'function') return;
  for (const { message } of warnings) {
    proc.emitWarning(message, { type: 'BetterOfficeSaveWarning', code: 'BETTEROFFICE_SAVE_WARNING' });
  }
}
