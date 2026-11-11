import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import type { CommandStatus } from '@betteroffice/xlsx/worker';

/** What a refused structural edit would have broken, as the refusal banner names it. */
export type XlsxRefusalCause = 'chart' | 'table' | 'pivot' | 'fill' | 'other';

const CAUSES: ReadonlySet<string> = new Set<XlsxRefusalCause>(['chart', 'table', 'pivot', 'fill', 'other']);

export const REFUSAL_CAUSE_KEYS: Record<XlsxRefusalCause, TranslationKey> = {
  chart: 'commands.reasons.editUnsupportedChart',
  table: 'commands.reasons.editUnsupportedTable',
  pivot: 'commands.reasons.editUnsupportedPivot',
  fill: 'commands.reasons.editUnsupportedFill',
  other: 'commands.reasons.editUnsupported',
};

/**
 * The cause a refused command's status names. The workbook session attaches it to the
 * structural refusals it reads from the engine; any other refusal is `other`.
 */
export function refusalCause(status: CommandStatus): XlsxRefusalCause {
  const cause = (status.result as { failure?: { cause?: unknown } } | undefined)?.failure?.cause;
  return typeof cause === 'string' && CAUSES.has(cause) ? cause as XlsxRefusalCause : 'other';
}

/** What a structural refusal would have broken, read from the engine's own message. */
function structuralCause(message: string): XlsxRefusalCause {
  if (/^(?:chart |anchor \d+ of )/.test(message) || message === 'range move cannot safely preserve charts') return 'chart';
  if (message === 'range move cannot safely preserve Excel tables' ||
    /^cannot move sheet \d+ tracks: table .+ covers the moved tracks$/.test(message)) return 'table';
  const part = /^(.+?) references cells this (?:edit would move|move would relocate), and it cannot be rewritten$/
    .exec(message)?.[1].toLowerCase().replace(/^\//, '');
  if (!part) return 'other';
  if (part.startsWith('xl/charts/') || part.startsWith('xl/drawings/')) return 'chart';
  if (part.startsWith('xl/tables/')) return 'table';
  if (part.startsWith('xl/pivottables/') || part.startsWith('xl/pivotcache/')) return 'pivot';
  return 'other';
}

/** The engine refusals `applyOps` reports for a structural edit it will not perform. */
const APPLY_OPS_REFUSALS: readonly RegExp[] = [
  /^formula at sheet \d+, [A-Z]+[1-9]\d* cannot be safely rewritten$/,
  /^(?:defined name|chart reference in) .+ cannot be safely rewritten$/,
  /^.+ references cells this edit would move, and it cannot be rewritten$/,
  /^anchor \d+ of .+ is held by two sheets that disagree on where it sits$/,
  /^chart .+ cannot change how it follows a grid edit$/,
  /^chart .+ cannot be resized: a one-cell extent is not written back$/,
  /^chart .+ cannot change its anchor kind$/,
  /^cannot move sheet \d+ tracks: .+$/,
];

/** The engine refusals `moveRange` reports for a range move it will not perform. */
const MOVE_RANGE_REFUSALS: readonly RegExp[] = [
  /^range move cannot safely preserve .+$/,
  /^.+ contains worksheet features a cell move cannot safely preserve$/,
  /^.+ references cells this move would relocate, and it cannot be rewritten$/,
];

/**
 * The structural refusal an engine error names, for a caller holding the raw error instead of a
 * command status. These are the workbook session's `classifyRefusal` patterns for structural edits;
 * only those messages match, so an unexpected failure still reads as a failure.
 */
export function engineStructuralRefusal(
  error: unknown, method?: 'applyOps' | 'moveRange'
): { cause: XlsxRefusalCause; message: string } | null {
  if (!(error instanceof Error)) return null;
  const patterns = method === 'applyOps' ? APPLY_OPS_REFUSALS : method === 'moveRange'
    ? MOVE_RANGE_REFUSALS : [...APPLY_OPS_REFUSALS, ...MOVE_RANGE_REFUSALS];
  return patterns.some((pattern) => pattern.test(error.message))
    ? { cause: structuralCause(error.message), message: error.message } : null;
}

/** The worker commands that change the workbook, whose refusal the banner reports as a declined edit. */
const EDIT_COMMANDS: ReadonlySet<string> = new Set([
  'editCell', 'editCells', 'applyEdits', 'applyOps', 'patchRangeStyle', 'setNumberFormat', 'applyFormat',
  'moveChart', 'moveRange', 'fillRange', 'undo', 'redo', 'acceptProposal', 'rejectProposal',
]);

/** `applyOps` toolbar commands that read or capture rather than edit. */
const READ_TOOLBAR_COMMANDS: ReadonlySet<string> = new Set(['refreshProposals', 'paintFormat']);

function editsWorkbook(input: { kind: string; payload?: unknown } | null | undefined, draft: unknown): boolean {
  if (!input) return draft !== null && draft !== undefined;
  if (!EDIT_COMMANDS.has(input.kind)) return false;
  const toolbar = (input.payload as { toolbar?: { id?: unknown } } | undefined)?.toolbar;
  return !(input.kind === 'applyOps' && typeof toolbar?.id === 'string' && READ_TOOLBAR_COMMANDS.has(toolbar.id));
}

/**
 * Whether a refused worker command is an edit the workbook declined rather than cell input that
 * could not be written: an `unsupported` refusal of an edit, or any refusal of an edit command that
 * carried no cell draft. Navigation, read-only toggles and background reads keep their input-failure
 * report, and so do cancellations queued behind failed input.
 */
export function refusedEdit(
  status: CommandStatus, draft: unknown, input?: { kind: string; payload?: unknown } | null
): boolean {
  return status.status === 'refused' && editsWorkbook(input, draft) &&
    (status.reason === 'unsupported' || draft === null || draft === undefined);
}
