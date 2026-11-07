import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import type { CommandStatus } from '@betteroffice/xlsx/worker';

/** What a refused structural edit would have broken, as the refusal banner names it. */
export type XlsxRefusalCause = 'chart' | 'table' | 'pivot' | 'other';

const CAUSES: ReadonlySet<string> = new Set<XlsxRefusalCause>(['chart', 'table', 'pivot', 'other']);

export const REFUSAL_CAUSE_KEYS: Record<XlsxRefusalCause, TranslationKey> = {
  chart: 'commands.reasons.editUnsupportedChart',
  table: 'commands.reasons.editUnsupportedTable',
  pivot: 'commands.reasons.editUnsupportedPivot',
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

/** The worker commands that change the workbook, whose refusal the banner reports as a declined edit. */
const EDIT_COMMANDS: ReadonlySet<string> = new Set([
  'editCell', 'editCells', 'applyEdits', 'applyOps', 'patchRangeStyle', 'setNumberFormat', 'applyFormat',
  'moveChart', 'moveRange', 'undo', 'redo', 'acceptProposal', 'rejectProposal',
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
