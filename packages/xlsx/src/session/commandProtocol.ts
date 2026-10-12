import type { XlsxRangeTarget } from '../edits';
import type {
  CommandAck as SharedCommandAck, CommandConnection, CommandEnvelope as SharedCommandEnvelope,
  CommandOutcome, CommandTicket, EchoHooks, FrameEnvelope, SubmitOptions, TicketOutcome,
} from '../../../../shared/office-session/commands/types';
import type { CalculationStatus, CellPoint, MergedRange, SelectionFormatting, Viewport, WorkbookHandle } from '../wasm/loader';
import type { WorkbookFrame, WorkbookSheetSummary, WorkbookWireFrame } from './methods';
import { validWorkbookToolbarArgs } from './commandToolbar';
import { validateWorkbookReplayEnvelope, type WorkbookReplayMethod, type WorkbookReplayOp } from './replay';

export interface CommandIdentity {
  session: string;
  incarnation: string;
}

export interface CommandBasis {
  epoch: number;
  version: string;
}

export interface CommandEnvelope<K extends string, P> extends CommandIdentity {
  seq: number;
  basis: CommandBasis;
  afterSeq?: number;
  kind: K;
  payload: P;
}

export interface CommandStatus<R = unknown> {
  seq: number;
  status: 'applied' | 'noop' | 'refused' | 'cancelled';
  reason?: string;
  committed?: boolean;
  affectedSeqs?: number[];
  result?: R;
}

export interface CommandAck<S, E, F> {
  throughSeq: number;
  version: string;
  selection?: S;
  editState?: Partial<E>;
  statuses: CommandStatus[];
  frame?: F;
}

export interface WorkbookCommandSelection extends XlsxRangeTarget {
  activeCell?: CellPoint;
}

export interface WorkbookCommandContext extends CommandIdentity {
  basis: CommandBasis;
}

interface WorkbookIntent {
  inputTime?: number;
  undoGroupId?: string;
  selection?: WorkbookCommandSelection;
}

type CommandArgs<K extends WorkbookReplayMethod> = K extends 'propose'
  ? [string, string | null, (Parameters<WorkbookHandle['propose']>[2][number] & { target?: XlsxRangeTarget })[]]
  : Extract<WorkbookReplayOp, { method: K }>['args'];

type ReplayPayload<K extends WorkbookReplayMethod> = Extract<WorkbookReplayOp, { method: K }> extends infer O
  ? O extends WorkbookReplayOp ? Pick<O, 'calculation'> & { args: CommandArgs<K> } & WorkbookIntent &
    { target?: XlsxRangeTarget } & (K extends 'setActiveSheet' ? { expectVersion?: string } : {}) : never
  : never;

export type WorkbookCommandKind = WorkbookReplayMethod | 'select' | 'setReadOnly' | 'applyUpdate';

export type WorkbookToolbarPayload = {
  id: 'bold' | 'italic' | 'strikethrough' | 'paintFormat' | 'clear' | 'fontFamily' | 'fontSize' |
    'fontSizeStep' | 'decimalPlaces' | 'textColor' | 'fillColor' | 'borderColor' | 'numberFormat' |
    'borderPreset' | 'borderStyle' | 'merge' | 'horizontalAlignment' | 'verticalAlignment' | 'textWrapping';
  args: unknown;
  target: XlsxRangeTarget;
} | { id: 'refreshProposals'; args: unknown; target?: never };

export interface WorkbookFormatTokenResult {
  formatToken: string;
}

export type WorkbookCommandPayload<K extends WorkbookCommandKind> = K extends WorkbookReplayMethod
  ? ReplayPayload<K> | (K extends 'applyOps' ? WorkbookIntent & { toolbar: WorkbookToolbarPayload } : never) |
  (K extends 'applyFormat' ? WorkbookIntent & { target: XlsxRangeTarget; formatToken: string } | WorkbookIntent & {
    target: XlsxRangeTarget; capturedFrom: XlsxRangeTarget;
  } : never)
  : K extends 'select' ? { selection: WorkbookCommandSelection }
  : K extends 'applyUpdate' ? { update: Uint8Array }
  : { readOnly: boolean };

export interface WorkbookCommandFrameRequest {
  sheet: number;
  viewport: Viewport;
  zoom: number;
  dpr: number;
  generation: number;
  requestGeneration?: number;
}

export type WorkbookCommand = {
  [K in WorkbookCommandKind]: CommandEnvelope<K, WorkbookCommandPayload<K>> & {
    frame?: WorkbookCommandFrameRequest;
  }
}[WorkbookCommandKind];

export interface WorkbookEditState {
  historyReset?: boolean;
  canUndo: boolean;
  canRedo: boolean;
  formatting: SelectionFormatting;
  formattingAvailable: boolean;
  selectionEditable: boolean;
  mergedRanges: MergedRange[];
  calculation: CalculationStatus;
  sheets: WorkbookSheetSummary[];
  activeSheet: number;
  readOnly: boolean;
}

interface CommandFrameTags extends CommandIdentity {
  coversThroughSeq: number;
  requestGeneration: number;
  zoom: number;
  dpr: number;
}

export type WorkbookCommandFrame = WorkbookFrame & CommandFrameTags;
export type WorkbookCommandWireFrame = WorkbookWireFrame & CommandFrameTags;
export type WorkbookAck = CommandAck<WorkbookCommandSelection, WorkbookEditState, WorkbookCommandFrame>;
export type WorkbookWireAck = CommandAck<WorkbookCommandSelection, WorkbookEditState, WorkbookCommandWireFrame> & {
  shared?: WorkbookSharedWireAck;
};

export interface WorkbookEditPayload {
  payload: WorkbookCommand['payload'];
  frame?: WorkbookCommandFrameRequest;
  compatibility?: WorkbookCommand;
  dispatchOrder?: number;
}

export type WorkbookSharedCommand = SharedCommandEnvelope<WorkbookEditPayload>;
export type WorkbookSharedOutcome = CommandOutcome & { workbook?: CommandStatus };
export type WorkbookSharedAck<F = WorkbookCommandFrame> = Omit<SharedCommandAck<
  WorkbookCommandSelection, Partial<WorkbookEditState>, F
>, 'outcomes'> & { engineVersion: string; outcomes: WorkbookSharedOutcome[];
  compatibility?: CommandAck<WorkbookCommandSelection, WorkbookEditState, F> };
export type WorkbookSharedWireAck = WorkbookSharedAck<WorkbookCommandWireFrame>;
export type WorkbookEditConnection = Omit<CommandConnection<
  WorkbookEditPayload, WorkbookCommandSelection, Partial<WorkbookEditState>, WorkbookCommandFrame
>, 'initial' | 'client'> & {
  initial: WorkbookSharedAck;
  client: Omit<CommandConnection<WorkbookEditPayload, WorkbookCommandSelection,
    Partial<WorkbookEditState>, WorkbookCommandFrame>['client'], 'call'> & {
    call: {
      edit(command: WorkbookSharedCommand): Promise<WorkbookSharedAck>;
      checkpoint(): Promise<import('../../../../shared/office-session/commands/types').Checkpoint>;
    };
  };
};
export type WorkbookEditFrame = FrameEnvelope<WorkbookCommandFrame>;
export type WorkbookTicketOutcome = TicketOutcome & { workbook?: CommandStatus };
export type WorkbookCommandTicket = Omit<CommandTicket, 'ack'> & { readonly ack: Promise<WorkbookTicketOutcome> };
export type WorkbookSubmitOptions<X = unknown> = SubmitOptions<X>;
export type WorkbookEchoHooks<X = unknown> = EchoHooks<X, WorkbookCommandFrame>;

function record(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== 'object') return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === null || prototype === Object.prototype;
}

function integer(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function finite(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

function batchValue(value: unknown): boolean {
  let remaining = 1_000_000;
  let textRemaining = 16_777_216;
  const parents = new Set<unknown>();
  const visit = (field: unknown, depth: number): boolean => {
    if (--remaining < 0 || depth > 64) return false;
    if (typeof field === 'string') {
      textRemaining -= field.length;
      return textRemaining >= 0;
    }
    if (field === null || field === undefined || typeof field === 'boolean') return true;
    if (typeof field === 'number') return finite(field);
    if ((!Array.isArray(field) && !record(field)) || parents.has(field)) return false;
    parents.add(field);
    let valid = true;
    if (Array.isArray(field)) {
      if (field.length > remaining) valid = false;
      else for (let index = 0; index < field.length; index++) {
        if (!Object.prototype.hasOwnProperty.call(field, index) || !visit(field[index], depth + 1)) { valid = false; break; }
      }
    } else {
      for (const key in field) {
        textRemaining -= key.length;
        if (textRemaining < 0 || !visit(field[key], depth + 1)) { valid = false; break; }
      }
    }
    parents.delete(field);
    return valid;
  };
  return visit(value, 0);
}

function point(value: unknown): boolean {
  return record(value) && integer(value.row) && integer(value.col);
}

function target(value: unknown): boolean {
  return record(value) && typeof value.sheetId === 'string' && record(value.range) &&
    (value.range.kind === 'a1' ? typeof value.range.a1 === 'string' :
      value.range.kind === 'rowCol' && point(value.range.start) && point(value.range.end));
}

function selection(value: unknown): boolean {
  return target(value) && record(value) && (value.activeCell === undefined || point(value.activeCell));
}

export function isWorkbookCommandFrameRequest(value: unknown): value is WorkbookCommandFrameRequest {
  return record(value) && integer(value.sheet) && integer(value.generation) &&
    (value.requestGeneration === undefined || integer(value.requestGeneration)) &&
    finite(value.zoom) && value.zoom > 0 && finite(value.dpr) && value.dpr > 0 &&
    record(value.viewport) && finite(value.viewport.x) && finite(value.viewport.y) &&
    finite(value.viewport.width) && value.viewport.width > 0 &&
    finite(value.viewport.height) && value.viewport.height > 0;
}

export function validateWorkbookCommandEnvelope(value: unknown): asserts value is WorkbookCommand {
  if (!record(value) || typeof value.session !== 'string' || typeof value.incarnation !== 'string' ||
    !integer(value.seq) || value.seq === 0 || !record(value.basis) || !integer(value.basis.epoch) ||
    typeof value.basis.version !== 'string' || (value.afterSeq !== undefined &&
      (!integer(value.afterSeq) || value.afterSeq === 0 || value.afterSeq >= value.seq))) {
    throw new TypeError('Malformed workbook command envelope');
  }
}

export function validateWorkbookCommandPayload(command: WorkbookCommand): void {
  const payload: unknown = command.payload;
  const { frame } = command;
  let valid = record(payload) && (payload.selection === undefined || selection(payload.selection));
  if (valid && record(payload)) {
    valid = (payload.inputTime === undefined || finite(payload.inputTime)) &&
      (payload.undoGroupId === undefined || typeof payload.undoGroupId === 'string') &&
      (payload.target === undefined || target(payload.target));
    if (command.kind === 'setActiveSheet' && payload.expectVersion !== undefined) {
      valid &&= typeof payload.expectVersion === 'string';
    }
    if (command.kind === 'select') valid &&= selection(payload.selection);
    else if (command.kind === 'applyUpdate') valid &&= payload.update instanceof Uint8Array;
    else if (command.kind === 'setReadOnly') valid &&= typeof payload.readOnly === 'boolean';
    else if (command.kind === 'applyOps' && 'toolbar' in payload) {
      valid &&= !('args' in payload) && record(payload.toolbar) &&
        (payload.toolbar.id === 'refreshProposals' ? payload.toolbar.target === undefined : target(payload.toolbar.target)) &&
        validWorkbookToolbarArgs(payload.toolbar.id, payload.toolbar.args);
    } else if (command.kind === 'applyFormat' && 'formatToken' in payload) {
      valid &&= !('args' in payload) && !('capturedFrom' in payload) && target(payload.target) &&
        typeof payload.formatToken === 'string' && payload.formatToken.length > 0;
    } else if (command.kind === 'applyFormat' && 'capturedFrom' in payload) {
      valid &&= target(payload.target) && target(payload.capturedFrom);
    } else {
      if (command.kind === 'applyOps' || command.kind === 'propose') {
        const entries = Array.isArray(payload.args) ? payload.args[command.kind === 'applyOps' ? 0 : 2] : undefined;
        valid &&= Array.isArray(entries) && entries.every((entry) => record(entry) &&
          (entry.target === undefined || target(entry.target)));
      }
      try {
        if (command.kind === 'applyEdits' && !batchValue(payload.args)) throw new WorkbookCommandRefusal('invalid-command');
        validateWorkbookReplayEnvelope({ sequence: command.seq,
          calculation: payload.calculation ?? { nowSerial: 0, randSeed: 0 },
          op: { method: command.kind, args: payload.args } });
      } catch (error) {
        if (command.kind !== 'applyEdits' && (!(error instanceof Error) || error.name !== 'WorkbookReplayValidationError')) throw error;
        valid = false;
      }
    }
  }
  if (frame !== undefined) {
    valid &&= isWorkbookCommandFrameRequest(frame);
  }
  if (!valid) throw new WorkbookCommandRefusal('invalid-command');
}

export class WorkbookCommandRefusal extends Error {
  constructor(readonly reason: string, readonly result?: unknown) {
    super(reason);
    this.name = 'WorkbookCommandRefusal';
  }
}
