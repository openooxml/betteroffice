import { SessionFailure } from '../../../../shared/office-session/types';
import type { XlsxEditStep, XlsxRangeAddress, XlsxRangeTarget } from '../edits';
import { StaleProposalError, type CapturedFormat, type CellPoint, type SheetInfo, type WorkbookCalculationContext, type WorkbookHandle } from '../wasm/loader';
import {
  validateWorkbookCommandEnvelope, validateWorkbookCommandPayload, WorkbookCommandRefusal,
  type CommandBasis, type CommandIdentity, type CommandStatus, type WorkbookCommand,
  type WorkbookCommandSelection, type WorkbookEditState, type WorkbookFormatTokenResult, type WorkbookWireAck,
} from './commandProtocol';
import { resolveWorkbookToolbar } from './commandToolbar';
import { applyWorkbookReplayOp, workbookReplayRefused, type WorkbookReplayMethod, type WorkbookReplayOp } from './replay';

const STATUS_CACHE_SIZE = 256;
const FORMAT_TOKEN_CACHE_SIZE = 64;
const MAX_VALIDATION_STEPS = 128;
const MAX_VALIDATION_CELLS = 100_000;
const MAX_VALIDATION_TEXT = 1_048_576;
const MAX_VALIDATION_BYTES = 16 * 1024 * 1024;

export function translateWorkbookSheetIds<T>(value: T, from: string[], to: string[], strict = false): T {
  const sheetId = (id: string) => {
    const translated = to[from.indexOf(id)];
    if (strict && translated === undefined) throw new WorkbookCommandRefusal('target-sheet-missing');
    return translated ?? id;
  };
  const translate = (item: unknown): unknown => {
    if (Array.isArray(item)) return item.map(translate);
    if (!item || typeof item !== 'object') return item;
    return Object.fromEntries(Object.entries(item).map(([key, field]) => [key,
      key === 'sheetId' && typeof field === 'string' ? sheetId(field) :
      (key === 'sheetIds' || key === 'changedSheets') && Array.isArray(field) ? field.map(sheetId) :
      translate(field)]));
  };
  return translate(value) as T;
}

function column(col: number): string {
  let text = '';
  for (let at = col + 1; at > 0; at = Math.floor((at - 1) / 26)) {
    text = String.fromCharCode(65 + (at - 1) % 26) + text;
  }
  return text;
}

function cellAddress(cell: CellPoint): string {
  if (!Number.isSafeInteger(cell.row) || cell.row < 0 || cell.row >= 1_048_576 ||
    !Number.isSafeInteger(cell.col) || cell.col < 0 || cell.col >= 16_384) {
    throw new WorkbookCommandRefusal('missing-target');
  }
  return `${column(cell.col)}${cell.row + 1}`;
}

function corners(range: XlsxRangeAddress): { start: CellPoint; end: CellPoint } {
  if (range.kind === 'rowCol') {
    cellAddress(range.start);
    cellAddress(range.end);
    if (range.start.row > range.end.row || range.start.col > range.end.col) {
      throw new WorkbookCommandRefusal('missing-target');
    }
    return { start: range.start, end: range.end };
  }
  const match = /^\$?([A-Za-z]+)\$?([1-9]\d*)(?::\$?([A-Za-z]+)\$?([1-9]\d*))?$/.exec(range.a1);
  if (!match) throw new WorkbookCommandRefusal('missing-target');
  const point = (letters: string, row: string): CellPoint => {
    let col = 0;
    for (const letter of letters.toUpperCase()) col = col * 26 + letter.charCodeAt(0) - 64;
    const cell = { row: Number(row) - 1, col: col - 1 };
    cellAddress(cell);
    return cell;
  };
  return corners({ kind: 'rowCol', start: point(match[1], match[2]),
    end: point(match[3] ?? match[1], match[4] ?? match[2]) });
}

export function resolveCommandTarget(
  handle: WorkbookHandle, target: XlsxRangeTarget, sheetIds = handle.sheetInfo().sheetIds
) {
  const sheet = sheetIds.indexOf(target.sheetId);
  if (sheet < 0) throw new WorkbookCommandRefusal('target-sheet-missing');
  const { start, end } = corners(target.range);
  return { sheet, start, end, range: `${cellAddress(start)}:${cellAddress(end)}` };
}

function validateSelection(handle: WorkbookHandle, state: WorkbookCommandState, selection: WorkbookCommandSelection): void {
  const { start, end } = resolveCommandTarget(handle, selection, state.sheetIds);
  if (selection.activeCell) {
    cellAddress(selection.activeCell);
    if (selection.activeCell.row < start.row || selection.activeCell.row > end.row ||
      selection.activeCell.col < start.col || selection.activeCell.col > end.col) {
      throw new WorkbookCommandRefusal('missing-target');
    }
  }
}

function sheetTarget(handle: WorkbookHandle, sheet: number, range: XlsxRangeAddress,
  sheetIds = handle.sheetInfo().sheetIds): XlsxRangeTarget {
  const sheetId = sheetIds[sheet];
  if (sheetId === undefined) throw new WorkbookCommandRefusal('missing-target');
  const target = { sheetId, range };
  resolveCommandTarget(handle, target, sheetIds);
  return target;
}

function cellTarget(handle: WorkbookHandle, sheet: number, row: number, col: number,
  sheetIds = handle.sheetInfo().sheetIds): XlsxRangeTarget {
  return sheetTarget(handle, sheet, { kind: 'rowCol', start: { row, col }, end: { row, col } }, sheetIds);
}

function checkedBatch(handle: WorkbookHandle, steps: XlsxEditStep[]): void {
  const expectVersion = handle.version();
  const encoder = new TextEncoder();
  const requestBytes = encoder.encode(JSON.stringify({ expectVersion, steps: [] })).length;
  let chunk: XlsxEditStep[] = [];
  let cells = 0;
  let text = 0;
  let bytes = requestBytes;
  const flush = () => {
    if (chunk.length === 0) return;
    const result = handle.validateEdits({ expectVersion, steps: chunk });
    if (!result.ok) throw new WorkbookCommandRefusal(result.failure.code, result);
    chunk = [];
    cells = text = 0;
    bytes = requestBytes;
  };
  for (const step of steps) {
    const { start, end } = corners(step.target.range);
    const visits = (end.row - start.row + 1) * (end.col - start.col + 1);
    const units = step.op === 'setCellInputs' ? step.inputs.flat().reduce((sum, input) => sum + input.length, 0) : 0;
    const probe: XlsxEditStep = { op: 'patchStyle', target: step.target, patch: {} };
    const size = encoder.encode(JSON.stringify(probe)).length + 1;
    if (chunk.length === MAX_VALIDATION_STEPS || cells + visits > MAX_VALIDATION_CELLS ||
      text + units > MAX_VALIDATION_TEXT || bytes + size > MAX_VALIDATION_BYTES) flush();
    if (visits > MAX_VALIDATION_CELLS || units > MAX_VALIDATION_TEXT || bytes + size > MAX_VALIDATION_BYTES) {
      throw new WorkbookCommandRefusal('limit-exceeded');
    }
    chunk.push(probe);
    cells += visits;
    text += units;
    bytes += size;
  }
  flush();
}

function inputSteps(handle: WorkbookHandle, edits: { sheet: number; row: number; col: number; input: string }[],
  sheetIds = handle.sheetInfo().sheetIds): XlsxEditStep[] {
  const cells = new Map<string, typeof edits[number]>();
  for (const edit of edits) {
    cellTarget(handle, edit.sheet, edit.row, edit.col, sheetIds);
    cells.set(`${edit.sheet}:${edit.row}:${edit.col}`, edit);
  }
  const ordered = [...cells.values()].sort((a, b) => a.sheet - b.sheet || a.row - b.row || a.col - b.col);
  const steps: XlsxEditStep[] = [];
  const rectangles = new Map<string, { step: Extract<XlsxEditStep, { op: 'setCellInputs' }>; cells: number; text: number }>();
  for (let at = 0; at < ordered.length;) {
    const first = ordered[at++];
    let last = first;
    const inputs = [first.input];
    let text = first.input.length;
    while (at < ordered.length && ordered[at].sheet === first.sheet &&
      ordered[at].row === first.row && ordered[at].col === last.col + 1 &&
      inputs.length < MAX_VALIDATION_CELLS && text + ordered[at].input.length <= MAX_VALIDATION_TEXT) {
      last = ordered[at++];
      inputs.push(last.input);
      text += last.input.length;
    }
    const key = `${first.sheet}:${first.col}:${last.col}`;
    const previous = rectangles.get(key);
    if (previous?.step.target.range.kind === 'rowCol' && previous.step.target.range.end.row === first.row - 1 &&
      previous.cells + inputs.length <= MAX_VALIDATION_CELLS && previous.text + text <= MAX_VALIDATION_TEXT) {
      previous.step.target.range.end.row = first.row;
      previous.step.inputs.push(inputs);
      previous.cells += inputs.length;
      previous.text += text;
    } else {
      const step: Extract<XlsxEditStep, { op: 'setCellInputs' }> = { op: 'setCellInputs',
        target: sheetTarget(handle, first.sheet, { kind: 'rowCol',
          start: { row: first.row, col: first.col }, end: { row: first.row, col: last.col } }, sheetIds), inputs: [inputs] };
      steps.push(step);
      rectangles.set(key, { step, cells: inputs.length, text });
    }
  }
  return steps;
}

function captureCommandFormat(handle: WorkbookHandle, sheet: number, range: string) {
  try { return handle.captureFormat(sheet, range); }
  catch (error) {
    if (error instanceof Error && !(error instanceof WebAssembly.RuntimeError) &&
      /^sheet \d+ is not an editable worksheet$/.test(error.message)) {
      throw new WorkbookCommandRefusal('unsupported');
    }
    throw error;
  }
}

function admitFormatting(handle: WorkbookHandle, sheet: number, range: string): void {
  const target = resolveCommandTarget(handle, sheetTarget(handle, sheet, { kind: 'a1', a1: range }));
  const catalog = handle.readCells({ ranges: [] });
  if (!catalog.ok) throw new Error(catalog.failure.message);
  if (!catalog.sheets[sheet].editable) {
    captureCommandFormat(handle, sheet, 'A1');
    throw new WorkbookCommandRefusal('locked-target');
  }
  if ((target.end.row - target.start.row + 1) * (target.end.col - target.start.col + 1) > 100_000) {
    throw new WorkbookCommandRefusal('limit-exceeded');
  }
}

function resolveReplayOp(handle: WorkbookHandle, state: WorkbookCommandState, command: WorkbookCommand): WorkbookReplayOp {
  if (command.kind === 'select' || command.kind === 'setReadOnly' || command.kind === 'applyUpdate') {
    throw new WorkbookCommandRefusal('invalid-command');
  }
  const payload = command.payload;
  if (command.kind === 'applyOps' && 'toolbar' in payload) {
    const { toolbar } = payload;
    if (!toolbar.target) throw new WorkbookCommandRefusal('invalid-command');
    const target = resolveCommandTarget(handle, toolbar.target, state.sheetIds);
    if (toolbar.id === 'clear' &&
      (target.end.row - target.start.row + 1) * (target.end.col - target.start.col + 1) > MAX_VALIDATION_CELLS) {
      throw new WorkbookCommandRefusal('limit-exceeded');
    }
    if (toolbar.id !== 'clear') admitFormatting(handle, target.sheet, target.range);
    const formatting = toolbar.id === 'clear' ? {} : handle.selectionFormatting(target.sheet, target.range);
    const merged = toolbar.id === 'merge' ? handle.mergedRanges(target.sheet, target.range) : [];
    const op = resolveWorkbookToolbar(toolbar, target, formatting, merged);
    if (!op) throw new WorkbookCommandRefusal('invalid-command');
    return op;
  }
  if (command.kind === 'applyFormat' && 'formatToken' in payload) {
    const format = state.formats.get(payload.formatToken);
    if (!format) throw new WorkbookCommandRefusal('invalid-command');
    const destination = resolveCommandTarget(handle, payload.target, state.sheetIds);
    return { method: 'applyFormat', args: [destination.sheet, destination.range, format] };
  }
  if (command.kind === 'applyFormat' && 'capturedFrom' in payload) {
    const destination = resolveCommandTarget(handle, payload.target, state.sheetIds);
    const source = resolveCommandTarget(handle, payload.capturedFrom, state.sheetIds);
    if ((source.end.row - source.start.row + 1) * (source.end.col - source.start.col + 1) > 100_000) {
      throw new WorkbookCommandRefusal('limit-exceeded');
    }
    return { method: 'applyFormat', args: [destination.sheet, destination.range,
      captureCommandFormat(handle, source.sheet, source.range)] };
  }
  if (!('args' in payload)) throw new WorkbookCommandRefusal('invalid-command');
  const op = { method: command.kind, args: [...payload.args], calculation: payload.calculation } as WorkbookReplayOp;
  const numericSheet = () => {
    if (state.bases.get(command.basis.version) !== state.sheetRevision) {
      throw new WorkbookCommandRefusal('stale-target');
    }
  };
  if (op.method === 'applyOps') {
    const sheets = [...state.sheetIds];
    let nextSheetId = state.nextSheetId;
    op.args[0] = op.args[0].map((value) => {
      const { target = payload.target, ...operation } = value as Record<string, unknown> & { target?: XlsxRangeTarget };
      if (target) {
        const resolved = resolveCommandTarget(handle, target, sheets);
        if (operation.type === 'removeSheet') operation.index = resolved.sheet;
        else if (operation.type === 'moveSheet') operation.from = resolved.sheet;
        else if ('sheet' in operation) operation.sheet = resolved.sheet;
        else throw new WorkbookCommandRefusal('invalid-command');
      } else if ('sheet' in operation || 'index' in operation || operation.type === 'moveSheet' || operation.type === 'setDefinedNames' || operation.type === 'restoreSheet') {
        numericSheet();
        if (sheets.length !== state.sheetIds.length || sheets.some((id, index) => id !== state.sheetIds[index])) {
          throw new WorkbookCommandRefusal('stale-target');
        }
      }
      if (operation.type === 'addSheet') sheets.splice(operation.index as number, 0, `command-sheet:${nextSheetId++}`);
      else if (operation.type === 'removeSheet') sheets.splice(operation.index as number, 1);
      else if (operation.type === 'moveSheet') {
        const [moved] = sheets.splice(operation.from as number, 1);
        sheets.splice(operation.to as number, 0, moved);
      }
      return operation;
    });
    return op;
  }
  if (op.method === 'propose') {
    op.args[2] = op.args[2].map((value) => {
      const { target = payload.target, ...edit } = value as typeof value & { target?: XlsxRangeTarget };
      if (target) edit.sheet = resolveCommandTarget(handle, target, state.sheetIds).sheet;
      else numericSheet();
      return edit;
    });
    return op;
  }
  if (payload.target) {
    const target = resolveCommandTarget(handle, payload.target, state.sheetIds);
    switch (op.method) {
      case 'editCell':
        if (target.start.row !== target.end.row || target.start.col !== target.end.col) {
          throw new WorkbookCommandRefusal('invalid-command');
        }
        op.args = [target.sheet, target.start.row, target.start.col, op.args[3]];
        break;
      case 'editCells': op.args[0] = target.sheet; break;
      case 'patchRangeStyle':
      case 'setNumberFormat':
      case 'applyFormat': op.args[0] = target.sheet; op.args[1] = target.range; break;
      case 'setActiveSheet': op.args[0] = target.sheet; break;
      case 'moveChart':
      case 'moveRange': op.args[0] = target.sheet; break;
      default: throw new WorkbookCommandRefusal('invalid-command');
    }
  } else if (['editCell', 'editCells', 'patchRangeStyle', 'setNumberFormat', 'applyFormat', 'setActiveSheet', 'moveChart', 'moveRange'].includes(op.method)) numericSheet();
  return op;
}

function preflight(handle: WorkbookHandle, state: WorkbookCommandState, op: WorkbookReplayOp): void {
  switch (op.method) {
    case 'applyEdits':
      if (op.args[0].expectVersion !== handle.version()) throw new WorkbookCommandRefusal('stale-version');
      for (const step of op.args[0].steps) resolveCommandTarget(handle, step.target, state.sheetIds);
      op.args[0] = translateWorkbookSheetIds(op.args[0], state.sheetIds, handle.sheetInfo().sheetIds);
      break;
    case 'editCell': {
      const [sheet, row, col] = op.args;
      checkedBatch(handle, [{ op: 'patchStyle', target: cellTarget(handle, sheet, row, col, state.info?.sheetIds), patch: {} }]);
      break;
    }
    case 'editCells': {
      const [sheet, edits] = op.args;
      checkedBatch(handle, inputSteps(handle, edits.map((edit) => ({ ...edit, sheet })), state.info?.sheetIds));
      break;
    }
    case 'applyOps':
      for (const value of op.args[0]) {
        const operation = value as Record<string, unknown>;
        const dimension = operation.type === 'setColWidth' ? operation.width :
          operation.type === 'setRowHeight' ? operation.height : undefined;
        const maximum = operation.type === 'setColWidth' ? 255 : 409.5;
        if (dimension !== undefined && dimension !== null &&
          (typeof dimension !== 'number' || !Number.isFinite(dimension) || dimension < 0 || dimension > maximum)) {
          throw new WorkbookCommandRefusal('invalid-step');
        }
      }
      break;
    case 'patchRangeStyle':
    case 'setNumberFormat':
      admitFormatting(handle, op.args[0], op.args[1]);
      break;
    case 'propose':
      if (!handle.isProposalsAvailable()) throw new WorkbookCommandRefusal('unsupported');
      checkedBatch(handle, inputSteps(handle, op.args[2], state.info?.sheetIds));
      break;
    case 'acceptProposal':
      if (!handle.isProposalsAvailable()) throw new WorkbookCommandRefusal('unsupported');
      if (!handle.listProposals().some((proposal) => proposal.id === op.args[0])) {
        throw new WorkbookCommandRefusal('missing-target');
      }
      break;
    case 'applyFormat':
      admitFormatting(handle, op.args[0], op.args[1]);
      if (op.args[2].rows === 0 || op.args[2].columns === 0 ||
        op.args[2].formats.length !== op.args[2].rows * op.args[2].columns) {
        throw new WorkbookCommandRefusal('invalid-command');
      }
      break;
    case 'setActiveSheet':
    case 'moveChart':
    case 'moveRange': sheetTarget(handle, op.args[0], { kind: 'a1', a1: 'A1' }); break;
  }
}

export function classifyRefusal(error: unknown, method?: WorkbookReplayMethod): WorkbookCommandRefusal | undefined {
  if (error instanceof WorkbookCommandRefusal) return error;
  if (method === 'acceptProposal' && error instanceof StaleProposalError) {
    return new WorkbookCommandRefusal('stale-proposal', { cells: error.cells, targets: error.targets });
  }
  if (!(error instanceof Error) || error instanceof WebAssembly.RuntimeError) return undefined;
  if (method === 'applyOps' && error.message === 'structural operations are unavailable in collaborative mode') {
    return new WorkbookCommandRefusal('unsupported');
  }
  if (error.message === 'calculation context is unavailable for collaborative workbooks') {
    return new WorkbookCommandRefusal('unsupported');
  }
  if (['applyOps', 'patchRangeStyle', 'setNumberFormat', 'applyFormat', 'propose', 'acceptProposal'].includes(method ?? '') &&
    error.message === 'number format table is full') return new WorkbookCommandRefusal('limit-exceeded');
  if (method === 'applyOps' && (
    error.message === 'invalid collaborative workbook state: too many sheet merge edits' ||
    error.message === 'invalid collaborative workbook state: too many sheet hyperlink edits'
  )) {
    return new WorkbookCommandRefusal('limit-exceeded');
  }
  if (method === 'applyOps' && (
    /^formula at sheet \d+, [A-Z]+[1-9]\d* cannot be safely rewritten$/.test(error.message) ||
    /^(?:defined name|chart reference in) .+ cannot be safely rewritten$/.test(error.message) ||
    /^.+ references cells this edit would move, and it cannot be rewritten$/.test(error.message) ||
    /^anchor \d+ of .+ is held by two sheets that disagree on where it sits$/.test(error.message) ||
    /^chart .+ cannot change how it follows a grid edit$/.test(error.message) ||
    /^chart .+ cannot be resized: a one-cell extent is not written back$/.test(error.message) ||
    /^chart .+ cannot change its anchor kind$/.test(error.message)
  )) return new WorkbookCommandRefusal('unsupported');
  if (method === 'moveRange' && (
    /^range move cannot safely preserve .+$/.test(error.message) ||
    /^.+ contains worksheet features a cell move cannot safely preserve$/.test(error.message) ||
    /^.+ references cells this move would relocate, and it cannot be rewritten$/.test(error.message)
  )) return new WorkbookCommandRefusal('unsupported');
  if (method === 'applyOps' &&
    /^chart frame .+ does not hold the chart part and anchor this op was recorded against; its drawing has changed$/.test(error.message)) {
    return new WorkbookCommandRefusal('stale-target');
  }
  if ((method === 'editCell' || method === 'editCells' || method === 'propose' || method === 'acceptProposal' || method === 'moveRange') &&
    /^(?:cell number must be finite|cell text exceeds Excel's length limit|formula exceeds the length limit)$/.test(error.message)) {
    return new WorkbookCommandRefusal('invalid-step');
  }
  if ((method === 'applyOps' || method === 'patchRangeStyle' || method === 'setNumberFormat' || method === 'applyFormat') &&
    (/^(?:font family must not be empty|font size must be between 1 and 400|captured format dimensions do not match its cells)$/.test(error.message) ||
      /^invalid color .*; expected #rrggbb$/.test(error.message))) {
    return new WorkbookCommandRefusal('invalid-step');
  }
  if (method === 'applyOps' && (
    /^the anchor of .+ would leave the sheet grid$/.test(error.message) ||
    /^defined name .+ has a formula above the length limit$/.test(error.message) ||
    /^(?:chart reference exceeds the supported length|chart reference contains a character xml cannot carry)$/.test(error.message)
  )) return new WorkbookCommandRefusal('invalid-step');
  if ((method === 'applyOps' || method === 'moveChart') &&
    /^(?:anchor offset is out of range|anchor extent is not positive|anchor corners are inverted or coincident|anchor cell is outside the worksheet grid|anchor extends beyond the worksheet grid)$/.test(error.message)) {
    return new WorkbookCommandRefusal('invalid-step');
  }
  if ((method === 'applyOps' && /^bad ops: /.test(error.message)) ||
    (method === 'applyFormat' && /^bad format args: /.test(error.message)) ||
    (method === 'moveRange' && /^bad (?:move range args|range|move destination): /.test(error.message)) ||
    (method === 'applyEdits' && /^malformed request: /.test(error.message))) {
    return new WorkbookCommandRefusal('invalid-command');
  }
  if ((method === 'applyOps' || method === 'moveChart' || method === 'moveRange') && (
    /^(?:sheet(?: index)?|row|column) \d+ (?:is )?out of range$/.test(error.message) ||
    /^cell row \d+, column \d+ is out of range$/.test(error.message) ||
    /^sheet \d+ is not an editable worksheet$/.test(error.message) ||
    /^invalid sheet name: /.test(error.message) ||
    /^duplicate sheet name: /.test(error.message) ||
    /^(?:row|column) operation (?:count must be positive|exceeds sheet bounds)$/.test(error.message) ||
    /^range \d+x\d+ exceeds the \d+-cell copy cap$/.test(error.message) ||
    /^(?:workbook has no sheets|sheet contains too many hyperlinks|hyperlink must have a destination|hyperlink field exceeds its length limit)$/.test(error.message) ||
    /^workbook contains (?:a non-finite cell number|cell text above Excel's length limit|a formula above the length limit|an invalid cell style index|an invalid column width|an invalid row height|overlapping merged ranges)$/.test(error.message) ||
    /^no chart frame on this sheet is named /.test(error.message) ||
    /^chart .+ is pinned to the sheet and cannot be moved$/.test(error.message) ||
    /^(?:row|column) insertion would (?:discard content at|push a chart anchor past) the sheet boundary$/.test(error.message) ||
    /^(?:freeze pane is out of range|range start must be above and left of range end|restore sheet operations are internal)$/.test(error.message)
  )) return new WorkbookCommandRefusal('invalid-step');
  return undefined;
}

export interface WorkbookCommandState {
  historyReset?: boolean;
  identity: CommandIdentity;
  epoch: number;
  throughSeq: number;
  selection: WorkbookCommandSelection;
  readOnly: boolean;
  statuses: Map<number, CommandStatus>;
  bases: Map<string, number>;
  sheetRevision: number;
  editState?: WorkbookEditState;
  info?: SheetInfo;
  editableSheets?: boolean[];
  formattingDirty?: boolean;
  mergesDirty?: boolean;
  catalogDirty?: boolean;
  formattingSelection?: string;
  publishedSelection?: string;
  sheetIds: string[];
  nextSheetId: number;
  calculation: WorkbookCalculationContext | null;
  formats: Map<string, CapturedFormat>;
  historySelections: ({
    before: WorkbookCommandSelection; after: WorkbookCommandSelection;
    beforeSheets: string[]; afterSheets: string[];
  } | undefined)[];
}

export function createWorkbookCommandState(
  identity: CommandIdentity, basis: CommandBasis, selection: WorkbookCommandSelection,
  sheetIds: string[], calculation: WorkbookCalculationContext | null = null
): WorkbookCommandState {
  return { identity, epoch: basis.epoch, throughSeq: 0, selection, readOnly: false,
    statuses: new Map(), bases: new Map([[basis.version, 0]]), sheetRevision: 0, historySelections: [],
    sheetIds: [...sheetIds], nextSheetId: sheetIds.length, calculation, formats: new Map() };
}

export function projectEditState(handle: WorkbookHandle, state: WorkbookCommandState): Partial<WorkbookEditState> | undefined {
  const previous = state.editState;
  const catalogChanged = !previous || state.catalogDirty;
  if (!state.info || catalogChanged) state.info = handle.sheetInfo();
  const info = state.info;
  if (!state.editableSheets || catalogChanged) {
    const catalog = handle.readCells({ ranges: [] });
    if (!catalog.ok) throw new Error(catalog.failure.message);
    state.editableSheets = catalog.sheets.map((sheet) => sheet.editable);
  }
  const history = handle.historyState();
  const selected = JSON.stringify(state.selection);
  const selectionChanged = state.formattingSelection !== selected;
  let formatting = previous?.formatting;
  let mergedRanges = previous?.mergedRanges;
  let formattingAvailable = previous?.formattingAvailable ?? false;
  let selectionEditable = previous?.selectionEditable ?? false;
  if (!formatting || state.formattingDirty || selectionChanged) {
    const target = resolveCommandTarget(handle, state.selection, state.sheetIds);
    const { sheet, start, end } = target;
    selectionEditable = state.editableSheets[sheet];
    formattingAvailable = selectionEditable && (end.row - start.row + 1) * (end.col - start.col + 1) <= 100_000;
    formatting = formattingAvailable ? handle.selectionFormatting(sheet, target.range) : {};
  }
  if (!mergedRanges || state.mergesDirty || selectionChanged) {
    const target = resolveCommandTarget(handle, state.selection, state.sheetIds);
    mergedRanges = handle.mergedRanges(target.sheet, target.range);
  }
  state.formattingSelection = selected;
  state.formattingDirty = state.mergesDirty = state.catalogDirty = false;
  const next: WorkbookEditState = { canUndo: history.canUndo, canRedo: history.canRedo,
    formatting, formattingAvailable, selectionEditable, mergedRanges: mergedRanges ?? [],
    calculation: handle.calculationStatus(),
    sheets: catalogChanged ? state.sheetIds.map((id, index) => ({ id, index, name: info.sheetNames[index] })) : previous!.sheets,
    activeSheet: info.activeSheet, readOnly: state.readOnly,
    ...(state.historyReset ? { historyReset: true } : {}) };
  const changed: Partial<WorkbookEditState> = {};
  for (const key of Object.keys(next) as (keyof WorkbookEditState)[]) {
    if (!previous || previous[key] !== next[key] && JSON.stringify(previous[key]) !== JSON.stringify(next[key])) {
      Object.assign(changed, { [key]: next[key] });
    }
  }
  state.editState = next;
  return Object.keys(changed).length > 0 ? changed : undefined;
}

function applyIntent(handle: WorkbookHandle, state: WorkbookCommandState, command: WorkbookCommand): CommandStatus {
  validateWorkbookCommandPayload(command);
  if (command.frame && command.frame.sheet >= handle.sheetCount()) throw new WorkbookCommandRefusal('missing-target');
  if (command.kind === 'applyUpdate') {
    const before = handle.version();
    try { handle.applyUpdate(command.payload.update); }
    catch (error) {
      if (error instanceof Error && !(error instanceof WebAssembly.RuntimeError) && handle.version() === before) {
        throw new WorkbookCommandRefusal('invalid-update');
      }
      throw error;
    }
    state.formattingDirty = state.mergesDirty = state.catalogDirty = true;
    return { seq: command.seq, status: handle.version() === before ? 'noop' : 'applied',
      committed: handle.version() !== before, result: { remoteUpdate: true, proposals: handle.listProposals() } };
  }
  if (command.kind === 'select') {
    validateSelection(handle, state, command.payload.selection);
    const changed = JSON.stringify(state.selection) !== JSON.stringify(command.payload.selection);
    state.selection = structuredClone(command.payload.selection);
    return { seq: command.seq, status: changed ? 'applied' : 'noop', committed: false };
  }
  if (command.kind === 'setReadOnly') {
    const changed = state.readOnly !== command.payload.readOnly;
    state.readOnly = command.payload.readOnly;
    return { seq: command.seq, status: changed ? 'applied' : 'noop', committed: false };
  }
  const refreshProposals = command.kind === 'applyOps' && 'toolbar' in command.payload &&
    command.payload.toolbar.id === 'refreshProposals';
  if (command.kind === 'setActiveSheet' && command.payload.expectVersion !== undefined &&
    command.payload.expectVersion !== handle.version()) throw new WorkbookCommandRefusal('stale-version');
  if (state.readOnly && command.kind !== 'setActiveSheet' && !refreshProposals) throw new WorkbookCommandRefusal('read-only');
  if (command.payload.selection) validateSelection(handle, state, command.payload.selection);
  if (refreshProposals) {
    return { seq: command.seq, status: 'applied', committed: false, result: { proposals: handle.listProposals() } };
  }
  if (command.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'paintFormat') {
    const { target } = command.payload.toolbar;
    if (!target) throw new WorkbookCommandRefusal('invalid-command');
    const resolved = resolveCommandTarget(handle, target, state.sheetIds);
    admitFormatting(handle, resolved.sheet, resolved.range);
    const format = captureCommandFormat(handle, resolved.sheet, resolved.range);
    const bytes = globalThis.crypto.getRandomValues(new Uint8Array(16));
    const token = `format:${Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('')}`;
    state.formats.set(token, format);
    if (state.formats.size > FORMAT_TOKEN_CACHE_SIZE) state.formats.delete(state.formats.keys().next().value!);
    if (command.payload.selection) state.selection = structuredClone(command.payload.selection);
    const result: WorkbookFormatTokenResult = { formatToken: token };
    return { seq: command.seq, status: 'applied', committed: false, result };
  }
  const op = resolveReplayOp(handle, state, command);
  preflight(handle, state, op);
  if (op.method === 'setActiveSheet' && command.payload.selection &&
    state.sheetIds[op.args[0]] !== command.payload.selection.sheetId) throw new WorkbookCommandRefusal('invalid-command');
  const before = handle.version();
  const beforeActiveSheet = state.info?.activeSheet ?? handle.sheetInfo().activeSheet;
  const history = handle.historyState();
  const beforeSelection = structuredClone(state.selection);
  const beforeSheetIds = [...state.sheetIds];
  let result: ReturnType<typeof applyWorkbookReplayOp>;
  try {
    if (op.calculation) handle.setCalculationContext(op.calculation);
    result = applyWorkbookReplayOp(handle, op);
  } catch (error) {
    const refusal = classifyRefusal(error, op.method);
    if (refusal && handle.version() === before) {
      if (op.calculation) handle.setCalculationContext(state.calculation);
      throw refusal;
    }
    throw new SessionFailure(error instanceof WebAssembly.RuntimeError ? 'trap' : 'crash',
      `Workbook command failed at sequence ${command.seq} (${command.kind})`,
      error instanceof Error ? error.message : String(error));
  }
  if (workbookReplayRefused(result)) {
    if (op.method !== 'applyEdits' || handle.version() !== before) {
      throw new SessionFailure('crash', 'Workbook command failed after dispatch');
    }
    const refused = result as Extract<ReturnType<WorkbookHandle['applyEdits']>, { ok: false }>;
    if (op.calculation) handle.setCalculationContext(state.calculation);
    throw new WorkbookCommandRefusal(refused.failure.code, refused);
  }
  if (op.calculation) state.calculation = { ...op.calculation };
  const committed = handle.version() !== before;
  const afterHistory = handle.historyState();
  if (committed && op.method === 'applyOps') {
    for (const value of op.args[0]) {
      const operation = value as { type: string; index: number; from: number; to: number };
      if (operation.type === 'addSheet') state.sheetIds.splice(operation.index, 0, `command-sheet:${state.nextSheetId++}`);
      else if (operation.type === 'removeSheet') state.sheetIds.splice(operation.index, 1);
      else if (operation.type === 'moveSheet') {
        const [moved] = state.sheetIds.splice(operation.from, 1);
        state.sheetIds.splice(operation.to, 0, moved);
      }
    }
  }
  if (op.method === 'undo' || op.method === 'redo') {
    if (committed) {
      const depth = op.method === 'undo' ? history.undoDepth : afterHistory.undoDepth;
      const anchor = state.historySelections[depth];
      if (anchor) {
        state.sheetIds = [...(op.method === 'undo' ? anchor.beforeSheets : anchor.afterSheets)];
        state.selection = structuredClone(op.method === 'undo' ? anchor.before : anchor.after);
        const sheet = state.sheetIds.indexOf(state.selection.sheetId);
        if (sheet >= 0) handle.setActiveSheet(sheet);
      }
    }
  } else if (command.payload.selection) state.selection = structuredClone(command.payload.selection);
  if (committed) {
    if (op.method === 'undo' || op.method === 'redo') {
      state.formattingDirty = state.mergesDirty = true;
      state.catalogDirty = true;
    } else if (op.method === 'applyOps') {
      const types = op.args[0].map((value) => (value as { type: string }).type);
      const movesCells = types.some((type) => ['insertRows', 'deleteRows', 'insertCols', 'deleteCols', 'addSheet', 'removeSheet', 'moveSheet'].includes(type));
      state.formattingDirty = movesCells || types.some((type) =>
        ['setCell', 'patchRangeStyle', 'setRangeNumberFormat', 'applyRangeFormat', 'restoreColStyles'].includes(type));
      state.mergesDirty = movesCells || types.some((type) => ['mergeCells', 'unmergeCells'].includes(type));
      state.catalogDirty = types.some((type) => ['addSheet', 'removeSheet', 'moveSheet', 'renameSheet'].includes(type));
    } else if (op.method === 'patchRangeStyle' || op.method === 'setNumberFormat' ||
      op.method === 'applyFormat' || op.method === 'acceptProposal' || op.method === 'moveRange' ||
      op.method === 'applyEdits' && op.args[0].steps.some((step) => step.op === 'patchStyle' || step.op === 'setNumberFormat')) {
      state.formattingDirty = true;
    }
    if (beforeSheetIds.length !== state.sheetIds.length || beforeSheetIds.some((id, index) => id !== state.sheetIds[index])) {
      state.sheetRevision += 1;
    }
  }
  const info = !state.info || state.catalogDirty || op.method === 'setActiveSheet' ||
    op.method === 'undo' || op.method === 'redo' ? handle.sheetInfo() : state.info;
  state.info = info;
  if (!state.sheetIds.includes(state.selection.sheetId) ||
    op.method === 'setActiveSheet' && !command.payload.selection && beforeActiveSheet !== info.activeSheet) {
    state.selection = { sheetId: state.sheetIds[info.activeSheet], range: { kind: 'a1', a1: 'A1' } };
  }
  if (op.method !== 'undo' && op.method !== 'redo') {
    if (committed && afterHistory.undoDepth > history.undoDepth) {
      if (state.historySelections.length > history.undoDepth + 1) state.historySelections.length = history.undoDepth + 1;
      state.historySelections[afterHistory.undoDepth] = { before: beforeSelection, after: structuredClone(state.selection),
        beforeSheets: beforeSheetIds, afterSheets: [...state.sheetIds] };
    }
  }
  const stateChanged = op.method === 'setActiveSheet' ? beforeActiveSheet !== info.activeSheet ||
    JSON.stringify(beforeSelection) !== JSON.stringify(state.selection) :
    op.method === 'propose' || op.method === 'acceptProposal' || op.method === 'rejectProposal' && result === true;
  return { seq: command.seq, status: committed || stateChanged ? 'applied' : 'noop', committed,
    ...(result === undefined ? {} : { result: translateWorkbookSheetIds(result, info.sheetIds, state.sheetIds) }) };
}

export function applyWorkbookCommand(
  handle: WorkbookHandle, state: WorkbookCommandState, command: WorkbookCommand
): WorkbookWireAck {
  validateWorkbookCommandEnvelope(command);
  if (command.session !== state.identity.session || command.incarnation !== state.identity.incarnation ||
    command.basis.epoch !== state.epoch) throw new Error('Workbook command identity or epoch mismatch');
  const cached = state.statuses.get(command.seq);
  if (cached) return { throughSeq: state.throughSeq, version: handle.version(), statuses: [cached] };
  if (command.seq !== state.throughSeq + 1) throw new Error(`Expected workbook command sequence ${state.throughSeq + 1}`);
  const dependency = command.afterSeq === undefined ? undefined : state.statuses.get(command.afterSeq);
  const refused = command.afterSeq !== undefined && !dependency ? 'dependency-expired' : undefined;
  const cancelled = dependency?.status === 'refused' || dependency?.status === 'cancelled';
  const reply = applyWorkbookCommandIntent(handle, state, command, refused, cancelled);
  state.throughSeq = command.seq;
  state.statuses.set(command.seq, structuredClone(reply.statuses[0]));
  if (state.statuses.size > STATUS_CACHE_SIZE) state.statuses.delete(state.statuses.keys().next().value!);
  return reply;
}

export function applyWorkbookCommandIntent(
  handle: WorkbookHandle, state: WorkbookCommandState, command: WorkbookCommand,
  refusal?: string, cancelled = false
): WorkbookWireAck {
  let status: CommandStatus;
  try {
    if (refusal) throw new WorkbookCommandRefusal(refusal);
    if (cancelled) {
      status = { seq: command.seq, status: 'cancelled', reason: 'dependency-refused', committed: false,
        affectedSeqs: [command.seq] };
    } else {
      if (!state.bases.has(command.basis.version)) throw new WorkbookCommandRefusal('stale-basis');
      status = applyIntent(handle, state, command);
    }
  } catch (error) {
    const refused = classifyRefusal(error);
    if (!refused) throw error instanceof SessionFailure ? error : new SessionFailure(
      error instanceof WebAssembly.RuntimeError ? 'trap' : 'crash', 'Workbook command failed',
      error instanceof Error ? error.message : String(error));
    status = { seq: command.seq, status: 'refused', reason: refused.reason, committed: false,
      affectedSeqs: [command.seq], ...(refused.result === undefined ? {} : {
        result: translateWorkbookSheetIds(refused.result, handle.sheetInfo().sheetIds, state.sheetIds),
      }) };
  }
  const version = handle.version();
  if (state.bases.has(command.basis.version)) {
    const revision = state.bases.get(command.basis.version)!;
    state.bases.delete(command.basis.version);
    state.bases.set(command.basis.version, revision);
  }
  state.bases.set(version, state.sheetRevision);
  if (state.bases.size > STATUS_CACHE_SIZE) state.bases.delete(state.bases.keys().next().value!);
  const selected = JSON.stringify(state.selection);
  const selection = selected !== state.publishedSelection || status.status === 'refused' || status.status === 'cancelled'
    ? structuredClone(state.selection) : undefined;
  if (status.status !== 'refused' && status.status !== 'cancelled') state.publishedSelection = selected;
  let editState: Partial<WorkbookEditState> | undefined;
  try { editState = projectEditState(handle, state); }
  catch (error) {
    throw new SessionFailure(error instanceof WebAssembly.RuntimeError ? 'trap' : 'crash',
      'Workbook command state projection failed', error instanceof Error ? error.message : String(error));
  }
  return { throughSeq: command.seq, version, statuses: [status],
    ...(selection ? { selection } : {}), ...(editState ? { editState } : {}) };
}
