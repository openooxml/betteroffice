import type { MethodPolicies } from '../../../../shared/office-session';
import type { WorkbookCalculationContext, WorkbookHandle } from '../wasm/loader';
import type { WorkbookSession } from './client';

export const WORKBOOK_REPLAY_MUTATORS = {
  editCell: true,
  editCells: true,
  applyEdits: true,
  applyOps: true,
  patchRangeStyle: true,
  setNumberFormat: true,
  applyFormat: true,
  moveChart: true,
  undo: true,
  redo: true,
  setActiveSheet: true,
} satisfies Partial<Record<keyof WorkbookHandle, true>>;

export type WorkbookReplayMethod = keyof typeof WORKBOOK_REPLAY_MUTATORS;

export type WorkbookReplayOp = {
  [K in WorkbookReplayMethod]: { method: K; args: Parameters<WorkbookHandle[K]> }
}[WorkbookReplayMethod];

export interface WorkbookReplayEnvelope {
  sequence: number;
  calculation: WorkbookCalculationContext;
  op: WorkbookReplayOp;
}

export interface WorkbookReplayReply {
  sequence: number;
  revision: number;
  version: number;
  result: ReturnType<WorkbookHandle[WorkbookReplayMethod]>;
}

export type WorkbookInternalSessionMethods = {
  replay(envelope: WorkbookReplayEnvelope): WorkbookReplayReply;
};

export const WORKBOOK_INTERNAL_SESSION_METHODS = { replay: true } as const;

export const WORKBOOK_INTERNAL_SESSION_POLICIES: MethodPolicies<WorkbookInternalSessionMethods> = {
  replay: { lane: 'input', mutates: true, userInput: true, reorderable: false },
};

export const workbookSessionInternals = new WeakMap<WorkbookSession, {
  replay(envelope: WorkbookReplayEnvelope): Promise<WorkbookReplayReply>;
  editPeerAttached: boolean;
}>();

function record(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== 'object') return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === null || prototype === Object.prototype;
}

function integer(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function finite(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

function point(value: unknown): boolean {
  return record(value) && integer(value.row) && integer(value.col);
}

function target(value: unknown): boolean {
  if (!record(value) || typeof value.sheetId !== 'string' || !record(value.range)) return false;
  return value.range.kind === 'a1' ? typeof value.range.a1 === 'string' :
    value.range.kind === 'rowCol' && point(value.range.start) && point(value.range.end);
}

function strings(value: unknown): boolean {
  return Array.isArray(value) && Array.from(value).every((item) => typeof item === 'string');
}

function matrix(value: unknown): boolean {
  return Array.isArray(value) && Array.from(value).every(strings);
}

function optionalFields(
  value: Record<string, unknown>, fields: Record<string, (field: unknown) => boolean>
): boolean {
  return Object.entries(fields).every(([key, valid]) => value[key] === undefined || valid(value[key]));
}

function oneOf(values: readonly string[]): (value: unknown) => boolean {
  return (value) => typeof value === 'string' && values.includes(value);
}

function style(value: unknown): boolean {
  return record(value) && optionalFields(value, {
    bold: (field) => typeof field === 'boolean',
    italic: (field) => typeof field === 'boolean',
    strikethrough: (field) => typeof field === 'boolean',
    fontFamily: (field) => typeof field === 'string',
    fontSize: finite,
    textColor: (field) => typeof field === 'string',
    fillColor: (field) => typeof field === 'string',
    border: (field) => record(field) && optionalFields(field, {
      preset: oneOf(['all', 'inner', 'horizontal', 'vertical', 'outer', 'left', 'top', 'right', 'bottom', 'none']),
      style: oneOf(['solid', 'dashed', 'dotted', 'double']),
      color: (color) => typeof color === 'string',
    }),
    horizontalAlignment: oneOf(['left', 'center', 'right']),
    verticalAlignment: oneOf(['top', 'middle', 'bottom']),
    textWrapping: oneOf(['overflow', 'wrap', 'clip']),
    clear: (field) => Array.isArray(field) && Array.from(field).every(oneOf([
      'bold', 'italic', 'strikethrough', 'fontFamily', 'fontSize', 'textColor', 'fillColor',
      'borders', 'horizontalAlignment', 'verticalAlignment', 'textWrapping',
    ])),
  });
}

function numberFormat(value: unknown): boolean {
  return oneOf([
    'automatic', 'plainText', 'number', 'percent', 'scientific', 'currency', 'date', 'time',
    'increaseDecimal', 'decreaseDecimal',
  ])(value) || (record(value) && value.type === 'custom' && typeof value.pattern === 'string');
}

function guard(value: unknown): boolean {
  return record(value) && optionalFields(value, {
    formula: (field) => field === null || typeof field === 'string',
    displayText: (field) => typeof field === 'string',
    value: (field) => record(field) && (
      field.kind === 'empty' ||
      (field.kind === 'number' && finite(field.value)) ||
      (field.kind === 'text' && typeof field.value === 'string') ||
      (field.kind === 'bool' && typeof field.value === 'boolean') ||
      (field.kind === 'error' && oneOf([
        '#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#SPILL!', '#CALC!',
      ])(field.value))
    ),
  });
}

function editRequest(value: unknown): boolean {
  if (!record(value) || typeof value.expectVersion !== 'string' || !Array.isArray(value.steps)) {
    return false;
  }
  return optionalFields(value, {
    source: oneOf(['host', 'agent']),
    history: oneOf(['separate', 'none']),
    calculation: (field) => record(field) &&
      Object.keys(field).every((key) => key === 'nowSerial') &&
      optionalFields(field, { nowSerial: finite }),
  }) && Array.from(value.steps).every((step) => {
    if (!record(step) || !target(step.target) || !optionalFields(step, {
      expect: (field) => record(field) && Array.isArray(field.cells) &&
        Array.from(field.cells).every((row) => Array.isArray(row) && Array.from(row).every(guard)),
    })) return false;
    switch (step.op) {
      case 'setCellInputs': return matrix(step.inputs);
      case 'setFormulas': return matrix(step.formulas);
      case 'setNumberFormat': return numberFormat(step.format);
      case 'patchStyle': return style(step.patch);
      default: return false;
    }
  });
}

function capturedCellFormat(value: unknown): boolean {
  return record(value) && record(value.font) && (value.fill === 'None' || record(value.fill)) &&
    record(value.border) && record(value.numberFormat) &&
    typeof value.numberFormat.kind === 'string' && record(value.alignment);
}

function validOp(value: unknown): boolean {
  if (!record(value) || !Array.isArray(value.args)) return false;
  const args = Array.from(value.args);
  switch (value.method) {
    case 'editCell':
      return args.length === 4 && args.slice(0, 3).every(integer) && typeof args[3] === 'string';
    case 'editCells':
      return args.length === 2 && integer(args[0]) && Array.isArray(args[1]) &&
        Array.from(args[1]).every((edit) => point(edit) && record(edit) && typeof edit.input === 'string');
    case 'applyEdits': return args.length === 1 && editRequest(args[0]);
    case 'applyOps':
      return args.length === 1 && Array.isArray(args[0]) &&
        Array.from(args[0]).every((op) => record(op) && typeof op.type === 'string');
    case 'patchRangeStyle':
      return args.length === 3 && integer(args[0]) && typeof args[1] === 'string' && style(args[2]);
    case 'setNumberFormat':
      return args.length === 3 && integer(args[0]) && typeof args[1] === 'string' && numberFormat(args[2]);
    case 'applyFormat':
      return args.length === 3 && integer(args[0]) && typeof args[1] === 'string' &&
        record(args[2]) && integer(args[2].rows) && integer(args[2].columns) &&
        Array.isArray(args[2].formats) && Array.from(args[2].formats).every(capturedCellFormat);
    case 'moveChart':
      return args.length === 4 && integer(args[0]) && typeof args[1] === 'string' &&
        args.slice(2).every(finite);
    case 'undo':
    case 'redo': return args.length === 0;
    case 'setActiveSheet': return args.length === 1 && integer(args[0]);
    default: return false;
  }
}

export function validateWorkbookReplayEnvelope(value: unknown): asserts value is WorkbookReplayEnvelope {
  if (!record(value) || !integer(value.sequence) || value.sequence === 0 ||
    !record(value.calculation) || !finite(value.calculation.nowSerial) ||
    !integer(value.calculation.randSeed) || value.calculation.randSeed > 0xffff_ffff ||
    !validOp(value.op)) {
    const error = new TypeError('Malformed workbook replay envelope');
    error.name = 'WorkbookReplayValidationError';
    throw error;
  }
}

export function workbookReplayRefused(result: unknown): boolean {
  return record(result) && result.ok === false;
}

export function applyWorkbookReplayOp(
  handle: WorkbookHandle, op: WorkbookReplayOp
): WorkbookReplayReply['result'] {
  switch (op.method) {
    case 'editCell': return handle.editCell(...op.args);
    case 'editCells': return handle.editCells(...op.args);
    case 'applyEdits': return handle.applyEdits(...op.args);
    case 'applyOps': return handle.applyOps(...op.args);
    case 'patchRangeStyle': return handle.patchRangeStyle(...op.args);
    case 'setNumberFormat': return handle.setNumberFormat(...op.args);
    case 'applyFormat': return handle.applyFormat(...op.args);
    case 'moveChart': return handle.moveChart(...op.args);
    case 'undo': return handle.undo(...op.args);
    case 'redo': return handle.redo(...op.args);
    case 'setActiveSheet': return handle.setActiveSheet(...op.args);
  }
}
