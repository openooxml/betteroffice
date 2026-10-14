import type { CellPoint, MergedRange, RangeStylePatch, SelectionFormatting } from '../wasm/loader';
import { WorkbookCommandRefusal, type WorkbookToolbarPayload } from './commandProtocol';
import type { WorkbookReplayOp } from './replay';

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
const MAX_MERGE_OPERATIONS = 4096;

export function validWorkbookToolbarArgs(id: unknown, args: unknown): boolean {
  if (id === 'bold' || id === 'italic' || id === 'strikethrough' || id === 'paintFormat' ||
    id === 'clear' || id === 'refreshProposals') return args === null;
  if ((id === 'merge' || id === 'textWrapping') && args === null) return true;
  if (!args || typeof args !== 'object' || Array.isArray(args)) return false;
  const bound = args as Record<string, unknown>;
  const choice = (values: string[]) => typeof bound.value === 'string' && values.includes(bound.value);
  switch (id) {
    case 'fontFamily': return typeof bound.family === 'string' && bound.family.trim() !== '';
    case 'fontSize': return typeof bound.points === 'number' && Number.isFinite(bound.points) &&
      bound.points >= 1 && bound.points <= 400;
    case 'fontSizeStep':
    case 'decimalPlaces': return bound.direction === 'increase' || bound.direction === 'decrease';
    case 'textColor':
    case 'fillColor':
    case 'borderColor': return typeof bound.color === 'string' && /^#[0-9a-f]{6}$/i.test(bound.color);
    case 'numberFormat': return choice(['automatic', 'plainText', 'number', 'percent', 'scientific', 'currency', 'date', 'time', 'custom']);
    case 'borderPreset': return choice(['all', 'inner', 'horizontal', 'vertical', 'outer', 'left', 'top', 'right', 'bottom', 'none']);
    case 'borderStyle': return choice(['solid', 'dashed', 'dotted', 'double']);
    case 'merge': return choice(['all', 'horizontal', 'vertical', 'unmerge']);
    case 'horizontalAlignment': return choice(['left', 'center', 'right']);
    case 'verticalAlignment': return choice(['top', 'middle', 'bottom']);
    case 'textWrapping': return choice(['overflow', 'wrap', 'clip']);
    default: return false;
  }
}

export function resolveWorkbookToolbar(
  toolbar: WorkbookToolbarPayload,
  target: { sheet: number; range: string; start: CellPoint; end: CellPoint },
  formatting: SelectionFormatting, merged: MergedRange[]
): WorkbookReplayOp | undefined {
  const { id, args } = toolbar;
  const { sheet, range, start, end } = target;
  const bound = args as Record<string, never>;
  const style = (patch: RangeStylePatch): WorkbookReplayOp => ({ method: 'patchRangeStyle', args: [sheet, range, patch] });
  switch (id) {
    case 'clear': {
      const edits: { row: number; col: number; input: string }[] = [];
      for (let row = start.row; row <= end.row; row++) {
        for (let col = start.col; col <= end.col; col++) edits.push({ row, col, input: '' });
      }
      return { method: 'editCells', args: [sheet, edits] };
    }
    case 'bold': return style({ bold: formatting.bold !== true });
    case 'italic': return style({ italic: formatting.italic !== true });
    case 'strikethrough': return style({ strikethrough: formatting.strikethrough !== true });
    case 'fontFamily': return style({ fontFamily: bound.family });
    case 'fontSize': return style({ fontSize: bound.points });
    case 'fontSizeStep': {
      const size = formatting.fontSize ?? 10;
      const next = bound.direction === 'increase' ? FONT_SIZES.find((entry) => entry > size) ?? size + 1 :
        [...FONT_SIZES].reverse().find((entry) => entry < size) ?? Math.max(1, size - 1);
      return style({ fontSize: next });
    }
    case 'textColor': return style({ textColor: bound.color });
    case 'fillColor': return style({ fillColor: bound.color });
    case 'horizontalAlignment': return style({ horizontalAlignment: bound.value });
    case 'verticalAlignment': return style({ verticalAlignment: bound.value });
    case 'textWrapping': return style({ textWrapping: args === null ?
      formatting.textWrapping === 'wrap' ? 'overflow' : 'wrap' : bound.value });
    case 'borderPreset': return style({ border: { preset: bound.value,
      style: formatting.borderStyle ?? 'solid', color: formatting.borderColor ?? '#000000' } });
    case 'borderStyle': return style({ border: { style: bound.value } });
    case 'borderColor': return style({ border: { color: bound.color } });
    case 'numberFormat': return { method: 'setNumberFormat', args: [sheet, range,
      bound.value === 'custom' ? { type: 'custom', pattern: formatting.numberFormatPattern ?? '0.00' } : bound.value] };
    case 'decimalPlaces': return { method: 'setNumberFormat', args: [sheet, range,
      bound.direction === 'increase' ? 'increaseDecimal' : 'decreaseDecimal'] };
    case 'merge': {
      const action = args === null ? merged.length > 0 ? 'unmerge' : 'all' : bound.value as string;
      const count = action === 'horizontal' ? end.row - start.row + 1 :
        action === 'vertical' ? end.col - start.col + 1 : action === 'all' ? 1 : merged.length;
      if (count > MAX_MERGE_OPERATIONS) throw new WorkbookCommandRefusal('limit-exceeded');
      const ops: unknown[] = [];
      if (action === 'all') ops.push({ type: 'mergeCells', sheet, range: { start, end } });
      else if (action === 'horizontal') {
        for (let row = start.row; row <= end.row; row++) {
          ops.push({ type: 'mergeCells', sheet, range: { start: { row, col: start.col }, end: { row, col: end.col } } });
        }
      } else if (action === 'vertical') {
        for (let col = start.col; col <= end.col; col++) {
          ops.push({ type: 'mergeCells', sheet, range: { start: { row: start.row, col }, end: { row: end.row, col } } });
        }
      } else {
        for (const range of merged) ops.push({ type: 'unmergeCells', sheet, range });
      }
      return { method: 'applyOps', args: [ops] };
    }
    default: return undefined;
  }
}
