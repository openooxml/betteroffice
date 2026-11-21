import { MAX_COLLABORATION_FRAME_BYTES } from '../../../../shared/collaboration-limits';
import type { AwarenessCursor } from './types';

export const DEFAULT_MAX_FRAME_BYTES = MAX_COLLABORATION_FRAME_BYTES;
export const OPEN_SYNC_CLIENT = Number.MAX_SAFE_INTEGER - 1;
export const XLSX_MAX_ROWS = 1_048_576;
export const XLSX_MAX_COLUMNS = 16_384;
export const MAX_AWARENESS_SHEET_LENGTH = 256;

export function validAwarenessCursor(value: unknown, _maxFrameBytes?: number): value is AwarenessCursor {
  if (typeof value !== 'object' || value === null) return false;
  const cursor = value as AwarenessCursor;
  const cell = (point: AwarenessCursor['anchor'] | undefined) => !!point &&
    Number.isSafeInteger(point.row) && point.row >= 0 && Number.isSafeInteger(point.col) && point.col >= 0;
  return typeof cursor.sheet === 'string' && cursor.sheet.length > 0 &&
    cursor.sheet.length <= MAX_AWARENESS_SHEET_LENGTH && cell(cursor.anchor) && cell(cursor.head);
}
