import { DocumentToolError } from './types';

export function editLimit(value = 32): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > 500) throw new DocumentToolError('INVALID_ARGUMENT', 'maxEdits must be an integer from 1 to 500.');
  return value;
}
