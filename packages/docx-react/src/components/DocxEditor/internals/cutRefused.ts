import { clipboardLeftOutInError } from '@betteroffice/docx/yrs';

/** @internal A cut refused because the clipboard cannot carry these selected embed kinds. */
export class DocxCutRefusedError extends Error {
  readonly leftOut: readonly string[];

  constructor(leftOut: readonly string[]) {
    super(`Cut refused: the clipboard cannot carry ${leftOut.join(', ')}`);
    this.name = 'DocxCutRefusedError';
    this.leftOut = leftOut;
  }
}

/** @internal The refusal a failed cut read stands for, or null when the error is about something else. */
export function cutRefusalFrom(error: unknown): DocxCutRefusedError | null {
  if (error instanceof DocxCutRefusedError) return error;
  const leftOut = clipboardLeftOutInError(error);
  return leftOut ? new DocxCutRefusedError(leftOut) : null;
}
