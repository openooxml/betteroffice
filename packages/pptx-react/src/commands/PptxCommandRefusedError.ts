import type { PresentationCommandIntent, PresentationCommandOutcome } from '@betteroffice/pptx/worker';

/** A refused input that left the document unchanged and the editor usable. */
export class PptxCommandRefusedError extends Error {
  override readonly name = 'PptxCommandRefusedError';
  readonly code = 'command-refused';
  constructor(readonly operation: PresentationCommandIntent['kind'], readonly outcome: PresentationCommandOutcome) {
    super(`Presentation ${operation} refused: ${outcome.reason}`);
  }
}
