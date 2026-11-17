export class SuggestedFormattingOverContainerError extends Error {
  readonly code = 'suggested-formatting-over-container';

  constructor(message: string) {
    super(message);
    this.name = 'SuggestedFormattingOverContainerError';
  }
}

export function formattingFailure(error: unknown): unknown {
  if (error instanceof Error && 'code' in error && error.code === 'suggested-formatting-over-container') {
    return new SuggestedFormattingOverContainerError(error.message);
  }
  return error;
}
