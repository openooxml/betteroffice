export class PptxAsyncEditorUnavailableError extends Error {
  readonly code = 'async-editor-unavailable';
  constructor(message: string) {
    super(message);
    this.name = 'PptxAsyncEditorUnavailableError';
  }
}
