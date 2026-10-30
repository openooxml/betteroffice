export class XlsxWorkerEditorChunkLoadError extends Error {
  readonly code = 'chunk-load-failed';
  constructor(readonly cause: unknown) {
    super('Failed to load the editor code. Reload the page to retry.');
    this.name = 'XlsxWorkerEditorChunkLoadError';
  }
}
