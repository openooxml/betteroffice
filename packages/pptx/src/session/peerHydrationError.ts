export class PptxPeerHydrationError extends Error {
  readonly refusal: { code: string };
  constructor(readonly code: string, message: string, readonly cause?: unknown) {
    super(message);
    this.name = 'PptxPeerHydrationError';
    this.refusal = { code };
  }
}

export class PptxPeerNotReadyError extends Error {
  readonly code = 'peer-not-ready';
  constructor() {
    super('Presentation editor is not ready');
    this.name = 'PptxPeerNotReadyError';
  }
}

export class PptxWorkerEditorFailedError extends Error {
  readonly code = 'editor-failed';
  constructor(readonly cause: unknown) {
    super(`Presentation worker editor failed: ${cause instanceof Error ? cause.message : String(cause)}`);
    this.name = 'PptxWorkerEditorFailedError';
  }
}

export class PptxWorkerEditorDisposedError extends Error {
  readonly code = 'editor-disposed';
  constructor() {
    super('Presentation worker editor is disposed');
    this.name = 'PptxWorkerEditorDisposedError';
  }
}

export class PptxWorkerEditorCollaborationError extends Error {
  readonly code = 'collaboration-unavailable';
  constructor() {
    super('Collaboration is unavailable in the presentation worker editor');
    this.name = 'PptxWorkerEditorCollaborationError';
  }
}
