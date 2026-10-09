export type DocxWorkerErrorStage = 'open' | 'layout' | 'render';

const STAGE_TEXT: Record<DocxWorkerErrorStage, string> = {
  open: 'opening',
  layout: 'laying out',
  render: 'rendering',
};

/**
 * The document worker failed and a fresh worker could not take over. The editor shows the error
 * instead of the pages and does not open the document on the main thread; the original failure is
 * the `cause`.
 */
export class DocxWorkerError extends Error {
  constructor(readonly stage: DocxWorkerErrorStage, cause?: unknown) {
    super(
      `The document worker failed while ${STAGE_TEXT[stage]} the document` +
        (cause instanceof Error && cause.message ? `: ${cause.message}` : ''),
      { cause }
    );
    this.name = 'DocxWorkerError';
  }
}
