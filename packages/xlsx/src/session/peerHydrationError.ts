export type WorkbookPeerHydrationErrorCode =
  | 'mutation-outside-replay'
  | 'version-mismatch'
  | 'missing-hydration'
  | 'missing-module';

export class WorkbookPeerHydrationError extends Error {
  readonly refusal: { code: WorkbookPeerHydrationErrorCode };

  constructor(readonly code: WorkbookPeerHydrationErrorCode, message: string) {
    super(message);
    this.name = 'WorkbookPeerHydrationError';
    this.refusal = { code };
  }
}
