export const LEGACY_SYNC_API_PROP = 'legacySyncApi' as const;

export class PeerlessSyncApiError extends Error {
  readonly code = 'sync-api-unavailable';
  constructor(readonly member: string, alternative = 'the async editor API') {
    super(`${member} requires ${LEGACY_SYNC_API_PROP}={true}. Use ${alternative} instead.`);
    this.name = 'DocxSyncApiError';
  }
}

export class DocxEditorNotReadyError extends Error {
  readonly code = 'editor-not-ready';
  constructor() {
    super('The editor is not ready');
    this.name = 'DocxEditorNotReadyError';
  }
}
