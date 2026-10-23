export const LEGACY_SYNC_API_PROP = 'legacySyncApi' as const;

export class PeerlessSyncApiError extends Error {
  readonly code = 'sync-api-unavailable';
  constructor(readonly member: string) {
    super(`${member} requires ${LEGACY_SYNC_API_PROP}={true}. Use the async editor API on the worker-only editor.`);
    this.name = 'DocxSyncApiError';
  }
}
