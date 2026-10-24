export const LEGACY_SYNC_API_PROP = 'legacySyncApi' as const;

export class PeerlessSyncApiError extends Error {
  readonly code = 'sync-api-unavailable';
  constructor(readonly member: string, alternative = 'the async editor API') {
    super(`${member} requires ${LEGACY_SYNC_API_PROP}={true}. Use ${alternative} instead.`);
    this.name = 'DocxSyncApiError';
  }
}
