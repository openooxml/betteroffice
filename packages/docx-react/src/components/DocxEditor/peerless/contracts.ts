export const LEGACY_SYNC_API_PROP = 'legacySyncApi' as const;

export class PeerlessSyncApiError extends Error {
  constructor(member: string) {
    super(`${member} requires ${LEGACY_SYNC_API_PROP}={true}. Use the async editor API on a peer-less editor.`);
    this.name = 'PeerlessSyncApiError';
  }
}
