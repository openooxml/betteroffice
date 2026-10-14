import { LEGACY_SYNC_API } from '../runtimeOptions';

export class XlsxAsyncApiSyncAccessError extends Error {
  readonly code = 'sync-api-unavailable';
  constructor(readonly member: string, readonly alternative: string) {
    super(`${member} is unavailable with the default async API; use ${alternative} or set ${LEGACY_SYNC_API}={true} for the deprecated sync API`);
    this.name = 'XlsxAsyncApiSyncAccessError';
  }
}
