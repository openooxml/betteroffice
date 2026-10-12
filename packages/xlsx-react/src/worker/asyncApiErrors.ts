export class XlsxAsyncApiSyncAccessError extends Error {
  readonly code = 'sync-api-unavailable';
  constructor(readonly member: string, readonly alternative: string) {
    super(`${member} is unavailable with asyncApi; use ${alternative}`);
    this.name = 'XlsxAsyncApiSyncAccessError';
  }
}
