import type { YrsSession } from '@betteroffice/docx/yrs';

const requests = new WeakMap<YrsSession, () => void>();

export function registerQueuedOpeningInput(session: YrsSession, request: () => void): () => void {
  requests.set(session, request);
  return () => {
    if (requests.get(session) === request) requests.delete(session);
  };
}

export function requestQueuedOpeningInput(session: YrsSession): void {
  const request = requests.get(session);
  if (!request) return;
  requests.delete(session);
  request();
}
