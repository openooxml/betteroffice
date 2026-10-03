import { useState } from 'react';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { presentedWorkerVersion } from './layoutProvenance';
import { workerOpenReplicaPending } from './workerOpenReplica';

/**
 * Whether the loaded document is a viewer session: opened read-only in the worker, so it has no
 * main-thread edit peer. A session opened for editing, or switched to editing or out of worker-open
 * once, keeps the editing input when it is switched to viewing.
 */
export function useViewerSession(workerOpen: boolean, workerProposals: boolean, generation: number): boolean {
  const [opened, setOpened] = useState({ generation, viewer: workerOpen && workerProposals });
  if (opened.generation !== generation) {
    const next = { generation, viewer: workerOpen && workerProposals };
    setOpened(next);
    return next.viewer;
  }
  if (opened.viewer && !(workerOpen && workerProposals)) {
    setOpened({ generation, viewer: false });
    return false;
  }
  return opened.viewer;
}

/** Whether a viewer session still reads from the worker: false once its document fell back here. */
export function viewerReadsWorker(
  queries: { readonly displayList: object } | null | undefined,
  session: YrsSession | null | undefined
): boolean {
  return !(queries && presentedWorkerVersion(queries) === null && session && !workerOpenReplicaPending(session));
}
