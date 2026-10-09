import { useState } from 'react';

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
