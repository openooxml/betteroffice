import type { Comment } from '@betteroffice/docx/types/content';
import type { YrsSession } from '@betteroffice/docx/yrs';

export interface WorkerOpenSave {
  available(): boolean;
  save(comments: Comment[], peer?: YrsSession): Promise<ArrayBuffer>;
}

const savers = new WeakMap<YrsSession, WorkerOpenSave>();

export function registerWorkerOpenSave(session: YrsSession, save: WorkerOpenSave): () => void {
  savers.set(session, save);
  return () => {
    if (savers.get(session) === save) savers.delete(session);
  };
}

export function workerOpenSave(session: YrsSession): WorkerOpenSave | null {
  return savers.get(session) ?? null;
}
