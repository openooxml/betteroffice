import type { PagedEditorRef } from '../PagedEditor';

export function isWorkerViewer(editor: PagedEditorRef | null | undefined): boolean {
  return editor?.isWorkerViewer?.() === true;
}
