import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { DocxDisplayRange, DocxFindDisplayMatch, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { FindOptions } from '@betteroffice/docx/utils/findReplace';
import type { ScrollToParaIdOptions } from '@betteroffice/docx/utils';
import type { DocxSelectionInfo } from '../../DocxEditor';
import { isPresented, presentedWorkerVersion } from './layoutProvenance';
import { readAt } from './viewerReads';

export type ViewerNavigationTarget =
  | { kind: 'paragraphTarget'; paraId: string }
  | { kind: 'commentTarget'; commentId: string }
  | { kind: 'revisionTarget'; revisionId: string };

export interface ViewerRefReadAccess {
  read: ResidentEngineWorkerClient['documentRead'];
  story: string;
  host(): HTMLElement | null | undefined;
  queries(): DisplayListQueries | null | undefined;
  awaitFrame(previous: DisplayListQueries | null | undefined, timeoutMs: number): Promise<DisplayListQueries | null | undefined>;
  current(): boolean;
  selection(): DocxDisplayRange | null;
}

async function beforeDeadline<T>(pending: Promise<T>, deadline: number): Promise<T | null> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([pending, new Promise<null>((resolve) => {
      timer = setTimeout(() => resolve(null), Math.max(0, deadline - Date.now()));
    })]);
  } finally {
    clearTimeout(timer);
  }
}

function frameVersion(access: ViewerRefReadAccess, queries: DisplayListQueries | null | undefined): string | null {
  return queries && isPresented(access.host(), queries.displayList) ? presentedWorkerVersion(queries) : null;
}

export async function readViewerSelectionInfo(access: ViewerRefReadAccess): Promise<DocxSelectionInfo | null> {
  const deadline = Date.now() + 10_000;
  for (let attempt = 0; attempt < 5 && access.current(); attempt += 1) {
    const selection = access.selection();
    const queries = access.queries();
    const expectVersion = frameVersion(access, queries);
    if (!selection || expectVersion === null) return null;
    const outcome = await beforeDeadline(readAt(access.read, {
      kind: 'selectionInfo', story: access.story, ...selection, expectVersion,
    }), deadline);
    if (!access.current() || !outcome) return null;
    const next = access.selection();
    if (!next) return null;
    if (next.anchor !== selection.anchor || next.head !== selection.head) continue;
    if (outcome.status === 'ok' && frameVersion(access, access.queries()) === expectVersion) return outcome.value;
    if (attempt === 4 || !await beforeDeadline(access.awaitFrame(queries, deadline - Date.now()), deadline)) return null;
  }
  return null;
}

export async function navigateViewer(
  access: ViewerRefReadAccess,
  target: ViewerNavigationTarget,
  apply: (range: DocxDisplayRange, options?: ScrollToParaIdOptions) => void,
  options?: ScrollToParaIdOptions
): Promise<boolean> {
  const deadline = Date.now() + 10_000;
  for (let attempt = 0; attempt < 5 && access.current(); attempt += 1) {
    const queries = access.queries();
    const expectVersion = frameVersion(access, queries);
    if (expectVersion === null) return false;
    const outcome = await beforeDeadline(readAt(access.read, { ...target, story: access.story, expectVersion }), deadline);
    if (!access.current() || !outcome) return false;
    if (outcome.status === 'ok' && frameVersion(access, access.queries()) === expectVersion) {
      if (!outcome.value) return false;
      apply(outcome.value, options);
      return true;
    }
    if (attempt === 4 || !await beforeDeadline(access.awaitFrame(queries, deadline - Date.now()), deadline)) return false;
  }
  return false;
}

export async function readViewerFindMatches(
  access: ViewerRefReadAccess,
  searchText: string,
  options: FindOptions
): Promise<{ version: string; matches: DocxFindDisplayMatch[] } | null> {
  const deadline = Date.now() + 10_000;
  for (let attempt = 0; attempt < 5 && access.current(); attempt += 1) {
    const queries = access.queries();
    const expectVersion = frameVersion(access, queries);
    if (expectVersion === null) return null;
    const outcome = await beforeDeadline(readAt(access.read, {
      kind: 'findMatches', searchText, options, expectVersion,
    }), deadline);
    if (!access.current() || !outcome) return null;
    if (outcome.status === 'ok' && outcome.value && frameVersion(access, access.queries()) === expectVersion) {
      return { version: expectVersion, matches: outcome.value };
    }
    if (attempt === 4 || !await beforeDeadline(access.awaitFrame(queries, deadline - Date.now()), deadline)) return null;
  }
  return null;
}
