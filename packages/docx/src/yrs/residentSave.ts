import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import {
  dirtyProjectionStory,
  mergeDocxHostMetadata,
  saveEditorDocument,
  type EditorSaveRecord,
} from './editorSave';
import { wrapOpenedEditSession } from './yrsSessionFacade';
import type { EditSession } from './wasm/index';
import { yrsToDocument } from './yrsToDocument';

/** Earlier saves of the worker's document and the projection the next one starts from. @internal */
export interface ResidentSaveRecord extends EditorSaveRecord {
  base?: Document;
  /** Revision `base` reflects; changes since are projected again. */
  revision?: number;
}

/**
 * Saves the worker's opened session as the editor saves its own: the stories
 * changed since the last save projected over that save's projection (every
 * story when none are dirty), with `host`'s metadata (the open's when omitted),
 * and `record` updated. Revisions come from `storiesChangedSince`, the resident
 * session's revision stream, with comment ID changes also projecting the body. @internal
 */
export async function saveResidentDocument(
  raw: EditSession,
  clientId: number,
  storiesChangedSince: (since: number) => { revision: number; stories: string[]; comments?: boolean },
  source: Uint8Array,
  hostJson: string,
  host: Document | undefined,
  comments: Comment[],
  record: ResidentSaveRecord
): Promise<ArrayBuffer> {
  const opened = wrapOpenedEditSession(raw, clientId, source, hostJson);
  const base = record.base ?? opened.session.materializeDocx();
  if (!base?.originalBuffer) throw new Error('The resident worker holds no opened package');
  const changes = record.revision === undefined ? undefined : storiesChangedSince(record.revision);
  const storyIds = new Set(changes?.stories.map(dirtyProjectionStory) ?? []);
  if (changes?.comments === true) {
    storyIds.add(dirtyProjectionStory('body'));
  }
  const projected = yrsToDocument(
    opened.session,
    mergeDocxHostMetadata(base, host ?? opened.host.document),
    storyIds.size > 0 ? { storyIds } : undefined
  );
  const buffer = await saveEditorDocument(opened.session, projected, comments, record);
  projected.originalBuffer = buffer;
  record.base = projected;
  record.revision = storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
  return buffer;
}
