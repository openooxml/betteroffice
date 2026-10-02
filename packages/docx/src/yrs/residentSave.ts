/** The editor's save of a resident worker's opened document, run in the worker. @internal */

import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import { mergeDocxHostMetadata, saveEditorDocument, type EditorSaveRecord } from './editorSave';
import { wrapOpenedEditSession } from './yrsSessionFacade';
import type { EditSession } from './wasm/index';
import { yrsToDocument } from './yrsToDocument';

/**
 * The saves of one opened document so far, and the last one's projection: the
 * base the next save projects from, as the editor's cached projection is. @internal
 */
export interface ResidentSaveRecord extends EditorSaveRecord {
  base?: Document;
}

/**
 * Saves a resident session's edit session `raw`, opened from `source` with the
 * `open` reply `hostJson`, as the editor saves its replica: `host` is
 * {@link hostSaveMetadata} of the editor's document and `record` the earlier
 * saves, updated. @internal
 */
export async function saveResidentDocument(
  raw: EditSession,
  clientId: number,
  source: Uint8Array,
  hostJson: string,
  host: Document,
  comments: Comment[],
  record: ResidentSaveRecord
): Promise<ArrayBuffer> {
  const session = wrapOpenedEditSession(raw, clientId, source, hostJson);
  const base = record.base ?? session.materializeDocx();
  if (!base?.originalBuffer) throw new Error('The resident worker holds no opened package');
  const projected = yrsToDocument(session, mergeDocxHostMetadata(base, host));
  const buffer = await saveEditorDocument(session, projected, comments, record);
  projected.originalBuffer = buffer;
  record.base = projected;
  return buffer;
}
