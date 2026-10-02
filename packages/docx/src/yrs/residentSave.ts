/** The editor's save of a resident worker's opened document, run in the worker. @internal */

import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import { mergeDocxHostMetadata, saveEditorDocument, type EditorSaveRecord } from './editorSave';
import { wrapOpenedEditSession } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import { yrsToDocument } from './yrsToDocument';

/** The saves of one opened document so far, and the bytes the last one wrote. @internal */
export interface ResidentSaveRecord extends EditorSaveRecord {
  original?: ArrayBuffer;
}

/**
 * Saves `resident`, opened from `source` with the `open` reply `hostJson`, as
 * the editor saves its replica: `host` is {@link hostSaveMetadata} of the
 * editor's document and `record` the earlier saves, updated. @internal
 */
export async function saveResidentDocument(
  resident: ResidentEngineSession,
  source: Uint8Array,
  hostJson: string,
  host: Document,
  comments: Comment[],
  record: ResidentSaveRecord
): Promise<ArrayBuffer> {
  const { session: raw, clientId } = resident.editSession();
  const session = wrapOpenedEditSession(raw, clientId, source, hostJson);
  const materialized = session.materializeDocx();
  if (!materialized?.originalBuffer) throw new Error('The resident worker holds no opened package');
  const base = record.original ? { ...materialized, originalBuffer: record.original } : materialized;
  const projected = yrsToDocument(session, mergeDocxHostMetadata(base, host));
  const buffer = await saveEditorDocument(session, projected, comments, record);
  record.original = buffer;
  return buffer;
}
