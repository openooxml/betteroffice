import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import { mergeDocxHostMetadata, saveEditorDocument, type EditorSaveRecord } from './editorSave';
import { wrapOpenedEditSession } from './yrsSessionFacade';
import type { EditSession } from './wasm/index';
import { yrsToDocument } from './yrsToDocument';

/** Earlier saves of the worker's document and the projection the next one starts from. @internal */
export interface ResidentSaveRecord extends EditorSaveRecord {
  base?: Document;
}

/**
 * Saves the worker's opened session as the editor saves its own: projected
 * from the last save's projection with `host`'s metadata (the open's when
 * omitted), and `record` updated. @internal
 */
export async function saveResidentDocument(
  raw: EditSession,
  clientId: number,
  source: Uint8Array,
  hostJson: string,
  host: Document | undefined,
  comments: Comment[],
  record: ResidentSaveRecord
): Promise<ArrayBuffer> {
  const opened = wrapOpenedEditSession(raw, clientId, source, hostJson);
  const base = record.base ?? opened.session.materializeDocx();
  if (!base?.originalBuffer) throw new Error('The resident worker holds no opened package');
  const projected = yrsToDocument(
    opened.session,
    mergeDocxHostMetadata(base, host ?? opened.host.document)
  );
  const buffer = await saveEditorDocument(opened.session, projected, comments, record);
  projected.originalBuffer = buffer;
  record.base = projected;
  return buffer;
}
