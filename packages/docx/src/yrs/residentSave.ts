import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import {
  DirtyProjectionStories,
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
}

/** Saves peer marks when supplied, otherwise worker marks; clears worker marks on success. @internal */
export async function saveResidentDocument(
  raw: EditSession,
  clientId: number,
  source: Uint8Array,
  hostJson: string,
  host: Document | undefined,
  comments: Comment[],
  record: ResidentSaveRecord,
  dirty: DirtyProjectionStories,
  stories?: readonly string[]
): Promise<ArrayBuffer> {
  const opened = wrapOpenedEditSession(raw, clientId, source, hostJson);
  const base = record.base ?? opened.session.materializeDocx();
  if (!base?.originalBuffer) throw new Error('The resident worker holds no opened package');
  const pending = dirty.capture();
  const selected = new DirtyProjectionStories();
  for (const story of stories ?? pending.stories) selected.add(story);
  const projected = yrsToDocument(
    opened.session,
    mergeDocxHostMetadata(base, host ?? opened.host.document),
    selected.projectionOptions()
  );
  const buffer = await saveEditorDocument(opened.session, projected, comments, record);
  projected.originalBuffer = buffer;
  record.base = projected;
  pending.clear();
  return buffer;
}
