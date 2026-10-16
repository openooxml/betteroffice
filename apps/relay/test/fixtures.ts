import * as Y from 'yjs';
import { decodeMessages, encodeUpdate } from '../../../packages/docx/src/collaboration/protocol';

export function documentFrame(text = 'hello', client = 17) {
  const doc = new Y.Doc();
  doc.clientID = client;
  doc.getText('body').insert(0, text);
  const frame = encodeUpdate(Y.encodeStateAsUpdate(doc));
  doc.destroy();
  return frame;
}

export function rehydrate(frames: Uint8Array[]) {
  const doc = new Y.Doc();
  for (const frame of frames) for (const message of decodeMessages(frame)) {
    if (message.type === 'update' || message.type === 'sync-step-2') Y.applyUpdate(doc, message.update);
  }
  return doc;
}
