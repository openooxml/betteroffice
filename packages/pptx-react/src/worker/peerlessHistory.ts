import type { PresentationCommandAck } from '@betteroffice/pptx';
import type { KeyboardEvent } from 'react';
import { commitComposition, type PeerlessViewStore } from './peerlessViewStore';

export function peerlessHistory(store: PeerlessViewStore, direction: 'undo' | 'redo',
  inputTime = performance.now()): Promise<PresentationCommandAck> | null {
  if (store.getSnapshot().editState.readOnly) return null;
  void commitComposition(store, '', inputTime).catch(() => {});
  return store.dispatch({ kind: direction, payload: { inputTime, undoGroup: 'history' } });
}

export function peerlessHistoryKey(event: KeyboardEvent<HTMLDivElement>, store: PeerlessViewStore): void {
  if (event.defaultPrevented || event.target !== event.currentTarget || event.altKey ||
    !(event.ctrlKey || event.metaKey)) return;
  const key = event.key.toLowerCase();
  const direction = key === 'z' ? event.shiftKey ? 'redo' : 'undo' : key === 'y' && !event.shiftKey ? 'redo' : null;
  if (!direction) return;
  event.preventDefault();
  if (!event.repeat) void peerlessHistory(store, direction, event.timeStamp)?.catch(() => {});
}
