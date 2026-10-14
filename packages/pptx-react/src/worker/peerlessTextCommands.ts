import type { PresentationCommandAck, TextStyle } from '@betteroffice/pptx';
import type { PptxTextEnvironment } from '../commands/evaluate';
import type { FormattingAction } from '../components/toolbarTypes';
import { peerlessInputReference, type PeerlessViewSnapshot, type PeerlessViewStore } from './peerlessViewStore';

export function peerlessTextContext(view: PeerlessViewSnapshot): PptxTextEnvironment | null {
  if (!view.selection || !view.editState.toolbar.canEditText) return null;
  const { formatting, alignment } = view.editState.toolbar;
  return {
    kind: view.selection.kind === 'shape' || view.selection.kind === 'object' ? 'shape'
      : view.selection.anchor === view.selection.focus ? 'caret' : 'range',
    bold: formatting?.bold ?? 'mixed', italic: formatting?.italic ?? 'mixed',
    underline: formatting?.underline == null ? 'mixed' : formatting.underline !== 'none',
    fontFamily: formatting?.fontFamily ?? null, fontSize: formatting?.fontSizePt ?? null,
    color: formatting?.color ?? null, alignment,
  };
}

export function formatPeerlessText(store: PeerlessViewStore, action: FormattingAction |
  { type: 'fontSizeStep'; direction: 'increase' | 'decrease'; sizes: readonly number[] },
  inputTime = performance.now()): Promise<PresentationCommandAck> | null {
  const view = store.getSnapshot();
  const at = peerlessInputReference(store, true);
  if (!at || view.editState.readOnly || view.composition) return null;
  const payload = { at, inputTime, undoGroup: 'format' };
  if (typeof action !== 'string' && action.type === 'fontSizeStep') {
    return store.dispatch({ kind: 'formatText', payload: { ...payload, patch: {},
      fontSizeStep: { direction: action.direction, sizes: [...action.sizes] } } });
  }
  if (typeof action !== 'string' && action.type === 'align') {
    return store.dispatch({ kind: 'setParagraphAlignment', payload: { ...payload, alignment: action.value } });
  }
  const formatting = view.editState.toolbar.formatting;
  let patch: TextStyle;
  switch (action) {
    case 'bold': patch = { bold: formatting?.bold !== true }; break;
    case 'italic': patch = { italic: formatting?.italic !== true }; break;
    case 'underline': patch = { underline: formatting?.underline != null && formatting.underline !== 'none' ? 'none' : 'sng' }; break;
    default:
      switch (action.type) {
        case 'fontFamily': patch = { fontFamily: action.value }; break;
        case 'fontSize': patch = { fontSizePt: action.value }; break;
        case 'textColor': patch = { color: action.value }; break;
        default: return null;
      }
  }
  return store.dispatch({ kind: 'formatText', payload: { ...payload, patch,
    ...(typeof action === 'string' ? { toggle: action } : {}) } });
}
