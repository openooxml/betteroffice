import type { DocxCommandPayloads } from '@betteroffice/docx/yrs';
import type { PeerlessSelectionIntent } from './selection';

/** @internal Relative mutation payloads omit main-side selection positions. */
export type PeerlessTextIntent = {
  [K in 'insertText' | 'replaceSelection' | 'deleteBackward' | 'deleteForward' | 'paragraphBreak']:
    { kind: K; payload: DocxCommandPayloads[K] };
}['insertText' | 'replaceSelection' | 'deleteBackward' | 'deleteForward' | 'paragraphBreak'];

/** @internal */
export type PeerlessPartIntent = { [K in 'openPart' | 'ensurePart' | 'removePart']: { kind: K; payload: DocxCommandPayloads[K] } }['openPart' | 'ensurePart' | 'removePart'];

type PeerlessRulerKind = 'setSectionMargins' | 'setParagraphAttrs' | 'setTabStops' | 'removeTabStop';
type PeerlessRulerIntent = { [K in PeerlessRulerKind]: { kind: K; payload: DocxCommandPayloads[K] } }[PeerlessRulerKind];

type PeerlessPropertyKind = 'setHyperlink' | 'setObjectProperties' | 'setSectionProperties' | 'setNoteProperties' | 'setWatermark' | 'resizeObject' | 'moveObject';
export type PeerlessPropertyIntent = { [K in PeerlessPropertyKind]: { kind: K; payload: DocxCommandPayloads[K] } }[PeerlessPropertyKind];

type PeerlessControlKind = 'setContentControlValue' | 'addRepeatingItem' | 'removeRepeatingItem';
export type PeerlessControlIntent = { [K in PeerlessControlKind]: { kind: K; payload: DocxCommandPayloads[K] } }[PeerlessControlKind];

type PeerlessTableKind = 'insertTableRow' | 'insertTableColumn' | 'deleteTableRow' | 'deleteTableColumn' | 'deleteTable' | 'mergeTableCells' | 'splitTableCell' | 'setTableCellShading' | 'setTableBorders' | 'setTableProperties';
export type PeerlessTableIntent = { [K in PeerlessTableKind]: { kind: K; payload: K extends 'setTableProperties' ? DocxCommandPayloads[K] : Omit<DocxCommandPayloads[K], 'selection'> } }[PeerlessTableKind];

export type PeerlessInputIntent = PeerlessProposalIntent | PeerlessRevisionIntent | { [K in 'applyEdits' | 'addComment' | 'updateComment' | 'removeComment']: { kind: K; payload: DocxCommandPayloads[K] } }['applyEdits' | 'addComment' | 'updateComment' | 'removeComment'] | PeerlessControlIntent | PeerlessRulerIntent | PeerlessPropertyIntent | PeerlessPartIntent | PeerlessSelectionIntent | PeerlessTextIntent | PeerlessTableIntent | { kind: 'undo' | 'redo'; payload: Record<string, never> } | { kind: 'replaceMatches'; payload: DocxCommandPayloads['replaceMatches'] } | { [K in 'setRenderEnvironment' | 'setFontEnvironment']: { kind: K; payload: DocxCommandPayloads[K] } }['setRenderEnvironment' | 'setFontEnvironment'] | { [K in 'insertImage' | 'insertTable' | 'insertPageBreak' | 'insertSectionBreak' | 'pasteFragment']: { kind: K; payload: Omit<DocxCommandPayloads[K], 'selection'> } }['insertImage' | 'insertTable' | 'insertPageBreak' | 'insertSectionBreak' | 'pasteFragment'];

type PeerlessProposalKind = 'propose' | 'setProposalStates' | 'withdrawProposals';
export type PeerlessProposalIntent = { [K in PeerlessProposalKind]: { kind: K; payload: DocxCommandPayloads[K] } }[PeerlessProposalKind];
/** Without a revisionId the decision targets the change at the selection its predecessor leaves. */
export type PeerlessRevisionIntent = { [K in 'acceptRevision' | 'rejectRevision']: { kind: K; payload: Partial<Omit<DocxCommandPayloads[K], 'expectVersion'>> } }['acceptRevision' | 'rejectRevision'];

const implicitVerticalGoals = new WeakSet<PeerlessInputIntent>();

export function hasImplicitVerticalGoal(intent: PeerlessInputIntent): boolean {
  return implicitVerticalGoals.has(intent);
}

/** @internal Native-event ownership and trailing-composition deduplication state. */
export interface PeerlessInputEventState {
  composing: boolean;
  compositionId: number | null;
  lastHandledEventId: number | null;
  trailingCompositionText: string | null;
}

/** @internal The caller gives overlapping browser events the same eventId. */
export type PeerlessInputEvent = { eventId: number; inputTime: number } & (
  | { source: 'beforeinput'; inputType: string; data: string | null; isComposing: boolean }
  | { source: 'keydown'; key: string; code: string; shift: boolean; alt: boolean; ctrl: boolean; meta: boolean; goalX?: number }
  | { source: 'input'; inputType: string; value: string; isComposing: boolean }
  | { source: 'paste'; text: string }
);

/** @internal */
export interface PeerlessNormalizedInput {
  state: PeerlessInputEventState;
  intents: readonly PeerlessInputIntent[];
  preventDefault: boolean;
  fence: boolean;
}

/** @internal */
export function normalizeInputEvent(
  _event: PeerlessInputEvent,
  _state: PeerlessInputEventState
): PeerlessNormalizedInput {
  const event = _event;
  const state = _state;
  const result = (intents: PeerlessInputIntent[], fence = false): PeerlessNormalizedInput => ({
    state: intents.length ? { ...state, lastHandledEventId: event.eventId, trailingCompositionText: null } : state,
    intents, preventDefault: intents.length > 0, fence,
  });
  if (state.composing || ('isComposing' in event && event.isComposing)) return result([]);
  if (state.lastHandledEventId === event.eventId) {
    return { state, intents: [], preventDefault: true, fence: false };
  }
  if (state.trailingCompositionText !== null &&
    ((event.source === 'beforeinput' && event.data === state.trailingCompositionText &&
      /^(insertText|insertFromComposition|insertCompositionText)$/.test(event.inputType)) ||
    (event.source === 'input' && event.value === state.trailingCompositionText))) {
    return { state: { ...state, lastHandledEventId: event.eventId, trailingCompositionText: null },
      intents: [], preventDefault: true, fence: true };
  }
  if (event.source === 'paste') {
    return result(pasteIntents(event.text), true);
  }
  if (event.source === 'keydown') {
    const { key, shift: extend, alt, ctrl, meta } = event;
    if ((ctrl || meta) && !alt && (key.toLowerCase() === 'z' || key.toLowerCase() === 'y')) return result([{ kind: key.toLowerCase() === 'y' || extend ? 'redo' : 'undo', payload: {} }], true);
    if ((ctrl || meta) && key.toLowerCase() === 'a') return result([
      { kind: 'navigate', payload: { unit: 'document-boundary', direction: 'backward', extend: false } },
      { kind: 'navigate', payload: { unit: 'document-boundary', direction: 'forward', extend: true } },
    ], true);
    if (key === 'ArrowLeft' || key === 'ArrowRight') return result([{ kind: 'navigate', payload: {
      unit: meta ? 'line' : ctrl || alt ? 'word' : 'character',
      direction: key === 'ArrowLeft' ? 'backward' : 'forward', extend,
    } }], true);
    if (key === 'ArrowUp' || key === 'ArrowDown') {
      const intent: PeerlessInputIntent = { kind: 'navigate', payload:
        meta || ctrl || alt ? {
          unit: meta ? 'document-boundary' : 'paragraph',
          direction: key === 'ArrowUp' ? 'backward' : 'forward', extend,
        } : { unit: 'line', direction: key === 'ArrowUp' ? 'up' : 'down', extend, goalX: event.goalX ?? 0 },
      };
      if (!meta && !ctrl && !alt && event.goalX === undefined) implicitVerticalGoals.add(intent);
      return result([intent], true);
    }
    if (key === 'Home' || key === 'End') return result([{ kind: 'navigate', payload: {
      unit: ctrl || meta ? 'document-boundary' : 'line',
      direction: key === 'Home' ? 'backward' : 'forward', extend,
    } }], true);
    if (key === 'Backspace' || key === 'Delete') {
      const backward = key === 'Backspace';
      const deletion: PeerlessTextIntent = { kind: backward ? 'deleteBackward' : 'deleteForward', payload: {} };
      return result(ctrl || alt || meta ? [{ kind: 'navigate', payload: {
        unit: meta ? 'line' : 'word', direction: backward ? 'backward' : 'forward', extend: true,
      } }, deletion] : [deletion], ctrl || alt || meta);
    }
    if (key === 'Enter' && !ctrl && !meta && !alt) return result([{ kind: 'paragraphBreak', payload: {} }], true);
    return result([]);
  }
  const text = event.source === 'beforeinput' ? event.data : event.value;
  switch (event.inputType) {
    case 'historyUndo': return result([{ kind: 'undo', payload: {} }], true);
    case 'historyRedo': return result([{ kind: 'redo', payload: {} }], true);
    case 'insertText': case 'insertFromComposition':
      return text ? result([{ kind: 'insertText', payload: { text } }]) : result([]);
    case 'insertFromPaste': case 'insertFromDrop':
      return result(pasteIntents(text ?? ''), true);
    case 'insertReplacementText':
      return text ? result([{ kind: 'replaceSelection', payload: { text } }], true) : result([]);
    case 'insertParagraph': case 'insertLineBreak':
      return result([{ kind: 'paragraphBreak', payload: {} }], true);
    case 'deleteContentBackward':
      return result([{ kind: 'deleteBackward', payload: {} }]);
    case 'deleteContentForward': case 'deleteByCut':
      return result([{ kind: 'deleteForward', payload: {} }], event.inputType === 'deleteByCut');
    case 'deleteWordBackward': case 'deleteWordForward':
    case 'deleteSoftLineBackward': case 'deleteSoftLineForward': {
      const backward = event.inputType.endsWith('Backward');
      return result([
        { kind: 'navigate', payload: { unit: event.inputType.startsWith('deleteWord') ? 'word' : 'line',
          direction: backward ? 'backward' : 'forward', extend: true } },
        { kind: backward ? 'deleteBackward' : 'deleteForward', payload: {} },
      ], true);
    }
    default:
      return event.source === 'input' && text && !event.inputType
        ? result([{ kind: 'insertText', payload: { text } }]) : result([]);
  }
}

function pasteIntents(text: string): PeerlessInputIntent[] {
  if (!text) return [];
  const lines = text.split(/\r\n|[\r\n\u000B\u2028\u2029]/u);
  const intents: PeerlessInputIntent[] = [{ kind: 'replaceSelection', payload: { text: lines[0]! } }];
  for (const line of lines.slice(1)) {
    intents.push({ kind: 'paragraphBreak', payload: {} });
    if (line) intents.push({ kind: 'insertText', payload: { text: line } });
  }
  return intents;
}
