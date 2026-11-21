import { isKeyReleasedFromInput, takeClipboardGesture } from '../../../../../../shared/keyboard/nativeGesture';
import { Suspense, lazy, useCallback, useLayoutEffect, useRef, useSyncExternalStore, type ReactElement, type Ref, type RefObject } from 'react';
import { effectiveZoom } from '@betteroffice/docx/layout/render';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import type { DocxClipboardFragment, DocxCommandSelection } from '@betteroffice/docx/yrs';
import { scrollIntoViewDelta, scrollViewport } from '../internals/viewportBand';
import { parseClipboardHtml } from '../internals/clipboardHtml';
import { clipboardFragmentPlainText, pasteOptionsOffered, plainTextClipboardFragment } from '../pasteOptions/pasteText';
import { mergeFormattingFragment, type PasteFormattingOption } from '../pasteOptions/clipboardFragmentTransforms';
import { beginComposition, updateComposition, endComposition, type PeerlessCompositionEndResult, type PeerlessCompositionState } from './composition';
import type { CommandTicket } from '../../../../../../shared/office-session/commands/types';
import { admitInputIntents, admitIntent, appliedTicket, submitFormattingIntent, type PeerlessAdmissionResult, type PeerlessInputAdmission } from './inputAdmission';
import { normalizeInputEvent, selectsAll, type PeerlessInputEvent, type PeerlessInputEventState } from './inputEvents';
import type { PeerlessViewStore } from './viewStore';
import { useTranslation } from '../../../i18n';
import { usePasteOptions } from '../pasteOptions/usePasteOptions';

const LazyPasteOptionsControl = lazy(() => import('../pasteOptions/PasteOptionsControl'));

/** @internal Records a submitted fragment paste so Paste Options can offer it once the worker applied that paste. */
export type PeerlessPasteRecorder = (fragment: DocxClipboardFragment, submission: PeerlessAdmissionResult) => void;

type PastePoint = DocxCommandSelection['focus'];

type PasteRecord = {
  fragment: DocxClipboardFragment;
  offered: readonly PasteFormattingOption[];
  ticket: CommandTicket | null;
  end: PastePoint | null;
  version: number;
  incarnation: number;
  reveal: 'show' | 'anchor' | null;
};

function samePoint(a: PastePoint, b: PastePoint): boolean {
  return a.story === b.story && a.paraId === b.paraId && a.offset === b.offset;
}

function caretPoint(selection: DocxCommandSelection | null | undefined): PastePoint | null {
  return selection && samePoint(selection.anchor, selection.focus) ? selection.focus : null;
}

/** @internal */
export interface PeerlessInputProps {
  admission: PeerlessInputAdmission;
  viewStore: PeerlessViewStore;
  textareaRef?: Ref<HTMLTextAreaElement>;
  autoFocus?: boolean;
  className?: string;
  ariaLabel?: string;
  onCompositionChange?(state: PeerlessCompositionState | null): void;
  onRecoverableText?(text: string): void;
  /** Fills `data` synchronously when the selection's clipboard payload is ready. */
  onClipboard?(cut: boolean, data?: DataTransfer | null): Promise<void>;
  onPasteFragment?(fragment: DocxClipboardFragment, inputTime: number): boolean;
  pasteOptions?: boolean;
  onReplacePasteFragment?(fragment: DocxClipboardFragment, selection: DocxCommandSelection, variant?: 'plain' | 'merge'): PeerlessAdmissionResult | false;
  pasteRecorder?: RefObject<PeerlessPasteRecorder | null>;
  caretHostRef?: RefObject<HTMLDivElement | null>;
  zoom?: number;
  onNavigateCell?(direction: 'forward' | 'backward'): void;
}

/** @internal */
export function PeerlessInput(props: PeerlessInputProps): ReactElement | null {
  const { t } = useTranslation();
  const { admission, viewStore, textareaRef, autoFocus, className, ariaLabel } = props;
  const textarea = useRef<HTMLTextAreaElement | null>(null);
  const composing = useRef(false);
  const scrolledBeforeFirstCaret = useRef(false);
  useLayoutEffect(() => {
    const keys = new Set(['PageDown', 'PageUp', ' ', 'ArrowDown', 'ArrowUp', 'Home', 'End']);
    const onIntent = (event: Event) => {
      const host = props.caretHostRef?.current;
      if (event.type === 'keydown' && (!keys.has((event as KeyboardEvent).key) || event.target instanceof HTMLTextAreaElement)) return;
      const target = event.target;
      if (host && target instanceof Node && (target.contains(host) || host.contains(target))) scrolledBeforeFirstCaret.current = true;
    };
    const types = ['wheel', 'touchmove', 'pointerdown', 'keydown', 'scroll'] as const;
    for (const type of types) document.addEventListener(type, onIntent, { capture: true, passive: true });
    return () => { for (const type of types) document.removeEventListener(type, onIntent, true); };
  }, [props.caretHostRef]);
  const revealed = useRef<string | undefined>(undefined);
  const revealHold = useRef<{ seq?: number } | null>(null);
  const callbacks = useRef(props);
  callbacks.current = props;
  const getReadOnly = useCallback(() => admission.getSnapshot().readOnly, [admission]);
  const readOnly = useSyncExternalStore(admission.subscribeReadOnly, getReadOnly, getReadOnly);
  const snapshot = useSyncExternalStore(viewStore.subscribe, viewStore.getSnapshot, viewStore.getServerSnapshot);
  const getRecovering = useCallback(() => viewStore.getRecovering?.() === true, [viewStore]);
  const subscribeRecovering = useCallback((listener: () => void) =>
    viewStore.subscribeRecovering?.(listener) ?? (() => {}), [viewStore]);
  const recovering = useSyncExternalStore(subscribeRecovering, getRecovering, getRecovering);
  const pasteRecordRef = useRef<PasteRecord | null>(null);
  const setPasteRecord = useCallback((record: PasteRecord | null): void => {
    pasteRecordRef.current = record;
  }, []);
  const positionRef = useRef<(() => void) | null>(null);
  const dropPasteRecord = useCallback((record: PasteRecord): void => {
    if (pasteRecordRef.current !== record) return;
    setPasteRecord(null);
    pasteOptionsRef.current.dismiss();
  }, [setPasteRecord]);
  const awaitPaste = useCallback((record: PasteRecord, submission: PeerlessAdmissionResult): void => {
    setPasteRecord(record);
    if (submission.status === 'submitted' && admission.getSnapshot().tail !== submission.ticket) {
      dropPasteRecord(record);
      return;
    }
    void appliedTicket(submission).then((ticket) => {
      if (pasteRecordRef.current !== record) return;
      if (!ticket) {
        dropPasteRecord(record);
        return;
      }
      record.ticket = ticket;
      positionRef.current?.();
    });
  }, [admission, dropPasteRecord, setPasteRecord]);
  const choosePasteOption = useCallback((option: PasteFormattingOption): boolean => {
    const record = pasteRecordRef.current;
    if (!record?.end || !record.offered.includes(option)) return false;
    const admitted = admission.getSnapshot();
    const view = viewStore.getSnapshot();
    const caret = caretPoint(view.workerSelection);
    if (viewStore.getRecovering?.() === true || view.identity.incarnation !== record.incarnation ||
      view.version !== record.version || !caret || !samePoint(caret, record.end) ||
      admitted.originalUnackedIntents !== 0 || admitted.pendingInputs !== 0) {
      dropPasteRecord(record);
      return false;
    }
    const variant = option === 'merge'
      ? mergeFormattingFragment(record.fragment)
      : option === 'text-only'
        ? plainTextClipboardFragment(clipboardFragmentPlainText(record.fragment))
        : structuredClone(record.fragment);
    const end = record.end;
    const replacing: PasteRecord = { ...record, ticket: null, end: null, reveal: 'anchor' };
    setPasteRecord(replacing);
    const outcome = callbacks.current.onReplacePasteFragment?.(variant, {
      anchor: { ...end },
      focus: { ...end },
    }, option === 'text-only' ? 'plain' : option === 'merge' ? 'merge' : undefined);
    if (!outcome) {
      dropPasteRecord(replacing);
      return false;
    }
    awaitPaste(replacing, outcome);
    return true;
  }, [admission, awaitPaste, dropPasteRecord, setPasteRecord, viewStore]);
  const pasteOptions = usePasteOptions({
    enabled: props.pasteOptions !== false && !readOnly,
    onChoose: choosePasteOption,
    focusEditor: () => { textarea.current?.focus(); },
    surface: textarea,
    onHide: () => { setPasteRecord(null); },
  });
  const pasteOptionsRef = useRef(pasteOptions);
  pasteOptionsRef.current = pasteOptions;
  const recordPaste = useCallback<PeerlessPasteRecorder>((fragment, submission) => {
    pasteOptionsRef.current.dismiss();
    setPasteRecord(null);
    if (callbacks.current.pasteOptions === false) return;
    if (viewStore.getRecovering?.() === true) return;
    const view = viewStore.getSnapshot();
    const offered = pasteOptionsOffered(fragment);
    if (offered.length === 0) return;
    awaitPaste({
      fragment: structuredClone(fragment), offered, ticket: null, end: null,
      version: view.version, incarnation: view.identity.incarnation, reveal: 'show',
    }, submission);
  }, [awaitPaste, setPasteRecord, viewStore]);
  const pasteRecorder = props.pasteRecorder;
  useLayoutEffect(() => {
    if (!pasteRecorder) return;
    pasteRecorder.current = recordPaste;
    return () => { if (pasteRecorder.current === recordPaste) pasteRecorder.current = null; };
  }, [pasteRecorder, recordPaste]);
  const attach = useCallback((element: HTMLTextAreaElement | null) => {
    textarea.current = element;
    if (typeof textareaRef === 'function') textareaRef(element);
    else if (textareaRef) textareaRef.current = element;
  }, [textareaRef]);

  const focusOnMount = useRef(autoFocus);
  useLayoutEffect(() => {
    if (!focusOnMount.current) return;
    focusOnMount.current = false;
    textarea.current?.focus({ preventScroll: true });
  }, []);

  useLayoutEffect(() => {
    const element = textarea.current;
    if (!element) return;
    const position = () => {
      const state = viewStore.getSnapshot();
      const pasteRecord = pasteRecordRef.current;
      const pasteSeq = pasteRecord?.ticket?.seq;
      if (pasteRecord && (viewStore.getRecovering?.() === true || pasteRecord.incarnation !== state.identity.incarnation)) {
        dropPasteRecord(pasteRecord);
      } else if (pasteRecord && !pasteRecord.end && pasteSeq !== undefined && state.throughSeq >= pasteSeq) {
        const tail = admission.getSnapshot().tail;
        const own = state.throughSeq === pasteSeq && (tail === undefined || tail === pasteRecord.ticket);
        const caret = caretPoint(state.workerSelection);
        if (own && caret && state.version !== pasteRecord.version) {
          pasteRecord.end = { ...caret };
          pasteRecord.version = state.version;
        } else if (!own || state.version !== pasteRecord.version) {
          dropPasteRecord(pasteRecord);
        }
      }
      const geometry = state.geometry;
      const cellFocus = !geometry?.caret && !geometry?.selectionFocus ? geometry?.selectionRects.at(-1) : undefined;
      const focus = geometry?.caret ?? geometry?.selectionFocus ?? cellFocus;
      if (composing.current || !geometry || !focus || geometry.coversThroughSeq < state.throughSeq) return;
      const parent = element.parentElement;
      const host = props.caretHostRef?.current ?? (parent?.hasAttribute('data-docx-caret-host') ? parent : null);
      const bounds = host?.getBoundingClientRect();
      const zoom = props.zoom ?? 1;
      const inHost = !!host && element.parentElement === host;
      const scale = inHost ? 1 : effectiveZoom(element);
      const top = (bounds?.top ?? 0) + focus.y * zoom;
      const height = Math.max(1, focus.height * zoom);
      element.style.position = inHost ? 'absolute' : 'fixed';
      element.style.left = `${((inHost ? 0 : bounds?.left ?? 0) + focus.x * zoom) / scale}px`;
      element.style.top = `${(inHost ? focus.y * zoom : top) / scale}px`;
      element.style.height = `${height / scale}px`;
      const shown = pasteRecordRef.current;
      const caret = caretPoint(state.workerSelection);
      if (shown?.end && caret && samePoint(caret, shown.end) && state.version === shown.version) {
        const anchorPos = { left: (bounds?.left ?? 0) + focus.x * zoom + 6, top: top + height + 4 };
        if (shown.reveal === 'show') pasteOptionsRef.current.show(anchorPos, shown.offered);
        else if (shown.reveal === 'anchor' || pasteOptionsRef.current.visible) pasteOptionsRef.current.setAnchor(anchorPos);
        shown.reveal = null;
      }
      const selection = state.workerSelection;
      const key = selection ? JSON.stringify([selection.anchor.story, selection.anchor.paraId, selection.anchor.offset,
        selection.focus.story, selection.focus.paraId, selection.focus.offset, cellFocus]) : '';
      const scroller = host ? findVerticalScrollParentOrRoot(host) : null;
      const scrolledSinceOpen = revealed.current === undefined && scrolledBeforeFirstCaret.current;
      const changed = key !== revealed.current && !scrolledSinceOpen;
      revealed.current = key;
      const hold = revealHold.current;
      const held = !!hold && (hold.seq === undefined || geometry.coversThroughSeq <= hold.seq);
      if (hold?.seq !== undefined && geometry.coversThroughSeq >= hold.seq) revealHold.current = null;
      if (!changed || held || !scroller || document.activeElement !== element) return;
      const delta = scrollIntoViewDelta(scrollViewport(scroller), top, top + height, 24);
      if (delta) scroller.scrollTop += delta;
    };
    positionRef.current = position;
    position();
    window.addEventListener('resize', position);
    return () => {
      if (positionRef.current === position) positionRef.current = null;
      window.removeEventListener('resize', position);
    };
  }, [snapshot.geometry, snapshot.throughSeq, props.caretHostRef, props.zoom, viewStore]);

  useLayoutEffect(() => {
    if (!pasteOptions.visible) return;
    const follow = () => { positionRef.current?.(); };
    window.addEventListener('scroll', follow, true);
    return () => window.removeEventListener('scroll', follow, true);
  }, [pasteOptions.visible]);

  useLayoutEffect(() => {
    const record = pasteRecordRef.current;
    if (!record || record.incarnation === snapshot.identity.incarnation) return;
    setPasteRecord(null);
    pasteOptionsRef.current.dismiss();
  }, [snapshot.identity.incarnation, setPasteRecord]);

  useLayoutEffect(() => {
    if (!recovering) return;
    setPasteRecord(null);
    pasteOptionsRef.current.dismiss();
  }, [recovering, setPasteRecord]);

  useLayoutEffect(() => {
    const record = pasteRecordRef.current;
    if (!record?.end) return;
    const selection = snapshot.workerSelection;
    const atEnd = !!selection &&
      selection.anchor.story === selection.focus.story &&
      selection.anchor.paraId === selection.focus.paraId &&
      selection.anchor.offset === selection.focus.offset &&
      selection.focus.story === record.end.story &&
      selection.focus.paraId === record.end.paraId &&
      selection.focus.offset === record.end.offset;
    if (!atEnd || snapshot.version !== record.version) {
      if (snapshot.version === record.version && typeof document !== 'undefined' &&
        pasteOptionsRef.current.controlRef.current?.contains(document.activeElement) === true) return;
      setPasteRecord(null);
      pasteOptionsRef.current.dismiss();
    }
  }, [snapshot.workerSelection, snapshot.version, setPasteRecord]);

  useLayoutEffect(() => {
    const element = textarea.current;
    if (!element) return;
    let nextEventId = 0;
    let state: PeerlessInputEventState = {
      composing: false, compositionId: null, lastHandledEventId: null, trailingCompositionText: null,
    };
    let active: { id: number; phase: 'start' | 'beforeinput'; inputType: string | null; keyboard: boolean } | null = null;
    let composition: PeerlessCompositionState | null = null;
    let mounted = true;
    let escapeArmed = false;
    const clearEscape = () => { escapeArmed = false; };
    const isReadOnly = () => admission.getSnapshot().readOnly;
    const dispatch = (event: PeerlessInputEvent, native: Event) => {
      const normalized = normalizeInputEvent(event, state);
      state = normalized.state;
      if (normalized.intents.length > 0 && event.source !== 'paste') {
        setPasteRecord(null);
        pasteOptionsRef.current.dismiss();
      }
      if (normalized.preventDefault) native.preventDefault();
      admitInputIntents(admission, normalized.intents, {
        inputTime: event.inputTime, fence: normalized.fence,
        source: event.source === 'paste' ? 'paste' : 'typing',
      });
      const ticket = selectsAll(event) ? admission.getSnapshot().tail : undefined;
      const hold: { seq?: number } | null = ticket ? {} : null;
      revealHold.current = hold;
      const release = () => { if (revealHold.current === hold) revealHold.current = null; };
      void ticket?.ack.then(outcome => {
        if (outcome.seq === undefined || outcome.status === 'refused' || outcome.status === 'cancelled') release();
        else hold!.seq = outcome.seq;
      }, release);
    };
    let releasedKey = false;
    const keydown = (event: KeyboardEvent) => {
      releasedKey = false;
      const exit = escapeArmed && event.key === 'Tab';
      if (event.key !== 'Shift') clearEscape();
      if (event.defaultPrevented || state.composing || event.isComposing || event.keyCode === 229) return;
      if (isKeyReleasedFromInput(event)) { active = null; releasedKey = true; return; }
      if (event.key === 'Escape') { escapeArmed = true; active = null; return; }
      if (exit || (event.key === 'Tab' && isReadOnly())) { active = null; return; }
      state = { ...state, trailingCompositionText: null };
      const key = event.key.toLowerCase();
      if ((event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey && ['b', 'i', 'u'].includes(key)) {
        event.preventDefault();
        const inputType = key === 'b' ? 'formatBold' : key === 'i' ? 'formatItalic' : 'formatUnderline';
        active = { id: ++nextEventId, phase: 'start', inputType, keyboard: true };
        state = { ...state, lastHandledEventId: active.id };
        if (event.repeat) return;
        submitFormattingIntent(admission, { kind: 'toggleMark', payload: {
          mark: { type: key === 'b' ? 'bold' : key === 'i' ? 'italic' : 'underline' },
        } }, { inputTime: event.timeStamp });
        return;
      }
      if (event.key === 'Delete' && event.shiftKey && !event.altKey && !event.ctrlKey && !event.metaKey &&
        callbacks.current.onClipboard && !isReadOnly()) {
        const selection = viewStore.getSnapshot().workerSelection;
        const ranged = !!selection && (selection.anchor.story !== selection.focus.story ||
          selection.anchor.paraId !== selection.focus.paraId || selection.anchor.offset !== selection.focus.offset);
        if (ranged || admission.getSnapshot().pendingInputs > 0) {
          event.preventDefault();
          active = null;
          void callbacks.current.onClipboard(true, null);
          return;
        }
      }
      const tab = event.key === 'Tab' && !event.altKey && !event.ctrlKey && !event.metaKey;
      const editState = viewStore.getSnapshot().editState ?? admission.session.state.editState;
      const tableTab = tab && !!editState?.tableContext;
      const numPr = editState?.toolbar?.paragraphProperties.numPr as { numId?: number } | null | undefined;
      if (tab && event.shiftKey && !tableTab && !admission.getSnapshot().tail && !(numPr?.numId && numPr.numId > 0)) { active = null; return; }
      const inputType = tab ? 'insertText' : event.key === 'Backspace' ? 'deleteContentBackward'
        : event.key === 'Delete' ? 'deleteContentForward'
        : event.key === 'Enter' ? 'insertParagraph' : null;
      active = { id: ++nextEventId, phase: 'start', inputType, keyboard: true };
      let pageDistance: number | undefined;
      if ((event.key === 'PageUp' || event.key === 'PageDown') && !event.altKey && !event.ctrlKey && !event.metaKey) {
        const snapshot = viewStore.getSnapshot();
        if (snapshot.geometry?.caret || snapshot.workerSelection || admission.session.state.selection || admission.getSnapshot().pendingInputs > 0) {
          const host = callbacks.current.caretHostRef?.current ?? element;
          const viewport = scrollViewport(findVerticalScrollParentOrRoot(host));
          pageDistance = viewport.height * viewport.zoom / (effectiveZoom(host) * (callbacks.current.zoom ?? 1));
        }
      }
      if (tableTab) {
        event.preventDefault();
        state = { ...state, lastHandledEventId: active.id };
        if (callbacks.current.onNavigateCell) callbacks.current.onNavigateCell(event.shiftKey ? 'backward' : 'forward');
        else admitInputIntents(admission, [{ kind: 'navigateCell', payload: {
          direction: event.shiftKey ? 'backward' : 'forward', extend: false, createLastRow: !event.shiftKey && !isReadOnly(),
        } }], { inputTime: event.timeStamp, source: 'selection', fence: true });
      } else if (tab) {
        event.preventDefault();
        state = { ...state, lastHandledEventId: active.id };
        admitIntent(admission, { kind: 'insertText', payload: { text: '\t' } }, {
          inputTime: event.timeStamp, source: 'typing', fence: true,
          tabDirection: event.shiftKey ? 'backward' : 'forward',
        });
      }
      else dispatch({ source: 'keydown', eventId: active.id, inputTime: event.timeStamp,
        key: event.key, code: event.code, shift: event.shiftKey, alt: event.altKey,
        ctrl: event.ctrlKey, meta: event.metaKey, pageDistance,
      }, event);
      if (isReadOnly() && (inputType !== null || (!(event.ctrlKey || event.metaKey || event.altKey) && event.key.length === 1))) {
        event.preventDefault();
      }
    };
    const keyup = () => { if (active?.keyboard) active = null; };
    const compatible = (inputType: string) => active !== null &&
      (!inputType || active.inputType === null || active.inputType === inputType ||
        (active.inputType === 'insertParagraph' && inputType === 'insertLineBreak') ||
        (active.inputType === 'deleteContentBackward' && inputType.endsWith('Backward')) ||
        (active.inputType === 'deleteContentForward' && inputType.endsWith('Forward')));
    const beforeinput = (event: InputEvent) => {
      clearEscape();
      const inputType = event.inputType ?? '';
      if (releasedKey && inputType.startsWith('history')) { event.preventDefault(); return; }
      if (!active || active.phase !== 'start' || !compatible(inputType)) {
        active = { id: ++nextEventId, phase: 'start', inputType, keyboard: false };
      }
      active.phase = 'beforeinput';
      active.inputType = inputType;
      const mark = inputType === 'formatBold' ? 'bold' : inputType === 'formatItalic' ? 'italic'
        : inputType === 'formatUnderline' ? 'underline' : null;
      if (mark && !state.composing) {
        event.preventDefault();
        if (state.lastHandledEventId !== active.id) {
          submitFormattingIntent(admission, { kind: 'toggleMark', payload: { mark: { type: mark } } },
            { inputTime: event.timeStamp });
          state = { ...state, lastHandledEventId: active.id };
        }
        return;
      }
      dispatch({ source: 'beforeinput', eventId: active.id, inputTime: event.timeStamp,
        inputType, data: event.data, isComposing: event.isComposing,
      }, event);
      if (isReadOnly() && !state.composing) event.preventDefault();
    };
    const input = (event: InputEvent) => {
      clearEscape();
      const inputType = event.inputType ?? '';
      const eventId = compatible(inputType) ? active!.id : ++nextEventId;
      dispatch({ source: 'input', eventId, inputTime: event.timeStamp,
        inputType, value: element.value, isComposing: event.isComposing,
      }, event);
      active = null;
      if (!state.composing) element.value = '';
      else if (composition) {
        composition = updateComposition(composition, element.value);
        callbacks.current.onCompositionChange?.(composition);
      }
    };
    const paste = (event: ClipboardEvent) => {
      clearEscape();
      setPasteRecord(null);
      pasteOptionsRef.current.dismiss();
      if (state.composing) return;
      event.preventDefault();
      const gesture = takeClipboardGesture(element);
      if (gesture === 'ignore') return;
      active = { id: ++nextEventId, phase: 'start', inputType: 'insertFromPaste', keyboard: false };
      const html = gesture === 'plain' ? '' : event.clipboardData?.getData('text/html') ?? '';
      if (html) {
        const fragment = parseClipboardHtml(html, { formatted: true });
        if (fragment && callbacks.current.onPasteFragment?.(fragment, event.timeStamp) === true) {
          element.value = '';
          return;
        }
      }
      dispatch({ source: 'paste', eventId: active.id, inputTime: event.timeStamp,
        text: event.clipboardData?.getData('text/plain') ?? '',
      }, event);
      element.value = '';
    };
    const compositionstart = () => {
      clearEscape();
      if (composition || isReadOnly()) return;
      const id = ++nextEventId;
      composition = beginComposition(admission, id);
      composing.current = true;
      state = { ...state, composing: true, compositionId: id, trailingCompositionText: null };
      active = null;
      callbacks.current.onCompositionChange?.(composition);
    };
    const compositionupdate = (event: CompositionEvent) => {
      if (!composition) return;
      composition = updateComposition(composition, event.data ?? element.value);
      callbacks.current.onCompositionChange?.(composition);
    };
    const finishComposition = (text: string) => {
      if (!composition) return;
      const current = composition;
      composition = null;
      composing.current = false;
      state = { ...state, composing: false, compositionId: null,
        trailingCompositionText: text || null, lastHandledEventId: null };
      active = null;
      const recover = (committed: PeerlessCompositionState) => {
        if (mounted) callbacks.current.onCompositionChange?.({ ...committed, phase: 'refused', recoverableText: text });
        if (text) callbacks.current.onRecoverableText?.(text);
      };
      let result: PeerlessCompositionEndResult;
      try {
        result = endComposition(admission, current, text);
      } catch {
        element.value = '';
        recover({ ...current, text });
        return;
      }
      element.value = '';
      callbacks.current.onCompositionChange?.(result.state);
      if (result.status === 'committed') {
        const committed = result.state;
        void result.ticket.ack.then((outcome) => {
          if (outcome.status !== 'refused' && outcome.status !== 'cancelled' && outcome.committed !== false) return;
          recover(committed);
        }, () => recover(committed));
      } else if (result.state.recoverableText) callbacks.current.onRecoverableText?.(result.state.recoverableText);
    };
    const compositionend = (event: CompositionEvent) => finishComposition(event.data ?? composition?.text ?? element.value);
    const blur = () => { clearEscape(); if (composition) finishComposition(composition.text); active = null; };
    const copy = (event: ClipboardEvent) => {
      if (takeClipboardGesture(element) === 'ignore') { event.preventDefault(); return; }
      if (!callbacks.current.onClipboard || state.composing) return;
      event.preventDefault();
      void callbacks.current.onClipboard(event.type === 'cut', event.clipboardData);
    };
    element.addEventListener('keydown', keydown);
    element.addEventListener('keyup', keyup);
    element.addEventListener('beforeinput', beforeinput);
    element.addEventListener('input', input);
    element.addEventListener('paste', paste);
    element.addEventListener('copy', copy);
    element.addEventListener('cut', copy);
    element.addEventListener('compositionstart', compositionstart);
    element.addEventListener('compositionupdate', compositionupdate);
    element.addEventListener('compositionend', compositionend);
    element.addEventListener('blur', blur);
    element.ownerDocument.addEventListener('pointerdown', clearEscape, true);
    return () => {
      mounted = false;
      try {
        if (composition) {
          const current = composition;
          composition = null;
          try {
            endComposition(admission, current, '');
          } catch {} finally {
            if (current.text) callbacks.current.onRecoverableText?.(current.text);
          }
        }
      } finally {
        element.removeEventListener('keydown', keydown);
        element.removeEventListener('keyup', keyup);
        element.removeEventListener('beforeinput', beforeinput);
        element.removeEventListener('input', input);
        element.removeEventListener('paste', paste);
        element.removeEventListener('copy', copy);
        element.removeEventListener('cut', copy);
        element.removeEventListener('compositionstart', compositionstart);
        element.removeEventListener('compositionupdate', compositionupdate);
        element.removeEventListener('compositionend', compositionend);
        element.removeEventListener('blur', blur);
        element.ownerDocument.removeEventListener('pointerdown', clearEscape, true);
      }
    };
  }, [admission, viewStore, setPasteRecord]);

  return <>
  <textarea
    ref={attach}
    className={className ?? 'paged-editor__input'}
    aria-label={ariaLabel ?? t('editor.documentInput')}
    readOnly={readOnly}
    rows={1}
    spellCheck
    autoCapitalize="sentences"
    autoCorrect="on"
    style={{ position: 'fixed', left: 0, top: 0, width: 1, height: 1, opacity: 0, pointerEvents: 'none',
      padding: 0, border: 0, resize: 'none', overflow: 'hidden', caretColor: 'transparent' }}
  />
  {pasteOptions.visible && pasteOptions.anchor && (
    <Suspense fallback={null}>
      <LazyPasteOptionsControl
        anchor={pasteOptions.anchor}
        menuOpen={pasteOptions.menuOpen}
        selected={pasteOptions.selected}
        options={pasteOptions.offered}
        onMenuOpenChange={pasteOptions.setMenuOpen}
        onChoose={pasteOptions.choose}
        onCloseMenu={pasteOptions.closeMenu}
        controlRef={pasteOptions.controlRef}
      />
    </Suspense>
  )}
  </>;
}
