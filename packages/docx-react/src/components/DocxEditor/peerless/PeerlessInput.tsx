import { useCallback, useLayoutEffect, useRef, useSyncExternalStore, type ReactElement, type Ref, type RefObject } from 'react';
import { effectiveZoom } from '@betteroffice/docx/layout/render';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import type { DocxClipboardFragment } from '@betteroffice/docx/yrs';
import { scrollIntoViewDelta, scrollViewport } from '../internals/viewportBand';
import { parseClipboardHtml } from '../internals/clipboardHtml';
import { beginComposition, updateComposition, endComposition, type PeerlessCompositionEndResult, type PeerlessCompositionState } from './composition';
import { admitInputIntents, submitFormattingIntent, type PeerlessInputAdmission } from './inputAdmission';
import { normalizeInputEvent, type PeerlessInputEvent, type PeerlessInputEventState } from './inputEvents';
import type { PeerlessViewStore } from './viewStore';

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
  caretHostRef?: RefObject<HTMLDivElement | null>;
  zoom?: number;
  onNavigateCell?(direction: 'forward' | 'backward'): void;
}

/** @internal */
export function PeerlessInput(props: PeerlessInputProps): ReactElement | null {
  const { admission, viewStore, textareaRef, autoFocus, className, ariaLabel } = props;
  const textarea = useRef<HTMLTextAreaElement | null>(null);
  const composing = useRef(false);
  const revealed = useRef<string | undefined>(undefined);
  const callbacks = useRef(props);
  callbacks.current = props;
  const getReadOnly = useCallback(() => admission.getSnapshot().readOnly, [admission]);
  const readOnly = useSyncExternalStore(admission.subscribeReadOnly, getReadOnly, getReadOnly);
  const snapshot = useSyncExternalStore(viewStore.subscribe, viewStore.getSnapshot, viewStore.getServerSnapshot);
  const attach = useCallback((element: HTMLTextAreaElement | null) => {
    textarea.current = element;
    if (typeof textareaRef === 'function') textareaRef(element);
    else if (textareaRef) textareaRef.current = element;
  }, [textareaRef]);

  useLayoutEffect(() => {
    const element = textarea.current;
    if (!element) return;
    const position = () => {
      const state = viewStore.getSnapshot();
      const geometry = state.geometry;
      if (composing.current || !geometry?.caret || geometry.coversThroughSeq < state.throughSeq) return;
      const host = props.caretHostRef?.current;
      const bounds = host?.getBoundingClientRect();
      const zoom = props.zoom ?? 1;
      const scale = effectiveZoom(element);
      const top = (bounds?.top ?? 0) + geometry.caret.y * zoom;
      const height = Math.max(1, geometry.caret.height * zoom);
      element.style.left = `${((bounds?.left ?? 0) + geometry.caret.x * zoom) / scale}px`;
      element.style.top = `${top / scale}px`;
      element.style.height = `${height / scale}px`;
      const selection = state.workerSelection;
      const key = selection ? JSON.stringify([selection.anchor.story, selection.anchor.paraId, selection.anchor.offset,
        selection.focus.story, selection.focus.paraId, selection.focus.offset]) : '';
      const changed = key !== revealed.current;
      revealed.current = key;
      if (!changed || !host || state.editState?.readOnly || document.activeElement !== element) return;
      const scroller = findVerticalScrollParentOrRoot(host);
      const delta = scrollIntoViewDelta(scrollViewport(scroller), top, top + height, 24);
      if (delta) scroller.scrollTop += delta;
    };
    position();
    window.addEventListener('scroll', position, true);
    window.addEventListener('resize', position);
    return () => { window.removeEventListener('scroll', position, true); window.removeEventListener('resize', position); };
  }, [snapshot.geometry, snapshot.throughSeq, props.caretHostRef, props.zoom, viewStore]);

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
    const isReadOnly = () => admission.getSnapshot().readOnly;
    const dispatch = (event: PeerlessInputEvent, native: Event) => {
      const normalized = normalizeInputEvent(event, state);
      state = normalized.state;
      if (normalized.preventDefault) native.preventDefault();
      admitInputIntents(admission, normalized.intents, {
        inputTime: event.inputTime, fence: normalized.fence,
        source: event.source === 'paste' ? 'paste' : 'typing',
      });
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || state.composing || event.isComposing || event.keyCode === 229) return;
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
      const tableTab = event.key === 'Tab' && !event.altKey && !event.ctrlKey && !event.metaKey &&
        !!(viewStore.getSnapshot().editState ?? admission.session.state.editState)?.tableContext;
      const inputType = tableTab ? 'insertText' : event.key === 'Backspace' ? 'deleteContentBackward'
        : event.key === 'Delete' ? 'deleteContentForward'
        : event.key === 'Enter' ? 'insertParagraph' : null;
      active = { id: ++nextEventId, phase: 'start', inputType, keyboard: true };
      if (tableTab) {
        event.preventDefault();
        state = { ...state, lastHandledEventId: active.id };
        if (callbacks.current.onNavigateCell) callbacks.current.onNavigateCell(event.shiftKey ? 'backward' : 'forward');
        else admitInputIntents(admission, [{ kind: 'navigateCell', payload: {
          direction: event.shiftKey ? 'backward' : 'forward', extend: false, createLastRow: !event.shiftKey && !isReadOnly(),
        } }], { inputTime: event.timeStamp, source: 'selection', fence: true });
      }
      else dispatch({ source: 'keydown', eventId: active.id, inputTime: event.timeStamp,
        key: event.key, code: event.code, shift: event.shiftKey, alt: event.altKey,
        ctrl: event.ctrlKey, meta: event.metaKey,
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
      const inputType = event.inputType ?? '';
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
      if (state.composing) return;
      event.preventDefault();
      active = { id: ++nextEventId, phase: 'start', inputType: 'insertFromPaste', keyboard: false };
      const html = event.clipboardData?.getData('text/html') ?? '';
      if (html) {
        const fragment = parseClipboardHtml(html);
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
    const blur = () => { if (composition) finishComposition(composition.text); active = null; };
    const copy = (event: ClipboardEvent) => {
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
      }
    };
  }, [admission, viewStore]);

  return <textarea
    ref={attach}
    className={className ?? 'paged-editor__peerless-input'}
    aria-label={ariaLabel ?? 'Document input'}
    autoFocus={autoFocus}
    readOnly={readOnly}
    rows={1}
    spellCheck
    autoCapitalize="sentences"
    autoCorrect="on"
    style={{ position: 'fixed', left: 0, top: 0, width: 1, height: 1, opacity: 0,
      padding: 0, border: 0, resize: 'none', overflow: 'hidden', caretColor: 'transparent' }}
  />;
}
