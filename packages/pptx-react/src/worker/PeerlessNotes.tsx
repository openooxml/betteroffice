import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { useTranslation } from '../i18n';
import { invokePeerless } from './peerlessObjectCommands';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export function PeerlessNotes({ run, onError }: { run: PeerlessPresentation; onError(error: Error): void }) {
  const { t } = useTranslation();
  const subscribe = useCallback((listener: () => void) => run.store.subscribeEditState(listener), [run]);
  const snapshot = useCallback(() => run.store.getSnapshot().editState, [run]);
  const view = useSyncExternalStore(subscribe, snapshot, snapshot);
  const slideId = view.notes?.slideId;
  const stored = view.notes?.text ?? '';
  const [draft, setDraft] = useState({ slideId, text: stored });
  const pending = useRef(new Map<string, { count: number; text: string }>());
  useEffect(() => { if (!slideId || !pending.current.has(slideId)) setDraft({ slideId, text: stored }); }, [stored, slideId]);
  const value = draft.slideId === slideId ? draft.text : (slideId && pending.current.get(slideId)?.text) ?? stored;
  return !slideId ? null : <label style={{ display: 'flex', flexDirection: 'column', padding: 8 }}>
    {t('notes.panelLabel')}
    <textarea data-testid="pptx-notes-textarea" aria-label={t('notes.panelLabel')} placeholder={t('notes.placeholder')}
      disabled={view.readOnly || !run.current || Boolean(run.failure)} value={value}
      onChange={(event) => {
        const text = event.target.value;
        const target = slideId;
        setDraft({ slideId: target, text });
        const entry = pending.current.get(target) ?? { count: 0, text };
        entry.count++;
        entry.text = text;
        pending.current.set(target, entry);
        void invokePeerless(run, { method: 'setSlideNotes', args: [target, text] }).then((ack) => {
          if (ack.outcomes.some((outcome) => outcome.status === 'refused')) throw new Error(ack.outcomes[0]?.reason ?? 'Notes refused');
        }).catch((error: unknown) => {
          if (run.current && !run.failure) onError(error instanceof Error ? error : new Error(String(error)));
        }).finally(() => {
          if (--entry.count) return;
          pending.current.delete(target);
          const notes = run.store.getSnapshot().editState.notes;
          if (run.current && notes?.slideId === target) setDraft({ slideId: target, text: notes.text ?? '' });
        });
      }} />
  </label>;
}
