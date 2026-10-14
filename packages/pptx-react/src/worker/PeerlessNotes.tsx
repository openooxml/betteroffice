import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { useTranslation } from '../i18n';
import { invokePeerless } from './peerlessObjectCommands';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export function PeerlessNotes({ run, onError }: { run: PeerlessPresentation; onError(error: Error): void }) {
  const { t } = useTranslation();
  const subscribe = useCallback((listener: () => void) => run.store.subscribeEditState(listener), [run]);
  const snapshot = useCallback(() => run.store.getSnapshot().editState, [run]);
  const view = useSyncExternalStore(subscribe, snapshot, snapshot);
  const [draft, setDraft] = useState(view.notes?.text ?? '');
  const pending = useRef(0);
  const slideId = view.notes?.slideId;
  useEffect(() => { if (!pending.current) setDraft(view.notes?.text ?? ''); }, [view.notes?.text, slideId]);
  return !slideId ? null : <label style={{ display: 'flex', flexDirection: 'column', padding: 8 }}>
    {t('notes.panelLabel')}
    <textarea data-testid="pptx-notes-textarea" aria-label={t('notes.panelLabel')} placeholder={t('notes.placeholder')}
      disabled={view.readOnly || !run.current || Boolean(run.failure)} value={draft}
      onChange={(event) => {
        const text = event.target.value;
        setDraft(text);
        pending.current++;
        void invokePeerless(run, { method: 'setSlideNotes', args: [slideId, text] }).then((ack) => {
          if (ack.outcomes.some((outcome) => outcome.status === 'refused')) throw new Error(ack.outcomes[0]?.reason ?? 'Notes refused');
        }).catch((error: unknown) => {
          if (run.current && !run.failure) onError(error instanceof Error ? error : new Error(String(error)));
        }).finally(() => {
          if (--pending.current === 0 && run.current) setDraft(run.store.getSnapshot().editState.notes?.text ?? '');
        });
      }} />
  </label>;
}
