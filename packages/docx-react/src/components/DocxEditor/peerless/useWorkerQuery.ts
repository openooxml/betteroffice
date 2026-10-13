import { useEffect, useState, useSyncExternalStore } from 'react';
import type { ResidentDocumentRead } from '@betteroffice/docx/yrs';
import type { createPeerlessShellController, PeerlessWorkerQuery } from './shellController';

export function useWorkerQuery<K extends ResidentDocumentRead['kind']>(controller: ReturnType<typeof createPeerlessShellController>, request: (ResidentDocumentRead & { kind: K }) | null) {
  const state = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const [result, setResult] = useState<{ view: typeof state.view; query: PeerlessWorkerQuery<K> } | null>(null);
  const key = JSON.stringify(request);
  useEffect(() => {
    let current = true;
    if (state.status !== 'ready' || key === 'null') return;
    void controller.query<K>(JSON.parse(key)).then(query => {
      if (current) setResult({ view: state.view, query });
    }, () => { if (current) setResult(null); });
    return () => { current = false; };
  }, [controller, state.status, state.view?.identity, state.view?.version, state.view?.throughSeq, state.view?.geometry?.frameEpoch, key]);
  return key !== 'null' && state.status === 'ready' && result?.view?.identity === state.view?.identity && result?.view?.version === state.view?.version &&
    result?.view?.throughSeq === state.view?.throughSeq ? result?.query ?? null : null;
}
