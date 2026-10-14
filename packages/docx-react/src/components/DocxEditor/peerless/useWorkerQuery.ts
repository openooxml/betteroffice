import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { ResidentDocumentRead } from '@betteroffice/docx/yrs';
import type { createPeerlessShellController, PeerlessWorkerQuery } from './shellController';

export function useWorkerQuery<K extends ResidentDocumentRead['kind']>(controller: ReturnType<typeof createPeerlessShellController>, request: (ResidentDocumentRead & { kind: K }) | null) {
  const state = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const [result, setResult] = useState<{ view: typeof state.view; query: PeerlessWorkerQuery<K> } | null>(null);
  const latestRequest = useRef(request);
  latestRequest.current = request;
  const key = JSON.stringify(request, (_key, value) => ArrayBuffer.isView(value)
    ? { $type: value.constructor.name, $bytes: Array.from(new Uint8Array(value.buffer, value.byteOffset, value.byteLength)) } : value);
  useEffect(() => {
    let current = true;
    const currentRequest = latestRequest.current;
    if (state.status !== 'ready' || !currentRequest) return;
    void controller.query<K>(currentRequest).then(query => {
      if (current) setResult({ view: state.view, query });
    }, () => { if (current) setResult(null); });
    return () => { current = false; };
  }, [controller, state.status, state.view?.identity, state.view?.version, state.view?.throughSeq, state.view?.geometry?.frameEpoch, key]);
  return key !== 'null' && state.status === 'ready' && result?.view?.identity === state.view?.identity && result?.view?.version === state.view?.version &&
    result?.view?.throughSeq === state.view?.throughSeq ? result?.query ?? null : null;
}
