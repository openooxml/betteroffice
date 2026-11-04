import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { DocxDocumentRead, DocxReviewSnapshot } from '@betteroffice/docx/yrs';
import type { createPeerlessShellController, PeerlessWorkerQuery } from './shellController';
import { sameShellValue } from './shellSnapshot';
import type { PeerlessViewSnapshot } from './viewStore';

export function useWorkerQuery<K extends DocxDocumentRead['kind']>(controller: ReturnType<typeof createPeerlessShellController>, request: (DocxDocumentRead & { kind: K }) | null) {
  const frameBound = request?.kind === 'contentControlHandles';
  const state = useSyncExternalStore(frameBound ? controller.subscribe : controller.subscribeShell,
    frameBound ? controller.getSnapshot : controller.getShellSnapshot, frameBound ? controller.getSnapshot : controller.getShellSnapshot);
  const [result, setResult] = useState<{ view: PeerlessViewSnapshot | null; query: PeerlessWorkerQuery<K>; key: string } | null>(null);
  const latestResult = useRef(result);
  const latestRequest = useRef(request);
  latestRequest.current = request;
  const key = JSON.stringify(request, (_key, value) => ArrayBuffer.isView(value)
    ? { $type: value.constructor.name, $bytes: Array.from(new Uint8Array(value.buffer, value.byteOffset, value.byteLength)) } : value) ?? 'null';
  useEffect(() => {
    let current = true;
    const currentRequest = latestRequest.current;
    if (state.status !== 'ready' || !currentRequest) return;
    let serial = 0;
    let frameEpoch = controller.getSnapshot().view?.geometry?.frameEpoch;
    const load = () => {
      const requestSerial = ++serial;
      const view = controller.getSnapshot().view;
      void controller.query<K>(currentRequest).then(query => {
        if (!current || requestSerial !== serial) return;
        const previous = latestResult.current;
        const next = { view, query, key };
        latestResult.current = next;
        if (!previous || previous.key !== key || !sameQueryView(previous.view, view) || !sameQuery(previous.query, query, currentRequest.kind)) setResult(next);
      }, () => { if (current && requestSerial === serial) { latestResult.current = null; setResult(null); } });
    };
    load();
    const remove = currentRequest.kind === 'reviewSnapshot' ? controller.subscribe(() => {
      const latest = controller.getSnapshot();
      if (latest.status !== 'ready' || !sameQueryView(state.view, latest.view) || frameEpoch === latest.view?.geometry?.frameEpoch) return;
      frameEpoch = latest.view?.geometry?.frameEpoch;
      load();
    }) : undefined;
    return () => { current = false; remove?.(); };
  }, [controller, state.status, state.view?.identity, state.view?.version, state.view?.throughSeq, state.view?.geometry, key]);
  const latest = latestResult.current;
  return key !== 'null' && state.status === 'ready' && latest?.key === key && sameQueryView(latest.view, state.view) ? latest.query : null;
}

function sameQueryView(first: PeerlessViewSnapshot | null, second: PeerlessViewSnapshot | null) {
  return first?.version === second?.version && first?.throughSeq === second?.throughSeq && sameShellValue(first?.identity, second?.identity);
}

function sameQuery(first: PeerlessWorkerQuery, second: PeerlessWorkerQuery, kind: DocxDocumentRead['kind']) {
  if (first.version !== second.version || !sameShellValue(first.selection, second.selection)) return false;
  if (kind === 'reviewSnapshot') {
    const left = first.value as DocxReviewSnapshot | null;
    const right = second.value as DocxReviewSnapshot | null;
    if (left && right) return sameShellValue({ ...left, frameEpoch: 0 }, { ...right, frameEpoch: 0 });
  }
  return sameShellValue(first.value, second.value);
}
