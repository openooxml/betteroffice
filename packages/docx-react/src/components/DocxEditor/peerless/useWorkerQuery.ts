import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { DocxDocumentRead, DocxReviewSnapshot } from '@betteroffice/docx/yrs';
import type { createPeerlessShellController, PeerlessWorkerQuery } from './shellController';
import { sameShellValue } from './shellSnapshot';
import type { PeerlessViewSnapshot } from './viewStore';
import { reviewVersionOf } from './workerReview';

const SELECTION_SCOPED = new Set<DocxDocumentRead['kind']>(['properties', 'sectionProperties']);

const withBytes = (_key: string, value: unknown) => ArrayBuffer.isView(value)
  ? { $type: value.constructor.name, $bytes: Array.from(new Uint8Array(value.buffer, value.byteOffset, value.byteLength)) } : value;

/**
 * Reads `liveRequest` once per document version; selection-scoped kinds also re-read for each selection, content control handles for each layout.
 * While `held`, the request and selection last seen unheld stay in effect.
 */
export function useWorkerQuery<K extends DocxDocumentRead['kind']>(controller: ReturnType<typeof createPeerlessShellController>,
  liveRequest: (DocxDocumentRead & { kind: K }) | null, { held = false }: { held?: boolean } = {}) {
  const frameBound = liveRequest?.kind === 'contentControlHandles';
  const state = useSyncExternalStore(frameBound ? controller.subscribe : controller.subscribeShell,
    frameBound ? controller.getSnapshot : controller.getShellSnapshot, frameBound ? controller.getSnapshot : controller.getShellSnapshot);
  const [result, setResult] = useState<{ view: PeerlessViewSnapshot | null; query: PeerlessWorkerQuery<K>; key: string } | null>(null);
  const latestResult = useRef(result);
  const shown = state.view;
  const liveKey = JSON.stringify(liveRequest, withBytes) ?? 'null';
  const live = { request: liveRequest, requestKey: liveKey, key: liveRequest && SELECTION_SCOPED.has(liveRequest.kind)
    ? `${liveKey}|${JSON.stringify(shown?.workerSelection ?? null, withBytes)}` : liveKey };
  const settled = useRef(live);
  if (!held) settled.current = live;
  const { request, requestKey, key } = held ? settled.current : live;
  const latestRequest = useRef(request);
  latestRequest.current = request;
  const layout = frameBound ? `${shown?.geometry?.documentVersion}:${shown?.geometry?.layoutRevision}` : null;
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
      const held = latestResult.current;
      if (held?.key === key && reviewPlaced(latest.view, held.query.value as DocxReviewSnapshot | null)) return;
      load();
    }) : undefined;
    return () => { current = false; remove?.(); };
  }, [controller, state.status, shown?.identity.session, shown?.identity.incarnation, shown?.identity.epoch, shown?.version, layout, key]);
  const latest = latestResult.current;
  return requestKey !== 'null' && state.status === 'ready' && latest?.key === key && sameQueryView(latest.view, state.view) ? latest.query : null;
}

function sameQueryView(first: PeerlessViewSnapshot | null, second: PeerlessViewSnapshot | null) {
  return first?.version === second?.version && sameShellValue(first?.identity, second?.identity);
}

function reviewPlaced(view: PeerlessViewSnapshot | null, review: DocxReviewSnapshot | null) {
  if (!view?.geometry || !review) return false;
  const { version, previewVersion } = reviewVersionOf(view);
  return review.version === version && review.previewVersion === previewVersion && view.geometry.documentVersion === version &&
    review.layoutRevision === view.geometry.layoutRevision;
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
