import type { ResidentEngineWorkerClient, DocxReviewSnapshot } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { PeerlessViewSnapshot } from './viewStore';

/** The document and proposal preview versions the worker reports for `view`. */
export function reviewVersionOf(view: PeerlessViewSnapshot): { version: string | undefined; previewVersion: number } {
  return { version: view.editState?.review?.version ?? view.editState?.documentVersion ?? view.geometry?.documentVersion,
    previewVersion: view.editState?.review?.previewVersion ?? 0 };
}

/** Whether `review` describes the document and proposal preview `view` presents, in any frame of it. */
export function reviewDescribes(view: PeerlessViewSnapshot, review: DocxReviewSnapshot): boolean {
  const frame = view.geometry;
  const { version, previewVersion } = reviewVersionOf(view);
  return !!frame && view.presentedThroughSeq >= view.throughSeq && frame.coversThroughSeq >= view.throughSeq &&
    frame.documentVersion === version && review.version === version && review.previewVersion === previewVersion &&
    review.proposals.version === version && review.proposals.previewVersion === previewVersion &&
    review.proposalGeometry.version === version && review.proposalGeometry.previewVersion === previewVersion;
}

export function reviewMatches(view: PeerlessViewSnapshot, review: DocxReviewSnapshot): boolean {
  return reviewDescribes(view, review) && review.frameEpoch === view.geometry!.frameEpoch && review.layoutRevision === view.geometry!.layoutRevision;
}

export async function readPresentedReview(options: {
  binding: PeerlessEditSessionBinding;
  current(): PeerlessEditSessionBinding | null;
  subscribe(listener: () => void): () => void;
  read: ResidentEngineWorkerClient['documentRead'];
  readAt?: ResidentEngineWorkerClient['documentReadAt'];
  timeoutMs: number;
}): Promise<DocxReviewSnapshot> {
  const current = options.binding;
  const identity = current.viewStore.getSnapshot().identity;
  const assertCurrent = () => {
    const latest = current.viewStore.getSnapshot().identity;
    if (options.current() !== current || latest.session !== identity.session || latest.incarnation !== identity.incarnation || latest.epoch !== identity.epoch) {
      throw new Error('Document was replaced');
    }
  };
  let remove = () => {};
  let timer: ReturnType<typeof setTimeout>;
  const expired = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error('Review snapshot timed out')), options.timeoutMs);
  });
  let stopReplacement = () => {};
  const replaced = new Promise<never>((_, reject) => {
    stopReplacement = options.subscribe(() => { try { assertCurrent(); } catch (error) { reject(error); } });
  });
  const interrupted = Promise.race([expired, replaced]);
  const wait = (view: PeerlessViewSnapshot) => new Promise<void>((resolve, reject) => {
    const check = () => {
      try {
        assertCurrent();
        if (current.viewStore.getSnapshot() === view) return;
        remove(); resolve();
      } catch (error) { remove(); reject(error); }
    };
    remove = options.subscribe(check);
    check();
  });
  try {
    while (true) {
      assertCurrent();
      const view = current.viewStore.getSnapshot();
      const version = view.editState?.review?.version ?? view.editState?.documentVersion ?? view.geometry?.documentVersion;
      const request = { kind: 'reviewSnapshot' as const, expectPreviewVersion: view.editState?.review?.previewVersion ?? 0 };
      if (version && view.geometry) {
        const result = await Promise.race([options.readAt ? options.readAt(request, version) : options.read(request).then(result => ({ status: 'ok' as const, ...result })), interrupted]);
        assertCurrent();
        const latest = current.viewStore.getSnapshot();
        if (result.status === 'ok' && result.version === version && result.value && reviewMatches(latest, result.value)) return result.value;
        if (latest !== view) continue;
      }
      await Promise.race([wait(view), interrupted]);
    }
  } finally { clearTimeout(timer!); remove(); stopReplacement(); }
}
