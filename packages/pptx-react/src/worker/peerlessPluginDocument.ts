import type { PresentationCommandIntent } from '@betteroffice/pptx/worker';
import type { PluginInvocation } from '../../../../shared/plugin-host/runtime';
import type { PptxPluginDocument } from '../plugins/createPptxPluginHost';
import type { PptxPluginSnapshot } from '../plugins/types';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export function peerlessPluginDocument(run: PeerlessPresentation) {
  const owners = new Map<number, PluginInvocation<PptxPluginSnapshot>>();
  const document: PptxPluginDocument = {
    version: () => run.client.state.version,
    onUpdate(listener) {
      let throughSeq = run.client.state.throughSeq;
      const off = run.client.subscribe((ack) => {
        const changes = ack.outcomes.filter((outcome) => outcome.seq > throughSeq && outcome.committed);
        const invocations = changes.map((outcome) => owners.get(outcome.seq));
        let notify = () => listener(new Uint8Array(), 'remote');
        if (changes.length && invocations.every((invocation) => invocation !== undefined)) {
          for (const invocation of new Set(invocations)) {
            const next = notify;
            notify = () => invocation!.commit(next);
          }
        }
        throughSeq = Math.max(throughSeq, ack.throughSeq);
        notify();
        for (const seq of owners.keys()) if (seq <= throughSeq) owners.delete(seq);
      });
      return () => { off(); owners.clear(); };
    },
  };
  return {
    document,
    dispatch(invocation: PluginInvocation<PptxPluginSnapshot>, intent: PresentationCommandIntent) {
      let sequence: number | undefined;
      const off = run.store.subscribePresentation(() => {
        const pending = run.store.getSnapshot().pending.find((record) => record.intent.payload === intent.payload);
        if (!pending) return;
        sequence = pending.seq;
        owners.set(sequence, invocation);
      });
      return run.store.dispatch(intent).finally(() => {
        off();
        if (sequence !== undefined) owners.delete(sequence);
      });
    },
  };
}
