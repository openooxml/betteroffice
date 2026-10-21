import type { CollaborationPipeChannel } from '../../../../shared/collaboration-pipe/main';
import type { PipeCommandPayload, PipeOutputBatch } from '../../../../shared/collaboration-pipe/protocol';
import type { PptxPresenceCursor, PptxPresencePeer } from '../collaboration/types';
import type { PresentationCommandClient } from './commandProtocol';

type Channel = CollaborationPipeChannel<PptxPresenceCursor, PptxPresencePeer>;

export function createPresentationCollaborationPipe() {
  const endpoints = new Map<number, { receive: Parameters<Channel['open']>[0]; lost: (reason: string) => void }>();
  let nextProvider = 0;
  let retired: string | undefined;
  return {
    channel(post: (payload: PipeCommandPayload<PptxPresenceCursor>) => Promise<unknown>): Channel {
      return { open(receive, lost) {
        const provider = ++nextProvider;
        if (retired) queueMicrotask(() => lost(retired!));
        else endpoints.set(provider, { receive, lost });
        return {
          post(ops) {
            if (retired || !endpoints.has(provider)) return Promise.reject(new Error(retired ?? 'Collaboration replica is retired'));
            return post({ provider, ops });
          },
          close() { endpoints.delete(provider); },
        };
      } };
    },
    receive(batches: readonly PipeOutputBatch<PptxPresencePeer>[]) {
      for (const batch of batches) endpoints.get(batch.provider)?.receive(batch.out);
    },
    lose(reason: string) {
      retired = reason;
      const active = [...endpoints.values()];
      endpoints.clear();
      for (const endpoint of active) endpoint.lost(reason);
    },
  };
}

type Router = ReturnType<typeof createPresentationCollaborationPipe>;
const pipeKey = Symbol.for('betteroffice.pptx.collaborationPipe');
type PipeClient = PresentationCommandClient & { [pipeKey]?: Router };

/** @internal */
export const presentationCollaborationPipes = {
  get(client: PresentationCommandClient): Router | undefined { return (client as PipeClient)[pipeKey]; },
  set(client: PresentationCommandClient, pipe: Router): void {
    Object.defineProperty(client, pipeKey, { value: pipe, configurable: true });
  },
  delete(client: PresentationCommandClient): void { delete (client as PipeClient)[pipeKey]; },
};
