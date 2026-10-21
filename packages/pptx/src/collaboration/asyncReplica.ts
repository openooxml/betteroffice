import type { CollaborationPipeChannel } from '../../../../shared/collaboration-pipe/main';
import type { AsyncCollaborationReplica, CollaborationReplica, PptxPresenceCursor, PptxPresencePeer } from './types';

const pipes = new WeakMap<object, CollaborationPipeChannel<PptxPresenceCursor, PptxPresencePeer>>();

/** @internal */
export function attachCollaborationPipe(replica: AsyncCollaborationReplica, channel: CollaborationPipeChannel<PptxPresenceCursor, PptxPresencePeer>): void {
  pipes.set(replica, channel);
}

/** @internal */
export function collaborationPipeFor(replica: AsyncCollaborationReplica | CollaborationReplica): CollaborationPipeChannel<PptxPresenceCursor, PptxPresencePeer> | undefined {
  return pipes.get(replica);
}
