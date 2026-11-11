// @betteroffice/docx declaration collaboration/replica.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsSelection, YrsSession } from '../yrs';
import type { CollaborationReplica } from './types';
export interface WorkerCollaborationTarget {
    invalidate(update: Uint8Array, selection: YrsSelection | null): void;
}
export declare function createDirectCollaborationReplica(session: YrsSession): CollaborationReplica;
export declare function createWorkerCollaborationReplica(session: YrsSession, worker: WorkerCollaborationTarget): CollaborationReplica;
