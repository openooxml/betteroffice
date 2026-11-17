// @betteroffice/docx declaration yrs/residentCollaboration.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { type CollaborationTextInsertion } from '../collaboration/types';
import type { PipeOutputBatch } from '../../../../shared/collaboration-pipe/protocol';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxCommandHooksFactory } from './residentCommandProtocol';
import type { RetainedFrame } from '../layout/render/frameDelta';
import { type DisplayListRect } from '../layout/render/displayListQueries';
export interface DocxPresencePeer {
    clientId: number;
    clock: number;
    cursor?: {
        story: string;
        anchor: Uint8Array;
        head: Uint8Array;
    };
    inferred?: {
        story: string;
        paraId: string;
        offset: number;
    };
}
export interface DocxPresenceGeometry {
    version: string;
    layoutRevision: number;
    frameEpoch: number;
    peers: Array<{
        clientId: number;
        clock: number;
        status: 'ok' | 'unresolved';
        caret: (DisplayListRect & {
            pageId: string;
        }) | null;
        rects: Array<DisplayListRect & {
            pageId: string;
        }>;
    }>;
}
export interface ResidentCollaborationUpdates {
    updates: ArrayBuffer[];
    inferences: Array<{
        seq: number;
        insertion: CollaborationTextInsertion;
    }>;
    /** This session's own state-vector clock after the drained edits and merges. */
    clock: number;
    /** Output of the providers this worker runs, for the threads that own their transports. */
    pipe?: PipeOutputBatch[];
}
type Hooks = ReturnType<DocxCommandHooksFactory>;
export declare function createResidentCollaboration(engine: ResidentEngineSession, options?: {
    onPipeOutput?: () => void;
}): {
    drain(): (ResidentCollaborationUpdates & {
        transfer: ArrayBuffer[];
    }) | undefined;
    wrap(hooks: Hooks): Hooks;
    dispose(): void;
};
export declare function readResidentCollaborationState(engine: ResidentEngineSession, stateVector?: Uint8Array): Uint8Array;
export declare function readResidentPresence(engine: ResidentEngineSession, frame: RetainedFrame, peers: readonly DocxPresencePeer[], pages: readonly [number, number], layoutRevision: number): Promise<DocxPresenceGeometry>;
export {};
