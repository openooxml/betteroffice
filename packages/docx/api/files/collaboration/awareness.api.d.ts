// @betteroffice/docx declaration collaboration/awareness.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { CollaborationCursor, CollaborationPeer, CollaborationResolvedUser, CollaborationTextInsertion } from './types';
export declare const MAX_AWARENESS_STRING_LENGTH = 1024;
export declare const MAX_AWARENESS_CURSOR_BYTES = 1024;
export declare const MAX_AWARENESS_ENTRIES_PER_UPDATE = 1024;
export declare const MAX_AWARENESS_TRACKED_PEERS = 1024;
export declare function resolvePresenceColor(clientId: number, color: unknown): string;
export interface AwarenessUpdateEntry {
    clientId: number;
    clock: number;
    state: {
        user: CollaborationResolvedUser;
        cursor: CollaborationCursor | null;
    } | null;
}
export interface AwarenessRecord {
    clientId: number;
    clock: number;
    state: AwarenessUpdateEntry['state'];
    lastSeenAt: number;
    cursorMovedAt: number;
    inferredCursor: CollaborationTextInsertion | null;
}
export interface AwarenessReduction {
    records: ReadonlyMap<number, AwarenessRecord>;
    peersChanged: boolean;
}
export declare function reduceAwarenessEntries(current: ReadonlyMap<number, AwarenessRecord>, entries: readonly AwarenessUpdateEntry[], localClientId: number, now: number): AwarenessReduction;
export declare function reduceTypingInference(current: ReadonlyMap<number, AwarenessRecord>, inference: CollaborationTextInsertion, localClientId: number, now: number): AwarenessReduction;
export declare function expireAwarenessRecords(current: ReadonlyMap<number, AwarenessRecord>, now: number, maxAgeMs: number): AwarenessReduction;
export declare function awarenessPeers(records: ReadonlyMap<number, AwarenessRecord>): CollaborationPeer[];
