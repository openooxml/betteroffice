// @betteroffice/docx declaration yrs/proposals.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Host proposals: tracked changes a host proposes by paragraph anchor and search, grouped by the
 * host's proposal ids. A round resolves against one version and applies as one atomic batch
 * outside undo history. Decisions only change how the proposals' revisions render: the
 * revisions stay in the document, and the document version and history are untouched until the
 * host withdraws the proposals.
 */
import type { DocxEditFailure, DocxEditFailureCode, DocxEditReceipt, DocxEditRequest, DocxEditResult, DocxEditSuggestion, DocxFindTextRequest, DocxFindTextResult, DocxReadParagraphsRequest, DocxReadParagraphsResult } from './edits';
import type { DocxParagraphAnchor, DocxParagraphAnchorResult, DocxSessionParagraphAnchor } from './paragraphIdentity';
export type DocxProposalState = 'proposed' | 'accepted' | 'rejected';
/** Which non-overlapping matches of a search a proposal replaces; numbers count from 1. */
export type DocxOccurrence = 'first' | 'all' | number;
/**
 * One proposal. `replaceText` with `replaceWith: ''` deletes; with `search: ''` it fills a
 * paragraph whose accepted text is empty (an inline atom counts as text), where the empty search
 * matches once, at offset 0. `insertText` offsets are UTF-16 units of the paragraph's accepted
 * text.
 */
export type DocxProposalInput = {
    id: string;
    paragraph: DocxParagraphAnchor;
    suggest: DocxEditSuggestion;
} & ({
    op: 'replaceText';
    search: string;
    replaceWith: string;
    occurrence?: DocxOccurrence;
    at?: {
        offset: number;
    };
} | {
    op: 'insertText';
    at: 'start' | 'end' | {
        offset: number;
    };
    text: string;
});
export interface DocxProposalRequest {
    expectVersion: string;
    expectPreviewVersion?: number;
    proposals: readonly DocxProposalInput[];
}
export interface DocxProposalRecord {
    id: string;
    state: DocxProposalState;
    stateVersion?: number;
    decisionState?: 'accepted' | 'rejected';
    paragraph: DocxSessionParagraphAnchor;
    revisionIds: readonly string[];
    changed: boolean;
}
export interface DocxProposalSnapshot {
    version: string;
    revision?: number;
    /** Increments whenever a decision changes. */
    previewVersion: number;
    proposals: readonly DocxProposalRecord[];
    withdrawn?: readonly DocxProposalRecord[];
}
/** @internal */
export interface DocxProposalRegistryState {
    revision?: number;
    withdrawn?: readonly DocxProposalRecord[];
    previewVersion: number;
    entries: {
        record: DocxProposalRecord;
        key: string;
        suggest: {
            author: string;
            date: string;
        };
    }[];
}
export type DocxProposalFailure = Omit<DocxEditFailure, 'code'> & {
    code: DocxEditFailureCode | 'stale-preview' | 'unknown-proposal' | 'proposal-id-conflict';
    proposalId?: string;
};
export type DocxProposalResult = {
    ok: true;
    snapshot: DocxProposalSnapshot;
    withdrawnIds?: readonly string[];
} | {
    ok: false;
    version: string;
    failure: DocxProposalFailure;
};
/** Missing or remotely settled ids withdraw as a no-op, preserving accepted text. */
export interface DocxProposalWithdrawRequest {
    expectVersion: string;
    expectPreviewVersion?: number;
    ids: readonly string[];
    /** Release ownership while keeping ordinary tracked revisions. Off by default. */
    keepRevisions?: boolean;
}
export interface DocxProposalStateRequest {
    expectVersion: string;
    expectPreviewVersion: number;
    changes: readonly {
        id: string;
        state: DocxProposalState;
    }[];
}
/** The session reads and writes proposals are built on. @internal */
export interface DocxProposalSession {
    version(): string;
    resolveParagraphAnchor(anchor: DocxParagraphAnchor): DocxParagraphAnchorResult;
    findText(request: DocxFindTextRequest): DocxFindTextResult;
    readParagraphs(request: DocxReadParagraphsRequest): DocxReadParagraphsResult;
    applyEdits(request: DocxEditRequest): DocxEditResult;
    listRevisions(): readonly {
        revisionId: string;
        kind: string;
    }[];
    revisionStamps?(ids: readonly string[]): Record<string, readonly {
        author: string;
        date: string;
    }[]>;
    /**
     * Accepts and rejects revisions for good, outside undo history; unknown ids are skipped. Returns
     * the refusal, with the document unchanged, while it holds remote updates not integrated yet.
     */
    settleRevisions(accept: readonly string[], reject: readonly string[]): DocxEditFailure | undefined;
    /** Runs `read` with story projections shared across its reads; omitted, reads run unshared. */
    sharedReads?<R>(read: () => R): R;
}
/** @internal */
export interface DocxProposalRegistry {
    propose(request: DocxProposalRequest): DocxProposalResult;
    setStates(request: DocxProposalStateRequest): DocxProposalResult;
    withdraw(request: DocxProposalWithdrawRequest): DocxProposalResult;
    snapshot(): DocxProposalSnapshot;
    /** @internal */
    reconcile(): void;
    /** @internal */
    exportState(): DocxProposalRegistryState;
    /** @internal */
    mirror(mirror: {
        version: string;
        proposals: DocxProposalRegistryState;
    } | null): void;
    subscribe(listener: (snapshot: DocxProposalSnapshot) => void): () => void;
    /** Forgets every proposal, as when the session opens another document. */
    reset(): void;
    destroy(): void;
}
/** The revision decisions a snapshot shows, for `YrsRenderEnv.revisionPreview`. */
export declare function proposalRevisionPreview(snapshot: DocxProposalSnapshot): Readonly<Record<string, 'accepted' | 'rejected'>> | undefined;
/** @internal */
export type ProposalRoundOutcome = {
    ok: false;
    failure: DocxProposalFailure;
    version?: string;
} | {
    ok: true;
    planned: Array<{
        anchor: DocxSessionParagraphAnchor;
        first: number;
        count: number;
    }>;
    receipts: DocxEditReceipt[];
    changedStories: string[];
};
/** @internal */
export interface ProposalWithdrawal {
    owned: readonly string[];
    accept: readonly string[];
    reject: readonly string[];
    proposalIds?: Readonly<Record<string, string>>;
    /** Per owned revision, the stamp its proposal suggested; a revision holding another is refused. */
    suggested: Readonly<Record<string, {
        author: string;
        date: string;
    }>>;
    /** Refuses as `stale-version` when the document is no longer at this version. */
    expectVersion?: string;
}
/** @internal */
export type ProposalWithdrawalOutcome = {
    ok: true;
} | {
    ok: false;
    failure: DocxProposalFailure;
};
/** Runs within the caller's shared-read scope. @internal */
export declare function executeProposalRound(session: DocxProposalSession, fresh: readonly DocxProposalInput[], expectVersion: string): ProposalRoundOutcome;
/** @internal */
export declare function executeProposalWithdrawal(session: DocxProposalSession, { owned, accept, reject, proposalIds, suggested, expectVersion }: ProposalWithdrawal): ProposalWithdrawalOutcome;
/** The session holds update notifications until `propose` returns, so they see the round. */
export declare function createProposalRegistry(session: DocxProposalSession): DocxProposalRegistry;
