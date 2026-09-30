/**
 * Host proposals: tracked changes a host proposes by paragraph anchor and search, grouped by the
 * host's proposal ids. A round resolves against one version and applies as one atomic batch
 * outside undo history. Decisions only change how the proposals' revisions render: the
 * revisions stay in the document, and the document version and history are untouched until the
 * host withdraws the proposals.
 */

import type {
  DocxEditFailure,
  DocxEditFailureCode,
  DocxEditReceipt,
  DocxEditRequest,
  DocxEditResult,
  DocxEditStep,
  DocxEditSuggestion,
  DocxFindTextRequest,
  DocxFindTextResult,
  DocxReadParagraphsRequest,
  DocxReadParagraphsResult,
  DocxTextMatch,
} from './edits';
import type {
  DocxParagraphAnchor,
  DocxParagraphAnchorResult,
  DocxSessionParagraphAnchor,
} from './paragraphIdentity';

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
} & (
  | {
      op: 'replaceText';
      search: string;
      replaceWith: string;
      occurrence?: DocxOccurrence;
    }
  | {
      op: 'insertText';
      at: 'start' | 'end' | { offset: number };
      text: string;
    }
);

export interface DocxProposalRequest {
  expectVersion: string;
  proposals: readonly DocxProposalInput[];
}

export interface DocxProposalRecord {
  id: string;
  state: DocxProposalState;
  paragraph: DocxSessionParagraphAnchor;
  revisionIds: readonly string[];
  changed: boolean;
}

export interface DocxProposalSnapshot {
  version: string;
  /** Increments whenever a decision changes. */
  previewVersion: number;
  proposals: readonly DocxProposalRecord[];
}

export type DocxProposalFailure = Omit<DocxEditFailure, 'code'> & {
  code: DocxEditFailureCode | 'stale-preview' | 'unknown-proposal' | 'proposal-id-conflict';
  proposalId?: string;
};

export type DocxProposalResult =
  | { ok: true; snapshot: DocxProposalSnapshot }
  | { ok: false; version: string; failure: DocxProposalFailure };

/** Ids that name no proposal are ignored, so a retried withdrawal changes nothing. */
export interface DocxProposalWithdrawRequest {
  expectVersion: string;
  ids: readonly string[];
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
  listRevisions(): readonly { revisionId: string; kind: string }[];
  revisionStamps?(ids: readonly string[]): Record<string, readonly { author: string; date: string }[]>;
  /** Accepts and rejects revisions for good, outside undo history; unknown ids are skipped. */
  settleRevisions(accept: readonly string[], reject: readonly string[]): void;
  /** Runs `read` with story projections shared across its reads; omitted, reads run unshared. */
  sharedReads?<R>(read: () => R): R;
}

/** @internal */
export interface DocxProposalRegistry {
  propose(request: DocxProposalRequest): DocxProposalResult;
  setStates(request: DocxProposalStateRequest): DocxProposalResult;
  withdraw(request: DocxProposalWithdrawRequest): DocxProposalResult;
  snapshot(): DocxProposalSnapshot;
  subscribe(listener: (snapshot: DocxProposalSnapshot) => void): () => void;
  /** Forgets every proposal, as when the session opens another document. */
  reset(): void;
  destroy(): void;
}

const SEARCH_LIMIT = 10_000;
const MAX_OFFSET = 0xffffffff;

/** The revision decisions a snapshot shows, for `YrsRenderEnv.revisionPreview`. */
export function proposalRevisionPreview(
  snapshot: DocxProposalSnapshot
): Readonly<Record<string, 'accepted' | 'rejected'>> | undefined {
  const decided = new Map<string, 'accepted' | 'rejected'>();
  for (const proposal of snapshot.proposals) {
    if (proposal.state === 'proposed') continue;
    for (const revisionId of proposal.revisionIds) decided.set(revisionId, proposal.state);
  }
  if (decided.size === 0) return undefined;
  return Object.fromEntries([...decided].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)));
}

function canonical(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonical);
  if (value === null || typeof value !== 'object') return value;
  return Object.fromEntries(
    Object.keys(value)
      .sort()
      .filter((key) => (value as Record<string, unknown>)[key] !== undefined)
      .map((key) => [key, canonical((value as Record<string, unknown>)[key])])
  );
}

/** What makes two proposals with one id the same proposal; authorship metadata does not. */
function proposalKey(input: DocxProposalInput): string {
  const edit =
    input.op === 'replaceText'
      ? {
          op: input.op,
          search: input.search,
          replaceWith: input.replaceWith,
          occurrence:
            input.occurrence === undefined || input.occurrence === 1 ? 'first' : input.occurrence,
        }
      : { op: input.op, at: input.at, text: input.text };
  return JSON.stringify(canonical({ paragraph: input.paragraph, ...edit }));
}

function checkedInput(input: DocxProposalInput): void {
  if (!input || typeof input !== 'object') throw new TypeError('a proposal must be an object');
  const fields: Record<string, unknown> = { ...input };
  const text = (key: string) => {
    if (typeof fields[key] !== 'string') throw new TypeError(`proposal ${key} must be a string`);
  };
  text('id');
  if (!input.paragraph || typeof input.paragraph !== 'object') {
    throw new TypeError('proposal paragraph must be a paragraph anchor');
  }
  if (!input.suggest || typeof input.suggest !== 'object') {
    throw new TypeError('proposal suggest must carry author and date');
  }
  if (input.op === 'replaceText') {
    text('search');
    text('replaceWith');
  } else if (input.op === 'insertText') {
    text('text');
  } else {
    throw new TypeError(`unknown proposal op ${JSON.stringify((input as { op?: unknown }).op)}`);
  }
}

/** Keeps the left-most matches that do not overlap an earlier kept one. */
function nonOverlapping(matches: readonly DocxTextMatch[]): DocxTextMatch[] {
  const kept: DocxTextMatch[] = [];
  let end = -1;
  for (const match of matches) {
    if (match.range.start.offset < end) continue;
    kept.push(match);
    end = match.range.end.offset;
  }
  return kept;
}

/**
 * The first pair of changing steps from different proposals whose ranges touch. The engine keeps
 * them apart but merges adjacent suggestions into one revision, which one decision could not split.
 */
function adjoining(
  steps: readonly DocxEditStep[],
  owners: readonly number[],
  inert: readonly boolean[]
): [number, number] | null {
  const spans = steps.flatMap((step, index) =>
    step.op === 'replaceText' && step.target.kind === 'range' && !inert[index]
      ? [{ index, range: step.target }]
      : []
  );
  for (const left of spans) {
    for (const right of spans) {
      if (
        owners[left.index] !== owners[right.index] &&
        left.range.story === right.range.story &&
        left.range.start.paraId === right.range.start.paraId &&
        left.range.end.offset === right.range.start.offset
      ) {
        return left.index < right.index ? [left.index, right.index] : [right.index, left.index];
      }
    }
  }
  return null;
}

type Located = { anchor: DocxSessionParagraphAnchor } | { failure: DocxProposalFailure };

type Planned = { steps: DocxEditStep[] } | { failure: DocxProposalFailure };

/** @internal */
export type ProposalRoundOutcome =
  | { ok: false; failure: DocxProposalFailure; version?: string }
  | {
      ok: true;
      planned: Array<{ anchor: DocxSessionParagraphAnchor; first: number; count: number }>;
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
  suggested?: Readonly<Record<string, { author: string; date: string }>>;
  /** Refuses as `stale-version` when the document is no longer at this version. */
  expectVersion?: string;
}

/** @internal */
export type ProposalWithdrawalOutcome = { ok: true } | { ok: false; failure: DocxProposalFailure };

/** Runs within the caller's shared-read scope. @internal */
export function executeProposalRound(
  session: DocxProposalSession,
  fresh: readonly DocxProposalInput[],
  expectVersion: string
): ProposalRoundOutcome {
  const refuse = (
    failure: DocxProposalFailure,
    version = session.version()
  ): ProposalRoundOutcome => ({ ok: false, failure, version });
  if (expectVersion !== session.version()) {
    return refuse({
      code: 'stale-version',
      message: 'the document changed since the expected version was read',
    });
  }
  const locate = (input: DocxProposalInput): Located => {
    const failure = (code: DocxProposalFailure['code'], message: string): Located => ({
      failure: { code, message, proposalId: input.id },
    });
    const resolved = session.resolveParagraphAnchor(input.paragraph);
    switch (resolved.status) {
      case 'found':
        return resolved.anchor.kind === 'session'
          ? { anchor: resolved.anchor }
          : failure(
              'unsupported',
              `proposal ${input.id} targets a paragraph outside the editable stories`
            );
      case 'missing':
        return failure(
          'missing-target',
          `proposal ${input.id} targets a paragraph that was not found`
        );
      case 'ambiguous':
        return failure(
          'ambiguous-target',
          `proposal ${input.id} targets a paragraph id that ${resolved.candidates.length} paragraphs carry`
        );
      default:
        return failure(
          'unsupported',
          `proposal ${input.id} targets an anchor this session cannot resolve (${resolved.reason})`
        );
    }
  };

  const plan = (input: DocxProposalInput, anchor: DocxSessionParagraphAnchor): Planned => {
    const failure = (code: DocxProposalFailure['code'], message: string): Planned => ({
      failure: { code, message, proposalId: input.id },
    });
    const { story, paraId } = anchor;
    const suggest = { author: input.suggest.author, date: input.suggest.date };
    if (input.op === 'insertText') {
      if (typeof input.at === 'string') {
        if (input.at !== 'start' && input.at !== 'end') {
          return failure('invalid-step', `proposal ${input.id} inserts at an unknown boundary`);
        }
        return {
          steps: [
            {
              op: 'insertText',
              target: { kind: 'paragraph', story, paraId },
              at: input.at,
              text: input.text,
              suggest,
            },
          ],
        };
      }
      const offset = input.at?.offset;
      if (!Number.isInteger(offset) || offset < 0 || offset > MAX_OFFSET) {
        return failure(
          'invalid-step',
          `proposal ${input.id} inserts at an offset that is not a non-negative integer`
        );
      }
      const at = { paraId, offset };
      return {
        steps: [
          {
            op: 'insertText',
            target: { kind: 'range', story, start: at, end: at, view: 'accepted' },
            at: 'start',
            text: input.text,
            suggest,
          },
        ],
      };
    }
    const occurrence = input.occurrence ?? 'first';
    if (
      occurrence !== 'first' &&
      occurrence !== 'all' &&
      !(Number.isSafeInteger(occurrence) && occurrence >= 1)
    ) {
      return failure(
        'invalid-step',
        `proposal ${input.id} names occurrence ${String(
          occurrence
        )}; use 'first', 'all' or a number from 1`
      );
    }
    if (input.search === '') return fill(input, anchor, occurrence);
    const found = session.findText({
      text: input.search,
      within: { kind: 'paragraph', story, paraId },
      view: 'accepted',
      limit: SEARCH_LIMIT,
    });
    if (!found.ok) return { failure: { ...found.failure, proposalId: input.id } };
    const matches = nonOverlapping(found.matches);
    const selected =
      occurrence === 'all'
        ? matches
        : matches.slice(occurrence === 'first' ? 0 : occurrence - 1).slice(0, 1);
    if (found.truncated) {
      return failure(
        'limit-exceeded',
        `proposal ${input.id} matches more than ${SEARCH_LIMIT} times`
      );
    }
    if (selected.length === 0) {
      return failure(
        'missing-target',
        matches.length === 0
          ? `proposal ${input.id}: ${JSON.stringify(input.search)} was not found in its paragraph`
          : `proposal ${input.id}: occurrence ${occurrence} of ${JSON.stringify(
              input.search
            )} was not found; its paragraph has ${matches.length}`
      );
    }
    return {
      steps: selected.map((match) => ({
        op: 'replaceText',
        target: { kind: 'range', ...match.range },
        text: input.replaceWith,
        suggest,
      })),
    };
  };

  /** The one empty match of an empty search: offset 0 of a paragraph whose accepted text is empty. */
  const fill = (
    input: DocxProposalInput & { op: 'replaceText' },
    { story, paraId }: DocxSessionParagraphAnchor,
    occurrence: DocxOccurrence
  ): Planned => {
    const read = session.readParagraphs({ story, paraIds: [paraId], view: 'accepted' });
    if (!read.ok) return { failure: { ...read.failure, proposalId: input.id } };
    const empty = read.paragraphs.length === 1 && read.paragraphs[0]!.text === '';
    if (!empty || (typeof occurrence === 'number' && occurrence > 1)) {
      return {
        failure: {
          code: 'missing-target',
          message: empty
            ? `proposal ${input.id}: occurrence ${occurrence} of empty text was not found; its paragraph has 1`
            : `proposal ${input.id} fills a paragraph that is not empty`,
          proposalId: input.id,
        },
      };
    }
    const at = { paraId, offset: 0 };
    return {
      steps: [
        {
          op: 'replaceText',
          target: { kind: 'range', story, start: at, end: at, view: 'accepted' },
          text: input.replaceWith,
          suggest: { author: input.suggest.author, date: input.suggest.date },
        },
      ],
    };
  };

  const steps: DocxEditStep[] = [];
  const owners: number[] = [];
  const inert: boolean[] = [];
  const planned: Array<{ anchor: DocxSessionParagraphAnchor; first: number; count: number }> = [];
  for (const [index, input] of fresh.entries()) {
    const located = locate(input);
    if ('failure' in located) return refuse(located.failure);
    const built = plan(input, located.anchor);
    if ('failure' in built) return refuse(built.failure);
    planned.push({ anchor: located.anchor, first: steps.length, count: built.steps.length });
    for (const step of built.steps) {
      steps.push(step);
      owners.push(index);
      inert.push(input.op === 'replaceText' && input.replaceWith === input.search);
    }
  }
  const touching = adjoining(steps, owners, inert);
  if (touching) {
    const [earlier, later] = touching.map((step) => fresh[owners[step]!]!.id);
    return refuse({
      code: 'overlapping-steps',
      message: `proposal ${later} adjoins proposal ${earlier}; adjacent suggestions would share one revision`,
      proposalId: later,
    });
  }
  const result = session.applyEdits({
    expectVersion,
    history: 'none',
    steps,
  });
  if (!result.ok) {
    const { stepIndex, conflictingStepIndex, ...failure } = result.failure;
    const owner = stepIndex === undefined ? undefined : fresh[owners[stepIndex]!]?.id;
    const other =
      conflictingStepIndex === undefined
        ? undefined
        : fresh[owners[conflictingStepIndex]!]?.id;
    return refuse(
      {
        ...failure,
        ...(other !== undefined && other !== owner
          ? { message: `proposal ${owner} overlaps proposal ${other}` }
          : owner !== undefined
          ? { message: `proposal ${owner}: ${failure.message}` }
          : {}),
        ...(owner !== undefined ? { proposalId: owner } : {}),
      },
      result.version
    );
  }
  return { ok: true, planned, receipts: result.receipts, changedStories: result.changedStories };
}

/** @internal */
export function executeProposalWithdrawal(
  session: DocxProposalSession,
  { owned, accept, reject, proposalIds, suggested, expectVersion }: ProposalWithdrawal
): ProposalWithdrawalOutcome {
  if (expectVersion !== undefined && expectVersion !== session.version()) {
    return {
      ok: false,
      failure: {
        code: 'stale-version',
        message: 'the document changed since the expected version was read',
      },
    };
  }
  const owners = new Set(owned);
  const foreign = session
    .listRevisions()
    .find(
      ({ revisionId, kind }) =>
        owners.has(revisionId) && kind !== 'insertion' && kind !== 'deletion'
    );
  if (foreign) {
    const id = proposalIds?.[foreign.revisionId];
    return {
      ok: false,
      failure: {
        code: 'tracked-revision-conflict',
        message:
          id === undefined
            ? `revision ${foreign.revisionId} also marks a ${foreign.kind} change made outside the proposals`
            : `proposal ${id} shares revision ${foreign.revisionId} with a ${foreign.kind} change made outside the proposals`,
        ...(id === undefined ? {} : { proposalId: id }),
      },
    };
  }
  if (suggested && session.revisionStamps) {
    const stamps = session.revisionStamps(owned);
    for (const revisionId of owned) {
      const suggest = suggested[revisionId];
      if (
        suggest &&
        stamps[revisionId]?.some(
          ({ author, date }) => author !== suggest.author || date !== suggest.date
        )
      ) {
        const id = proposalIds?.[revisionId];
        return {
          ok: false,
          failure: {
            code: 'tracked-revision-conflict',
            message:
              id === undefined
                ? `revision ${revisionId} holds changes made outside the proposals`
                : `proposal ${id} shares revision ${revisionId} with changes made outside the proposals`,
            ...(id === undefined ? {} : { proposalId: id }),
          },
        };
      }
    }
  }
  if (accept.length > 0 || reject.length > 0) session.settleRevisions(accept, reject);
  return { ok: true };
}

/** The session holds update notifications until `propose` returns, so they see the round. */
export function createProposalRegistry(session: DocxProposalSession): DocxProposalRegistry {
  const records = new Map<
    string,
    { record: DocxProposalRecord; key: string; suggest: { author: string; date: string } }
  >();
  const listeners = new Set<(snapshot: DocxProposalSnapshot) => void>();
  let previewVersion = 0;
  let notifying = false;
  let renotify = false;

  const snapshot = (): DocxProposalSnapshot => ({
    version: session.version(),
    previewVersion,
    proposals: [...records.values()].map(({ record }) => ({
      ...record,
      paragraph: { ...record.paragraph },
      revisionIds: [...record.revisionIds],
    })),
  });

  /** Delivers the latest snapshot; a change made by a listener restarts delivery with a newer one. */
  const notify = (): void => {
    renotify = true;
    if (notifying) return;
    notifying = true;
    try {
      while (renotify) {
        renotify = false;
        if (listeners.size === 0) return;
        const current = snapshot();
        for (const listener of [...listeners]) {
          try {
            listener(current);
          } catch (error) {
            console.error('[yrs] a proposal listener threw', error);
          }
          if (renotify) break;
        }
      }
    } finally {
      notifying = false;
      renotify = false;
    }
  };

  const refuse = (
    failure: DocxProposalFailure,
    version = session.version()
  ): DocxProposalResult => ({
    ok: false,
    version,
    failure,
  });

  const propose = (request: DocxProposalRequest): DocxProposalResult => {
    if (!request || typeof request !== 'object' || !Array.isArray(request.proposals)) {
      throw new TypeError('a proposal request needs a proposals array');
    }
    const fresh: Array<{ input: DocxProposalInput; key: string }> = [];
    const seen = new Map<string, string>();
    for (const input of request.proposals) {
      checkedInput(input);
      if (input.id === '') {
        return refuse({
          code: 'invalid-step',
          message: 'a proposal id must not be empty',
          proposalId: '',
        });
      }
      const key = proposalKey(input);
      const known = records.get(input.id)?.key ?? seen.get(input.id);
      if (known !== undefined) {
        if (known === key) continue;
        return refuse({
          code: 'proposal-id-conflict',
          message: `proposal id ${input.id} already names a different change`,
          proposalId: input.id,
        });
      }
      seen.set(input.id, key);
      fresh.push({ input, key });
    }
    if (fresh.length === 0) return { ok: true, snapshot: snapshot() };
    if (request.expectVersion !== session.version()) {
      return refuse({
        code: 'stale-version',
        message: 'the document changed since the expected version was read',
      });
    }
    const result = executeProposalRound(
      session,
      fresh.map(({ input }) => input),
      request.expectVersion
    );
    if (!result.ok) return refuse(result.failure, result.version);
    for (const [index, { input, key }] of fresh.entries()) {
      const { anchor, first, count } = result.planned[index]!;
      const receipts = result.receipts.slice(first, first + count);
      records.set(input.id, {
        key,
        suggest: { author: input.suggest.author, date: input.suggest.date },
        record: {
          id: input.id,
          state: 'proposed',
          paragraph: {
            kind: 'session',
            sessionId: anchor.sessionId,
            story: anchor.story,
            paraId: anchor.paraId,
          },
          revisionIds: [...new Set(receipts.flatMap((receipt) => receipt.revisionIds))],
          changed: receipts.some((receipt) => receipt.changed),
        },
      });
    }
    notify();
    return { ok: true, snapshot: snapshot() };
  };

  const setStates = (request: DocxProposalStateRequest): DocxProposalResult => {
    if (!request || typeof request !== 'object' || !Array.isArray(request.changes)) {
      throw new TypeError('a proposal state request needs a changes array');
    }
    if (request.expectVersion !== session.version()) {
      return refuse({
        code: 'stale-version',
        message: 'the document changed since the expected version was read',
      });
    }
    if (request.expectPreviewVersion !== previewVersion) {
      return refuse({
        code: 'stale-preview',
        message: `the proposals are at preview version ${previewVersion}, not ${request.expectPreviewVersion}`,
      });
    }
    const next = new Map<string, DocxProposalState>();
    for (const change of request.changes) {
      if (!change || typeof change.id !== 'string')
        throw new TypeError('a state change needs an id');
      if (
        change.state !== 'proposed' &&
        change.state !== 'accepted' &&
        change.state !== 'rejected'
      ) {
        return refuse({
          code: 'invalid-step',
          message: `proposal ${change.id} cannot be set to ${JSON.stringify(change.state)}`,
          proposalId: change.id,
        });
      }
      if (!records.has(change.id)) {
        return refuse({
          code: 'unknown-proposal',
          message: `no proposal has id ${change.id}`,
          proposalId: change.id,
        });
      }
      const earlier = next.get(change.id);
      if (earlier !== undefined && earlier !== change.state) {
        return refuse({
          code: 'invalid-step',
          message: `proposal ${change.id} is set to two states`,
          proposalId: change.id,
        });
      }
      next.set(change.id, change.state);
    }
    let changed = false;
    for (const [id, state] of next) {
      const entry = records.get(id)!;
      if (entry.record.state === state) continue;
      entry.record = { ...entry.record, state };
      changed = true;
    }
    if (changed) {
      previewVersion += 1;
      notify();
    }
    return { ok: true, snapshot: snapshot() };
  };

  /**
   * Settles each withdrawn proposal as its decision previews it, so the document reads as the
   * preview did: accepted revisions apply, rejected and undecided ones are removed.
   */
  const withdraw = (request: DocxProposalWithdrawRequest): DocxProposalResult => {
    if (!request || typeof request !== 'object' || !Array.isArray(request.ids)) {
      throw new TypeError('a withdrawal request needs an ids array');
    }
    for (const id of request.ids) {
      if (typeof id !== 'string') throw new TypeError('a proposal id must be a string');
    }
    const withdrawn = [...new Set(request.ids)].filter((id) => records.has(id));
    if (withdrawn.length === 0) return { ok: true, snapshot: snapshot() };
    if (request.expectVersion !== session.version()) {
      return refuse({
        code: 'stale-version',
        message: 'the document changed since the expected version was read',
      });
    }
    const leaving = new Set(withdrawn);
    const kept = new Map<string, string>();
    for (const [id, { record }] of records) {
      if (leaving.has(id)) continue;
      for (const revisionId of record.revisionIds) kept.set(revisionId, id);
    }
    const owners = new Map<string, string>();
    for (const id of withdrawn) {
      const { record } = records.get(id)!;
      for (const revisionId of record.revisionIds) {
        const other = kept.get(revisionId);
        if (other !== undefined) {
          return refuse({
            code: 'tracked-revision-conflict',
            message: `proposal ${id} shares revision ${revisionId} with proposal ${other}`,
            proposalId: id,
          });
        }
        owners.set(revisionId, id);
      }
    }
    const shown = proposalRevisionPreview(snapshot()) ?? {};
    const settling = [...owners.keys()];
    const accept = settling.filter((revisionId) => shown[revisionId] === 'accepted');
    const reject = settling.filter((revisionId) => shown[revisionId] !== 'accepted');
    const result = executeProposalWithdrawal(session, {
      owned: settling,
      accept,
      reject,
      proposalIds: Object.fromEntries(owners),
      suggested: Object.fromEntries(
        [...owners].map(([revisionId, id]) => [revisionId, records.get(id)!.suggest])
      ),
    });
    if (!result.ok) return refuse(result.failure);
    let decided = false;
    for (const id of withdrawn) {
      decided ||= records.get(id)!.record.state !== 'proposed';
      records.delete(id);
    }
    if (decided) previewVersion += 1;
    notify();
    return { ok: true, snapshot: snapshot() };
  };

  return {
    propose(request) {
      const read = () => propose(request);
      return session.sharedReads ? session.sharedReads(read) : read();
    },
    setStates,
    withdraw,
    snapshot,
    subscribe(listener) {
      if (typeof listener !== 'function')
        throw new TypeError('proposal listener must be a function');
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    reset() {
      if (records.size === 0) return;
      if ([...records.values()].some(({ record }) => record.state !== 'proposed'))
        previewVersion += 1;
      records.clear();
      notify();
    },
    destroy() {
      records.clear();
      listeners.clear();
    },
  };
}
