import type { YrsLoc, YrsSession, YrsStorySegment } from './index';
import type { DocxTextRange, DocxTextView } from './edits';
import type { DocxParagraphAnchor } from './paragraphIdentity';
import { proposalRevisionPreview, type DocxOccurrence, type DocxProposalSnapshot } from './proposals';
import { createYrsSidebarProjection } from '../layout/render/yrsSidebarProjection';
import {
  createYrsPositionProjection,
  createYrsLocProjectionFromOutline,
  yrsLocToProjectedDisplayPosition,
  type YrsLocProjection,
  type YrsPositionOutline,
} from './yrsPositionProjection';
import { createYrsInputPositionMap, type YrsInputPositionMap } from './inputPositionMap';

/** @internal */
export interface AnchorResolutionFailure {
  ok: false;
  failure: {
    code:
      | 'stale-version'
      | 'missing-target'
      | 'ambiguous-target'
      | 'layout-unavailable'
      | 'unsupported'
      | 'unknown-proposal';
    message: string;
  };
}

/** @internal */
export type AnchorReader = Pick<
  YrsSession,
  | 'resolveParagraphAnchor'
  | 'hasStory'
  | 'paragraphSpans'
  | 'storySegments'
  | 'findText'
  | 'listRevisions'
>;

/** @internal */
export type ProposalGeometryReader = AnchorReader &
  Pick<
    YrsSession,
    'storyIds' | 'paragraphs' | 'paragraphIdCount' | 'locateParagraph' | 'version'
  > & {
    proposalRevisions?(ids: readonly string[]): readonly ProposalGeometryRevision[];
    positionOutline?(root: string): YrsPositionOutline | null;
  };

/** @internal */
export type ProposalGeometryRevision = Pick<
  ReturnType<AnchorReader['listRevisions']>[number],
  'revisionId' | 'kind' | 'story' | 'range'
>;

/** @internal */
export type AnchorGeometryTarget =
  | { kind: 'proposal'; id: string }
  | { kind: 'revision'; revisionId: string }
  | { kind: 'paragraph'; paragraph: DocxParagraphAnchor }
  | { kind: 'search'; paragraph: DocxParagraphAnchor; text: string; occurrence?: DocxOccurrence }
  | { kind: 'range'; version: string; range: DocxTextRange };

/** @internal */
export interface RawAnchorRange {
  start: YrsLoc;
  end: YrsLoc;
}

type AnchorResolution =
  | { ok: true; ranges: RawAnchorRange[]; paragraph: YrsLoc }
  | AnchorResolutionFailure;

/** Cached reads of one document version. */
interface VersionReads {
  version: string;
  revisions?: ReturnType<AnchorReader['listRevisions']>;
  proposalRevisions?: { ids: string; revisions: readonly ProposalGeometryRevision[] };
  spans: Map<string, ReturnType<AnchorReader['paragraphSpans']>>;
  segments: Map<string, readonly YrsStorySegment[]>;
  anchors: Map<string, ReturnType<AnchorReader['resolveParagraphAnchor']>>;
}

const versionReads = new WeakMap<AnchorReader, VersionReads>();

function readsAt(session: AnchorReader, version: string): VersionReads {
  let reads = versionReads.get(session);
  if (reads?.version !== version) {
    reads = { version, spans: new Map(), segments: new Map(), anchors: new Map() };
    versionReads.set(session, reads);
  }
  return reads;
}

function once<K, V>(cache: Map<K, V>, key: K, read: () => V): V {
  if (cache.has(key)) return cache.get(key)!;
  const value = read();
  cache.set(key, value);
  return value;
}

function revisionsAt(session: AnchorReader, version: string) {
  const reads = readsAt(session, version);
  return (reads.revisions ??= session.listRevisions());
}

function segmentsAt(session: AnchorReader, version: string, story: string) {
  return once(readsAt(session, version).segments, story, () => session.storySegments(story));
}

/** @internal */
export function anchorFailure(
  code: AnchorResolutionFailure['failure']['code'],
  message: string
): AnchorResolutionFailure {
  return { ok: false, failure: { code, message } };
}

/** @internal */
export function isBodyStory(story: string): boolean {
  return story === 'body' || story.startsWith('body:');
}

interface ViewSpan {
  raw: number;
  view: number;
  length: number;
}

interface ViewParagraph {
  paraId: string;
  inlineStart: number;
  rawLength: number;
  viewLength: number;
  spans: ViewSpan[];
}

const BLOCK_EMBEDS = new Set(['table', 'blockSdt', 'pageBreak', 'columnBreak']);

function viewParagraphs(segments: readonly YrsStorySegment[], view: DocxTextView): ViewParagraph[] {
  const paragraphs: ViewParagraph[] = [];
  let rawLength = 0;
  let viewLength = 0;
  let spans: ViewSpan[] = [];
  let leading = true;
  let inlineStart = 0;
  for (const segment of segments) {
    if (segment.kind === 'pilcrow') {
      paragraphs.push({ paraId: segment.paraId, inlineStart, rawLength, viewLength, spans });
      inlineStart = 0;
      rawLength = 0;
      viewLength = 0;
      spans = [];
      leading = true;
      continue;
    }
    if (leading && segment.kind === 'embed' && BLOCK_EMBEDS.has(segment.embedKind)) {
      rawLength += 1;
      inlineStart = rawLength;
      continue;
    }
    leading = false;
    const length = segment.kind === 'text' ? segment.text.length : 1;
    const hidden = segment.attributes[view === 'accepted' ? 'del' : 'ins'] != null;
    if (!hidden && length > 0) {
      spans.push({ raw: rawLength, view: viewLength, length });
      viewLength += length;
    }
    rawLength += length;
  }
  return paragraphs;
}

function rawRange(
  paragraphs: readonly ViewParagraph[],
  range: DocxTextRange
): { ok: true; range: RawAnchorRange } | AnchorResolutionFailure {
  const locate = (paraId: string): ViewParagraph | AnchorResolutionFailure => {
    const matches = paragraphs.filter((paragraph) => paragraph.paraId === paraId);
    if (matches.length > 1) {
      return anchorFailure('ambiguous-target', 'The paragraph cannot be resolved uniquely');
    }
    return matches[0] ?? anchorFailure('missing-target', 'The paragraph no longer exists');
  };
  const start = locate(range.start.paraId);
  if ('ok' in start) return start;
  const end = locate(range.end.paraId);
  if ('ok' in end) return end;
  if (
    !Number.isInteger(range.start.offset) ||
    !Number.isInteger(range.end.offset) ||
    range.start.offset < 0 ||
    range.end.offset < 0 ||
    range.start.offset > start.viewLength ||
    range.end.offset > end.viewLength ||
    paragraphs.indexOf(start) > paragraphs.indexOf(end) ||
    (start === end && range.start.offset > range.end.offset)
  ) {
    return anchorFailure('missing-target', 'The range is outside the paragraph text');
  }

  const boundary = (paragraph: ViewParagraph, offset: number, forward: boolean): number => {
    const span = paragraph.spans.find((candidate) =>
      forward
        ? offset >= candidate.view && offset < candidate.view + candidate.length
        : offset > candidate.view && offset <= candidate.view + candidate.length
    );
    return span
      ? span.raw + offset - span.view
      : forward
        ? paragraph.rawLength
        : paragraph.inlineStart;
  };
  const from = boundary(start, range.start.offset, true);
  const collapsed = start === end && range.start.offset === range.end.offset;
  return {
    ok: true,
    range: {
      start: { story: range.story, paraId: start.paraId, offset: from },
      end: {
        story: range.story,
        paraId: end.paraId,
        offset: collapsed ? from : boundary(end, range.end.offset, false),
      },
    },
  };
}

/** @internal */
export function textRangeToRaw(
  segments: readonly YrsStorySegment[],
  range: DocxTextRange
): { ok: true; range: RawAnchorRange } | AnchorResolutionFailure {
  return rawRange(viewParagraphs(segments, range.view), range);
}

function resolveParagraph(
  session: AnchorReader,
  anchor: DocxParagraphAnchor,
  version: string
): { ok: true; loc: YrsLoc; length: number } | AnchorResolutionFailure {
  const reads = readsAt(session, version);
  const resolved = once(reads.anchors, JSON.stringify(anchor), () =>
    session.resolveParagraphAnchor(anchor)
  );
  if (resolved.status === 'missing') {
    return anchorFailure('missing-target', 'The paragraph no longer exists');
  }
  if (resolved.status === 'ambiguous') {
    return anchorFailure('ambiguous-target', 'The paragraph cannot be resolved uniquely');
  }
  if (
    resolved.status === 'unsupported' ||
    resolved.anchor.kind !== 'session' ||
    !isBodyStory(resolved.anchor.story)
  ) {
    return anchorFailure('unsupported', 'The paragraph has no body display position');
  }
  const { story, paraId } = resolved.anchor;
  const spans = once(reads.spans, story, () => session.paragraphSpans(story)).filter(
    (paragraph) => paragraph.paraId === paraId
  );
  if (spans.length > 1) {
    return anchorFailure('ambiguous-target', 'The paragraph cannot be resolved uniquely');
  }
  if (!spans.length) return anchorFailure('missing-target', 'The paragraph no longer exists');
  return { ok: true, loc: { story, paraId, offset: 0 }, length: spans[0]!.length };
}

/** @internal */
export function hiddenRanges(
  session: AnchorReader,
  version: string,
  snapshot: DocxProposalSnapshot | null,
  revisions?: readonly ProposalGeometryRevision[]
): RawAnchorRange[] {
  const preview = snapshot ? proposalRevisionPreview(snapshot) : undefined;
  return hiddenRangesForPreview(session, version, preview, revisions);
}

/** @internal */
export function hiddenRangesForPreview(
  session: AnchorReader,
  version: string,
  preview: ReturnType<typeof proposalRevisionPreview>,
  revisions?: readonly ProposalGeometryRevision[]
): RawAnchorRange[] {
  if (!preview) return [];
  return (revisions ?? revisionsAt(session, version))
    .filter(
      ({ revisionId, kind }) =>
        (kind === 'deletion' && preview[revisionId] === 'accepted') ||
        (kind === 'insertion' && preview[revisionId] === 'rejected')
    )
    .map(({ story, range }) => ({
      start: { story, ...range.start },
      end: { story, ...range.end },
    }));
}

/** @internal */
export function resolveAnchorTarget(
  session: AnchorReader,
  target: AnchorGeometryTarget,
  version: string,
  snapshot: DocxProposalSnapshot | null = null,
  proposalRevisions?: readonly ProposalGeometryRevision[]
): AnchorResolution {
  if (target.kind === 'range') {
    if (target.version !== version) {
      return anchorFailure('stale-version', 'The document changed after that version');
    }
    if (!isBodyStory(target.range.story)) {
      return anchorFailure('unsupported', 'The range has no body display position');
    }
    if (!session.hasStory(target.range.story)) {
      return anchorFailure('missing-target', 'The story no longer exists');
    }
    const mapped = textRangeToRaw(segmentsAt(session, version, target.range.story), target.range);
    return mapped.ok
      ? {
          ok: true,
          ranges: [mapped.range],
          paragraph: { ...mapped.range.start, offset: 0 },
        }
      : mapped;
  }

  if (target.kind === 'proposal' || target.kind === 'revision') {
    const proposal =
      target.kind === 'proposal'
        ? snapshot?.proposals.find((record) => record.id === target.id)
        : null;
    if (target.kind === 'proposal' && !proposal) {
      return anchorFailure('unknown-proposal', 'The proposal is not registered in this document');
    }
    const revisions = (proposalRevisions ?? revisionsAt(session, version)).filter((revision) =>
      target.kind === 'revision'
        ? revision.revisionId === target.revisionId
        : proposal!.revisionIds.includes(revision.revisionId)
    );
    if (target.kind === 'revision' && !revisions.length) {
      return anchorFailure('missing-target', 'The revision no longer exists');
    }
    const ranges = revisions.map(({ story, range }) => ({
      start: { story, ...range.start },
      end: { story, ...range.end },
    }));
    if (ranges.some((range) => !isBodyStory(range.start.story))) {
      return anchorFailure('unsupported', 'The revision has no body display position');
    }
    if (proposal) {
      const paragraph = resolveParagraph(session, proposal.paragraph, version);
      return paragraph.ok ? { ok: true, ranges, paragraph: paragraph.loc } : paragraph;
    }
    return { ok: true, ranges, paragraph: { ...ranges[0]!.start, offset: 0 } };
  }

  const paragraph = resolveParagraph(session, target.paragraph, version);
  if (!paragraph.ok) return paragraph;
  if (target.kind === 'paragraph') {
    return {
      ok: true,
      ranges: [{ start: paragraph.loc, end: { ...paragraph.loc, offset: paragraph.length } }],
      paragraph: paragraph.loc,
    };
  }
  if (!target.text) return anchorFailure('missing-target', 'The search text is empty');
  const occurrence = target.occurrence ?? 'first';
  if (
    occurrence !== 'first' &&
    occurrence !== 'all' &&
    (!Number.isInteger(occurrence) || occurrence < 1)
  ) {
    return anchorFailure(
      'missing-target',
      'The occurrence must be first, all, or a positive integer'
    );
  }
  const { story, paraId } = paragraph.loc;
  const found = session.findText({
    text: target.text,
    within: { kind: 'paragraph', story, paraId },
    view: 'accepted',
    limit: 10000,
  });
  if (!found.ok) {
    const { code, message } = found.failure;
    return anchorFailure(
      code === 'stale-version' || code === 'missing-target' || code === 'ambiguous-target'
        ? code
        : 'unsupported',
      message
    );
  }
  const matches: DocxTextRange[] = [];
  for (const { range } of [...found.matches].sort(
    (a, b) => a.range.start.offset - b.range.start.offset
  )) {
    if (range.start.offset >= (matches.at(-1)?.end.offset ?? 0)) matches.push(range);
  }
  const selected =
    occurrence === 'all'
      ? matches
      : matches.slice(
          occurrence === 'first' ? 0 : occurrence - 1,
          occurrence === 'first' ? 1 : occurrence
        );
  if (found.truncated && (occurrence === 'all' || !selected.length)) {
    return anchorFailure('unsupported', 'The search returned too many matches');
  }
  if (!selected.length) {
    return anchorFailure(
      'missing-target',
      `The requested search occurrence (${occurrence}) is unavailable`
    );
  }
  const paragraphs = viewParagraphs(segmentsAt(session, version, story), 'accepted').filter(
    (entry) => entry.paraId === paraId
  );
  const ranges: RawAnchorRange[] = [];
  for (const range of selected) {
    const mapped = rawRange(paragraphs, range);
    if (!mapped.ok) return mapped;
    ranges.push(mapped.range);
  }
  return { ok: true, ranges, paragraph: paragraph.loc };
}

/** @internal */
export interface ProposalDisplayTarget {
  ranges: { from: number; to: number }[];
  paragraph: number | null;
}

/** @internal */
export type ProposalGeometryTarget = ({ ok: true } & ProposalDisplayTarget) | AnchorResolutionFailure;

/** @internal */
export interface ProposalGeometryMirror {
  version: string;
  previewVersion: number;
  proposals: string;
  targets: Record<string, ProposalGeometryTarget>;
  navigationTargets?: Record<string, ReturnType<typeof resolveNavigationTarget>>;
  hidden: { from: number; to: number }[];
}

const proposalSetIdentities = new WeakMap<DocxProposalSnapshot, string>();

/** @internal */
export function proposalSetIdentity(snapshot: DocxProposalSnapshot): string {
  let identity = proposalSetIdentities.get(snapshot);
  if (identity === undefined) {
    identity = JSON.stringify(
      snapshot.proposals.map(({ id, revisionIds, paragraph }) => [id, revisionIds, paragraph])
    );
    proposalSetIdentities.set(snapshot, identity);
  }
  return identity;
}

/** @internal */
export function computeProposalGeometryMirror(
  reader: ProposalGeometryReader,
  snapshot: DocxProposalSnapshot,
  includeNavigationTargets = true
): ProposalGeometryMirror {
  const version = reader.version();
  let revisions: readonly ProposalGeometryRevision[] | undefined;
  if (reader.proposalRevisions) {
    const ids = [...new Set(snapshot.proposals.flatMap(({ revisionIds }) => revisionIds))].sort();
    const key = JSON.stringify(ids);
    const reads = readsAt(reader, version);
    if (reads.proposalRevisions?.ids !== key) {
      reads.proposalRevisions = {
        ids: key,
        revisions: ids.length > 0 ? reader.proposalRevisions(ids) : [],
      };
    }
    revisions = reads.proposalRevisions.revisions;
  }
  const projections = new Map<string, YrsLocProjection | null>();
  const inputMaps = new Map<string, YrsInputPositionMap | null>();
  const projectionFor = (rootStory: string): YrsLocProjection | null =>
    once(projections, rootStory, () => {
      if (!reader.hasStory(rootStory)) return null;
      const outline = reader.positionOutline?.(rootStory);
      return outline ? createYrsLocProjectionFromOutline(outline) :
        createYrsPositionProjection(reader, rootStory, {
          segments: (story) => segmentsAt(reader, version, story) as YrsStorySegment[],
        });
    });
  const inputMap = (story: string): YrsInputPositionMap | null =>
    once(inputMaps, story, () =>
      reader.hasStory(story)
        ? createYrsInputPositionMap(
            story,
            once(readsAt(reader, version).spans, story, () => reader.paragraphSpans(story))
          )
        : null
    );
  const positionFor = (loc: YrsLoc): number | null =>
    yrsLocToProjectedDisplayPosition(reader, projectionFor, loc, 'body', inputMap);
  const display = (range: RawAnchorRange): { from: number; to: number } | null => {
    const from = positionFor(range.start);
    const to = positionFor(range.end);
    return from === null || to === null ? null : { from, to };
  };
  const targets = Object.fromEntries(
    snapshot.proposals.map(({ id }): [string, ProposalGeometryTarget] => {
      const resolved = resolveAnchorTarget(reader, { kind: 'proposal', id }, version, snapshot, revisions);
      if (!resolved.ok) return [id, resolved];
      const ranges: { from: number; to: number }[] = [];
      for (const range of resolved.ranges) {
        const mapped = display(range);
        if (!mapped) {
          return [id, anchorFailure('unsupported', 'The target has no body display position')];
        }
        ranges.push(mapped);
      }
      ranges.sort((a, b) => a.from - b.from || a.to - b.to);
      return [id, { ok: true, ranges, paragraph: positionFor(resolved.paragraph) }];
    })
  );
  return {
    version,
    previewVersion: snapshot.previewVersion,
    proposals: proposalSetIdentity(snapshot),
    targets,
    navigationTargets: includeNavigationTargets ? Object.fromEntries(snapshot.proposals.map(({ id, paragraph }) => [
      id, resolveNavigationTarget(reader, paragraph.story, paragraph.paraId),
    ])) : undefined,
    hidden: hiddenRanges(reader, version, snapshot, revisions)
      .map(display)
      .filter((range): range is { from: number; to: number } => range !== null),
  };
}

/** @internal */
export function resolveMirroredNavigationTarget(
  mirror: ProposalGeometryMirror | null,
  snapshot: DocxProposalSnapshot,
  story: string,
  paraId: string
): ReturnType<typeof resolveNavigationTarget> | null {
  if (
    !mirror?.navigationTargets || mirror.version !== snapshot.version ||
    mirror.previewVersion !== snapshot.previewVersion ||
    mirror.proposals !== proposalSetIdentity(snapshot)
  ) return null;
  const proposal = snapshot.proposals.find(({ paragraph }) =>
    paragraph.story === story && paragraph.paraId === paraId
  );
  return proposal && Object.hasOwn(mirror.navigationTargets, proposal.id)
    ? mirror.navigationTargets[proposal.id]!
    : null;
}

/** @internal */
export function resolveNavigationTarget(
  reader: ProposalGeometryReader,
  story: string,
  paraId: string
): { loc: YrsLoc; position: number } | 'missing-target' | 'ambiguous-target' | 'unsupported' {
  if (typeof story !== 'string' || typeof paraId !== 'string') return 'missing-target';
  if (!reader.hasStory(story)) return 'missing-target';
  const count = reader.paragraphIdCount(story, paraId);
  if (count === 0) return 'missing-target';
  if (count > 1) return 'ambiguous-target';
  const loc = { story, paraId, offset: 0 };
  const point = createYrsSidebarProjection(reader).locToDisplayPoint(loc);
  if (!point || point.hfRid) return 'unsupported';
  return { loc, position: point.position };
}
