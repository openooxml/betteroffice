import {
  proposalRevisionPreview,
  type DocxParagraphAnchor,
  type DocxTextRange,
  type DocxTextView,
  type YrsLoc,
  type YrsSession,
  type YrsStorySegment,
} from '@betteroffice/docx/yrs';
import { proposalSnapshot } from './proposalPreview';
import type { DocxAnchorGeometryResult, DocxGeometryTarget } from './types';

export type AnchorFailure = Extract<DocxAnchorGeometryResult, { ok: false }>;
export type AnchorSession = Pick<
  YrsSession,
  | 'resolveParagraphAnchor'
  | 'hasStory'
  | 'paragraphSpans'
  | 'storySegments'
  | 'findText'
  | 'listRevisions'
>;

export interface RawAnchorRange {
  start: YrsLoc;
  end: YrsLoc;
}

type AnchorResolution = { ok: true; ranges: RawAnchorRange[]; paragraph: YrsLoc } | AnchorFailure;

export function anchorFailure(
  code: AnchorFailure['failure']['code'],
  message: string
): AnchorFailure {
  return { ok: false, failure: { code, message } };
}

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
): { ok: true; range: RawAnchorRange } | AnchorFailure {
  const locate = (paraId: string): ViewParagraph | AnchorFailure => {
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

/** Maps view boundaries to paragraph-local session offsets, which count leading block embeds. */
export function textRangeToRaw(
  segments: readonly YrsStorySegment[],
  range: DocxTextRange
): { ok: true; range: RawAnchorRange } | AnchorFailure {
  return rawRange(viewParagraphs(segments, range.view), range);
}

function resolveParagraph(
  session: AnchorSession,
  anchor: DocxParagraphAnchor
): { ok: true; loc: YrsLoc; length: number } | AnchorFailure {
  const resolved = session.resolveParagraphAnchor(anchor);
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
  const spans = session.paragraphSpans(story).filter((paragraph) => paragraph.paraId === paraId);
  if (spans.length > 1) {
    return anchorFailure('ambiguous-target', 'The paragraph cannot be resolved uniquely');
  }
  if (!spans.length) return anchorFailure('missing-target', 'The paragraph no longer exists');
  return { ok: true, loc: { story, paraId, offset: 0 }, length: spans[0]!.length };
}

/** The revisions the proposal preview hides: accepted deletions and rejected insertions. */
export function hiddenRanges(session: AnchorSession): RawAnchorRange[] {
  const snapshot = proposalSnapshot(session);
  const preview = snapshot ? proposalRevisionPreview(snapshot) : undefined;
  if (!preview) return [];
  return session
    .listRevisions()
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

export function resolveAnchorTarget(
  session: AnchorSession,
  target: DocxGeometryTarget,
  version: string
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
    const mapped = textRangeToRaw(session.storySegments(target.range.story), target.range);
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
        ? proposalSnapshot(session)?.proposals.find((record) => record.id === target.id)
        : null;
    if (target.kind === 'proposal' && !proposal) {
      return anchorFailure('unknown-proposal', 'The proposal is not registered in this document');
    }
    const revisions = session
      .listRevisions()
      .filter((revision) =>
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
      const paragraph = resolveParagraph(session, proposal.paragraph);
      return paragraph.ok ? { ok: true, ranges, paragraph: paragraph.loc } : paragraph;
    }
    return { ok: true, ranges, paragraph: { ...ranges[0]!.start, offset: 0 } };
  }

  const paragraph = resolveParagraph(session, target.paragraph);
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
  const paragraphs = viewParagraphs(session.storySegments(story), 'accepted').filter(
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
