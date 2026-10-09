import type { Document } from '../types/document';
import type { HeaderFooterType } from '../types/content';
import type { DocxReadStoriesResult, DocxTextView } from './edits';
import type { YrsSession } from './index';

export type DocxSessionStoryKind =
  | 'body'
  | 'table-cell'
  | 'content-control'
  | 'header'
  | 'footer'
  | 'footnote'
  | 'endnote'
  | 'other';

const STORY_KINDS: ReadonlySet<string> = new Set<DocxSessionStoryKind>([
  'body',
  'table-cell',
  'content-control',
  'header',
  'footer',
  'footnote',
  'endnote',
  'other',
]);

/** A section that references a header or footer part. */
export interface DocxStoryUse {
  sectionIndex: number;
  variant: HeaderFooterType;
}

export interface DocxSessionStory {
  story: string;
  kind: DocxSessionStoryKind;
  /** The story a table cell or block content control sits in. */
  parent?: string;
  /** The header or footer part's relationship id. */
  relationshipId?: string;
  /** The sections referencing a header or footer part, inheritance applied. */
  uses?: DocxStoryUse[];
}

/** The host document's header and footer parts by relationship id, which story ids cannot tell apart. */
export interface DocxStoryParts {
  headers: ReadonlyMap<string, readonly DocxStoryUse[]>;
  footers: ReadonlyMap<string, readonly DocxStoryUse[]>;
}

/** Every story, the stories of these kinds, or these story ids; kind names and ids may mix. */
export type DocxSessionStorySelection = 'all' | readonly DocxSessionStoryKind[] | readonly string[];

export interface DocxReadStorySelectionRequest {
  stories: DocxSessionStorySelection;
  view: DocxTextView;
  /** Refuses with `stale-version` when the document is at another version. */
  expectVersion?: string;
}

const CELL = /:t\d+:r\d+c\d+$/;
const CONTROL = /:sdt\d+$/;

/** The document's header and footer parts with the sections using each. */
export function storyParts(document: Document | null | undefined): DocxStoryParts {
  const pkg = document?.package;
  const body = pkg?.document;
  const sections =
    body?.sections?.map((section) => section.properties) ??
    (body?.finalSectionProperties ? [body.finalSectionProperties] : []);
  const uses = (
    parts: ReadonlyMap<string, unknown> | undefined,
    key: 'headerReferences' | 'footerReferences'
  ) => {
    const byPart = new Map<string, DocxStoryUse[]>([...(parts?.keys() ?? [])].map((rId) => [rId, []]));
    sections.forEach((properties, sectionIndex) => {
      for (const { type, rId } of properties[key] ?? []) {
        byPart.get(rId)?.push({ sectionIndex, variant: type });
      }
    });
    return byPart;
  };
  return {
    headers: uses(pkg?.headers, 'headerReferences'),
    footers: uses(pkg?.footers, 'footerReferences'),
  };
}

/** Each story's kind and container, from its id and the host document's parts. */
export function describeStories(ids: readonly string[], parts: DocxStoryParts): DocxSessionStory[] {
  return ids.map((story): DocxSessionStory => {
    const nested = CELL.exec(story) ?? CONTROL.exec(story);
    if (nested) {
      return {
        story,
        kind: nested[0].startsWith(':sdt') ? 'content-control' : 'table-cell',
        parent: story.slice(0, nested.index),
      };
    }
    if (story === 'body') return { story, kind: 'body' };
    if (story.startsWith('fn:')) return { story, kind: 'footnote' };
    if (story.startsWith('en:')) return { story, kind: 'endnote' };
    if (story.startsWith('hf:')) {
      const relationshipId = story.slice(3);
      const header = parts.headers.get(relationshipId);
      const uses = header ?? parts.footers.get(relationshipId);
      if (uses) return { story, kind: header ? 'header' : 'footer', relationshipId, uses: [...uses] };
    }
    return { story, kind: 'other' };
  });
}

/** Reads the selected stories in one read; explicit ids the document lacks report `missing-target`. */
export function readStorySelection(
  session: Pick<YrsSession, 'storyIds' | 'readStories'>,
  request: DocxReadStorySelectionRequest,
  parts: DocxStoryParts
): DocxReadStoriesResult {
  const { stories: selection, view, expectVersion } = request;
  const version = expectVersion === undefined ? {} : { expectVersion };
  if (selection === 'all') return session.readStories({ view, ...version });
  const kinds = new Set(selection.filter((entry) => STORY_KINDS.has(entry)));
  const stories = new Set(selection.filter((entry) => !STORY_KINDS.has(entry)));
  if (kinds.size > 0) {
    for (const info of describeStories(session.storyIds(), parts)) {
      if (kinds.has(info.kind)) stories.add(info.story);
    }
  }
  return session.readStories({ stories: [...stories].sort(), view, ...version });
}
