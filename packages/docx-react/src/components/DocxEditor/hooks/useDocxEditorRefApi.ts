import { useImperativeHandle } from 'react';
import type { Comment } from '@betteroffice/docx/types/content';
import type { Document } from '@betteroffice/docx/types/document';
import type {
  YrsInlineFormatDelta,
  YrsLoc,
  YrsParagraph,
  YrsSession,
  YrsStoryRange,
} from '@betteroffice/docx/yrs';
import { createStyleResolver } from '@betteroffice/docx/styles';
import type { DocxInput, ScrollToParaIdOptions } from '@betteroffice/docx/utils';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import type { CommentIdAllocator } from '../commentFactories';
import { createComment } from '../commentFactories';
import type { SelectionState } from '../types';

type LocatedParagraph = {
  story: string;
  paragraph: YrsParagraph;
};

function bodyStoryIds(session: YrsSession): string[] {
  return session.storyIds().filter((story) => story === 'body' || story.startsWith('body:'));
}

function locateParagraph(session: YrsSession, paraId: string): LocatedParagraph | null {
  for (const story of bodyStoryIds(session)) {
    const paragraph = session.paragraphs(story).find((candidate) => candidate.paraId === paraId);
    if (paragraph) return { story, paragraph };
  }
  return null;
}

function uniqueMatchOffset(text: string, search: string): number | null {
  const offset = text.indexOf(search);
  return offset >= 0 && text.indexOf(search, offset + 1) < 0 ? offset : null;
}

function paragraphRange(
  session: YrsSession,
  paraId: string,
  search?: string
): YrsStoryRange | null {
  const located = locateParagraph(session, paraId);
  if (!located) return null;
  const { story, paragraph } = located;
  let start = 0;
  let end = paragraph.text.length;
  if (search !== undefined) {
    if (search === '') start = end;
    else {
      const offset = uniqueMatchOffset(paragraph.text, search);
      if (offset == null) return null;
      start = offset;
      end = offset + search.length;
    }
  }
  return {
    story,
    start: { paraId, offset: start },
    end: { paraId, offset: end },
  };
}

function storyOffset(session: YrsSession, loc: YrsLoc): number {
  return session.locateParagraph(loc.story, loc.paraId).start + loc.offset;
}

function normalizeSelection(session: YrsSession): YrsStoryRange | null {
  const selection = session.selection();
  if (!selection || selection.anchor.story !== selection.head.story) return null;
  const anchorOffset = storyOffset(session, selection.anchor);
  const headOffset = storyOffset(session, selection.head);
  const [start, end] =
    anchorOffset <= headOffset ? [selection.anchor, selection.head] : [selection.head, selection.anchor];
  return {
    story: start.story,
    start: { paraId: start.paraId, offset: start.offset },
    end: { paraId: end.paraId, offset: end.offset },
  };
}

function textForRange(session: YrsSession, range: YrsStoryRange): string {
  const paragraphs = session.paragraphs(range.story);
  const startIndex = paragraphs.findIndex((paragraph) => paragraph.paraId === range.start.paraId);
  const endIndex = paragraphs.findIndex((paragraph) => paragraph.paraId === range.end.paraId);
  if (startIndex < 0 || endIndex < startIndex) return '';
  if (startIndex === endIndex) {
    return paragraphs[startIndex].text.slice(range.start.offset, range.end.offset);
  }
  return [
    paragraphs[startIndex].text.slice(range.start.offset),
    ...paragraphs.slice(startIndex + 1, endIndex).map((paragraph) => paragraph.text),
    paragraphs[endIndex].text.slice(0, range.end.offset),
  ].join('\n');
}

function formattingDelta(marks: Parameters<DocxEditorRef['applyFormatting']>[0]['marks']) {
  const delta: YrsInlineFormatDelta = {};
  if (marks.bold !== undefined) delta.bold = marks.bold;
  if (marks.italic !== undefined) delta.italic = marks.italic;
  if (marks.underline !== undefined) {
    delta.underline = marks.underline
      ? typeof marks.underline === 'object'
        ? { style: marks.underline.style }
        : true
      : null;
  }
  if (marks.strike !== undefined) delta.strike = marks.strike;
  if (marks.color !== undefined) {
    delta.color = marks.color.rgb
      ? { rgb: marks.color.rgb }
      : marks.color.themeColor
        ? { themeColor: marks.color.themeColor }
        : null;
  }
  if (marks.highlight !== undefined) delta.highlight = marks.highlight || null;
  if (marks.fontSize !== undefined) delta.fontSize = marks.fontSize > 0 ? marks.fontSize : null;
  if (marks.fontFamily !== undefined) {
    const ascii = marks.fontFamily.ascii ?? marks.fontFamily.hAnsi;
    delta.fontFamily = ascii
      ? { ascii, hAnsi: marks.fontFamily.hAnsi ?? ascii }
      : null;
  }
  return delta;
}

/** Owns the public imperative surface without exposing an editor-view handle. */
export function useDocxEditorRefApi({
  ref,
  document,
  documentFromYrs,
  historyStateRef,
  pagedEditorRef,
  handleSave,
  handleDirectPrint,
  zoom,
  setZoom,
  scrollPageInfo,
  loadParsedDocument,
  loadBuffer,
  comments,
  setComments,
  setShowCommentsSidebar,
  contentChangeSubscribersRef,
  selectionChangeSubscribersRef,
  getCachedStyleResolver,
  commentIdAllocator,
}: {
  ref: React.ForwardedRef<DocxEditorRef>;
  document: Document | null;
  documentFromYrs: () => Document | null;
  historyStateRef: React.RefObject<Document | null>;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  handleSave: () => Promise<ArrayBuffer | null>;
  handleDirectPrint: () => void;
  zoom: number;
  setZoom: (zoom: number) => void;
  scrollPageInfo: { currentPage: number; totalPages: number; visible: boolean };
  loadParsedDocument: (doc: Document) => void;
  loadBuffer: (buffer: DocxInput) => Promise<void>;
  comments: Comment[];
  setComments: React.Dispatch<React.SetStateAction<Comment[]>>;
  setShowCommentsSidebar: React.Dispatch<React.SetStateAction<boolean>>;
  contentChangeSubscribersRef: React.RefObject<Set<(doc: Document) => void>>;
  selectionChangeSubscribersRef: React.RefObject<Set<(state: SelectionState | null) => void>>;
  getCachedStyleResolver: (
    styles: Parameters<typeof createStyleResolver>[0]
  ) => ReturnType<typeof createStyleResolver>;
  commentIdAllocator: CommentIdAllocator;
}) {
  useImperativeHandle(
    ref,
    () => ({
      getDocument: () => pagedEditorRef.current?.getDocument() ?? documentFromYrs() ?? document,
      getEditorRef: () => pagedEditorRef.current,
      save: handleSave,
      setZoom,
      getZoom: () => zoom,
      focus: () => pagedEditorRef.current?.focus(),
      getCurrentPage: () => scrollPageInfo.currentPage,
      getTotalPages: () => scrollPageInfo.totalPages,
      scrollToPage: (pageNumber) => pagedEditorRef.current?.scrollToPage(pageNumber),
      scrollToPosition: (displayPosition) =>
        pagedEditorRef.current?.scrollToPosition(displayPosition),
      openPrintPreview: handleDirectPrint,
      print: handleDirectPrint,
      loadDocument: loadParsedDocument,
      loadDocumentBuffer: loadBuffer,

      addComment: (options) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const range = session ? paragraphRange(session, options.paraId, options.search) : null;
        if (!editor || !session || !range || textForRange(session, range).length === 0) return null;
        const comment = createComment(commentIdAllocator, options.text, options.author);
        session.applyRawOps(range.story, [
          {
            op: 'setComment',
            id: String(comment.id),
            ranges: [
              [
                storyOffset(session, { story: range.story, ...range.start }),
                storyOffset(session, { story: range.story, ...range.end }),
              ],
            ],
            author: options.author,
            date: comment.date,
            body: comment.content,
          },
        ]);
        editor.syncYrsInputState(true);
        setComments((previous) => [...previous, comment]);
        setShowCommentsSidebar(true);
        return comment.id;
      },

      replyToComment: (commentId, text, authorName) => {
        if (!comments.some((comment) => comment.id === commentId)) return null;
        const reply = createComment(commentIdAllocator, text, authorName, commentId);
        setComments((previous) => [...previous, reply]);
        return reply.id;
      },

      resolveComment: (commentId) => {
        setComments((previous) =>
          previous.map((comment) =>
            comment.id === commentId ? { ...comment, done: true } : comment
          )
        );
      },

      proposeChange: (options) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        if (!editor || !session || (!options.search && !options.replaceWith)) return false;
        const range = paragraphRange(session, options.paraId, options.search);
        if (!range) return false;
        if (options.search) {
          const context = session.selectionContext(range);
          if (context.inInsertion || context.inDeletion) return false;
        }
        session.replaceRange(range, options.replaceWith, {
          name: options.author,
          date: new Date().toISOString(),
        });
        editor.syncYrsInputState(true);
        setShowCommentsSidebar(true);
        return true;
      },

      applyFormatting: (options) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const range = session ? paragraphRange(session, options.paraId, options.search) : null;
        if (!editor || !session || !range) return false;
        if (textForRange(session, range).length > 0) session.formatRange(range, formattingDelta(options.marks));
        editor.syncYrsInputState(true);
        return true;
      },

      setParagraphStyle: (options) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const range = session ? paragraphRange(session, options.paraId) : null;
        if (!editor || !session || !range) return false;
        const currentDocument = historyStateRef.current;
        const resolver = currentDocument?.package.styles
          ? getCachedStyleResolver(currentDocument.package.styles)
          : null;
        if (resolver && !resolver.hasParagraphStyle(options.styleId)) return false;
        session.applyParagraphStyle(range, options.styleId);
        editor.syncYrsInputState(true);
        return true;
      },

      insertBreak: (options) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const located = session ? locateParagraph(session, options.paraId) : null;
        if (!editor || !session || !located || located.story !== 'body') return false;
        const at = {
          story: located.story,
          paraId: located.paragraph.paraId,
          offset: located.paragraph.text.length,
        };
        if (options.type === 'page') session.insertPageBreak(at);
        else if (options.type === 'sectionNextPage') {
          session.insertSectionBreak(at, 'nextPage');
        } else if (options.type === 'sectionContinuous') {
          session.insertSectionBreak(at, 'continuous');
        } else return false;
        editor.syncYrsInputState(true);
        return true;
      },

      getPageContent: (pageNumber) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const page = editor?.getLayout()?.pages[pageNumber - 1];
        if (!editor || !session || !page) return null;
        const seen = new Set<string>();
        const paragraphs: Array<{ paraId: string; text: string; styleId?: string }> = [];
        for (const fragment of page.fragments) {
          if (fragment.kind !== 'paragraph' || fragment.pmStart == null) continue;
          const loc =
            editor.displayPositionToYrsLoc(fragment.pmStart) ??
            editor.displayPositionToYrsLoc(fragment.pmStart + 1);
          if (!loc || seen.has(loc.paraId)) continue;
          const paragraph = session
            .paragraphs(loc.story)
            .find((candidate) => candidate.paraId === loc.paraId);
          if (!paragraph) continue;
          seen.add(paragraph.paraId);
          const styleId = paragraph.properties.pStyle;
          paragraphs.push({
            paraId: paragraph.paraId,
            text: paragraph.text,
            ...(typeof styleId === 'string' ? { styleId } : {}),
          });
        }
        return {
          pageNumber,
          text: paragraphs.map((paragraph) => `[${paragraph.paraId}] ${paragraph.text}`).join('\n'),
          paragraphs,
        };
      },

      scrollToParaId: (paraId: string, options?: ScrollToParaIdOptions) =>
        pagedEditorRef.current?.scrollToParaId(paraId, options) ?? false,
      scrollToCommentId: (commentId) =>
        pagedEditorRef.current?.scrollToCommentId(commentId) ?? false,
      scrollToChangeId: (revisionId) =>
        pagedEditorRef.current?.scrollToChangeId(revisionId) ?? false,
      highlightRange: (from, to) => pagedEditorRef.current?.highlightRange(from, to),

      findInDocument: (query, options) => {
        const session = pagedEditorRef.current?.getYrsSession();
        if (!session || !query) return [];
        const caseSensitive = options?.caseSensitive ?? false;
        const needle = caseSensitive ? query : query.toLowerCase();
        const limit = options?.limit ?? 20;
        const results: ReturnType<DocxEditorRef['findInDocument']> = [];
        for (const story of bodyStoryIds(session)) {
          for (const paragraph of session.paragraphs(story)) {
            if (results.length >= limit) return results;
            const haystack = caseSensitive ? paragraph.text : paragraph.text.toLowerCase();
            const offset = haystack.indexOf(needle);
            if (offset < 0 || haystack.indexOf(needle, offset + 1) >= 0) continue;
            results.push({
              paraId: paragraph.paraId,
              match: paragraph.text.slice(offset, offset + query.length),
              before: paragraph.text.slice(Math.max(0, offset - 40), offset),
              after: paragraph.text.slice(offset + query.length, offset + query.length + 40),
            });
          }
        }
        return results;
      },

      getSelectionInfo: () => {
        const session = pagedEditorRef.current?.getYrsSession();
        const range = session ? normalizeSelection(session) : null;
        if (!session || !range) return null;
        const paragraphs = session.paragraphs(range.story);
        const startParagraph = paragraphs.find(
          (paragraph) => paragraph.paraId === range.start.paraId
        );
        const endParagraph = paragraphs.find((paragraph) => paragraph.paraId === range.end.paraId);
        if (!startParagraph || !endParagraph) return null;
        return {
          paraId: startParagraph.paraId,
          selectedText: textForRange(session, range),
          paragraphText: startParagraph.text,
          before: startParagraph.text.slice(0, range.start.offset),
          after:
            startParagraph.paraId === endParagraph.paraId
              ? startParagraph.text.slice(range.end.offset)
              : endParagraph.text.slice(range.end.offset),
        };
      },

      getComments: () => comments,

      onContentChange: (listener) => {
        contentChangeSubscribersRef.current.add(listener);
        return () => contentChangeSubscribersRef.current.delete(listener);
      },
      onSelectionChange: (listener) => {
        selectionChangeSubscribersRef.current.add(listener);
        return () => selectionChangeSubscribersRef.current.delete(listener);
      },
    }),
    [
      document,
      documentFromYrs,
      zoom,
      scrollPageInfo,
      scrollPageInfo,
      handleSave,
      handleDirectPrint,
      loadParsedDocument,
      loadBuffer,
      comments,
    ]
  );
}
