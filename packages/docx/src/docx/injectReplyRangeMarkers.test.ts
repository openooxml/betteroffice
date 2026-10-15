import { describe, expect, it } from 'bun:test';
import type { BlockContent, Comment, ParagraphContent } from '../types/content';
import { injectReplyRangeMarkers, injectTCReplyRangeMarkers } from './injectReplyRangeMarkers';

const run: ParagraphContent = { type: 'run', formatting: {}, content: [{ type: 'text', text: 'Commented' }] };

function paragraph(...content: ParagraphContent[]): BlockContent {
  return { type: 'paragraph', formatting: {}, content };
}

function comment(id: number, parentId?: number): Comment {
  return { id, author: 'A', content: [], ...(parentId === undefined ? {} : { parentId }) };
}

const start = (id: number): ParagraphContent => ({ type: 'commentRangeStart', id });
const end = (id: number): ParagraphContent => ({ type: 'commentRangeEnd', id });

describe('reply range markers', () => {
  it('adds a range for a reply that has none', () => {
    const content = [paragraph(start(1), run, end(1))];
    expect(injectReplyRangeMarkers(content, [comment(1), comment(2, 1)])).toBe(true);
    expect(content).toEqual([paragraph(start(1), start(2), run, end(1), end(2))]);
  });

  it('leaves a reply that already has its own range', () => {
    const content = [paragraph(start(1), start(2), run, end(1), end(2))];
    const before = structuredClone(content);
    expect(injectReplyRangeMarkers(content, [comment(1), comment(2, 1)])).toBe(false);
    expect(content).toEqual(before);
  });

  it('leaves a reply whose range sits in another paragraph or a table', () => {
    const content: BlockContent[] = [
      paragraph(start(1), run, end(1)),
      {
        type: 'table',
        rows: [{ type: 'tableRow', cells: [{ type: 'tableCell', content: [paragraph(start(2), run, end(2))] }] }],
      } as BlockContent,
    ];
    const before = structuredClone(content);
    expect(injectReplyRangeMarkers(content, [comment(1), comment(2, 1)])).toBe(false);
    expect(content).toEqual(before);
  });

  it('leaves a tracked-change reply that already has its own range', () => {
    const insertion = {
      type: 'insertion',
      info: { id: 7, author: 'A', date: '2026-01-01T00:00:00Z' },
      content: [run],
    } as ParagraphContent;
    const content = [paragraph(start(9), insertion, end(9))];
    const before = structuredClone(content);
    expect(injectTCReplyRangeMarkers(content, [comment(9, 7)])).toBe(false);
    expect(content).toEqual(before);
  });
});
