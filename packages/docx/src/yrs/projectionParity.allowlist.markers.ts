import type { ProjectionAllowance } from './projectionParity.allowlist';

export const markers: readonly ProjectionAllowance[] = [
  { kind: 'field', path: 'deferred', reason: 'Comment references use retained source runs.' },
  { kind: 'commentRanges', path: 'deferred', reason: 'Comment ranges await native lowering.' },
  { kind: 'bookmarks', path: 'deferred', reason: 'Bookmarks await native lowering.' },
  { kind: 'ins', path: 'deferred', reason: 'Insertion and move-to marks await native lowering.' },
  { kind: 'del', path: 'deferred', reason: 'Deletion and move-from marks await native lowering.' },
  { kind: 'rPrChange', path: 'deferred', reason: 'Run property revisions await native lowering.' },
  { kind: 'pPrChange', path: 'deferred', reason: 'Paragraph property revisions await native lowering.' },
  { kind: 'revisions', path: 'deferred', reason: 'Move range boundaries use retained source markers.' },
  { kind: 'rawInline', path: 'deferred', reason: 'Raw inline markers await native lowering.' },
  { kind: 'field', path: 'blocks[].content[].content[]', reason: 'Comment reference lowering is deferred.' },
  { kind: 'commentRanges', path: 'blocks[].content[]', reason: 'Comment side ranges are deferred.' },
  { kind: 'bookmarks', path: 'blocks[].content[]', reason: 'Bookmark markers are deferred.' },
  { kind: 'ins', path: 'blocks[].content[]', reason: 'Insertion and move-to marks are deferred.' },
  { kind: 'del', path: 'blocks[].content[]', reason: 'Deletion and move-from marks are deferred.' },
  { kind: 'rPrChange', path: 'blocks[].content[].propertyChanges', reason: 'Run property revisions are deferred.' },
];
