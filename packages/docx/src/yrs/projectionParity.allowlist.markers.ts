import type { ProjectionAllowance } from './projectionParity.allowlist';

export const markers: readonly ProjectionAllowance[] = [
  { kind: 'rawInline', path: 'blocks[].content[]', reason: 'Raw inline marker lowering is deferred.' },
  { kind: 'field', path: 'blocks[].content[].content[]', reason: 'Comment reference lowering is deferred.' },
  { kind: 'commentRanges', path: 'blocks[].content[]', reason: 'Comment side ranges are deferred.' },
  { kind: 'bookmarks', path: 'blocks[].content[]', reason: 'Bookmark markers are deferred.' },
  { kind: 'ins', path: 'blocks[].content[]', reason: 'Insertion and move-to marks are deferred.' },
  { kind: 'del', path: 'blocks[].content[]', reason: 'Deletion and move-from marks are deferred.' },
  { kind: 'pPrChange', path: 'blocks[].propertyChanges', reason: 'Paragraph property revisions are deferred.' },
  { kind: 'rPrChange', path: 'blocks[].content[].propertyChanges', reason: 'Run property revisions are deferred.' },
];
