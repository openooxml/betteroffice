import type { ProjectionAllowance } from './projectionParity.allowlist';

export const storiesAux: readonly ProjectionAllowance[] = [
  { kind: 'blockSdt', path: 'deferred', reason: 'Block controls await native lowering.' },
  { kind: 'headerFooterStory', path: 'deferred', reason: 'Headers and footers use retained source stories.' },
  { kind: 'noteStory', path: 'deferred', reason: 'Notes use retained source stories.' },
  { kind: 'sectionRanges', path: 'deferred', reason: 'Section ranges use retained source section data.' },
  { kind: 'blockSdt', path: 'blocks[]', reason: 'Deferred block controls can be absent or retain their source stories.' },
  { kind: 'headerFooterStory', path: 'blocks', reason: 'The deferred header/footer story can retain its source content.' },
  { kind: 'noteStory', path: 'blocks', reason: 'The deferred note story can retain its source content.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].content', reason: 'Sections retain source ranges until section assembly lands.' },
];
