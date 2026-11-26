import type { ProjectionAllowance } from './projectionParity.allowlist';

export const storiesAux: readonly ProjectionAllowance[] = [
  { kind: 'blockSdt', path: 'blocks[]', reason: 'Deferred block controls can be absent or retain their source stories.' },
  { kind: 'headerFooterStory', path: 'blocks', reason: 'The deferred header/footer story can retain its source content.' },
  { kind: 'noteStory', path: 'blocks', reason: 'The deferred note story can retain its source content.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].content', reason: 'Sections retain source ranges until section assembly lands.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].id', reason: 'Retained sections can lack the session closing paragraph ID.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].properties.columnCount', reason: 'Retained sections can lack the resolved column count.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].properties.equalWidth', reason: 'Retained sections can lack the resolved column width policy.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].properties.verticalAlign', reason: 'Retained sections can lack resolved vertical alignment.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].properties.orientation', reason: 'Retained sections can lack resolved page orientation.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].properties.sectionStart', reason: 'Retained sections can lack the resolved section break type.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].properties.headerDistance', reason: 'Retained sections can lack resolved header distance.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].properties.footerDistance', reason: 'Retained sections can lack resolved footer distance.' },
  { kind: 'sectionRanges', path: 'package.document.sections[].properties.gutter', reason: 'Retained sections can lack the resolved gutter.' },
];
