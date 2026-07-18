import type { RelationshipMap } from '../types';

/** Shared relationship URI constants; package relationship parsing is Rust-owned. */
export const RELATIONSHIP_TYPES = {
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  fontTable: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  webSettings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings',
  oleObject: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject',
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  diagramData: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData',
  officeDocument:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  coreProperties:
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProperties:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  customProperties:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
  customXml: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  commentsExtended: 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
  commentsIds: 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
  commentsExtensible:
    'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
} as const;

/** Pure relationship lookup retained by public XmlElement compatibility helpers. */
export function resolveTarget(map: RelationshipMap, rId: string): string | undefined {
  return map.get(rId)?.target;
}
