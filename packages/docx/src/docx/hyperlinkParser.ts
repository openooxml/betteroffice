/** Public hyperlink helpers and XmlElement compatibility adapter. */

import type {
  Hyperlink,
  MediaFile,
  RelationshipMap,
  Run,
  Theme,
} from '../types/document';
import { sanitizeHref } from '../utils/sanitizeHref';
import type { ChartPartsMap } from './chartParser';
import type { SmartArtContext } from './smartArtParser';
import type { StyleMap } from './styleParser';
import { getAttribute, getChildElements, getTextContent, type XmlElement } from './xmlParser';

export function parseHyperlink(
  node: XmlElement,
  rels: RelationshipMap | null,
  _styles: StyleMap | null = null,
  _theme: Theme | null = null,
  _media: Map<string, MediaFile> | null = null,
  _charts?: ChartPartsMap | null,
  _smartArt: SmartArtContext | null = null
): Hyperlink {
  const hyperlink: Hyperlink = { type: 'hyperlink', children: [] };
  const rId = getAttribute(node, 'r', 'id');
  if (rId) {
    hyperlink.rId = rId;
    const relationship = rels?.get(rId);
    if (relationship) hyperlink.href = sanitizeHref(relationship.target);
  }
  const anchor = getAttribute(node, 'w', 'anchor');
  if (anchor) {
    hyperlink.anchor = anchor.slice(0, 1024);
    if (!hyperlink.href) hyperlink.href = `#${hyperlink.anchor}`;
  }
  const tooltip = getAttribute(node, 'w', 'tooltip');
  if (tooltip) hyperlink.tooltip = tooltip.slice(0, 2048);
  const target = getAttribute(node, 'w', 'tgtFrame');
  if (target) hyperlink.target = target.slice(0, 255);
  const history = getAttribute(node, 'w', 'history');
  if (history === '1' || history === 'true') hyperlink.history = true;
  const location = getAttribute(node, 'w', 'docLocation');
  if (location) hyperlink.docLocation = location.slice(0, 2048);

  for (const child of getChildElements(node).slice(0, 10_000)) {
    if (child.name?.replace(/^.*:/, '') !== 'r') continue;
    const text = getTextContent(child);
    hyperlink.children.push({
      type: 'run',
      content: text ? [{ type: 'text', text }] : [],
    });
  }
  if (hyperlink.children.length) hyperlink.structuredChildren = [...hyperlink.children];
  return hyperlink;
}

export function getHyperlinkText(hyperlink: Hyperlink): string {
  return hyperlink.children
    .filter((child): child is Run => child.type === 'run')
    .flatMap((run) => run.content)
    .map((content) => (content.type === 'text' ? content.text : content.type === 'tab' ? '\t' : ''))
    .join('');
}

export function isExternalLink(hyperlink: Hyperlink): boolean {
  if (hyperlink.href) return /^https?:\/\/|^mailto:|^tel:|^ftp:/i.test(hyperlink.href);
  return !!hyperlink.rId && !hyperlink.anchor;
}

export function isInternalLink(hyperlink: Hyperlink): boolean {
  return !!hyperlink.anchor;
}

export function getHyperlinkUrl(hyperlink: Hyperlink): string | undefined {
  return hyperlink.href;
}

export function hasContent(hyperlink: Hyperlink): boolean {
  return hyperlink.children.some((child) => child.type === 'run');
}

export function getHyperlinkRuns(hyperlink: Hyperlink): Run[] {
  return hyperlink.children.filter((child): child is Run => child.type === 'run');
}

export function resolveHyperlinkUrl(
  hyperlink: Hyperlink,
  rels: RelationshipMap
): string | undefined {
  if (hyperlink.rId) hyperlink.href = sanitizeHref(rels.get(hyperlink.rId)?.target);
  if (hyperlink.anchor && !hyperlink.href) hyperlink.href = `#${hyperlink.anchor}`;
  return hyperlink.href;
}
