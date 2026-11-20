import type { YrsAuthor, YrsParagraphAttrs, YrsSession, YrsStoryRange } from './index';

/** The attributes that take a paragraph off its list. @internal */
export const LIST_OFF_ATTRS: Readonly<YrsParagraphAttrs> = Object.freeze({
  numPr: null, listIsBullet: null, listNumFmt: null, listMarker: null, listLevelNumFmts: null, listAbstractNumId: null,
  listStartOverride: null, listMarkerFontFamily: null, listMarkerSuffix: null, listMarkerHidden: null, listMarkerBold: null,
  listMarkerItalic: null, listMarkerColor: null, listMarkerFontSize: null, other: Object.freeze({ listDef: null }),
});

/** Whether a paragraph with `properties` is on a list shown as `kind`. */
function isOnListOfKind(properties: Readonly<Record<string, unknown>>, kind: 'bullet' | 'numbered'): boolean {
  const numPr = properties.numPr;
  const numId = numPr !== null && typeof numPr === 'object' ? (numPr as { numId?: unknown }).numId : undefined;
  if (typeof numId !== 'number' || numId === 0) return false;
  const bullet = properties.listIsBullet === true || properties.listNumFmt === 'bullet';
  return bullet === (kind === 'bullet');
}

/**
 * Bullets or Numbering over the paragraphs of `range`: every paragraph off its list when the first
 * is on a list of `kind`, otherwise the first onto the list `ensureListDefinition` gives it and the
 * others onto that list after it.
 * @internal
 */
export function applyListToggle(
  engine: Pick<YrsSession, 'paragraphs' | 'selectionContext' | 'ensureListDefinition' | 'setParagraphAttrs'>,
  range: YrsStoryRange,
  kind: 'bullet' | 'numbered',
  suggesting?: YrsAuthor
): void {
  const remove = isOnListOfKind(engine.selectionContext(range).paragraphProperties, kind);
  const paragraphs = engine.paragraphs(range.story);
  const start = paragraphs.findIndex(paragraph => paragraph.paraId === range.start.paraId);
  const end = paragraphs.findIndex(paragraph => paragraph.paraId === range.end.paraId);
  const lead = paragraphs[start]?.paraId;
  for (const { paraId } of paragraphs.slice(start, end + 1)) {
    const at = { story: range.story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } };
    const attrs = remove ? LIST_OFF_ATTRS
      : engine.ensureListDefinition(range.story, paraId, kind, paraId === lead ? undefined : lead);
    engine.setParagraphAttrs(at, attrs, suggesting);
  }
}
