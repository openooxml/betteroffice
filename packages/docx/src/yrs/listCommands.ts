import type { YrsAuthor, YrsParagraphAttrs, YrsSession, YrsStoryRange } from './index';

/** The attributes that take a paragraph off its list. @internal */
export const LIST_OFF_ATTRS: Readonly<YrsParagraphAttrs> = Object.freeze({
  numPr: null, listIsBullet: null, listNumFmt: null, listMarker: null, listLevelNumFmts: null, listAbstractNumId: null,
  listStartOverride: null, listMarkerFontFamily: null, listMarkerSuffix: null, listMarkerHidden: null, listMarkerBold: null,
  listMarkerItalic: null, listMarkerColor: null, listMarkerFontSize: null, other: Object.freeze({ listDef: null, listLevelIndents: null }),
});

/** Whether a paragraph with `properties` is on a list. */
function isOnList(properties: Readonly<Record<string, unknown>>): boolean {
  const numPr = properties.numPr;
  const numId = numPr !== null && typeof numPr === 'object' ? (numPr as { numId?: unknown }).numId : undefined;
  return typeof numId === 'number' && numId !== 0;
}

/** Whether a paragraph with `properties` is on a list shown as `kind`. */
function isOnListOfKind(properties: Readonly<Record<string, unknown>>, kind: 'bullet' | 'numbered'): boolean {
  if (!isOnList(properties)) return false;
  const bullet = properties.listIsBullet === true || properties.listNumFmt === 'bullet';
  return bullet === (kind === 'bullet');
}

/** Whether the style of a paragraph with `properties` numbers it. */
function isNumberedByStyle(properties: Readonly<Record<string, unknown>>): boolean {
  const original = properties._originalFormatting;
  const fromStyle = properties.numPrFromStyle
    ?? (original !== null && typeof original === 'object' ? (original as { numPrFromStyle?: unknown }).numPrFromStyle : undefined);
  const numId = fromStyle !== null && typeof fromStyle === 'object' ? (fromStyle as { numId?: unknown }).numId : undefined;
  return typeof numId === 'number' && numId !== 0;
}

/**
 * The attributes that take a paragraph with `properties` off its list: `LIST_OFF_ATTRS`, and for a
 * paragraph its style numbers the `numId` 0 that keeps it off that numbering in the saved file.
 * `styleNumbers` is the engine's answer for the style the paragraph has now; without it the flag
 * stored on the paragraph when the file was read decides.
 * @internal
 */
export function listOffAttrs(properties: Readonly<Record<string, unknown>>, styleNumbers?: boolean): YrsParagraphAttrs {
  return (styleNumbers ?? isNumberedByStyle(properties)) ? { ...LIST_OFF_ATTRS, numPr: { numId: 0, ilvl: 0 } } : LIST_OFF_ATTRS;
}

/**
 * Bullets or Numbering over the paragraphs of `range`: every paragraph off its list when the first
 * is on a list of `kind`, otherwise the first onto the list `ensureListDefinition` gives it and the
 * others onto that list after it.
 * @internal
 */
export function applyListToggle(
  engine: Pick<YrsSession, 'paragraphs' | 'selectionContext' | 'ensureListDefinition' | 'setParagraphAttrs'>
    & Partial<Pick<YrsSession, 'paragraphStyleNumbers'>>,
  range: YrsStoryRange,
  kind: 'bullet' | 'numbered',
  suggesting?: YrsAuthor
): void {
  const remove = isOnListOfKind(engine.selectionContext(range).paragraphProperties, kind);
  const paragraphs = engine.paragraphs(range.story);
  const start = paragraphs.findIndex(paragraph => paragraph.paraId === range.start.paraId);
  const end = paragraphs.findIndex(paragraph => paragraph.paraId === range.end.paraId);
  const lead = paragraphs[start]?.paraId;
  const styleNumbers = styleNumbering(engine, range.story);
  let first: YrsParagraphAttrs | undefined;
  for (const { paraId, properties } of paragraphs.slice(start, end + 1)) {
    const at = { story: range.story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } };
    let attrs: YrsParagraphAttrs;
    if (remove) attrs = listOffAttrs(properties, styleNumbers(paraId, properties));
    else if (paraId === lead) attrs = first = engine.ensureListDefinition(range.story, paraId, kind);
    else attrs = first && properties.numPr == null ? first : engine.ensureListDefinition(range.story, paraId, kind, lead);
    engine.setParagraphAttrs(at, attrs, suggesting);
  }
}

/**
 * The engine's answer to whether a paragraph's current style numbers it, asked once for each style:
 * the answer depends on the style alone, and each question walks the story.
 */
function styleNumbering(
  engine: Partial<Pick<YrsSession, 'paragraphStyleNumbers'>>,
  story: string
): (paraId: string, properties: Readonly<Record<string, unknown>>) => boolean | undefined {
  const answers = new Map<unknown, boolean | undefined>();
  return (paraId, properties) => {
    const style = properties.pStyle ?? null;
    if (!answers.has(style)) answers.set(style, engine.paragraphStyleNumbers?.(story, paraId));
    return answers.get(style);
  };
}
