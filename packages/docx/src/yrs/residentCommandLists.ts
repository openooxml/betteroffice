import type { YrsParagraphAttrs, YrsStoryRange } from './index';
import type { ResidentDocxCommandEngine } from './residentCommandProtocol';

export function planResidentListChanges(
  engine: ResidentDocxCommandEngine, range: YrsStoryRange,
  action: { kind: 'bullet' | 'numbered' } | { direction: 'increase' | 'decrease' }
): Array<{ range: YrsStoryRange; attrs: YrsParagraphAttrs }> {
  const paragraphs = engine.paragraphs(range.story);
  const start = paragraphs.findIndex(p => p.paraId === range.start.paraId);
  const end = paragraphs.findIndex(p => p.paraId === range.end.paraId);
  const numId = 'kind' in action ? action.kind === 'bullet' ? 1 : 2 : undefined;
  const current = engine.selectionContext(range).paragraphProperties.numPr as { numId?: number } | undefined;
  const remove = numId !== undefined && current?.numId === numId;
  const clear = { numPr: null, listIsBullet: null, listNumFmt: null, listMarker: null };
  return paragraphs.slice(start, end + 1).flatMap(paragraph => {
    const properties = paragraph.properties;
    const numPr = properties.numPr as { numId?: number; ilvl?: number } | null | undefined;
    let attrs: YrsParagraphAttrs;
    if ('kind' in action) {
      attrs = remove ? clear : { numPr: { numId: numId!, ilvl: numPr?.ilvl ?? 0 }, listIsBullet: action.kind === 'bullet', listNumFmt: action.kind === 'bullet' ? null : 'decimal', listMarker: null };
    } else if (numPr?.numId) {
      const level = numPr.ilvl ?? 0;
      if (action.direction === 'increase' && level >= 8) return [];
      attrs = action.direction === 'decrease' && level === 0 ? clear
        : { numPr: { numId: numPr.numId, ilvl: level + (action.direction === 'increase' ? 1 : -1) } };
      attrs = { ...attrs, indentLeft: null, indentFirstLine: null, other: { hangingIndent: null } };
    } else {
      const current = typeof properties.indentLeft === 'number' && Number.isFinite(properties.indentLeft) ? properties.indentLeft : 0;
      const next = action.direction === 'increase' ? current + 720 : Math.max(0, current - 720);
      if (next === current) return [];
      attrs = { indentLeft: next };
    }
    return [{ range: { story: range.story, start: { paraId: paragraph.paraId, offset: 0 }, end: { paraId: paragraph.paraId, offset: 0 } }, attrs }];
  });
}
