import type { YrsParagraphAttrs, YrsStoryRange } from './index';
import { LIST_OFF_ATTRS } from './listCommands';
import type { DocxWorkerCommandEngine } from './residentCommandProtocol';

export function planResidentListChanges(
  engine: DocxWorkerCommandEngine, range: YrsStoryRange,
  action: { direction: 'increase' | 'decrease'; preserveNumbering?: boolean }
): Array<{ range: YrsStoryRange; attrs: YrsParagraphAttrs }> {
  const paragraphs = engine.paragraphs(range.story);
  const start = paragraphs.findIndex(p => p.paraId === range.start.paraId);
  const end = paragraphs.findIndex(p => p.paraId === range.end.paraId);
  return paragraphs.slice(start, end + 1).flatMap(paragraph => {
    const properties = paragraph.properties;
    const numPr = properties.numPr as { numId?: number; ilvl?: number } | null | undefined;
    let attrs: YrsParagraphAttrs;
    if (numPr?.numId) {
      const level = numPr.ilvl ?? 0;
      if (action.direction === 'increase' && level >= 8) return [];
      if (action.direction === 'decrease' && level === 0 && action.preserveNumbering) return [];
      attrs = action.direction === 'decrease' && level === 0 ? LIST_OFF_ATTRS
        : { numPr: { numId: numPr.numId, ilvl: level + (action.direction === 'increase' ? 1 : -1) } };
      attrs = { ...attrs, indentLeft: null, indentFirstLine: null, other: { ...attrs.other, hangingIndent: null } };
    } else {
      const current = typeof properties.indentLeft === 'number' && Number.isFinite(properties.indentLeft) ? properties.indentLeft : 0;
      const next = action.direction === 'increase' ? current + 720 : Math.max(0, current - 720);
      if (next === current) return [];
      attrs = { indentLeft: next };
    }
    return [{ range: { story: range.story, start: { paraId: paragraph.paraId, offset: 0 }, end: { paraId: paragraph.paraId, offset: 0 } }, attrs }];
  });
}
