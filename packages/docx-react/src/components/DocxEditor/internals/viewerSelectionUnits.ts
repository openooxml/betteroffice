import type { DisplayListQueries, DisplayPrimitive } from '@betteroffice/docx/layout/render';
import { findWordBoundaries } from '@betteroffice/docx/utils';
import type { DocxDisplayRange } from '@betteroffice/docx/yrs';

function sameParagraph(primitive: DisplayPrimitive, paraId: string | undefined, blockId: unknown): boolean {
  if (paraId !== undefined) return primitive.paraId === paraId;
  return (primitive.blockKey ?? primitive.blockId) === blockId;
}

/**
 * The word or paragraph around body display position `position`, read from the presented
 * display list: its painted text and paragraph fragments. Null when the pages holding it are
 * not built here.
 */
export function displayListSelectionUnit(
  queries: DisplayListQueries,
  position: number,
  unit: 'word' | 'paragraph'
): DocxDisplayRange | null {
  if (!queries.isReady()) return null;
  const fragments = queries.paragraphRects(position);
  if (fragments.length === 0) return null;
  const from = Math.min(...fragments.map((fragment) => fragment.from));
  const to = Math.max(...fragments.map((fragment) => fragment.to));
  if (unit === 'paragraph') return from < to ? { anchor: from, head: to } : null;
  const { paraId, blockId } = fragments[0];
  const chars = new Array<string>(to - from).fill('￼');
  const pages = new Set(fragments.map((fragment) => fragment.pageIndex));
  for (const page of queries.displayList.pages) {
    if (!pages.has(page.pageIndex)) continue;
    for (const primitive of page.primitives) {
      if (primitive.kind !== 'text' && primitive.kind !== 'glyphRun') continue;
      const start = primitive.docStart;
      const end = primitive.docEnd;
      if (start === undefined || end === undefined || start < from || end > to) continue;
      if (end - start !== primitive.text.length || !sameParagraph(primitive, paraId, blockId)) continue;
      for (let index = 0; index < primitive.text.length; index += 1) {
        chars[start - from + index] = primitive.text[index]!;
      }
    }
  }
  const [start, end] = findWordBoundaries(chars.join(''), position - from);
  return start < end ? { anchor: from + start, head: from + end } : null;
}
