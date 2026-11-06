import type { SlideDisplayList, SlidePrimitive } from '../types';
import type { PaintFontCoverage, TextFamilies } from './canvas';

/** Every family a display list draws text with, mapped to the family its prepared face is installed under. */
export async function prepareTextFamilies(list: SlideDisplayList, coverage: PaintFontCoverage): Promise<TextFamilies> {
  const texts = new Map<string, string[]>();
  const covered = (family: string, bold?: boolean, italic?: boolean) => coverage.covers?.(family, bold === true, italic === true) ?? true;
  const visit = (primitives: readonly SlidePrimitive[]) => {
    for (const primitive of primitives) {
      if (primitive.kind === 'textBox') {
        for (const line of primitive.lines) {
          for (const run of line.runs) {
            if (!covered(run.fontFamily, run.bold, run.italic)) continue;
            const parts = texts.get(run.fontFamily) ?? [];
            parts.push(run.text);
            texts.set(run.fontFamily, parts);
          }
        }
      } else if (primitive.kind === 'chart' || primitive.kind === 'table') visit(primitive.primitives);
    }
  };
  visit(list.primitives);
  const prepared = new Map<string, string>();
  await Promise.all([...texts].map(async ([family, parts]) => {
    prepared.set(family, await coverage.prepare(family, parts.join('\n')));
  }));
  return (family, bold, italic) => {
    if (!covered(family, bold, italic)) return family;
    const drawn = prepared.get(family);
    if (drawn === undefined) throw new Error(`Text in ${family} was painted without preparing its font`);
    return drawn === family ? family : coverage.latest?.(family) ?? drawn;
  };
}
