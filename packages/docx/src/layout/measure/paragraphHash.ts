/**
 * Stable identity hash for a paragraph block, used as the Rust measure
 * memo's key so blocks that would lay out identically share one measurement.
 * Every field that can change the rendered result contributes a labelled
 * token; the assembled token list is serialized as JSON so no separator
 * character can leak in from content.
 *
 * @packageDocumentation
 * @public
 */

import type { ParagraphBlock } from '../pagination/types';

/**
 * Build the stable identity string for a paragraph block.
 *
 * @public
 */
export function hashParagraphBlock(block: ParagraphBlock): string {
  const tokens: unknown[] = [];

  for (const run of block.runs) {
    if (run.kind === 'text') {
      tokens.push(['t', run.text, run.fontFamily, run.fontSize, run.bold, run.italic]);
    } else if (run.kind === 'tab') {
      tokens.push(['tab', run.width]);
    } else if (run.kind === 'image') {
      tokens.push(['img', run.width, run.height]);
    } else if (run.kind === 'lineBreak') {
      tokens.push(['br']);
    }
  }

  const attrs = block.attrs;
  if (attrs) {
    if (attrs.alignment) tokens.push(['align', attrs.alignment]);
    if (attrs.indent) {
      tokens.push([
        'indent',
        attrs.indent.left,
        attrs.indent.right,
        attrs.indent.firstLine,
        attrs.indent.hanging,
      ]);
    }
    if (attrs.spacing) {
      tokens.push([
        'spacing',
        attrs.spacing.before,
        attrs.spacing.after,
        attrs.spacing.line,
        attrs.spacing.lineRule,
      ]);
    }
    // empty paragraphs have no runs, so the default font is what drives their
    // line height; folding it in keeps two differently-sized empties apart.
    if (attrs.defaultFontSize != null) tokens.push(['dfs', attrs.defaultFontSize]);
    if (attrs.defaultFontFamily != null) tokens.push(['dff', attrs.defaultFontFamily]);
    // a paragraph border is an authorial difference on otherwise-identical
    // empties (e.g. a pBdr horizontal rule vs none) and must not be shared.
    const b = attrs.borders;
    if (b) {
      const side = (s?: { width?: number; style?: string; color?: string }) =>
        s ? [s.width ?? null, s.style ?? null, s.color ?? null] : null;
      tokens.push(['bdr', side(b.top), side(b.bottom), side(b.left), side(b.right)]);
    }
    if (attrs.suppressEmptyParagraphHeight) tokens.push(['sup']);
    // paragraph-mark revisions add painted glyphs/change bars, so identical
    // text with differing revision state must not collapse to one entry.
    if (attrs.pPrIns) tokens.push(['pins', attrs.pPrIns.revisionId]);
    if (attrs.pPrDel) tokens.push(['pdel', attrs.pPrDel.revisionId]);
  }

  return JSON.stringify(tokens);
}
