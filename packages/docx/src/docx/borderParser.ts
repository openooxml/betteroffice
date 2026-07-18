/**
 * Shared border parsing for every `CT_Border` field — table borders
 * (`w:tblBorders`/`w:tcBorders`), paragraph borders (`w:pBdr`), and page borders
 * (`w:pgBorders`). Mirrors {@link ../serializer/borderSerializer} on the parse
 * side: the table/paragraph/section/style parsers all delegate here so a border
 * is read identically everywhere (one place to fix, one place to harden).
 */

import { findChild, getAttribute, parseNumericAttribute, type XmlElement } from './xmlParser';
import type { BorderSpec, ColorValue, ParagraphFormatting, TableBorders } from '../types/document';

function parseColorValue(
  rgb: string | null,
  themeColor: string | null,
  themeTint: string | null,
  themeShade: string | null
): ColorValue {
  return {
    ...(rgb && rgb !== 'auto' ? { rgb } : {}),
    ...(rgb === 'auto' ? { auto: true } : {}),
    ...(themeColor ? { themeColor: themeColor as ColorValue['themeColor'] } : {}),
    ...(themeTint ? { themeTint } : {}),
    ...(themeShade ? { themeShade } : {}),
  };
}

/**
 * Parse a single border element (`<w:top>`, `<w:left>`, ...) into a
 * {@link BorderSpec}.
 *
 * A missing `w:val` defaults to `none` (the attribute is required by the schema,
 * so this only guards malformed input) and `w:color="auto"` maps to
 * `{ auto: true }` via {@link parseColorValue} rather than being stored as a
 * literal `rgb`. Returns `undefined` only when the element itself is absent.
 */
export function parseBorderSpec(element: XmlElement | null): BorderSpec | undefined {
  if (!element) return undefined;

  const style = (getAttribute(element, 'w', 'val') ?? 'none') as BorderSpec['style'];
  const border: BorderSpec = { style };

  // Size in eighths of a point.
  const sz = parseNumericAttribute(element, 'w', 'sz');
  if (sz !== undefined) border.size = sz;

  // Space from text/page edge in points.
  const space = parseNumericAttribute(element, 'w', 'space');
  if (space !== undefined) border.space = space;

  // Color — only set when an actual color attribute is present, so a plain
  // border doesn't carry an empty color object. `parseColorValue` handles
  // `auto` and the theme attributes.
  const colorVal = getAttribute(element, 'w', 'color');
  const themeColor = getAttribute(element, 'w', 'themeColor');
  const themeTint = getAttribute(element, 'w', 'themeTint');
  const themeShade = getAttribute(element, 'w', 'themeShade');
  if (colorVal || themeColor || themeTint || themeShade) {
    border.color = parseColorValue(colorVal, themeColor, themeTint, themeShade);
  }

  const shadow = getAttribute(element, 'w', 'shadow');
  if (shadow === '1' || shadow === 'true') border.shadow = true;

  const frame = getAttribute(element, 'w', 'frame');
  if (frame === '1' || frame === 'true') border.frame = true;

  return border;
}

/**
 * Parse a borders container (`w:tblBorders` or `w:tcBorders`) into
 * {@link TableBorders}. `left`/`right` fall back to the RTL `start`/`end`
 * aliases. Returns `undefined` when no side parsed.
 */
export function parseTableBorders(bordersElement: XmlElement | null): TableBorders | undefined {
  if (!bordersElement) return undefined;

  const borders: TableBorders = {};

  const top = parseBorderSpec(findChild(bordersElement, 'w', 'top'));
  if (top) borders.top = top;

  const bottom = parseBorderSpec(findChild(bordersElement, 'w', 'bottom'));
  if (bottom) borders.bottom = bottom;

  const left = parseBorderSpec(
    findChild(bordersElement, 'w', 'left') ?? findChild(bordersElement, 'w', 'start')
  );
  if (left) borders.left = left;

  const right = parseBorderSpec(
    findChild(bordersElement, 'w', 'right') ?? findChild(bordersElement, 'w', 'end')
  );
  if (right) borders.right = right;

  const insideH = parseBorderSpec(findChild(bordersElement, 'w', 'insideH'));
  if (insideH) borders.insideH = insideH;

  const insideV = parseBorderSpec(findChild(bordersElement, 'w', 'insideV'));
  if (insideV) borders.insideV = insideV;

  if (Object.keys(borders).length === 0) return undefined;

  return borders;
}

/**
 * Parse a paragraph borders container (`w:pBdr`) into
 * {@link ParagraphFormatting.borders}. Paragraph borders use the `between` and
 * `bar` sides instead of the table `insideH`/`insideV`. Returns `undefined`
 * when no side parsed.
 */
export function parseParagraphBorders(
  pBdr: XmlElement | null
): ParagraphFormatting['borders'] | undefined {
  if (!pBdr) return undefined;

  const borders: NonNullable<ParagraphFormatting['borders']> = {};
  for (const side of ['top', 'bottom', 'left', 'right', 'between', 'bar'] as const) {
    const spec = parseBorderSpec(findChild(pBdr, 'w', side));
    if (spec) borders[side] = spec;
  }

  return Object.keys(borders).length > 0 ? borders : undefined;
}
