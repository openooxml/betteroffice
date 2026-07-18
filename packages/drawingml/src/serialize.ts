/**
 * DrawingML serialization — a:-namespace color, fill, and outline emission
 * shared by format hosts. The host owns the surrounding envelope
 * (w:drawing/wp:anchor for docx, p:sp for pptx, ...).
 */

import type { ColorValue } from './color';
import type { ShapeFill, ShapeOutline } from './shape';

/**
 * Serialize a color value to DrawingML a:srgbClr or a:schemeClr
 *
 * @public
 */
export function serializeDrawingColor(color: ColorValue | undefined): string {
  if (!color) return '';
  if (color.rgb) {
    return `<a:srgbClr val="${color.rgb.replace('#', '')}"/>`;
  }
  if (color.themeColor) {
    let clr = `<a:schemeClr val="${color.themeColor}"`;
    if (color.themeTint) {
      clr += `><a:tint val="${color.themeTint}"/></a:schemeClr>`;
    } else if (color.themeShade) {
      clr += `><a:shade val="${color.themeShade}"/></a:schemeClr>`;
    } else {
      clr += `/>`;
    }
    return clr;
  }
  return '';
}

/**
 * Serialize shape fill to DrawingML
 *
 * @public
 */
export function serializeFill(fill: ShapeFill | undefined): string {
  if (!fill || fill.type === 'none') return '<a:noFill/>';
  if (fill.type === 'solid' && fill.color) {
    return `<a:solidFill>${serializeDrawingColor(fill.color)}</a:solidFill>`;
  }
  if (fill.type === 'gradient' && fill.gradient) {
    const g = fill.gradient;
    const stops = g.stops
      .map((s) => `<a:gs pos="${s.position}">${serializeDrawingColor(s.color)}</a:gs>`)
      .join('');
    const direction =
      g.type === 'linear' ? `<a:lin ang="${(g.angle || 0) * 60000}" scaled="1"/>` : '';
    return `<a:gradFill><a:gsLst>${stops}</a:gsLst>${direction}</a:gradFill>`;
  }
  return '';
}

/**
 * Serialize shape outline to DrawingML a:ln
 *
 * @public
 */
export function serializeOutline(outline: ShapeOutline | undefined): string {
  if (!outline) return '';
  const attrs: string[] = [];
  if (outline.width != null) attrs.push(`w="${outline.width}"`);
  if (outline.cap) attrs.push(`cap="${outline.cap}"`);

  const parts: string[] = [];
  if (outline.color) {
    parts.push(`<a:solidFill>${serializeDrawingColor(outline.color)}</a:solidFill>`);
  }
  if (outline.style && outline.style !== 'solid') {
    parts.push(`<a:prstDash val="${outline.style}"/>`);
  }
  if (outline.headEnd) {
    parts.push(
      `<a:headEnd type="${outline.headEnd.type}"${outline.headEnd.width ? ` w="${outline.headEnd.width}"` : ''}${outline.headEnd.length ? ` len="${outline.headEnd.length}"` : ''}/>`
    );
  }
  if (outline.tailEnd) {
    parts.push(
      `<a:tailEnd type="${outline.tailEnd.type}"${outline.tailEnd.width ? ` w="${outline.tailEnd.width}"` : ''}${outline.tailEnd.length ? ` len="${outline.tailEnd.length}"` : ''}/>`
    );
  }

  if (parts.length === 0 && attrs.length === 0) return '';
  return `<a:ln${attrs.length ? ' ' + attrs.join(' ') : ''}>${parts.join('')}</a:ln>`;
}
