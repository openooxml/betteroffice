import { presetGeometryToPath } from '@betteroffice/drawingml';

import { resolveColorValueToHex } from '../docx/drawingUtils';
import type {
  Paragraph,
  ParagraphContent,
  Run as DocumentRun,
  Shape,
  TextFormatting,
} from '../types/document';
import type { ShapeFillPaint } from '../types/content/shape';
import { emuToPixels } from '../utils/units';
import type {
  ParagraphBlock,
  Run,
  RunFormatting,
  ShapeBlock,
  ShapeBlockFill,
  ShapeBlockStroke,
  ShapeBlockTransform,
} from './pagination/types';
import { constrainImageToPage, nextBlockId } from './flow/toLayoutBlocks/shared';

/** Lower a document-model shape directly to the renderer contract. */
export function documentShapeToLayoutBlock(
  shape: Shape,
  pageContentHeight?: number
): ShapeBlock | null {
  const shapeType = shape.shapeType || 'rect';
  const geometryPath = shape.geometryPath?.length
    ? shape.geometryPath
    : presetGeometryToPath(shapeType);
  if (!geometryPath) return null;
  const constrained = constrainImageToPage(
    emuToPixels(shape.size?.width) || 100,
    emuToPixels(shape.size?.height) || 80,
    pageContentHeight
  );
  return {
    kind: 'shape',
    id: nextBlockId(),
    shapeType,
    geometryPath,
    fill: shapeFill(shape),
    stroke: shapeStroke(shape),
    transform: shapeTransform(shape),
    width: constrained.width,
    height: constrained.height,
    x: shape.offset ? emuToPixels(shape.offset.x) : undefined,
    y: shape.offset ? emuToPixels(shape.offset.y) : undefined,
    innerText: shape.textBody?.content.map(paragraphBlock),
    children: (shape.children ?? [])
      .map((child) => documentShapeToLayoutBlock(child, pageContentHeight))
      .filter((child): child is ShapeBlock => child != null),
    scene: shape.scene,
    effects: shape.effects,
    textBodyProperties: shape.textBodyProperties,
  };
}

function shapeFillPaint(paint: ShapeFillPaint | null | undefined): ShapeBlockFill | undefined {
  if (!paint?.kind) return undefined;
  if (paint.kind === 'pattern') {
    return {
      type: 'pattern',
      color: resolveColorValueToHex(paint.foregroundColor ?? paint.color),
      patternPreset: paint.patternPreset,
      foregroundColor: resolveColorValueToHex(paint.foregroundColor),
      backgroundColor: resolveColorValueToHex(paint.backgroundColor),
    };
  }
  if (paint.kind === 'picture') {
    const src = paint.picture?.src;
    return {
      type: 'picture',
      color: resolveColorValueToHex(paint.color),
      pictureRelId: paint.picture?.rId,
      pictureSrc: src && (src.startsWith('data:') || src.startsWith('blob:')) ? src : undefined,
      pictureSrcRect: paint.srcRect,
      pictureFillMode: paint.fillMode,
      pictureTile: paint.tile,
      pictureStretchRect: paint.stretchRect,
      pictureOpacity: paint.pictureOpacity,
    };
  }
  if (paint.kind === 'theme') {
    return {
      type: 'solid',
      color: resolveColorValueToHex(paint.color),
      themeRefIndex: paint.themeRefIndex,
    };
  }
  return undefined;
}

function shapeFill(shape: Shape): ShapeBlockFill | undefined {
  const detailed = shapeFillPaint(shape.fillPaint);
  if (detailed) return detailed;
  if (!shape.fill) return undefined;
  if (shape.fill.type === 'none') return { type: 'none' };
  if (shape.fill.type === 'gradient') {
    return {
      type: 'gradient',
      color: resolveColorValueToHex(shape.fill.color),
      gradientType: shape.fill.gradient?.type,
      gradientAngle: shape.fill.gradient?.angle,
      gradientStops: shape.fill.gradient?.stops.map((stop) => ({
        position: stop.position,
        color: resolveColorValueToHex(stop.color) ?? '#000000',
      })),
    };
  }
  return { type: shape.fill.type, color: resolveColorValueToHex(shape.fill.color) };
}

function shapeStroke(shape: Shape): ShapeBlockStroke | undefined {
  if (!shape.outline) return undefined;
  return {
    color: resolveColorValueToHex(shape.outline.color),
    width: shape.outline.width != null ? emuToPixels(shape.outline.width) : undefined,
    dash: shape.outline.style,
  };
}

function shapeTransform(shape: Shape): ShapeBlockTransform | undefined {
  const transform: ShapeBlockTransform = {};
  if (shape.transform?.rotation != null) transform.rotation = shape.transform.rotation;
  if (shape.transform?.flipH) transform.flipH = true;
  if (shape.transform?.flipV) transform.flipV = true;
  return Object.keys(transform).length > 0 ? transform : undefined;
}

function paragraphBlock(paragraph: Paragraph): ParagraphBlock {
  const sourceAlignment = paragraph.formatting?.alignment;
  const alignment =
    sourceAlignment === 'both'
      ? 'justify'
      : sourceAlignment === 'left' || sourceAlignment === 'center' || sourceAlignment === 'right'
        ? sourceAlignment
        : undefined;
  return {
    kind: 'paragraph',
    id: nextBlockId(),
    paraId: paragraph.paraId,
    runs: paragraph.content.flatMap(paragraphContentRuns),
    attrs: alignment ? { alignment } : {},
  };
}

function paragraphContentRuns(content: ParagraphContent): Run[] {
  if (content.type === 'run') return documentRuns(content);
  if (content.type === 'hyperlink') {
    return content.children.flatMap((child) => (child.type === 'run' ? documentRuns(child) : []));
  }
  return [];
}

function documentRuns(run: DocumentRun): Run[] {
  const formatting = runFormatting(run.formatting);
  const runs: Run[] = [];
  for (const content of run.content) {
    if (content.type === 'text' && content.text) {
      runs.push({ kind: 'text', text: content.text, ...formatting });
    } else if (content.type === 'tab') {
      runs.push({ kind: 'tab', ...formatting });
    } else if (content.type === 'break') {
      runs.push({ kind: 'lineBreak' });
    }
  }
  return runs;
}

function runFormatting(source: TextFormatting | undefined): RunFormatting {
  const formatting: RunFormatting = {};
  if (!source) return formatting;
  if (source.bold) formatting.bold = true;
  if (source.boldCs != null) formatting.boldCs = source.boldCs;
  if (source.italic) formatting.italic = true;
  if (source.italicCs != null) formatting.italicCs = source.italicCs;
  if (source.underline) formatting.underline = true;
  if (source.strike) formatting.strike = true;
  const color = resolveColorValueToHex(source.color);
  if (color) formatting.color = color;
  if (source.fontSize != null) formatting.fontSize = source.fontSize / 2;
  if (source.fontSizeCs != null) formatting.fontSizeCs = source.fontSizeCs / 2;
  if (source.fontFamily) formatting.fontSlots = { ...source.fontFamily };
  const family =
    source.fontFamily?.ascii ??
    source.fontFamily?.hAnsi ??
    source.fontFamily?.eastAsia ??
    source.fontFamily?.cs;
  if (family) formatting.fontFamily = family;
  if (source.cs != null) formatting.complexScript = source.cs;
  if (source.language) formatting.language = { ...source.language };
  return formatting;
}
