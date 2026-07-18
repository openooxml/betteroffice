/**
 * Run Conversion
 *
 * Converts ProseMirror inline content (text, tab, image, field, math, sdt,
 * hardBreak) into the layout engine's Run[] representation, with mark-driven
 * formatting (bold/italic/color/font/etc.) extracted from each child.
 */

import type { EditorTreeMark as Mark, EditorTreeNode as PMNode } from '../../../types/editorTree';
import type {
  Run,
  TextRun,
  TabRun,
  ImageRun,
  FieldRun,
  ParagraphBlock,
  ShapeBlock,
  ChartBlock,
  ShapeBlockFill,
  ShapeBlockStroke,
  ShapeBlockTransform,
  RunFormatting,
} from '../../pagination/types';
import type { InlineSdtWidget } from '../../pagination/types';
import type {
  LegacyChartAttrs,
  LegacyFontFamilyAttrs,
  LegacyFontSizeAttrs,
  LegacyParagraphAttrs,
  LegacyShapeAttrs,
  LegacyTextColorAttrs,
  LegacyUnderlineAttrs,
} from './legacyTypes';
import type {
  Paragraph as DocumentParagraph,
  ParagraphContent,
  Run as DocumentRun,
  Shape,
  TextFormatting,
  Theme,
} from '../../../types/document';
import type { ShapeFillPaint } from '../../../types/content/shape';
import { resolveColor, resolveHighlightToCss } from '../../../utils/colorResolver';
import { emuToPixels, halfPointsToPixels, halfPointsToPoints } from '../../../utils/units';
import { resolveColorValueToHex } from '../../../docx/drawingUtils';
import { presetGeometryToPath } from '@betteroffice/drawingml';
import type { PresetGeometryPathCommand } from '@betteroffice/drawingml';
import { twipsToPixels, constrainImageToPage, nextBlockId } from './shared';
import type { ToLayoutBlocksOptions } from './shared';

/**
 * Extract run formatting from ProseMirror marks.
 */
function extractRunFormatting(marks: readonly Mark[], theme?: Theme | null): RunFormatting {
  const formatting: RunFormatting = {};

  for (const mark of marks) {
    switch (mark.type.name) {
      case 'bold':
        formatting.bold = true;
        if (mark.attrs.cs === true || mark.attrs.complexScript === true) formatting.boldCs = true;
        break;

      case 'italic':
        formatting.italic = true;
        if (mark.attrs.cs === true || mark.attrs.complexScript === true) formatting.italicCs = true;
        break;

      case 'underline': {
        const attrs = mark.attrs as LegacyUnderlineAttrs;
        if (attrs.style || attrs.color) {
          const underlineColor = attrs.color ? resolveColor(attrs.color, theme) : undefined;
          formatting.underline = {
            style: attrs.style,
            color: underlineColor,
          };
        } else {
          formatting.underline = true;
        }
        break;
      }

      case 'strike':
        formatting.strike = true;
        break;

      case 'textColor': {
        const attrs = mark.attrs as LegacyTextColorAttrs;
        if (attrs.themeColor || attrs.rgb) {
          formatting.color = resolveColor(
            {
              rgb: attrs.rgb,
              themeColor: attrs.themeColor,
              themeTint: attrs.themeTint,
              themeShade: attrs.themeShade,
            },
            theme
          );
        }
        break;
      }

      case 'highlight':
        formatting.highlight = resolveHighlightToCss(mark.attrs.color as string);
        break;

      case 'fontSize': {
        const attrs = mark.attrs as LegacyFontSizeAttrs;
        // Keep both authored sizes. Script selection belongs to shaping, not
        // to the presence of a run-level rtl mark: Arabic in a bidi paragraph
        // commonly has no w:rtl on each individual run.
        const isRtl = marks.some((m) => m.type.name === 'rtl');
        const legacySize = isRtl && attrs.sizeCs != null ? attrs.sizeCs : attrs.size;
        if (legacySize != null) formatting.fontSize = legacySize / 2;
        if (attrs.sizeCs != null) formatting.fontSizeCs = attrs.sizeCs / 2;
        break;
      }

      case 'fontFamily': {
        const attrs = mark.attrs as LegacyFontFamilyAttrs & {
          hint?: 'default' | 'eastAsia' | 'cs' | null;
        };
        formatting.fontSlots = {
          ascii: attrs.ascii,
          hAnsi: attrs.hAnsi,
          eastAsia: attrs.eastAsia,
          cs: attrs.cs,
          asciiTheme: attrs.asciiTheme,
          hAnsiTheme: attrs.hAnsiTheme,
          eastAsiaTheme: attrs.eastAsiaTheme,
          csTheme: attrs.csTheme,
          hint: attrs.hint ?? undefined,
        };
        // Legacy consumers still need one family; the shaping input above is
        // authoritative and selects a slot per cluster.
        const isRtl = marks.some((m) => m.type.name === 'rtl');
        formatting.fontFamily =
          (isRtl ? attrs.cs : undefined) ||
          attrs.ascii ||
          attrs.hAnsi ||
          attrs.eastAsia ||
          attrs.cs;
        break;
      }

      // Batch B projects w:lang/w:cs and the independent complex-script
      // variants into marks with these names. Keep this lowering tolerant of
      // older schemas so Batch C can build on Batch A without redefining the
      // shared contract.
      case 'language':
        formatting.language = {
          latin: (mark.attrs.latin ?? mark.attrs.val) as string | undefined,
          eastAsia: mark.attrs.eastAsia as string | undefined,
          bidi: mark.attrs.bidi as string | undefined,
        };
        break;

      case 'complexScript':
        formatting.complexScript = mark.attrs.enabled !== false;
        if (typeof mark.attrs.bold === 'boolean') formatting.boldCs = mark.attrs.bold;
        if (typeof mark.attrs.italic === 'boolean') formatting.italicCs = mark.attrs.italic;
        break;

      case 'characterSpacing': {
        // The PM `characterSpacing` mark is a multi-attribute container for
        // four OOXML run-level properties: w:spacing (letter-spacing,
        // §17.3.2.35), w:position (baseline shift, §17.3.2.24), w:w
        // (horizontal text scale, §17.3.2.43), and w:kern (kerning
        // threshold, §17.3.2.18). All four are parsed into the PM mark and
        // rendered correctly in the hidden ProseMirror toDOM, but the
        // layout dropped every attribute except the one we explicitly
        // case'd, so painted runs lost the values.
        const attrs = mark.attrs as {
          spacing: number | null;
          position: number | null;
          scale: number | null;
          kerning: number | null;
        };
        if (attrs.spacing != null && attrs.spacing !== 0) {
          formatting.letterSpacing = twipsToPixels(attrs.spacing);
        }
        if (attrs.position != null && attrs.position !== 0) {
          // w:position is half-points; positive raises (CSS vertical-align
          // positive raises too).
          formatting.positionPx = halfPointsToPixels(attrs.position);
        }
        if (attrs.scale != null && attrs.scale !== 100) {
          formatting.horizontalScale = attrs.scale;
        }
        if (attrs.kerning != null && attrs.kerning > 0) {
          // w:kern is in half-points; convert to points so the painter can
          // gate `font-kerning` by comparing against the run's font size.
          formatting.kerningMinPt = halfPointsToPoints(attrs.kerning);
        }
        break;
      }

      case 'allCaps':
        formatting.allCaps = true;
        break;

      case 'smallCaps':
        formatting.smallCaps = true;
        break;

      case 'emboss':
        formatting.emboss = true;
        break;

      case 'imprint':
        formatting.imprint = true;
        break;

      case 'textShadow':
        formatting.textShadow = true;
        break;

      case 'textOutline':
        formatting.textOutline = true;
        break;

      case 'hidden':
        formatting.hidden = true;
        break;

      case 'rtl':
        formatting.rtl = true;
        break;

      case 'textEffect': {
        const effect = mark.attrs.effect as string | undefined;
        if (
          effect === 'blinkBackground' ||
          effect === 'lights' ||
          effect === 'antsBlack' ||
          effect === 'antsRed' ||
          effect === 'shimmer' ||
          effect === 'sparkle'
        ) {
          formatting.textEffect = effect;
        }
        break;
      }

      case 'modernTextEffects': {
        // Modern w14 effects ride the display-list contract losslessly; the
        // canvas backend owns the paint approximation.
        const effects = mark.attrs.effects as RunFormatting['modernEffects'] | null;
        if (effects) formatting.modernEffects = effects;
        break;
      }

      case 'emphasisMark': {
        // CJK emphasis marks (§17.3.2.12). The PM mark stores the variant
        // type as `attrs.type`; pass it through so the painter can look up
        // the matching CSS text-emphasis style.
        const t = mark.attrs.type as string | undefined;
        if (t === 'dot' || t === 'comma' || t === 'circle' || t === 'underDot') {
          formatting.emphasisMark = t;
        } else {
          // Unknown variant — fall back to dot (Word's default).
          formatting.emphasisMark = 'dot';
        }
        break;
      }

      case 'superscript':
        formatting.superscript = true;
        break;

      case 'subscript':
        formatting.subscript = true;
        break;

      case 'hyperlink': {
        const attrs = mark.attrs as { href: string; tooltip?: string };
        formatting.hyperlink = {
          href: attrs.href,
          tooltip: attrs.tooltip,
        };
        break;
      }

      case 'footnoteRef': {
        const attrs = mark.attrs as { id: string | number; noteType?: string };
        const id = typeof attrs.id === 'string' ? parseInt(attrs.id, 10) : attrs.id;
        if (attrs.noteType === 'endnote') {
          formatting.endnoteRefId = id;
        } else {
          formatting.footnoteRefId = id;
        }
        // A footnote/endnote reference anchor renders superscript by default:
        // Word's built-in FootnoteReference / EndnoteReference character style
        // sets w:vertAlign="superscript". The OOXML often omits an explicit
        // rStyle on the anchor run (e.g. Pandoc-generated documents author a
        // bare `<w:r><w:footnoteReference/></w:r>`), so apply the superscript
        // implicitly here to match Word / LibreOffice rather than depending on
        // the rStyle being present.
        formatting.superscript = true;
        break;
      }

      case 'comment': {
        const commentId = mark.attrs.commentId as number;
        if (commentId) {
          if (!formatting.commentIds) formatting.commentIds = [];
          formatting.commentIds.push(commentId);
        }
        break;
      }

      case 'insertion':
        formatting.isInsertion = true;
        formatting.changeAuthor = mark.attrs.author as string;
        formatting.changeDate = mark.attrs.date as string;
        formatting.changeRevisionId = mark.attrs.revisionId as number;
        break;

      case 'deletion':
        formatting.isDeletion = true;
        formatting.changeAuthor = mark.attrs.author as string;
        formatting.changeDate = mark.attrs.date as string;
        formatting.changeRevisionId = mark.attrs.revisionId as number;
        break;
    }
  }

  return formatting;
}

/**
 * Resolve the paragraph's style-cascaded run defaults into a `RunFormatting`
 * baseline that individual runs can inherit. Per ECMA-376 §17.3.2.27 a run
 * with a partial `w:rFonts` (e.g. only `w:eastAsia`) inherits the missing
 * sides from the paragraph style → basedOn chain → docDefaults; without
 * this, runs whose own mark omits `ascii`/`hAnsi` lose the style's font and
 * fall back to the painter's hardcoded Calibri stack (#392).
 */
function paragraphRunDefaults(pmAttrs: LegacyParagraphAttrs): RunFormatting {
  const dtf = pmAttrs.defaultTextFormatting as TextFormatting | undefined;
  if (!dtf) return {};
  const result: RunFormatting = {};
  if (dtf.fontFamily) {
    result.fontSlots = { ...dtf.fontFamily };
    const family =
      dtf.fontFamily.ascii || dtf.fontFamily.hAnsi || dtf.fontFamily.eastAsia || dtf.fontFamily.cs;
    if (family) result.fontFamily = family;
  }
  if (dtf.fontSize != null) result.fontSize = dtf.fontSize / 2;
  if (dtf.fontSizeCs != null) result.fontSizeCs = dtf.fontSizeCs / 2;
  if (dtf.boldCs != null) result.boldCs = dtf.boldCs;
  if (dtf.italicCs != null) result.italicCs = dtf.italicCs;
  if (dtf.cs != null) result.complexScript = dtf.cs;
  if (dtf.language) result.language = { ...dtf.language };
  return result;
}

/**
 * Hyperlinks inside TOC paragraphs render in the TOCx paragraph color, not
 * the Hyperlink character style's blue/underline. Strip the resolved
 * color/underline so the painter's link fallback doesn't fire; the PM doc
 * keeps the original marks so copy/paste out of a TOC carries the Hyperlink
 * styling like Word does. Applies to both text and field runs (a TOC entry's
 * page number is a PAGEREF field inside the entry's hyperlink).
 */
function stripTocHyperlinkStyle(formatting: RunFormatting): void {
  if (!formatting.hyperlink) return;
  formatting.hyperlink = { ...formatting.hyperlink, noDefaultStyle: true };
  delete formatting.color;
  delete formatting.underline;
}

function isContentLocked(lock: unknown): boolean {
  return lock === 'contentLocked' || lock === 'sdtContentLocked';
}

function inlineCheckboxWidgetFor(child: PMNode, childPos: number): InlineSdtWidget | undefined {
  const attrs = child.attrs as Record<string, unknown>;
  if (attrs.sdtType !== 'checkbox') return undefined;
  if (isContentLocked(attrs.lock) || attrs.dataBinding != null) return undefined;
  return {
    kind: 'checkbox',
    groupId: `sdt@${childPos}`,
    pos: childPos,
    tag: attrs.tag != null ? String(attrs.tag) : undefined,
    alias: attrs.alias != null ? String(attrs.alias) : undefined,
    checked: typeof attrs.checked === 'boolean' ? attrs.checked : undefined,
  };
}

const reportedUnsupportedShapeTypes = new Set<string>();

function parseGradientStops(value: unknown): ShapeBlockFill['gradientStops'] {
  if (typeof value !== 'string' || value.length === 0) return undefined;
  try {
    const parsed = JSON.parse(value) as unknown;
    if (!Array.isArray(parsed)) return undefined;
    return parsed
      .map((stop) => {
        if (typeof stop !== 'object' || stop == null) return null;
        const s = stop as { position?: unknown; color?: unknown };
        if (typeof s.position !== 'number' || typeof s.color !== 'string') return null;
        return { position: s.position, color: s.color };
      })
      .filter((stop): stop is { position: number; color: string } => stop != null);
  } catch {
    return undefined;
  }
}

/**
 * Build a ShapeBlockFill from the lossless `ShapeFillPaint` payload for the
 * paint kinds the legacy fillColor/fillType attrs cannot express (pattern
 * foreground/background, picture source/crop/tile/stretch, theme index).
 * Returns undefined for kinds the legacy path already carries losslessly.
 */
function shapeBlockFillFromPaint(
  paint: ShapeFillPaint | null | undefined
): ShapeBlockFill | undefined {
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
      // safe embedded sources only — the parser resolves data:/blob: media and
      // the canvas image resolver re-checks the scheme (defense in depth)
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

function shapeFillFromAttrs(attrs: LegacyShapeAttrs): ShapeBlockFill | undefined {
  const fromPaint = shapeBlockFillFromPaint(attrs.fillPaint);
  if (fromPaint) return fromPaint;
  const fillType = attrs.fillType ?? 'solid';
  if (fillType === 'none') return { type: 'none' };
  if (fillType === 'gradient') {
    return {
      type: 'gradient',
      color: attrs.fillColor,
      gradientType: attrs.gradientType,
      gradientAngle: attrs.gradientAngle,
      gradientStops: parseGradientStops(attrs.gradientStops),
    };
  }
  if (fillType === 'pattern' || fillType === 'picture') {
    return { type: fillType, color: attrs.fillColor };
  }
  return {
    type: 'solid',
    color: attrs.fillColor,
  };
}

function shapeStrokeFromAttrs(attrs: LegacyShapeAttrs): ShapeBlockStroke | undefined {
  if (attrs.outlineWidth == null && !attrs.outlineColor && !attrs.outlineStyle) return undefined;
  return {
    color: attrs.outlineColor,
    width: attrs.outlineWidth,
    dash: attrs.outlineStyle,
  };
}

function shapeTransformFromAttrs(attrs: LegacyShapeAttrs): ShapeBlockTransform | undefined {
  const transform: ShapeBlockTransform = {};
  if (typeof attrs.rotation === 'number') {
    transform.rotation = attrs.rotation;
  } else if (attrs.transform) {
    const rotateMatch = attrs.transform.match(/rotate\(([-\d.]+)deg\)/);
    if (rotateMatch) {
      transform.rotation = parseFloat(rotateMatch[1]);
    }
  }
  if (attrs.flipH || attrs.transform?.includes('scaleX(-1)')) transform.flipH = true;
  if (attrs.flipV || attrs.transform?.includes('scaleY(-1)')) transform.flipV = true;
  return Object.keys(transform).length > 0 ? transform : undefined;
}

function shapeGeometryPathFromAttrs(attrs: LegacyShapeAttrs): PresetGeometryPathCommand[] | null {
  return attrs.geometryPath?.length ? attrs.geometryPath : null;
}

function parseShapeChildren(value: unknown): Shape[] {
  if (typeof value !== 'string' || value.length === 0) return [];
  try {
    const parsed = JSON.parse(value) as unknown;
    return Array.isArray(parsed)
      ? parsed.filter((item): item is Shape => {
          return typeof item === 'object' && item != null && (item as Shape).type === 'shape';
        })
      : [];
  } catch {
    return [];
  }
}

function cssColorFromTextFormatting(formatting: TextFormatting | undefined): string | undefined {
  return resolveColorValueToHex(formatting?.color);
}

function runFormattingFromTextFormatting(formatting: TextFormatting | undefined): RunFormatting {
  const result: RunFormatting = {};
  if (!formatting) return result;
  if (formatting.bold) result.bold = true;
  if (formatting.boldCs != null) result.boldCs = formatting.boldCs;
  if (formatting.italic) result.italic = true;
  if (formatting.italicCs != null) result.italicCs = formatting.italicCs;
  if (formatting.underline) result.underline = true;
  if (formatting.strike) result.strike = true;
  const color = cssColorFromTextFormatting(formatting);
  if (color) result.color = color;
  if (formatting.fontSize != null) result.fontSize = formatting.fontSize / 2;
  if (formatting.fontSizeCs != null) result.fontSizeCs = formatting.fontSizeCs / 2;
  if (formatting.fontFamily) result.fontSlots = { ...formatting.fontFamily };
  const fontFamily =
    formatting.fontFamily?.ascii ||
    formatting.fontFamily?.hAnsi ||
    formatting.fontFamily?.eastAsia ||
    formatting.fontFamily?.cs;
  if (fontFamily) result.fontFamily = fontFamily;
  if (formatting.cs != null) result.complexScript = formatting.cs;
  if (formatting.language) result.language = { ...formatting.language };
  return result;
}

function imageTransformMetrics(
  width: number,
  height: number,
  transform: string | undefined
): Pick<ImageRun, 'rotationDeg' | 'flipH' | 'flipV' | 'rotationBounds'> {
  const rotate = transform?.match(/rotate\(\s*(-?\d+(?:\.\d+)?)deg\s*\)/i);
  const parsed = rotate ? Number(rotate[1]) : 0;
  const rotationDeg = Number.isFinite(parsed) ? ((parsed % 360) + 360) % 360 : 0;
  const radians = (rotationDeg * Math.PI) / 180;
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  const boundsWidth = width * cos + height * sin;
  const boundsHeight = width * sin + height * cos;
  return {
    rotationDeg: rotationDeg || undefined,
    flipH: transform?.includes('scaleX(-1)') || undefined,
    flipV: transform?.includes('scaleY(-1)') || undefined,
    rotationBounds:
      rotationDeg !== 0
        ? {
            width: boundsWidth,
            height: boundsHeight,
            offsetX: (boundsWidth - width) / 2,
            offsetY: (boundsHeight - height) / 2,
          }
        : undefined,
  };
}

function runsFromDocumentRun(run: DocumentRun): Run[] {
  const formatting = runFormattingFromTextFormatting(run.formatting);
  const runs: Run[] = [];
  for (const content of run.content) {
    switch (content.type) {
      case 'text':
        if (content.text) {
          runs.push({ kind: 'text', text: content.text, ...formatting });
        }
        break;
      case 'tab':
        runs.push({ kind: 'tab', ...formatting });
        break;
      case 'break':
        runs.push({ kind: 'lineBreak' });
        break;
    }
  }
  return runs;
}

function runsFromParagraphContent(content: ParagraphContent): Run[] {
  if (content.type === 'run') return runsFromDocumentRun(content);
  if (content.type === 'hyperlink') {
    return content.children.flatMap((child) =>
      child.type === 'run' ? runsFromDocumentRun(child) : []
    );
  }
  return [];
}

function documentParagraphToParagraphBlock(paragraph: DocumentParagraph): ParagraphBlock {
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
    runs: paragraph.content.flatMap(runsFromParagraphContent),
    attrs: {
      ...(alignment ? { alignment } : {}),
    },
  };
}

function shapeFillFromShape(shape: Shape): ShapeBlockFill | undefined {
  const fromPaint = shapeBlockFillFromPaint(shape.fillPaint);
  if (fromPaint) return fromPaint;
  const fill = shape.fill;
  if (!fill) return undefined;
  if (fill.type === 'none') return { type: 'none' };
  if (fill.type === 'gradient') {
    return {
      type: 'gradient',
      color: resolveColorValueToHex(fill.color),
      gradientType: fill.gradient?.type,
      gradientAngle: fill.gradient?.angle,
      gradientStops: fill.gradient?.stops.map((stop) => ({
        position: stop.position,
        color: resolveColorValueToHex(stop.color) ?? '#000000',
      })),
    };
  }
  return {
    type: fill.type,
    color: resolveColorValueToHex(fill.color),
  };
}

function shapeStrokeFromShape(shape: Shape): ShapeBlockStroke | undefined {
  const outline = shape.outline;
  if (!outline) return undefined;
  return {
    color: resolveColorValueToHex(outline.color),
    width: outline.width != null ? emuToPixels(outline.width) : undefined,
    dash: outline.style,
  };
}

function shapeTransformFromShape(shape: Shape): ShapeBlockTransform | undefined {
  const transform: ShapeBlockTransform = {};
  if (shape.transform?.rotation != null) transform.rotation = shape.transform.rotation;
  if (shape.transform?.flipH) transform.flipH = true;
  if (shape.transform?.flipV) transform.flipV = true;
  return Object.keys(transform).length > 0 ? transform : undefined;
}

function shapeToShapeBlock(shape: Shape, options: ToLayoutBlocksOptions): ShapeBlock | null {
  const shapeType = shape.shapeType || 'rect';
  const geometryPath = presetGeometryToPath(shapeType);
  if (!geometryPath) {
    if (!reportedUnsupportedShapeTypes.has(shapeType)) {
      reportedUnsupportedShapeTypes.add(shapeType);
      console.warn(
        `[openooxml] DrawingML preset geometry "${shapeType}" is deferred and was not emitted as a ShapeBlock.`
      );
    }
    return null;
  }

  const constrained = constrainImageToPage(
    emuToPixels(shape.size?.width) || 100,
    emuToPixels(shape.size?.height) || 80,
    options.pageContentHeight
  );
  return {
    kind: 'shape',
    id: nextBlockId(),
    shapeType,
    geometryPath,
    fill: shapeFillFromShape(shape),
    stroke: shapeStrokeFromShape(shape),
    transform: shapeTransformFromShape(shape),
    width: constrained.width,
    height: constrained.height,
    x: shape.offset ? emuToPixels(shape.offset.x) : undefined,
    y: shape.offset ? emuToPixels(shape.offset.y) : undefined,
    innerText: shape.textBody?.content.map(documentParagraphToParagraphBlock),
    children: shape.children
      ?.map((child) => shapeToShapeBlock(child, options))
      .filter((child): child is ShapeBlock => child != null),
  };
}

/**
 * Convert a PM shape atom to the upstream layout shape contract. Unsupported
 * preset geometries return null so the v1 path only emits shapes with a real
 * normalized geometry.
 */
export function shapeNodeToShapeBlock(
  node: PMNode,
  startPos: number,
  options: ToLayoutBlocksOptions
): ShapeBlock | null {
  const attrs = node.attrs as LegacyShapeAttrs;
  const shapeType = attrs.shapeType || 'rect';
  const geometryPath = shapeGeometryPathFromAttrs(attrs) ?? presetGeometryToPath(shapeType);
  if (!geometryPath) {
    if (!reportedUnsupportedShapeTypes.has(shapeType)) {
      reportedUnsupportedShapeTypes.add(shapeType);
      console.warn(
        `[openooxml] DrawingML preset geometry "${shapeType}" is deferred and was not emitted as a ShapeBlock.`
      );
    }
    return null;
  }

  const constrained = constrainImageToPage(
    attrs.width || 100,
    attrs.height || 80,
    options.pageContentHeight
  );
  const docEnd = startPos + node.nodeSize;

  return {
    kind: 'shape',
    id: nextBlockId(),
    shapeType,
    geometryPath,
    fill: shapeFillFromAttrs(attrs),
    stroke: shapeStrokeFromAttrs(attrs),
    transform: shapeTransformFromAttrs(attrs),
    width: constrained.width,
    height: constrained.height,
    children: parseShapeChildren(attrs.children)
      .map((child) => shapeToShapeBlock(child, options))
      .filter((child): child is ShapeBlock => child != null),
    docStart: startPos,
    docEnd,
    pmStart: startPos,
    pmEnd: docEnd,
  };
}

export function chartNodeToChartBlock(
  node: PMNode,
  startPos: number,
  options: ToLayoutBlocksOptions
): ChartBlock | null {
  const attrs = node.attrs as LegacyChartAttrs;
  if (!attrs.chartJson) return null;
  try {
    const chart = JSON.parse(attrs.chartJson) as ChartBlock['chart'];
    const constrained = constrainImageToPage(
      attrs.width || 320,
      attrs.height || 220,
      options.pageContentHeight
    );
    const docEnd = startPos + node.nodeSize;
    return {
      kind: 'chart',
      id: nextBlockId(),
      chart,
      width: constrained.width,
      height: constrained.height,
      docStart: startPos,
      docEnd,
      pmStart: startPos,
      pmEnd: docEnd,
    };
  } catch {
    return null;
  }
}

/**
 * Convert a paragraph node to runs.
 */
export function paragraphToRuns(
  node: PMNode,
  startPos: number,
  _options: ToLayoutBlocksOptions
): Run[] {
  const runs: Run[] = [];
  const offset = startPos + 1; // +1 for opening tag
  const theme = _options.theme;
  const paraDefaults = paragraphRunDefaults(node.attrs as LegacyParagraphAttrs);

  // Hyperlinks inside TOC paragraphs use the TOCx color, not the Hyperlink
  // character style's color — see `HyperlinkInfo.noDefaultStyle`.
  const styleId = (node.attrs as LegacyParagraphAttrs).styleId;
  const inTocParagraph = typeof styleId === 'string' && /^TOC\d*$/i.test(styleId);

  // Single dispatcher for one inline PM child. Recurses on `sdt` so nested
  // content controls keep contributing runs at the right pmStart/pmEnd.
  function pushRunsForChild(
    child: PMNode,
    childPos: number,
    inlineSdtWidget?: InlineSdtWidget
  ): void {
    if (child.isText && child.text) {
      const formatting = extractRunFormatting(child.marks, theme);
      if (inTocParagraph) stripTocHyperlinkStyle(formatting);
      const run: TextRun = {
        kind: 'text',
        text: child.text,
        ...paraDefaults,
        ...formatting,
        logicalOrder: runs.length,
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
        inlineSdtWidget,
      };
      runs.push(run);
    } else if (child.type.name === 'hardBreak') {
      runs.push({
        kind: 'lineBreak',
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
    } else if (child.type.name === 'tab') {
      const formatting = extractRunFormatting(child.marks, theme);
      const run: TabRun = {
        kind: 'tab',
        ...paraDefaults,
        ...formatting,
        logicalOrder: runs.length,
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'image') {
      const attrs = child.attrs;
      const constrained = constrainImageToPage(
        (attrs.width as number) || 100,
        (attrs.height as number) || 100,
        _options.pageContentHeight
      );
      // Carry the image's tracked-change marks so an inserted/deleted picture
      // paints in the revision color and resolves with the rest of the change.
      const changeFmt = extractRunFormatting(child.marks, theme);
      const transform = attrs.transform as string | undefined;
      const run: ImageRun = {
        kind: 'image',
        src: attrs.src as string,
        width: constrained.width,
        height: constrained.height,
        alt: attrs.alt as string | undefined,
        transform,
        ...imageTransformMetrics(constrained.width, constrained.height, transform),
        wrapType: attrs.wrapType as string | undefined,
        displayMode: attrs.displayMode as 'inline' | 'block' | 'float' | undefined,
        cssFloat: attrs.cssFloat as 'left' | 'right' | 'none' | undefined,
        distTop: attrs.distTop as number | undefined,
        distBottom: attrs.distBottom as number | undefined,
        distLeft: attrs.distLeft as number | undefined,
        distRight: attrs.distRight as number | undefined,
        position: attrs.position as ImageRun['position'] | undefined,
        cropTop: attrs.cropTop as number | undefined,
        cropRight: attrs.cropRight as number | undefined,
        cropBottom: attrs.cropBottom as number | undefined,
        cropLeft: attrs.cropLeft as number | undefined,
        opacity: attrs.opacity as number | undefined,
        isInsertion: changeFmt.isInsertion,
        isDeletion: changeFmt.isDeletion,
        changeAuthor: changeFmt.changeAuthor,
        changeDate: changeFmt.changeDate,
        changeRevisionId: changeFmt.changeRevisionId,
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'shape') {
      // Shape atoms are promoted to ShapeBlock by toLayoutBlocks so the paged
      // layout can place them as first-class DrawingML objects.
      return;
    } else if (child.type.name === 'chart') {
      return;
    } else if (child.type.name === 'field') {
      const ft = child.attrs.fieldType as string;
      const mappedType: FieldRun['fieldType'] =
        ft === 'PAGE'
          ? 'PAGE'
          : ft === 'NUMPAGES'
            ? 'NUMPAGES'
            : ft === 'DATE'
              ? 'DATE'
              : ft === 'TIME'
                ? 'TIME'
                : 'OTHER';
      // Field nodes carry the same character marks as text runs (the result
      // run's w:rPr). Without extracting them the painted page number would
      // fall back to the painter's hardcoded defaults instead of the footer
      // run's font/size/color — Word renders the field result with the run's
      // own formatting.
      const formatting = extractRunFormatting(child.marks, theme);
      if (inTocParagraph) stripTocHyperlinkStyle(formatting);
      // inert a11y identity: raw type token + instruction ride along so the
      // display list / mirror can announce what the field is. Never executed.
      const instruction = (child.attrs.instruction as string) || undefined;
      runs.push({
        kind: 'field',
        fieldType: mappedType,
        ...(ft && ft !== mappedType ? { rawType: ft } : {}),
        ...(instruction ? { instruction } : {}),
        fallback: (child.attrs.displayText as string) || '',
        ...paraDefaults,
        ...formatting,
        logicalOrder: runs.length,
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
    } else if (child.type.name === 'math') {
      const text = (child.attrs.plainText as string) || '[equation]';
      runs.push({
        kind: 'text',
        text,
        italic: true,
        fontFamily: 'Cambria Math',
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
    } else if (child.type.name === 'sdt') {
      const inlineWidget = inlineCheckboxWidgetFor(child, childPos) ?? inlineSdtWidget;
      const sdtInnerOffset = childPos + 1; // +1 for opening tag
      child.forEach((sdtChild, sdtChildOffset) => {
        pushRunsForChild(sdtChild, sdtInnerOffset + sdtChildOffset, inlineWidget);
      });
    }
  }

  node.forEach((child, childOffset) => {
    pushRunsForChild(child, offset + childOffset);
  });

  return runs;
}
