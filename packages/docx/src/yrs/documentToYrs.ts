/**
 * Direct Document -> yrs load projection.
 *
 * This is the Document-model twin of the coexistence seed path. It emits the
 * same flat story vocabulary as `seedFromPm`, but never constructs a
 * ProseMirror node: text becomes attributed Y.Text, paragraphs become pilcrow
 * embeds, tables/block SDTs become structural embeds with child stories, and
 * comments become root-story side-map entries.
 */

/* eslint-disable max-lines -- the complete load projection is intentionally co-located */

import { emuToPixels } from '../utils/units';
import { isWrapNone } from '../docx/wrapTypes';
import { resolveColorValueToHex } from '@betteroffice/drawingml';
import { sdtPropsToAttrs } from '../types/sdtAttributes';
import { createStyleResolver, type StyleResolver } from '../styles';
import type {
  BlockContent,
  Chart,
  ComplexField,
  Document,
  Hyperlink,
  Image,
  InlineSdt,
  MathEquation,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  Shape,
  SimpleField,
  SdtProperties,
  Table,
  TableBorders,
  TableCell,
  TableCellFormatting,
  TableRow,
  TextFormatting,
  Theme,
  TrackedChangeInfo,
} from '../types/document';
import { ensureHexPrefix, resolveColorToHex } from '../utils/colorResolver';
import { mergeTextFormatting } from '../utils/textFormattingMerge';
import { documentShapeToLayoutBlock } from '../layout/shapeBlocks';
import type { YrsRawOp, YrsSession } from './index';
import {
  blockSdtAttrsToPayload,
  blockSdtStoryId,
  dropNulls,
  footnoteStoryId,
  headerFooterStoryId,
  paraAttrsToPpr,
  tableAttrsToGrid,
  tableAttrsToTblPr,
  tableCellAttrsToTcPr,
  tableCellStoryId,
  tableRowAttrsToTrPr,
  type YrsAttrs,
} from './storyAttributes';

type Attrs = Record<string, unknown>;

interface MarkDescriptor {
  name: string;
  /** Complete PM-schema attrs, including non-null defaults. */
  attrs: Attrs;
}

interface TextUnit {
  kind: 'text';
  text: string;
  attrs: YrsAttrs;
  /** PM node width, used only for bookmark offsets. */
  pmSize: number;
  commentId?: number;
  marks: MarkDescriptor[];
}

interface EmbedUnit {
  kind: 'embed';
  embedKind: string;
  payload: Attrs;
  attrs: YrsAttrs;
  pmSize: number;
  commentId?: number;
  marks: MarkDescriptor[];
}

type InlineUnit = TextUnit | EmbedUnit;

interface StoryPlan {
  storyId: string;
  units: InlineUnit[];
  commentCoverage: Map<number, Array<[number, number]>>;
}

interface StoryOptions {
  includePageBreaks: boolean;
  appendBodyTail: boolean;
  seedComments: boolean;
  extraRunFormatting?: TextFormatting;
}

interface ProjectedCell {
  attrs: Attrs;
  content: BlockContent[];
  extraRunFormatting?: TextFormatting;
}

interface ProjectedRow {
  attrs: Attrs;
  cells: ProjectedCell[];
}

interface ProjectedTable {
  attrs: Attrs;
  rows: ProjectedRow[];
}

interface LoweringContext {
  styleResolver: StyleResolver | null;
  theme: Theme | null;
  plans: StoryPlan[];
}

const BOOLEAN_MARKS = new Set([
  'bold',
  'italic',
  'superscript',
  'subscript',
  'allCaps',
  'smallCaps',
  'emboss',
  'imprint',
  'textShadow',
  'textOutline',
  'hidden',
  'rtl',
]);

function stableStringify(value: unknown): string {
  if (value === null || value === undefined) return 'null';
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`;
  if (typeof value === 'object') {
    const object = value as Attrs;
    return `{${Object.keys(object)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${stableStringify(object[key])}`)
      .join(',')}}`;
  }
  return JSON.stringify(value);
}

function resolveTextFormatting(
  formatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null
): TextFormatting | undefined {
  if (!formatting || !styleResolver) return formatting;
  const styleFormatting = styleResolver.resolveRunStyle(formatting.styleId);
  return styleFormatting ? mergeTextFormatting(styleFormatting, formatting) : formatting;
}

function formattingToMarks(formatting: TextFormatting | undefined): MarkDescriptor[] {
  if (!formatting) return [];
  const marks: MarkDescriptor[] = [];
  const add = (name: string, attrs: Attrs = {}) => marks.push({ name, attrs });

  if (formatting.bold) add('bold');
  if (formatting.italic) add('italic');
  if (formatting.underline && formatting.underline.style !== 'none') {
    add('underline', {
      style: formatting.underline.style,
      color: formatting.underline.color ?? null,
    });
  }
  if (formatting.strike || formatting.doubleStrike) {
    add('strike', { double: formatting.doubleStrike || false });
  }
  if (formatting.color && !formatting.color.auto) {
    add('textColor', {
      rgb: formatting.color.rgb ?? null,
      themeColor: formatting.color.themeColor ?? null,
      themeTint: formatting.color.themeTint ?? null,
      themeShade: formatting.color.themeShade ?? null,
    });
  }

  const shadingFill = formatting.shading?.fill;
  const shadingHighlight =
    (!formatting.shading?.pattern || formatting.shading.pattern === 'clear') &&
    shadingFill?.rgb &&
    !shadingFill.auto
      ? ensureHexPrefix(shadingFill.rgb)
      : undefined;
  const highlight =
    formatting.highlight && formatting.highlight !== 'none'
      ? formatting.highlight
      : shadingHighlight;
  if (highlight) add('highlight', { color: highlight });

  if (formatting.fontSize != null || formatting.fontSizeCs != null) {
    add('fontSize', {
      size: formatting.fontSize ?? null,
      sizeCs: formatting.fontSizeCs ?? null,
    });
  }
  if (formatting.fontFamily) {
    add('fontFamily', {
      ascii: formatting.fontFamily.ascii ?? null,
      hAnsi: formatting.fontFamily.hAnsi ?? null,
      eastAsia: formatting.fontFamily.eastAsia ?? null,
      cs: formatting.fontFamily.cs ?? null,
      asciiTheme: formatting.fontFamily.asciiTheme ?? null,
      hAnsiTheme: formatting.fontFamily.hAnsiTheme ?? null,
      eastAsiaTheme: formatting.fontFamily.eastAsiaTheme ?? null,
      csTheme: formatting.fontFamily.csTheme ?? null,
    });
  }
  if (formatting.vertAlign === 'superscript') add('superscript');
  else if (formatting.vertAlign === 'subscript') add('subscript');
  if (formatting.allCaps) add('allCaps');
  if (formatting.smallCaps) add('smallCaps');
  if (
    formatting.spacing != null ||
    formatting.position != null ||
    formatting.scale != null ||
    formatting.kerning != null
  ) {
    add('characterSpacing', {
      spacing: formatting.spacing ?? null,
      position: formatting.position ?? null,
      scale: formatting.scale ?? null,
      kerning: formatting.kerning ?? null,
    });
  }
  if (formatting.emboss) add('emboss');
  if (formatting.imprint) add('imprint');
  if (formatting.shadow) add('textShadow');
  if (formatting.emphasisMark && formatting.emphasisMark !== 'none') {
    add('emphasisMark', { type: formatting.emphasisMark });
  }
  if (formatting.outline) add('textOutline');
  if (formatting.hidden) add('hidden');
  if (formatting.rtl) add('rtl');
  if (formatting.effect && formatting.effect !== 'none') {
    add('textEffect', { effect: formatting.effect });
  }
  if (formatting.modernEffects) {
    add('modernTextEffects', { effects: formatting.modernEffects });
  }
  if (formatting.styleId) add('runStyle', { styleId: formatting.styleId });
  return marks;
}

function marksToYrsAttrs(marks: readonly MarkDescriptor[]): YrsAttrs {
  const attrs: YrsAttrs = {};
  for (const mark of marks) {
    if (mark.name === 'comment' || mark.name === 'footnoteRef') continue;
    if (BOOLEAN_MARKS.has(mark.name)) {
      attrs[mark.name] = true;
    } else if (mark.name === 'highlight') {
      attrs.highlight = mark.attrs.color;
    } else if (mark.name === 'insertion' || mark.name === 'deletion') {
      attrs[mark.name === 'insertion' ? 'ins' : 'del'] = dropNulls({
        id: mark.attrs.revisionId,
        author: mark.attrs.author,
        date: mark.attrs.date,
      });
    } else {
      attrs[mark.name] = dropNulls(mark.attrs);
    }
  }
  return attrs;
}

function marksKey(marks: readonly MarkDescriptor[]): string {
  return marks
    .filter((mark) => mark.name !== 'hyperlink' && mark.name !== 'comment')
    .map((mark) => `${mark.name}:${JSON.stringify(mark.attrs)}`)
    .sort()
    .join('|');
}

function withMark(marks: readonly MarkDescriptor[], mark: MarkDescriptor): MarkDescriptor[] {
  return [...marks.filter((candidate) => candidate.name !== mark.name), mark];
}

function textUnit(text: string, marks: readonly MarkDescriptor[], commentId?: number): TextUnit {
  return {
    kind: 'text',
    text,
    attrs: marksToYrsAttrs(marks),
    pmSize: text.length,
    ...(commentId !== undefined ? { commentId } : {}),
    marks: [...marks],
  };
}

function embedUnit(
  embedKind: string,
  payload: Attrs,
  marks: readonly MarkDescriptor[] = [],
  commentId?: number,
  pmSize = 1
): EmbedUnit {
  return {
    kind: 'embed',
    embedKind,
    payload,
    attrs: marksToYrsAttrs(marks),
    pmSize,
    ...(commentId !== undefined ? { commentId } : {}),
    marks: [...marks],
  };
}

function runMarks(
  run: Run,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null
): MarkDescriptor[] {
  const runStyleFormatting = run.formatting?.styleId
    ? styleResolver?.getRunStyleOwnProperties(run.formatting.styleId)
    : undefined;
  return formattingToMarks(
    mergeTextFormatting(mergeTextFormatting(styleFormatting, runStyleFormatting), run.formatting)
  );
}

function imagePayload(image: Image): Attrs {
  const width = image.size?.width ? emuToPixels(image.size.width) : undefined;
  const height = image.size?.height ? emuToPixels(image.size.height) : undefined;
  const wrapType = image.wrap.type;
  const wrapText = image.wrap.wrapText;
  const hAlign = image.position?.horizontal?.alignment;
  let cssFloat: 'left' | 'right' | 'none' | undefined;
  if (wrapType === 'inline' || wrapType === 'topAndBottom') cssFloat = 'none';
  else if (wrapType === 'square' || wrapType === 'tight' || wrapType === 'through') {
    if (wrapText === 'left') cssFloat = 'right';
    else if (wrapText === 'right') cssFloat = 'left';
    else if (hAlign === 'left' || hAlign === 'right') cssFloat = hAlign;
    else cssFloat = 'none';
  } else cssFloat = 'none';

  const displayMode =
    wrapType === 'inline'
      ? 'inline'
      : wrapType === 'topAndBottom'
        ? 'block'
        : isWrapNone(wrapType) || (cssFloat && cssFloat !== 'none')
          ? 'float'
          : 'block';
  const transforms: string[] = [];
  if (image.transform?.rotation) transforms.push(`rotate(${image.transform.rotation}deg)`);
  if (image.transform?.flipH) transforms.push('scaleX(-1)');
  if (image.transform?.flipV) transforms.push('scaleY(-1)');

  let borderWidth: number | undefined;
  let borderColor: string | undefined;
  let borderStyle: string | undefined;
  if (image.outline?.width) {
    borderWidth = Math.round((image.outline.width / 914400) * 96 * 100) / 100;
    if (image.outline.color?.rgb) borderColor = `#${image.outline.color.rgb}`;
    const styles: Record<string, string> = {
      solid: 'solid',
      dot: 'dotted',
      dash: 'dashed',
      lgDash: 'dashed',
      dashDot: 'dashed',
      lgDashDot: 'dashed',
      lgDashDotDot: 'dashed',
      sysDot: 'dotted',
      sysDash: 'dashed',
      sysDashDot: 'dashed',
      sysDashDotDot: 'dashed',
    };
    borderStyle = image.outline.style ? styles[image.outline.style] || 'solid' : 'solid';
  }

  return dropNulls({
    src: image.src || '',
    alt: image.alt ?? null,
    title: image.title ?? null,
    width: width ?? null,
    height: height ?? null,
    rId: image.rId ?? null,
    wrapType,
    displayMode,
    cssFloat: cssFloat ?? null,
    transform: transforms.length > 0 ? transforms.join(' ') : null,
    distTop: image.wrap.distT != null ? emuToPixels(image.wrap.distT) : null,
    distBottom: image.wrap.distB != null ? emuToPixels(image.wrap.distB) : null,
    distLeft: image.wrap.distL != null ? emuToPixels(image.wrap.distL) : null,
    distRight: image.wrap.distR != null ? emuToPixels(image.wrap.distR) : null,
    position: image.position
      ? {
          horizontal: image.position.horizontal
            ? {
                relativeTo: image.position.horizontal.relativeTo,
                posOffset: image.position.horizontal.posOffset,
                align: image.position.horizontal.alignment,
              }
            : undefined,
          vertical: image.position.vertical
            ? {
                relativeTo: image.position.vertical.relativeTo,
                posOffset: image.position.vertical.posOffset,
                align: image.position.vertical.alignment,
              }
            : undefined,
        }
      : null,
    borderWidth: borderWidth ?? null,
    borderColor: borderColor ?? null,
    borderStyle: borderStyle ?? null,
    wrapText: wrapText ?? null,
    hlinkHref: image.hlinkHref ?? null,
    cropTop: image.crop?.top ?? null,
    cropRight: image.crop?.right ?? null,
    cropBottom: image.crop?.bottom ?? null,
    cropLeft: image.crop?.left ?? null,
    opacity: image.opacity ?? null,
    effectExtentTop: image.padding?.top ? emuToPixels(image.padding.top) : null,
    effectExtentBottom: image.padding?.bottom ? emuToPixels(image.padding.bottom) : null,
    effectExtentLeft: image.padding?.left ? emuToPixels(image.padding.left) : null,
    effectExtentRight: image.padding?.right ? emuToPixels(image.padding.right) : null,
    layoutInCell: image.layoutInCell ?? null,
    allowOverlap: image.allowOverlap ?? null,
  }) as Attrs;
}

function shapePayload(shape: Shape): Attrs {
  const width = shape.size?.width ? emuToPixels(shape.size.width) : 100;
  const height = shape.size?.height ? emuToPixels(shape.size.height) : 80;
  let fillColor: string | undefined;
  let fillType = 'solid';
  let gradientType: string | undefined;
  let gradientAngle: number | undefined;
  let gradientStops: string | undefined;
  if (shape.fill) {
    fillType = shape.fill.type;
    fillColor = resolveColorValueToHex(shape.fill.color);
    if (shape.fill.type === 'gradient' && shape.fill.gradient) {
      gradientType = shape.fill.gradient.type;
      gradientAngle = shape.fill.gradient.angle;
      gradientStops = JSON.stringify(
        shape.fill.gradient.stops.map((stop) => ({
          position: stop.position,
          color: resolveColorValueToHex(stop.color) ?? '#000000',
        }))
      );
    }
  }
  const transforms: string[] = [];
  if (shape.transform?.rotation) transforms.push(`rotate(${shape.transform.rotation}deg)`);
  if (shape.transform?.flipH) transforms.push('scaleX(-1)');
  if (shape.transform?.flipV) transforms.push('scaleY(-1)');
  const outlineWidth = shape.outline?.width
    ? Math.round((shape.outline.width / 914400) * 96 * 100) / 100
    : 1;
  return dropNulls({
    layoutBlockJson: JSON.stringify(documentShapeToLayoutBlock(shape)),
    shapeType: shape.shapeType || 'rect',
    geometryPath: shape.geometryPath ?? null,
    shapeId: shape.id ?? null,
    children: shape.children?.length ? JSON.stringify(shape.children) : null,
    width,
    height,
    fillPaint: shape.fillPaint ?? null,
    fillColor: fillColor ?? null,
    fillType,
    gradientType: gradientType ?? null,
    gradientAngle: gradientAngle ?? null,
    gradientStops: gradientStops ?? null,
    outlineWidth,
    outlineColor: resolveColorValueToHex(shape.outline?.color) ?? '#000000',
    outlineStyle: shape.outline?.style || 'solid',
    transform: transforms.length > 0 ? transforms.join(' ') : null,
    rotation: shape.transform?.rotation ?? null,
    flipH: shape.transform?.flipH ?? false,
    flipV: shape.transform?.flipV ?? false,
    displayMode: 'inline',
    wrapType: 'inline',
    behindDoc: false,
    decorative: false,
    hidden: false,
  }) as Attrs;
}

function chartPayload(chart: Chart): Attrs {
  return dropNulls({
    chartJson: JSON.stringify(chart),
    chartType: chart.chartType ?? null,
    title: chart.title ?? null,
    width: chart.size?.width ? emuToPixels(chart.size.width) : 320,
    height: chart.size?.height ? emuToPixels(chart.size.height) : 220,
    rId: chart.rId ?? null,
    path: chart.path ?? null,
  }) as Attrs;
}

function fieldPayload(
  field: SimpleField | ComplexField,
  styleFormatting?: TextFormatting
): {
  payload: Attrs;
  marks: MarkDescriptor[];
} {
  let displayText = '';
  let fieldFormatting: TextFormatting | undefined;
  const runs = field.type === 'simpleField' ? field.content : field.fieldResult;
  for (const child of runs) {
    if (child.type !== 'run') continue;
    for (const content of child.content) {
      if (content.type === 'text') displayText += content.text;
    }
    if (!fieldFormatting && child.formatting) fieldFormatting = child.formatting;
  }
  const formatting =
    fieldFormatting ?? (field.type === 'complexField' ? field.formatting : undefined);
  const displayMode = field.fieldTree?.displayMode ?? 'result';
  return {
    payload: {
      fieldType: field.fieldType,
      instruction: field.instruction,
      displayText,
      fieldKind: field.type === 'simpleField' ? 'simple' : 'complex',
      fldLock: field.fldLock ?? false,
      dirty: field.dirty ?? false,
      displayMode,
      hasCachedResult: displayText.length > 0,
      fieldData: JSON.stringify(field),
      modelKind: 'field',
    },
    marks: formattingToMarks(mergeTextFormatting(styleFormatting, formatting)),
  };
}

function mathPayload(math: MathEquation): Attrs {
  return {
    display: math.display,
    ommlXml: math.ommlXml,
    plainText: math.plainText || '',
  };
}

function hyperlinkMark(hyperlink: Hyperlink): MarkDescriptor {
  return {
    name: 'hyperlink',
    attrs: {
      href: hyperlink.href || (hyperlink.anchor ? `#${hyperlink.anchor}` : ''),
      tooltip: hyperlink.tooltip ?? null,
      rId: hyperlink.rId ?? null,
    },
  };
}

function noteRefUnit(
  id: number,
  noteType: 'footnote' | 'endnote',
  marks: readonly MarkDescriptor[],
  commentId?: number
): EmbedUnit {
  const noteMark: MarkDescriptor = {
    name: 'footnoteRef',
    attrs: { id: String(id), noteType },
  };
  const allMarks = [...marks, noteMark];
  return embedUnit(
    'noteRef',
    noteType === 'endnote' ? { endnoteRefId: id } : { footnoteRefId: id },
    allMarks,
    commentId,
    String(id).length
  );
}

function runContentToUnits(
  content: RunContent,
  marks: readonly MarkDescriptor[],
  commentId?: number
): InlineUnit[] {
  switch (content.type) {
    case 'text':
      return content.text ? [textUnit(content.text, marks, commentId)] : [];
    case 'tab':
      return [textUnit('\t', marks, commentId)];
    case 'break':
      return content.breakType === undefined || content.breakType === 'textWrapping'
        ? [embedUnit('break', {}, marks, commentId)]
        : [];
    case 'softHyphen':
      return [textUnit('\u00ad', marks, commentId)];
    case 'noBreakHyphen':
      return [textUnit('\u2011', marks, commentId)];
    case 'symbol': {
      const codePoint = Number.parseInt(content.char, 16);
      if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff) return [];
      const symbolMark: MarkDescriptor = {
        name: 'fontFamily',
        attrs: {
          ascii: content.font || null,
          hAnsi: content.font || null,
          eastAsia: content.font || null,
          cs: content.font || null,
          asciiTheme: null,
          hAnsiTheme: null,
          eastAsiaTheme: null,
          csTheme: null,
        },
      };
      return [textUnit(String.fromCodePoint(codePoint), withMark(marks, symbolMark), commentId)];
    }
    case 'commentReference':
      return [
        embedUnit('field', {
          fieldType: 'COMMENT',
          instruction: '',
          displayText: '',
          fieldKind: 'simple',
          fldLock: false,
          dirty: false,
          displayMode: 'result',
          hasCachedResult: false,
          modelKind: 'commentReference',
          ...(content.id !== undefined ? { commentId: content.id } : {}),
        }),
      ];
    case 'drawing':
      return [embedUnit('image', imagePayload(content.image))];
    case 'shape':
      return content.shape.textBody?.content.length
        ? []
        : [embedUnit('shape', shapePayload(content.shape))];
    case 'chart':
      return [embedUnit('chart', chartPayload(content.chart))];
    case 'footnoteRef':
      return [noteRefUnit(content.id, 'footnote', marks, commentId)];
    case 'endnoteRef':
      return [noteRefUnit(content.id, 'endnote', marks, commentId)];
    default:
      return [];
  }
}

function runToUnits(
  run: Run,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null,
  commentId?: number,
  extraMarks: readonly MarkDescriptor[] = []
): InlineUnit[] {
  const marks = [...runMarks(run, styleFormatting, styleResolver), ...extraMarks];
  return run.content.flatMap((content) => runContentToUnits(content, marks, commentId));
}

function hyperlinkToUnits(
  hyperlink: Hyperlink,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null,
  extraMarks: readonly MarkDescriptor[] = []
): InlineUnit[] {
  const units: InlineUnit[] = [];
  const link = hyperlinkMark(hyperlink);
  for (const child of hyperlink.structuredChildren ?? hyperlink.children) {
    if (child.type === 'run') {
      const marks = [...runMarks(child, styleFormatting, styleResolver), ...extraMarks, link];
      for (const content of child.content) units.push(...runContentToUnits(content, marks));
    } else if (child.type === 'simpleField' || child.type === 'complexField') {
      const field = fieldPayload(child, styleFormatting);
      units.push(embedUnit('field', field.payload, [...field.marks, ...extraMarks, link]));
    } else if (child.type === 'mathEquation') {
      units.push(embedUnit('math', mathPayload(child), [...extraMarks, link]));
    }
  }
  return units;
}

function trackedMark(
  info: TrackedChangeInfo,
  kind: 'insertion' | 'deletion',
  isMovePair: boolean
): MarkDescriptor {
  return {
    name: kind,
    attrs: {
      revisionId: info.id,
      author: info.author,
      date: info.date ?? null,
      isMovePair,
    },
  };
}

function trackedToUnits(
  content: Extract<ParagraphContent, { type: 'insertion' | 'deletion' | 'moveFrom' | 'moveTo' }>,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null,
  commentId?: number
): InlineUnit[] {
  const kind = content.type === 'insertion' || content.type === 'moveTo' ? 'insertion' : 'deletion';
  const mark = trackedMark(
    content.info,
    kind,
    content.type === 'moveFrom' || content.type === 'moveTo'
  );
  const units: InlineUnit[] = [];
  for (const child of content.content) {
    if (child.type === 'run') {
      units.push(...runToUnits(child, styleFormatting, styleResolver, commentId, [mark]));
    } else {
      const linked = hyperlinkToUnits(child, styleFormatting, styleResolver, [mark]);
      for (const unit of linked) {
        if (commentId !== undefined) unit.commentId = commentId;
      }
      units.push(...linked);
    }
  }
  return units;
}

function sdtPayload(
  sdt: InlineSdt,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null
): Attrs {
  const content: Array<
    | { kind: 'text'; text: string; attrs: YrsAttrs }
    | { kind: 'tab'; attrs: YrsAttrs }
    | { kind: string; payload: Attrs; attrs: YrsAttrs }
  > = [];
  const append = (unit: InlineUnit): void => {
    if (unit.kind === 'text') {
      const kind = unit.text === '\t' ? 'tab' : 'text';
      if (kind === 'tab') {
        content.push({ kind: 'tab', attrs: unit.attrs });
        return;
      }
      const previous = content[content.length - 1];
      if (
        previous &&
        'text' in previous &&
        stableStringify(previous.attrs) === stableStringify(unit.attrs)
      ) {
        previous.text += unit.text;
      } else {
        content.push({ kind: 'text', text: unit.text, attrs: unit.attrs });
      }
      return;
    }
    content.push({ kind: unit.embedKind, payload: unit.payload, attrs: unit.attrs });
  };

  for (const child of sdt.content) {
    if (child.type === 'run') {
      runToUnits(child, styleFormatting, styleResolver).forEach(append);
    } else if (child.type === 'hyperlink') {
      hyperlinkToUnits(child, styleFormatting, styleResolver).forEach(append);
    } else if (child.type === 'simpleField' || child.type === 'complexField') {
      const field = fieldPayload(child, styleFormatting);
      append(embedUnit('field', field.payload, field.marks));
    } else if (child.type === 'inlineSdt') {
      append(embedUnit('sdt', sdtPayload(child, styleFormatting, styleResolver)));
    } else if (child.type === 'mathEquation') {
      append(embedUnit('math', mathPayload(child)));
    }
  }
  return dropNulls({
    ...sdtPropsToAttrs(sdt.properties),
    propertiesJson: JSON.stringify(sdt.properties),
    content,
  }) as Attrs;
}

function paragraphStyleFormatting(
  paragraph: Paragraph,
  styleResolver: StyleResolver | null,
  extraRunFormatting?: TextFormatting
): TextFormatting | undefined {
  const styleFormatting = styleResolver
    ? styleResolver.resolveParagraphStyle(paragraph.formatting?.styleId).runFormatting
    : undefined;
  return mergeTextFormatting(styleFormatting, extraRunFormatting);
}

function runBoundary(
  run: Run,
  styleFormatting: TextFormatting | undefined,
  styleResolver: StyleResolver | null
): Attrs | null {
  const units = runToUnits(run, styleFormatting, styleResolver);
  if (units.some((unit) => unit.kind !== 'text' && unit.embedKind !== 'noteRef')) return null;
  const keys = units.map((unit) => marksKey(unit.marks));
  if (keys.some((key) => key !== keys[0])) return null;
  const text = units
    .map((unit) => {
      if (unit.kind === 'text') return unit.text;
      const id = unit.payload.footnoteRefId ?? unit.payload.endnoteRefId;
      return String(id ?? '');
    })
    .join('');
  const key = keys[0];
  return {
    text,
    ...(key !== undefined ? { marksKey: key } : {}),
    ...(run.formatting ? { formatting: run.formatting } : {}),
    ...(run.propertyChanges ? { propertyChanges: run.propertyChanges } : {}),
  };
}

function paragraphAttrs(
  paragraph: Paragraph,
  styleResolver: StyleResolver | null,
  units: readonly InlineUnit[],
  runBoundaries: Attrs[] | undefined
): Attrs {
  const formatting = paragraph.formatting;
  const styleId = formatting?.styleId;
  const attrs: Attrs = {
    paraId: paragraph.paraId ?? null,
    textId: paragraph.textId ?? null,
    styleId: styleId ?? null,
    numPr: formatting?.numPr ?? null,
    numPrFromStyle: formatting?.numPrFromStyle ?? null,
    listNumFmt: paragraph.listRendering?.numFmt ?? null,
    listIsBullet: paragraph.listRendering?.isBullet ?? null,
    listMarker: paragraph.listRendering?.marker ?? null,
    listMarkerHidden: paragraph.listRendering?.markerHidden || null,
    listMarkerFontFamily: paragraph.listRendering?.markerFontFamily || null,
    listMarkerFontSize: paragraph.listRendering?.markerFontSize || null,
    listMarkerSuffix: paragraph.listRendering?.markerSuffix || null,
    listLevelNumFmts: paragraph.listRendering?.levelNumFmts || null,
    listAbstractNumId: paragraph.listRendering?.abstractNumId ?? null,
    listStartOverride: paragraph.listRendering?.startOverride ?? null,
    _originalFormatting: formatting ?? null,
  };

  if (styleResolver) {
    const resolved = styleResolver.resolveParagraphStyle(styleId);
    const stylePpr = resolved.paragraphFormatting;
    attrs.alignment = formatting?.alignment ?? stylePpr?.alignment ?? null;
    attrs.spaceBefore = formatting?.spaceBefore ?? stylePpr?.spaceBefore ?? null;
    attrs.spaceAfter = formatting?.spaceAfter ?? stylePpr?.spaceAfter ?? null;
    attrs.lineSpacing = formatting?.lineSpacing ?? stylePpr?.lineSpacing ?? null;
    attrs.lineSpacingRule = formatting?.lineSpacingRule ?? stylePpr?.lineSpacingRule ?? null;
    attrs.spacingExplicit = formatting?.spacingExplicit || null;
    attrs.indentLeft = formatting?.indentLeft ?? stylePpr?.indentLeft ?? null;
    attrs.indentRight = formatting?.indentRight ?? stylePpr?.indentRight ?? null;
    const numberingRemoved =
      formatting?.numPr?.numId === 0 && stylePpr?.numPr && stylePpr.numPr.numId !== 0;
    const styleFirstLine = numberingRemoved ? undefined : stylePpr;
    attrs.indentFirstLine = formatting?.indentFirstLine ?? styleFirstLine?.indentFirstLine ?? null;
    attrs.hangingIndent = formatting?.hangingIndent ?? styleFirstLine?.hangingIndent ?? false;
    attrs.borders = formatting?.borders ?? stylePpr?.borders ?? null;
    attrs.shading = formatting?.shading ?? stylePpr?.shading ?? null;
    attrs.tabs = formatting?.tabs ?? stylePpr?.tabs ?? null;
    attrs.pageBreakBefore = formatting?.pageBreakBefore ?? stylePpr?.pageBreakBefore ?? null;
    attrs.keepNext = formatting?.keepNext ?? stylePpr?.keepNext ?? null;
    attrs.keepLines = formatting?.keepLines ?? stylePpr?.keepLines ?? null;
    attrs.contextualSpacing = formatting?.contextualSpacing ?? stylePpr?.contextualSpacing ?? null;
    attrs.outlineLevel = formatting?.outlineLevel ?? stylePpr?.outlineLevel ?? null;
    attrs.bidi = formatting?.bidi ?? stylePpr?.bidi ?? null;

    const defaultChar = styleResolver.getDefaultCharacterStyle()?.rPr;
    const styleRpr = defaultChar
      ? mergeTextFormatting(resolved.runFormatting, defaultChar)
      : resolved.runFormatting;
    attrs.defaultTextFormatting =
      mergeTextFormatting(
        styleRpr,
        resolveTextFormatting(formatting?.runProperties, styleResolver)
      ) ?? null;
    if (!formatting?.numPr && stylePpr?.numPr && stylePpr.numPr.numId !== 0) {
      attrs.numPr = stylePpr.numPr;
      attrs.numPrFromStyle = stylePpr.numPr;
    }
  } else {
    attrs.alignment = formatting?.alignment ?? null;
    attrs.spaceBefore = formatting?.spaceBefore ?? null;
    attrs.spaceAfter = formatting?.spaceAfter ?? null;
    attrs.lineSpacing = formatting?.lineSpacing ?? null;
    attrs.lineSpacingRule = formatting?.lineSpacingRule ?? null;
    attrs.spacingExplicit = formatting?.spacingExplicit || null;
    attrs.indentLeft = formatting?.indentLeft ?? null;
    attrs.indentRight = formatting?.indentRight ?? null;
    attrs.indentFirstLine = formatting?.indentFirstLine ?? null;
    attrs.hangingIndent = formatting?.hangingIndent ?? false;
    attrs.borders = formatting?.borders ?? null;
    attrs.shading = formatting?.shading ?? null;
    attrs.tabs = formatting?.tabs ?? null;
    attrs.pageBreakBefore = formatting?.pageBreakBefore ?? null;
    attrs.keepNext = formatting?.keepNext ?? null;
    attrs.keepLines = formatting?.keepLines ?? null;
    attrs.outlineLevel = formatting?.outlineLevel ?? null;
    attrs.bidi = formatting?.bidi ?? null;
    attrs.defaultTextFormatting = formatting?.runProperties ?? null;
  }

  if (paragraph.sectionProperties) {
    attrs._sectionProperties = paragraph.sectionProperties;
    const start = paragraph.sectionProperties.sectionStart;
    if (
      start === 'nextPage' ||
      start === 'continuous' ||
      start === 'oddPage' ||
      start === 'evenPage'
    ) {
      attrs.sectionBreakType = start;
    }
  }
  if (paragraph.renderedPageBreakBefore) attrs.renderedPageBreakBefore = true;
  if (paragraphStartsWithPageBreak(paragraph)) attrs.pageBreakBefore = true;
  if (paragraph.pPrIns) {
    attrs.pPrIns = {
      revisionId: paragraph.pPrIns.id,
      author: paragraph.pPrIns.author,
      date: paragraph.pPrIns.date ?? null,
    };
  }
  if (paragraph.pPrDel) {
    attrs.pPrDel = {
      revisionId: paragraph.pPrDel.id,
      author: paragraph.pPrDel.author,
      date: paragraph.pPrDel.date ?? null,
    };
  }
  if (paragraph.propertyChanges?.length) attrs.pPrChange = paragraph.propertyChanges;

  const bookmarks: Attrs[] = [];
  let contentIndex = 0;
  let pmOffset = 0;
  for (const content of paragraph.content) {
    if (content.type === 'bookmarkStart') {
      bookmarks.push({
        id: content.id,
        name: content.name,
        kind: 'start',
        offset: pmOffset,
        ...(content.colFirst !== undefined ? { colFirst: content.colFirst } : {}),
        ...(content.colLast !== undefined ? { colLast: content.colLast } : {}),
      });
    } else if (content.type === 'bookmarkEnd') {
      bookmarks.push({ id: content.id, kind: 'end', offset: pmOffset });
    } else {
      const count = unitsForParagraphContent(content);
      for (let index = 0; index < count; index += 1) {
        pmOffset += units[contentIndex]?.pmSize ?? 0;
        contentIndex += 1;
      }
    }
  }
  if (bookmarks.length > 0) attrs.bookmarks = bookmarks;
  if (runBoundaries?.length) attrs._originalRunBoundaries = runBoundaries;
  return attrs;
}

/* Set by paragraphUnits while paragraphAttrs calculates bookmark offsets. */
const paragraphContentUnitCounts = new WeakMap<object, number>();

function unitsForParagraphContent(content: ParagraphContent): number {
  return paragraphContentUnitCounts.get(content as object) ?? 0;
}

function paragraphUnits(
  paragraph: Paragraph,
  styleResolver: StyleResolver | null,
  extraRunFormatting?: TextFormatting
): { units: InlineUnit[]; ppr: Attrs } {
  const units: InlineUnit[] = [];
  const activeComments = new Set<number>();
  let boundaries: Attrs[] | undefined = [];
  const styleFormatting = paragraphStyleFormatting(paragraph, styleResolver, extraRunFormatting);

  for (const content of paragraph.content) {
    const start = units.length;
    const commentId = activeComments.values().next().value as number | undefined;
    if (content.type === 'commentRangeStart') activeComments.add(content.id);
    else if (content.type === 'commentRangeEnd') activeComments.delete(content.id);
    else if (content.type === 'run') {
      const boundary = runBoundary(content, styleFormatting, styleResolver);
      if (boundary && boundaries) boundaries.push(boundary);
      else boundaries = undefined;
      units.push(...runToUnits(content, styleFormatting, styleResolver, commentId));
    } else if (content.type === 'hyperlink') {
      boundaries = undefined;
      units.push(...hyperlinkToUnits(content, styleFormatting, styleResolver));
    } else if (content.type === 'simpleField' || content.type === 'complexField') {
      boundaries = undefined;
      const field = fieldPayload(content, styleFormatting);
      units.push(embedUnit('field', field.payload, field.marks));
    } else if (content.type === 'inlineSdt') {
      boundaries = undefined;
      units.push(
        embedUnit('sdt', sdtPayload(content, styleFormatting, styleResolver), [], undefined, 2)
      );
    } else if (
      content.type === 'insertion' ||
      content.type === 'deletion' ||
      content.type === 'moveFrom' ||
      content.type === 'moveTo'
    ) {
      boundaries = undefined;
      units.push(...trackedToUnits(content, styleFormatting, styleResolver, commentId));
    } else if (content.type === 'mathEquation') {
      boundaries = undefined;
      units.push(embedUnit('math', mathPayload(content)));
    } else if (content.type !== 'bookmarkStart' && content.type !== 'bookmarkEnd') {
      boundaries = undefined;
    }
    paragraphContentUnitCounts.set(content as object, units.length - start);
  }
  const attrs = paragraphAttrs(paragraph, styleResolver, units, boundaries);
  return { units, ppr: paraAttrsToPpr(attrs) };
}

type ParagraphToken = 'pageBreak' | 'visible';

function runTokens(run: Run, tokens: ParagraphToken[]): void {
  for (const content of run.content) {
    if (content.type === 'break' && content.breakType === 'page') tokens.push('pageBreak');
    else if (content.type !== 'text' || content.text.length > 0) tokens.push('visible');
  }
}

function inlineTokens(content: readonly ParagraphContent[], tokens: ParagraphToken[]): void {
  for (const item of content) {
    if (item.type === 'run') runTokens(item, tokens);
    else if (item.type === 'hyperlink') {
      for (const child of item.children) if (child.type === 'run') runTokens(child, tokens);
    } else if (item.type === 'simpleField') {
      for (const child of item.content) if (child.type === 'run') runTokens(child, tokens);
    } else if (item.type === 'complexField') {
      for (const child of [...item.fieldCode, ...item.fieldResult]) runTokens(child, tokens);
    } else if (item.type === 'inlineSdt') inlineTokens(item.content as ParagraphContent[], tokens);
    else if (
      item.type === 'insertion' ||
      item.type === 'deletion' ||
      item.type === 'moveFrom' ||
      item.type === 'moveTo'
    ) {
      for (const child of item.content) if (child.type === 'run') runTokens(child, tokens);
    } else if (item.type === 'mathEquation') tokens.push('visible');
  }
}

function paragraphStartsWithPageBreak(paragraph: Paragraph): boolean {
  const tokens: ParagraphToken[] = [];
  inlineTokens(paragraph.content, tokens);
  return tokens[0] === 'pageBreak';
}

function paragraphHasNonLeadingPageBreak(paragraph: Paragraph): boolean {
  const tokens: ParagraphToken[] = [];
  inlineTokens(paragraph.content, tokens);
  let leading = false;
  let visible = false;
  for (const token of tokens) {
    if (token === 'pageBreak') {
      if (visible || leading) return true;
      leading = true;
    } else visible = true;
  }
  return false;
}

type RowSpanInfo = { rowSpan: number; skip: boolean };

function calculateRowSpans(table: Table): Map<string, RowSpanInfo> {
  const result = new Map<string, RowSpanInfo>();
  const active = new Map<number, number>();
  table.rows.forEach((row, rowIndex) => {
    let column = 0;
    const cells = row.cells.map((cell) => {
      const current = column;
      column += cell.formatting?.gridSpan ?? 1;
      return { column: current, vMerge: cell.formatting?.vMerge, key: `${rowIndex}-${current}` };
    });
    const empty =
      cells.length > 0 &&
      cells.every((cell) => cell.vMerge === 'continue' && active.has(cell.column));
    if (empty) {
      for (const cell of cells) {
        active.delete(cell.column);
        result.set(cell.key, { rowSpan: 1, skip: false });
      }
      return;
    }
    for (const cell of cells) {
      if (cell.vMerge === 'restart') {
        active.set(cell.column, rowIndex);
        result.set(cell.key, { rowSpan: 1, skip: false });
      } else if (cell.vMerge === 'continue') {
        const start = active.get(cell.column);
        if (start === undefined) result.set(cell.key, { rowSpan: 1, skip: false });
        else {
          const owner = result.get(`${start}-${cell.column}`);
          if (owner) owner.rowSpan += 1;
          result.set(cell.key, { rowSpan: 1, skip: true });
        }
      } else {
        active.delete(cell.column);
        result.set(cell.key, { rowSpan: 1, skip: false });
      }
    }
  });
  return result;
}

function revisionAttrs(info: TrackedChangeInfo): Attrs {
  return { revisionId: info.id, author: info.author, date: info.date ?? null };
}

function cellBorders(
  formatting: TableCellFormatting | undefined,
  tableBorders: TableBorders | undefined,
  firstRow: boolean,
  lastRow: boolean,
  firstColumn: boolean,
  lastColumn: boolean
): TableBorders | undefined {
  const inherited = tableBorders
    ? {
        top: firstRow ? tableBorders.top : tableBorders.insideH,
        bottom: lastRow ? tableBorders.bottom : tableBorders.insideH,
        left: firstColumn ? tableBorders.left : tableBorders.insideV,
        right: lastColumn ? tableBorders.right : tableBorders.insideV,
      }
    : undefined;
  return inherited || formatting?.borders
    ? { ...(inherited ?? {}), ...(formatting?.borders ?? {}) }
    : undefined;
}

function projectCell(
  cell: TableCell,
  options: {
    isHeader: boolean;
    rowspan: number;
    gridWidth?: number;
    firstRow: boolean;
    lastRow: boolean;
    firstColumn: boolean;
    lastColumn: boolean;
    tableBorders?: TableBorders;
    defaultMargins?: { top?: number; bottom?: number; left?: number; right?: number };
    theme: Theme | null;
    tableBidi: boolean;
  }
): ProjectedCell {
  const formatting = cell.formatting;
  const backgroundColor = resolveColorToHex(formatting?.shading?.fill, options.theme);
  const width = formatting?.width?.value ?? options.gridWidth;
  const widthType = formatting?.width?.type ?? (options.gridWidth !== undefined ? 'pct' : null);
  const attrs: Attrs = {
    colspan: formatting?.gridSpan ?? 1,
    rowspan: options.rowspan,
    width: width ?? null,
    widthType,
    verticalAlign: formatting?.verticalAlign ?? null,
    backgroundColor: backgroundColor ?? null,
    borders:
      cellBorders(
        formatting,
        options.tableBorders,
        options.firstRow,
        options.lastRow,
        options.firstColumn,
        options.lastColumn
      ) ?? null,
    margins: formatting?.margins
      ? {
          top: formatting.margins.top?.value,
          bottom: formatting.margins.bottom?.value,
          left:
            formatting.margins.left?.value ??
            (options.tableBidi ? formatting.margins.end?.value : formatting.margins.start?.value),
          right:
            formatting.margins.right?.value ??
            (options.tableBidi ? formatting.margins.start?.value : formatting.margins.end?.value),
        }
      : (options.defaultMargins ?? null),
    textDirection: formatting?.textDirection ?? null,
    noWrap: formatting?.noWrap ?? false,
    _originalFormatting: formatting ?? null,
    _originalResolvedFill: backgroundColor ?? null,
  };
  if (cell.structuralChange) {
    const change = cell.structuralChange;
    const info = revisionAttrs(change.info);
    if (change.type === 'tableCellInsertion') attrs.cellMarker = { kind: 'ins', info };
    else if (change.type === 'tableCellDeletion') attrs.cellMarker = { kind: 'del', info };
    else if (change.type === 'tableCellMerge') {
      attrs.cellMarker = {
        kind: 'merge',
        info,
        vMerge: change.vMerge ?? 'cont',
        ...(change.vMergeOrig ? { vMergeOrig: change.vMergeOrig } : {}),
      };
    }
  }
  if (cell.propertyChanges?.length) attrs.tcPrChange = cell.propertyChanges;
  return {
    attrs: tableCellAttrsToTcPr(attrs, options.isHeader),
    content:
      cell.content.length > 0
        ? cell.content
        : [{ type: 'paragraph', content: [] } satisfies Paragraph],
  };
}

function projectRow(
  row: TableRow,
  table: Table,
  rowIndex: number,
  rowSpans: Map<string, RowSpanInfo>,
  tableBorders: TableBorders | undefined,
  defaultMargins: { top?: number; bottom?: number; left?: number; right?: number } | undefined,
  theme: Theme | null
): ProjectedRow {
  const attrs: Attrs = {
    height: row.formatting?.height?.value ?? null,
    heightRule: row.formatting?.heightRule ?? null,
    isHeader: !!row.formatting?.header,
    _originalFormatting: row.formatting ?? null,
  };
  if (row.structuralChange) {
    if (row.structuralChange.type === 'tableRowInsertion') {
      attrs.trIns = revisionAttrs(row.structuralChange.info);
    } else if (row.structuralChange.type === 'tableRowDeletion') {
      attrs.trDel = revisionAttrs(row.structuralChange.info);
    }
  }
  if (row.propertyChanges?.length) attrs.trPrChange = row.propertyChanges;

  const columnWidths = table.columnWidths;
  const totalWidth = columnWidths?.reduce((sum, width) => sum + width, 0) ?? 0;
  const totalColumns =
    columnWidths?.length ??
    Math.max(
      0,
      ...table.rows.map((candidate) =>
        candidate.cells.reduce((sum, cell) => sum + (cell.formatting?.gridSpan ?? 1), 0)
      )
    );
  let column = 0;
  const cells: ProjectedCell[] = [];
  row.cells.forEach((cell) => {
    const colspan = cell.formatting?.gridSpan ?? 1;
    const startColumn = column;
    const rowSpan = rowSpans.get(`${rowIndex}-${startColumn}`);
    let gridWidth: number | undefined;
    if (columnWidths && totalWidth > 0) {
      let cellWidth = 0;
      for (
        let index = 0;
        index < colspan && startColumn + index < columnWidths.length;
        index += 1
      ) {
        cellWidth += columnWidths[startColumn + index];
      }
      gridWidth = Math.round((cellWidth / totalWidth) * 100);
    }
    column += colspan;
    if (rowSpan?.skip) return;
    cells.push(
      projectCell(cell, {
        isHeader: rowIndex === 0 && !!table.formatting?.look?.firstRow,
        rowspan: rowSpan?.rowSpan ?? 1,
        gridWidth,
        firstRow: rowIndex === 0,
        lastRow: rowIndex === table.rows.length - 1,
        firstColumn: startColumn === 0,
        lastColumn: column === totalColumns,
        tableBorders,
        defaultMargins,
        theme,
        tableBidi: Boolean(table.formatting?.bidi),
      })
    );
  });
  if (cells.length === 0) {
    cells.push(
      projectCell(
        {
          type: 'tableCell',
          formatting: totalColumns > 1 ? { gridSpan: totalColumns } : undefined,
          content: [{ type: 'paragraph', content: [] }],
        },
        {
          isHeader: rowIndex === 0 && !!table.formatting?.look?.firstRow,
          rowspan: 1,
          gridWidth: totalWidth > 0 ? 100 : undefined,
          firstRow: rowIndex === 0,
          lastRow: rowIndex === table.rows.length - 1,
          firstColumn: true,
          lastColumn: true,
          tableBorders,
          defaultMargins,
          theme,
          tableBidi: Boolean(table.formatting?.bidi),
        }
      )
    );
  }
  return { attrs: tableRowAttrsToTrPr(attrs), cells };
}

function projectTable(
  table: Table,
  styleResolver: StyleResolver | null,
  theme: Theme | null
): ProjectedTable {
  const defaultStyle = styleResolver?.getDefaultTableStyle();
  const styleId = table.formatting?.styleId;
  const effectiveStyleId = styleId ?? defaultStyle?.styleId;
  const tableStyle = effectiveStyleId ? styleResolver?.getStyle(effectiveStyleId) : undefined;
  const borders =
    table.formatting?.borders ?? tableStyle?.tblPr?.borders ?? defaultStyle?.tblPr?.borders;
  const margins =
    table.formatting?.cellMargins ??
    tableStyle?.tblPr?.cellMargins ??
    defaultStyle?.tblPr?.cellMargins;
  const logicalLeft = table.formatting?.bidi ? margins?.end : margins?.start;
  const logicalRight = table.formatting?.bidi ? margins?.start : margins?.end;
  const defaultMargins = margins
    ? {
        top: margins.top?.value,
        bottom: margins.bottom?.value,
        left: margins.left?.value ?? logicalLeft?.value,
        right: margins.right?.value ?? logicalRight?.value,
      }
    : undefined;
  const basedOnStyleIds: string[] = [];
  const visited = new Set<string>();
  let inherited = tableStyle;
  while (inherited?.basedOn && basedOnStyleIds.length < 32) {
    if (visited.has(inherited.basedOn)) break;
    visited.add(inherited.basedOn);
    basedOnStyleIds.unshift(inherited.basedOn);
    inherited = styleResolver?.getStyle(inherited.basedOn);
  }
  const originalFormatting = {
    ...(table.formatting ?? {}),
    styleCascade: {
      ...(styleId ? { selectedStyleId: styleId } : {}),
      ...(defaultStyle?.styleId ? { defaultStyleId: defaultStyle.styleId } : {}),
      ...(basedOnStyleIds.length > 0 ? { basedOnStyleIds } : {}),
    },
  };
  const attrs: Attrs = {
    styleId: styleId ?? null,
    width: table.formatting?.width?.value ?? null,
    widthType: table.formatting?.width?.type ?? null,
    justification: table.formatting?.justification ?? null,
    columnWidths: table.columnWidths ?? null,
    tableLayout: table.formatting?.layout ?? null,
    floating: table.formatting?.floating ?? null,
    cellMargins: defaultMargins ?? null,
    look: table.formatting?.look ?? null,
    bidi: table.formatting?.bidi || null,
    _originalFormatting: originalFormatting,
  };
  if (table.propertyChanges?.length) attrs.tblPrChange = table.propertyChanges;
  const rowSpans = calculateRowSpans(table);
  return {
    attrs,
    rows: table.rows.map((row, rowIndex) =>
      projectRow(row, table, rowIndex, rowSpans, borders, defaultMargins, theme)
    ),
  };
}

function blockSdtAttrs(properties: SdtProperties): Attrs {
  return blockSdtAttrsToPayload(sdtPropsToAttrs(properties));
}

function addCommentCoverage(plan: StoryPlan): void {
  let offset = 0;
  for (const unit of plan.units) {
    const width = unit.kind === 'text' ? unit.text.length : 1;
    if (unit.commentId !== undefined && unit.commentId !== 0) {
      const intervals = plan.commentCoverage.get(unit.commentId);
      const previous = intervals?.[intervals.length - 1];
      if (previous && previous[1] === offset) previous[1] = offset + width;
      else if (intervals) intervals.push([offset, offset + width]);
      else plan.commentCoverage.set(unit.commentId, [[offset, offset + width]]);
    }
    offset += width;
  }
}

function visitStory(
  context: LoweringContext,
  storyId: string,
  sourceBlocks: readonly BlockContent[],
  options: StoryOptions
): void {
  const plan: StoryPlan = {
    storyId,
    units: [],
    commentCoverage: new Map(),
  };
  context.plans.push(plan);
  const blocks =
    sourceBlocks.length > 0 ? [...sourceBlocks] : [{ type: 'paragraph', content: [] } as Paragraph];
  let tableIndex = 0;
  let sdtIndex = 0;
  let paragraphIndex = 0;
  let lastKind: 'paragraph' | 'table' | 'blockSdt' | null = null;

  for (const block of blocks) {
    if (block.type === 'paragraph') {
      const paragraph = paragraphUnits(block, context.styleResolver, options.extraRunFormatting);
      plan.units.push(...paragraph.units);
      plan.units.push(
        embedUnit('pilcrow', {
          ...paragraph.ppr,
          paraId: block.paraId || `${storyId}:p${paragraphIndex}`,
        })
      );
      paragraphIndex += 1;
      if (options.includePageBreaks && paragraphHasNonLeadingPageBreak(block)) {
        plan.units.push(embedUnit('pageBreak', {}));
      }
      lastKind = 'paragraph';
      continue;
    }
    if (block.type === 'table') {
      const currentTable = tableIndex++;
      const table = projectTable(block, context.styleResolver, context.theme);
      const rows = table.rows.map((row, rowIndex) => ({
        trPr: row.attrs,
        cells: row.cells.map((cell, cellIndex) => ({
          tcPr: cell.attrs,
          story: tableCellStoryId(storyId, currentTable, rowIndex, cellIndex),
        })),
      }));
      plan.units.push(
        embedUnit('table', {
          tblPr: tableAttrsToTblPr(table.attrs),
          grid: tableAttrsToGrid(table.attrs),
          rows,
        })
      );
      table.rows.forEach((row, rowIndex) => {
        row.cells.forEach((cell, cellIndex) => {
          visitStory(
            context,
            tableCellStoryId(storyId, currentTable, rowIndex, cellIndex),
            cell.content,
            {
              includePageBreaks: false,
              appendBodyTail: false,
              seedComments: false,
              extraRunFormatting: cell.extraRunFormatting,
            }
          );
        });
      });
      lastKind = 'table';
      continue;
    }
    const currentSdt = sdtIndex++;
    const childStory = blockSdtStoryId(storyId, currentSdt);
    plan.units.push(
      embedUnit('blockSdt', {
        ...blockSdtAttrs(block.properties),
        story: childStory,
      })
    );
    visitStory(context, childStory, block.content, {
      includePageBreaks: options.includePageBreaks,
      appendBodyTail: false,
      seedComments: false,
    });
    lastKind = 'blockSdt';
  }

  if (options.appendBodyTail && (lastKind === 'table' || lastKind === 'blockSdt')) {
    plan.units.push(
      embedUnit('pilcrow', {
        hangingIndent: false,
        paraId: `${storyId}:p${paragraphIndex}`,
      })
    );
  }
  if (options.seedComments) addCommentCoverage(plan);
}

function unitsToRawOps(units: readonly InlineUnit[]): YrsRawOp[] {
  const ops: YrsRawOp[] = [{ op: 'delete', index: 0, len: 1 }];
  let index = 0;
  let text = '';
  let attrs: YrsAttrs = {};
  let attrsKey = stableStringify(attrs);
  const flush = () => {
    if (!text) return;
    ops.push({ op: 'insert', index, text, attrs });
    index += text.length;
    text = '';
  };
  for (const unit of units) {
    if (unit.kind === 'text') {
      const key = stableStringify(unit.attrs);
      if (text && key !== attrsKey) flush();
      attrs = unit.attrs;
      attrsKey = key;
      text += unit.text;
    } else {
      flush();
      ops.push({
        op: 'insertEmbed',
        index,
        kind: unit.embedKind,
        payload: unit.payload,
        attrs: unit.attrs,
      });
      index += 1;
    }
  }
  flush();
  return ops;
}

function seedPlan(session: YrsSession, plan: StoryPlan): void {
  session.applyRawOps(plan.storyId, unitsToRawOps(plan.units));
  const comments: YrsRawOp[] = [];
  for (const [id, ranges] of plan.commentCoverage) {
    comments.push({
      op: 'setComment',
      id: String(id),
      ranges,
      author: '',
      date: '',
      body: null,
    });
  }
  if (comments.length > 0) session.applyRawOps(plan.storyId, comments);
}

/**
 * Seeds every yrs-owned editable story directly from a parsed Document.
 *
 * Stories are `body`, `hf:{rId}`, `fn:{id}`, and recursively generated table
 * cell / block-SDT stories. The target session must not already contain any of
 * those story ids.
 *
 * @public
 */
export function documentToYrs(session: YrsSession, document: Document): void {
  const context: LoweringContext = {
    styleResolver: document.package.styles ? createStyleResolver(document.package.styles) : null,
    theme: document.package.theme ?? null,
    plans: [],
  };
  visitStory(context, 'body', document.package.document.content, {
    includePageBreaks: true,
    appendBodyTail: true,
    seedComments: true,
  });
  for (const [rId, part] of document.package.headers ?? []) {
    visitStory(context, headerFooterStoryId(rId), part.content, {
      includePageBreaks: false,
      appendBodyTail: false,
      seedComments: true,
    });
  }
  for (const [rId, part] of document.package.footers ?? []) {
    if (context.plans.some((plan) => plan.storyId === headerFooterStoryId(rId))) continue;
    visitStory(context, headerFooterStoryId(rId), part.content, {
      includePageBreaks: false,
      appendBodyTail: false,
      seedComments: true,
    });
  }
  for (const note of document.package.footnotes ?? []) {
    visitStory(context, footnoteStoryId(note.id), note.content, {
      includePageBreaks: false,
      appendBodyTail: false,
      seedComments: true,
    });
  }

  for (const plan of context.plans) session.createStory(plan.storyId, '', 'Normal', 'left');
  for (const plan of context.plans) seedPlan(session, plan);
}
