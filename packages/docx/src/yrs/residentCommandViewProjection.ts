import type {
  DocxCommandEchoContext,
  DocxEchoGlyphLine,
  DocxEchoRect,
  DocxCommandSelection,
  ResidentDocxCommandEngine,
} from './residentCommandProtocol';
import type { DocxCommandSelectionLayout } from './residentCommandSelection';
import type { DisplayPrimitive, GlyphRunPrimitive } from '../layout/render/displayList';
import type { RetainedFrame } from '../layout/render/frameDelta';
import { displayPrimitiveRect } from '../layout/render/displayListGeometry';
import { echoRound3, placeEchoGlyphs, type EchoGlyphIndex } from './residentEchoGlyphs';

/** @internal */
export const DOCX_ECHO_MAX_CONTEXT_BYTES = 16 * 1024;

const echoEncoder = new TextEncoder();

function utf8CodePointBytes(character: string): number {
  const point = character.codePointAt(0)!;
  return point <= 0x7f ? 1 : point <= 0x7ff ? 2 : point <= 0xffff ? 3 : 4;
}

/** @internal A glyph set with the ink extent of its glyphs, in font units. */
export interface EchoGlyphSource {
  key: string;
  glyphs: EchoGlyphIndex;
  ink: { upem: number; xMin: number; yMin: number; yMax: number; overhang: number };
}

/** @internal The worker's shaping facts for exact echo placement. */
export interface EchoGlyphProvider {
  glyphSet(fontId: number, size: number, kern: boolean): EchoGlyphSource | undefined;
  /** Ink box of one glyph in font units with the font's units per em; null without ink, undefined when unknown. */
  glyphInk(fontId: number, glyphId: number): { upem: number; xMin: number; xMax: number; yMin: number; yMax: number } | null | undefined;
  pixelRatio: number;
}

/** @internal */
export interface ProjectCommandEchoContextInput {
  engine: ResidentDocxCommandEngine;
  selection: DocxCommandSelection | null;
  layout: DocxCommandSelectionLayout;
  /** Without it the context carries no exact glyph line. */
  glyphs?: EchoGlyphProvider;
}

/** @internal Undefined declines echo, including when the serialized context exceeds the cap. */
export function projectCommandEchoContext(
  input: ProjectCommandEchoContextInput
): DocxCommandEchoContext | undefined {
  const { engine, selection, layout } = input;
  if (!selection || selection.anchor.story !== 'body' || selection.focus.story !== 'body' ||
    selection.anchor.paraId !== selection.focus.paraId) return undefined;
  const snapshot = engine.residentCaretSnapshot();
  const caret = snapshot.caretRect;
  if (!caret || snapshot.frameEpoch !== layout.frame.frameEpoch) return undefined;
  const activePage = layout.frame.pages[caret.pageIndex];
  if (!activePage || activePage.pageId.toString() !== caret.pageId || activePage.page.unbuilt) return undefined;
  let length: number;
  try {
    const span = engine.geometryReader.locateParagraph('body', selection.focus.paraId);
    length = span.end - span.start;
  } catch { return undefined; }
  if (!Number.isSafeInteger(length) || length <= 0 || length > DOCX_ECHO_MAX_CONTEXT_BYTES) return undefined;
  const read = engine.proposalEngine.readParagraphs({ story: 'body', paraIds: [selection.focus.paraId], view: 'accepted' });
  const paragraph = read.ok ? read.paragraphs[0] : undefined;
  if (!paragraph || paragraph.paraId !== selection.focus.paraId || paragraph.atoms.length || !paragraph.text ||
    echoEncoder.encode(paragraph.text).byteLength > DOCX_ECHO_MAX_CONTEXT_BYTES ||
    /[\u0000-\u001f\u007f\u0590-\u109f\u200e-\u200f\u202a-\u202e\u2066-\u2069]/u.test(paragraph.text)) return undefined;
  const starts = new Set<number>();
  let retainedBytes = 0;
  let retainedGlyphs = 0;
  const onCaretLine = (primitive: GlyphRunPrimitive) =>
    primitive.glyphs.some((glyph) => glyph.y >= caret.y && glyph.y <= caret.y + caret.height);
  for (const primitive of activePage.page.primitives) {
    if (primitive.kind !== 'glyphRun' || primitive.docStart === undefined || primitive.listMarker ||
      !onCaretLine(primitive)) continue;
    retainedBytes += primitive.text.length;
    retainedGlyphs += primitive.glyphs.length;
    if (retainedBytes > DOCX_ECHO_MAX_CONTEXT_BYTES || retainedGlyphs > DOCX_ECHO_MAX_CONTEXT_BYTES / 16) return undefined;
    const byteToUtf16 = new Map<number, number>();
    let byte = 0;
    let utf16 = 0;
    for (const character of primitive.text) {
      byteToUtf16.set(byte, utf16);
      byte += utf8CodePointBytes(character);
      utf16 += character.length;
    }
    byteToUtf16.set(byte, utf16);
    for (let index = 0; index < primitive.glyphs.length; index += 1) {
      const glyph = primitive.glyphs[index]!;
      if (glyph.y < caret.y || glyph.y > caret.y + caret.height) continue;
      const local = byteToUtf16.get(glyph.cluster);
      if (local !== undefined && Math.abs(glyph.x - caret.x) < 0.01) {
        starts.add(primitive.docStart + local - selection.focus.offset);
      }
      const next = primitive.glyphs[index + 1];
      const end = next ? byteToUtf16.get(next.cluster) : primitive.text.length;
      if (end !== undefined && glyph.advance !== undefined && Math.abs(glyph.x + glyph.advance - caret.x) < 0.01) {
        starts.add(primitive.docStart + end - selection.focus.offset);
      }
    }
  }
  if (starts.size !== 1) return undefined;
  const paragraphStart = starts.values().next().value!;
  const paragraphEnd = paragraphStart + paragraph.text.length;
  for (const retained of [activePage]) {
    if (retained.page.unbuilt) continue;
    const runs = retained.page.primitives.filter((primitive): primitive is GlyphRunPrimitive =>
      primitive.kind === 'glyphRun' && primitive.docStart !== undefined && primitive.docEnd !== undefined &&
      primitive.docStart >= paragraphStart && primitive.docEnd <= paragraphEnd && !primitive.listMarker &&
      onCaretLine(primitive));
    runs.sort((a, b) => a.docStart! - b.docStart!);
    if (!runs.length) continue;
    const start = runs[0]!.docStart!;
    const end = runs.at(-1)!.docEnd!;
    const lineOffset = start - paragraphStart;
    const text = paragraph.text.slice(lineOffset, end - paragraphStart);
    if (runs.map((run) => run.text).join('') !== text) continue;
    let cursor = start;
    let baseline: number | undefined;
    let right = -Infinity;
    const positions = new Map<number, number>();
    const styleRuns: DocxCommandEchoContext['styleRuns'][number][] = [];
    let supported = true;
    for (const run of runs) {
      if (run.docStart !== cursor || run.docEnd !== cursor + run.text.length || !run.glyphs.length ||
        run.rtl || run.bidiLevel || run.rotationDeg || run.horizontalScale && run.horizontalScale !== 100 ||
        run.allCaps || run.smallCaps || run.hidden || run.opacity !== undefined && run.opacity !== 1 ||
        run.wordSpacing || run.paintClip || run.clipGroup || run.modernEffects ||
        run.textShadow || run.textOutline || run.emphasisMark || run.textEffect || run.leaderGlyphs || run.revision) {
        supported = false;
        break;
      }
      const font = /^(.*?)\b(\d+(?:\.\d+)?)px\s+(.+)$/.exec(run.fallbackFont ?? '');
      if (!font || !Number.isFinite(run.size) || run.size <= 0) { supported = false; break; }
      const byteToUtf16 = new Map<number, number>();
      let byte = 0;
      let utf16 = 0;
      for (const character of run.text) {
        byteToUtf16.set(byte, utf16);
        byte += utf8CodePointBytes(character);
        utf16 += character.length;
      }
      byteToUtf16.set(byte, utf16);
      const offset = cursor - start;
      let lastX = -Infinity;
      for (const glyph of run.glyphs) {
        const local = byteToUtf16.get(glyph.cluster);
        baseline ??= glyph.y;
        if (local === undefined || glyph.advance === undefined || glyph.advance < 0 || glyph.bidiLevel ||
          glyph.y !== baseline || glyph.x < lastX || !Number.isFinite(glyph.x + glyph.y + glyph.advance)) {
          supported = false; break;
        }
        if (!positions.has(offset + local)) positions.set(offset + local, glyph.x);
        lastX = glyph.x;
        right = Math.max(right, glyph.x + glyph.advance);
      }
      positions.set(offset + run.text.length, right);
      styleRuns.push({ startUtf16: offset, endUtf16: offset + run.text.length, fontFamily: font[3]!,
        fontSizePx: run.size, color: run.color, bold: /\bbold\b|\b[6-9]00\b/.test(font[1]!),
        italic: /\bitalic\b|\boblique\b/.test(font[1]!), underline: false, strike: false,
        baseline: { x: run.glyphs[0]!.x, y: run.glyphs[0]!.y }, transform: [1, 0, 0, 1, 0, 0] });
      cursor = run.docEnd!;
    }
    if (!supported || baseline === undefined) continue;
    const offsets = [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)]
      .map((segment) => segment.index).concat(text.length);
    if (offsets.some((offset) => !positions.has(offset))) continue;
    const anchorGrapheme = offsets.indexOf(selection.anchor.offset - lineOffset);
    const focusGrapheme = offsets.indexOf(selection.focus.offset - lineOffset);
    if (anchorGrapheme < 0 || focusGrapheme < 0) return undefined;
    const size = Math.max(...styleRuns.map((run) => run.fontSizePx));
    const x = positions.get(0)!;
    const y = Math.max(0, baseline - size);
    const height = Math.min(retained.page.height - y, size * 1.5);
    let edge = retained.page.contentBounds
      ? retained.page.contentBounds.x + retained.page.contentBounds.width : retained.page.width;
    const active = new Set(runs);
    const omitted = new Set<number>();
    for (let index = 0; index < retained.page.primitives.length; index += 1) {
      const primitive = retained.page.primitives[index]!;
      if (primitive.kind === 'glyphRun' && active.has(primitive)) omitted.add(index);
      if (primitive.kind === 'decoration' && (primitive.deco === 'underline' || primitive.deco === 'strike') &&
        primitive.docStart !== undefined && primitive.docEnd !== undefined &&
        primitive.docStart >= start && primitive.docEnd <= end) {
        omitted.add(index);
        for (const style of styleRuns) {
          if (primitive.docStart < start + style.endUtf16 && primitive.docEnd > start + style.startUtf16) {
            if (primitive.deco === 'underline') style.underline = true;
            else style.strike = true;
          }
        }
      }
      if ((primitive.kind === 'image' || primitive.kind === 'shape') &&
        primitive.y < y + height && primitive.y + primitive.h > y && primitive.x + primitive.w > x) {
        edge = Math.min(edge, primitive.x);
      }
      if ((primitive.kind === 'text' || primitive.kind === 'glyphRun') && !omitted.has(index) &&
        (primitive.kind === 'text' ? Math.abs(primitive.baselineY - baseline) < size :
          primitive.glyphs.some((glyph) => Math.abs(glyph.y - baseline) < size))) supported = false;
    }
    if (!supported || edge < right || edge <= x || height <= 0) continue;
    const page = { pageId: retained.pageId.toString(), pageIndex: retained.pageIndex };
    let damageRect: DocxEchoRect = { x, y, width: edge - x, height };
    let glyphLine: DocxEchoGlyphLine | undefined;
    if (input.glyphs) {
      const projected = projectGlyphLine({ provider: input.glyphs, retained, runs, omitted, text, baseline, edge,
        paragraphText: paragraph.text, lineOffset, start, end, paragraphStart, paragraphEnd });
      if (!projected) return undefined;
      ({ line: glyphLine, rect: damageRect } = projected);
    }
    const context: DocxCommandEchoContext = {
      page, frameEpoch: layout.frame.frameEpoch, layoutRevision: layout.layoutRevision, serializedBytes: 0,
      text, styleRuns: mergeStyleRuns(styleRuns), selection: { anchorGrapheme, focusGrapheme },
      boundaries: offsets.map((utf16, grapheme) => ({ utf16, grapheme,
        caret: { x: positions.get(utf16)!, y, width: 1, height } })),
      transform: [1, 0, 0, 1, 0, 0], shaping: 'simple-ltr',
      damage: [{ page, rect: damageRect,
        backgroundId: `${layout.frame.frameEpoch}:${page.pageId}:echo`,
        omittedPrimitiveIds: [...omitted].sort((a, b) => a - b).map((index) => retained.primitiveIds[index]!.toString()) }],
      ...(glyphLine ? { glyphs: glyphLine } : {}),
    };
    let bytes = echoEncoder.encode(JSON.stringify(context)).byteLength;
    while (bytes !== context.serializedBytes) {
      context.serializedBytes = bytes;
      bytes = echoEncoder.encode(JSON.stringify(context)).byteLength;
    }
    return bytes <= DOCX_ECHO_MAX_CONTEXT_BYTES ? context : undefined;
  }
  return undefined;
}

type EchoStyleRun = DocxCommandEchoContext['styleRuns'][number];

/** Cluster-positioned lines carry one primitive per cluster; the echo needs one run per style. */
function mergeStyleRuns(runs: readonly EchoStyleRun[]): EchoStyleRun[] {
  const merged: EchoStyleRun[] = [];
  for (const run of runs) {
    const previous = merged.at(-1);
    if (previous && previous.endUtf16 === run.startUtf16 && previous.fontFamily === run.fontFamily &&
      previous.fontSizePx === run.fontSizePx && previous.color === run.color && previous.bold === run.bold &&
      previous.italic === run.italic && previous.underline === run.underline && previous.strike === run.strike &&
      previous.baseline.y === run.baseline.y && previous.transform.every((value, index) => value === run.transform[index])) {
      previous.endUtf16 = run.endUtf16;
    } else merged.push({ ...run });
  }
  return merged;
}

interface GlyphLineInput {
  provider: EchoGlyphProvider;
  retained: RetainedFrame['pages'][number];
  runs: readonly GlyphRunPrimitive[];
  omitted: ReadonlySet<number>;
  text: string;
  baseline: number;
  edge: number;
  paragraphText: string;
  lineOffset: number;
  start: number;
  end: number;
  paragraphStart: number;
  paragraphEnd: number;
}

const CSS_FONT = /^(italic )?(400|700) (\d+(?:\.\d+)?)px (.+), sans-serif$/;

/**
 * The echoed line in the engine's own placement: verifies that one glyph set reproduces every glyph
 * and pen position of the line, then bounds the ink a single-character edit can paint. Undefined
 * when the line, its neighbours or later paint fall outside what the echo can reproduce exactly.
 */
function projectGlyphLine(input: GlyphLineInput): { line: DocxEchoGlyphLine; rect: DocxEchoRect } | undefined {
  const { provider, retained, runs, text, baseline } = input;
  const first = runs[0]!;
  const facts = first.echo;
  const kern = facts?.kern ?? false;
  if (!facts || facts.primaryFont === false || runs.some(run => !run.echo || run.echo.primaryFont === false ||
    run.echo.x !== facts.x || run.echo.right !== facts.right || (run.echo.kern ?? false) !== kern ||
    (run.echo.kerning ?? 0) !== (facts.kerning ?? 0) ||
    run.echo.letterSpacing || run.echo.position || run.echo.scale !== undefined && run.echo.scale !== 100)) return undefined;
  if (runs.some((run) => run.fontId !== first.fontId || run.size !== first.size || run.color !== first.color ||
    run.fallbackFont !== first.fallbackFont) ||
    [...input.omitted].some((index) => retained.page.primitives[index]!.kind !== 'glyphRun')) return undefined;
  const font = CSS_FONT.exec(first.fallbackFont ?? '');
  const halfPoints = Math.round(first.size * 1.5);
  if (!font || echoRound3(halfPoints / 2 * 96 / 72) !== first.size) return undefined;
  const shapingSize = Math.fround(Math.fround(halfPoints / 2 * 96) / 72);
  const glyphIds: number[] = [];
  const xs: number[] = [];
  for (const run of runs) {
    let byte = 0;
    let glyph = 0;
    for (const char of run.text) {
      const placed = run.glyphs[glyph];
      if (!placed || placed.cluster !== byte || run.glyphs[glyph + 1]?.cluster === byte) return undefined;
      glyphIds.push(placed.id);
      xs.push(placed.x);
      byte += utf8CodePointBytes(char);
      glyph += 1;
    }
    if (glyph !== run.glyphs.length) return undefined;
  }
  const after = input.paragraphText.codePointAt(input.lineOffset + text.length);
  const next = after === undefined ? '' : String.fromCodePoint(after);
  const source = provider.glyphSet(first.fontId, shapingSize, kern);
  if (!source) return undefined;
  let previous: number | undefined;
  const count = source.glyphs.set.glyphIds.length;
  for (const character of input.paragraphText) {
    const index = source.glyphs.chars.get(character);
    if (index === undefined || previous !== undefined && source.glyphs.unsafe.has(previous * count + index)) return undefined;
    previous = index;
  }
  const page = retained.page;
  const placement = placeEchoGlyphs(source.glyphs, text, next, facts.x);
  if (!placement || placement.glyphIds.some((id, index) => id !== glyphIds[index]) ||
    placement.xs.some((value, index) => value !== xs[index]) || !Number.isFinite(facts.x + facts.right) ||
    input.edge < facts.right || facts.right <= facts.x) return undefined;
  const chosen = { x: facts.x };
  const right = facts.right;
  let inkEnd = chosen.x;
  [...text].forEach((char, index) => {
    if (char !== ' ') inkEnd = placement.xs[index]! + placement.advances[index]!;
  });
  // a filled line may be justified, which a shorter line would stretch
  if (inkEnd >= right - 0.5) return undefined;
  if (page.primitives.some((primitive, index) => !input.omitted.has(index) && primitive.docStart !== undefined &&
    primitive.docEnd !== undefined && primitive.docStart < input.end && primitive.docEnd > input.start)) return undefined;
  const lineRuns = (from: number, to: number, above: boolean) => page.primitives.filter((primitive): primitive is GlyphRunPrimitive =>
    primitive.kind === 'glyphRun' && !primitive.listMarker && primitive.docStart !== undefined && primitive.docEnd !== undefined &&
    primitive.docStart >= from && primitive.docEnd <= to && primitive.glyphs.length > 0 &&
    (above ? primitive.glyphs[0]!.y < baseline : primitive.glyphs[0]!.y > baseline))
    .sort((a, b) => a.docStart! - b.docStart!);
  let previousSlack: number | undefined;
  if (input.lineOffset > 0) {
    const previous = lineRuns(input.paragraphStart, input.start, true);
    if (!previous.length) return undefined;
    const last = previous.at(-1)!;
    const glyph = last.glyphs.at(-1)!;
    if (!last.echo) return undefined;
    previousSlack = last.echo.right - (glyph.x + (glyph.advance ?? 0));
  }
  let nextWord: number | undefined;
  if (input.lineOffset + text.length < input.paragraphText.length) {
    const following = lineRuns(input.end, input.paragraphEnd, false);
    nextWord = 0;
    for (const run of following) {
      if (run.glyphs[0]!.y !== following[0]!.glyphs[0]!.y || !/^[A-Za-z0-9]+$/.test(run.text)) break;
      nextWord += run.glyphs.reduce((width, glyph) => width + (glyph.advance ?? 0), 0);
    }
  }
  const ink = { xMin: Math.min(0, source.ink.xMin / source.ink.upem), yMin: Math.min(0, source.ink.yMin / source.ink.upem),
    yMax: Math.max(0, source.ink.yMax / source.ink.upem), overhang: Math.max(0, source.ink.overhang / source.ink.upem) };
  const ratio = provider.pixelRatio;
  const left = Math.floor((xs[0]! + ink.xMin * first.size) * ratio) - 1;
  const top = Math.floor((baseline - ink.yMax * first.size) * ratio) - 1;
  const inkRight = Math.ceil((Math.max(right, placement.end) + ink.overhang * first.size) * ratio) + 1;
  const bottom = Math.ceil((baseline - ink.yMin * first.size) * ratio) + 1;
  const device = {
    x: Math.max(0, left), y: Math.max(0, top),
    width: Math.min(Math.ceil(page.width * ratio), inkRight) - Math.max(0, left),
    height: Math.min(Math.ceil(page.height * ratio), bottom) - Math.max(0, top),
  };
  if (device.width <= 0 || device.height <= 0) return undefined;
  const band = { x: device.x / ratio, y: device.y / ratio, w: device.width / ratio, h: device.height / ratio };
  if (paintedOver(provider, page, input.omitted, band)) return undefined;
  return {
    line: { glyphSet: source.key, fontId: first.fontId, size: first.size,
      color: first.color, x: chosen.x, y: baseline, next, right,
      ...(facts.canEmpty === undefined ? {} : { canEmpty: facts.canEmpty }),
      ...(previousSlack === undefined ? {} : { previousSlack }), ...(nextWord === undefined ? {} : { nextWord }),
      device, pixelRatio: ratio },
    rect: { x: band.x, y: band.y, width: band.w, height: band.h },
  };
}

/** Whether anything painted after the echoed line reaches into its repaint band. */
function paintedOver(
  provider: EchoGlyphProvider,
  page: RetainedFrame['pages'][number]['page'],
  omitted: ReadonlySet<number>,
  band: { x: number; y: number; w: number; h: number },
): boolean {
  const margin = 1 / provider.pixelRatio;
  const meets = (rect: { x: number; y: number; w: number; h: number }) =>
    rect.x - margin < band.x + band.w && rect.x + rect.w + margin > band.x &&
    rect.y - margin < band.y + band.h && rect.y + rect.h + margin > band.y;
  const reaches = (primitive: DisplayPrimitive): boolean => {
    if (primitive.kind !== 'glyphRun') {
      const rect = displayPrimitiveRect(primitive);
      const slack = primitive.kind === 'shape' || primitive.kind === 'image' || primitive.kind === 'text'
        ? Math.max(rect.h, 16) : 0;
      return meets({ x: rect.x - slack, y: rect.y - slack, w: rect.w + 2 * slack, h: rect.h + 2 * slack });
    }
    const coarse = displayPrimitiveRect(primitive);
    if (!meets({ x: coarse.x - primitive.size, y: coarse.y - primitive.size,
      w: coarse.w + 2 * primitive.size, h: coarse.h + 2 * primitive.size })) return false;
    return primitive.glyphs.some((glyph) => {
      const ink = provider.glyphInk(primitive.fontId, glyph.id);
      if (ink === null) return false;
      if (!ink) return true;
      const scale = primitive.size / ink.upem;
      return meets({ x: glyph.x + ink.xMin * scale, y: glyph.y - ink.yMax * scale,
        w: (ink.xMax - ink.xMin) * scale, h: (ink.yMax - ink.yMin) * scale });
    });
  };
  const firstOmitted = Math.min(...omitted);
  for (let index = firstOmitted + 1; index < page.primitives.length; index += 1) {
    if (!omitted.has(index) && reaches(page.primitives[index]!)) return true;
  }
  for (const area of page.noteAreas ?? []) {
    for (const primitive of [...(area.separatorPrimitives ?? []), ...(area.primitives ?? [])]) {
      if (reaches(primitive)) return true;
    }
  }
  return (page.pageBorders ?? []).some((border) => border.zOrder !== 'back' &&
    meets({ x: border.x, y: border.y, w: border.w, h: border.h }));
}
