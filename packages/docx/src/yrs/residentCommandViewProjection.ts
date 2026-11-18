import type {
  DocxCommandEchoContext,
  DocxEchoGlyphFace,
  DocxEchoGlyphLine,
  DocxEchoRect,
  DocxCommandSelection,
  DocxWorkerCommandEngine,
} from './residentCommandProtocol';
import type { DocxCommandSelectionLayout } from './residentCommandSelection';
import type { DisplayPrimitive, GlyphRunPrimitive } from '../layout/render/displayList';
import type { RetainedFrame } from '../layout/render/frameDelta';
import { displayPrimitiveRect } from '../layout/render/displayListGeometry';
import {
  DOCX_ECHO_ALPHABET, DOCX_ECHO_JAMO, echoRound3, placeEchoFaceGlyphs, placeEchoLine, type EchoGlyphIndex,
} from './residentEchoGlyphs';

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
  /** The set of a font at a size, kerning gate and shaping language whose alphabet holds `chars`. */
  glyphSet(fontId: number, size: number, kern: boolean, language?: string, chars?: string): EchoGlyphSource | undefined;
  /** Ink box of one glyph in font units with the font's units per em; null without ink, undefined when unknown. */
  glyphInk(fontId: number, glyphId: number): { upem: number; xMin: number; xMax: number; yMin: number; yMax: number } | null | undefined;
  pixelRatio: number;
}

/** @internal */
export interface ProjectCommandEchoContextInput {
  engine: DocxWorkerCommandEngine;
  selection: DocxCommandSelection | null;
  layout: DocxCommandSelectionLayout;
  /** Without it the context carries no exact glyph line. */
  glyphs?: EchoGlyphProvider;
}

/** @internal Undefined declines echo, including when the serialized context exceeds the cap. */
const BLOCK_EMBEDS = new Set(['table', 'blockSdt', 'pageBreak', 'columnBreak']);

const revised = (values: Record<string, unknown> | undefined, keys: readonly string[]) =>
  keys.some((key) => values?.[key] !== undefined && values[key] !== null);

/**
 * A paragraph's accepted-view text and whether it holds inline atoms, read from its own segments
 * without building the story's view, with the offset units of the block embeds it starts with,
 * which the text leaves out and paragraph offsets count. Undefined wherever tracked changes touch
 * it or the mark before it (they can join paragraphs or hide text), or its id is not unique: the
 * caller then reads the story. An engine without segment reads also reads the story.
 */
export function readEchoParagraph(
  engine: Pick<DocxWorkerCommandEngine, 'geometryReader' | 'paragraphUnit' | 'paragraphSegments'>,
  story: string,
  paraId: string
): { text: string; hasAtoms: boolean; lead: number } | undefined {
  if (typeof engine.geometryReader.paragraphIdCount !== 'function' || typeof engine.paragraphUnit !== 'function' ||
    typeof engine.paragraphSegments !== 'function') return undefined;
  if (engine.geometryReader.paragraphIdCount(story, paraId) !== 1) return undefined;
  const unit = engine.paragraphUnit(story, paraId);
  if (unit < 0) return undefined;
  const units = engine.paragraphSegments(story, unit > 0 ? [unit - 1, unit] : [unit]);
  const own = units.at(-1);
  if (!own || units.length !== (unit > 0 ? 2 : 1)) return undefined;
  if (unit > 0) {
    const mark = units[0]!.at(-1);
    if (mark?.kind !== 'pilcrow' || revised(mark.attributes, ['ins', 'del']) ||
      revised(mark.properties, ['pPrIns', 'pPrDel'])) return undefined;
  }
  let text = '';
  let hasAtoms = false;
  let leading = true;
  let lead = 0;
  for (const segment of own) {
    if (revised(segment.attributes, ['ins', 'del'])) return undefined;
    if (segment.kind === 'text') {
      text += segment.text;
      leading = false;
    } else if (segment.kind === 'embed') {
      if (leading && BLOCK_EMBEDS.has(segment.embedKind)) { lead += 1; continue; }
      text += '\uFFFC';
      hasAtoms = true;
      leading = false;
    } else {
      if (segment.paraId !== paraId || revised(segment.properties, ['pPrIns', 'pPrDel'])) return undefined;
      return { text, hasAtoms, lead };
    }
  }
  return undefined;
}

/**
 * Offset units of the block embeds a paragraph starts with, for a paragraph the story read gave:
 * its text leaves them out too. Zero for an engine without segment reads; the line text check
 * below still declines a line the paragraph text does not match.
 */
function leadingBlockUnits(
  engine: Pick<DocxWorkerCommandEngine, 'paragraphUnit' | 'paragraphSegments'>,
  story: string,
  paraId: string
): number | undefined {
  if (typeof engine.paragraphUnit !== 'function' || typeof engine.paragraphSegments !== 'function') return 0;
  const unit = engine.paragraphUnit(story, paraId);
  const own = unit < 0 ? undefined : engine.paragraphSegments(story, [unit])[0];
  if (!own) return undefined;
  let lead = 0;
  for (const segment of own) {
    if (segment.kind !== 'embed' || !BLOCK_EMBEDS.has(segment.embedKind)) break;
    lead += 1;
  }
  return lead;
}

export function projectCommandEchoContext(
  input: ProjectCommandEchoContextInput
): DocxCommandEchoContext | undefined {
  const { engine, selection, layout } = input;
  // the body and its table cells, whose display positions are the body's
  if (!selection || !/^body(?::t\d+:r\d+c\d+)*$/.test(selection.focus.story) ||
    selection.anchor.story !== selection.focus.story || selection.anchor.paraId !== selection.focus.paraId) return undefined;
  const story = selection.focus.story;
  const snapshot = engine.residentCaretSnapshot();
  const caret = snapshot.caretRect;
  if (!caret || snapshot.frameEpoch !== layout.frame.frameEpoch) return undefined;
  const activePage = layout.frame.pages[caret.pageIndex];
  if (!activePage || activePage.pageId.toString() !== caret.pageId || activePage.page.unbuilt) return undefined;
  let length: number;
  try {
    const span = engine.geometryReader.locateParagraph(story, selection.focus.paraId);
    length = span.end - span.start;
  } catch { return undefined; }
  if (!Number.isSafeInteger(length) || length <= 0 || length > DOCX_ECHO_MAX_CONTEXT_BYTES) return undefined;
  const fast = readEchoParagraph(engine, story, selection.focus.paraId);
  const read = fast ? undefined
    : engine.proposalEngine.readParagraphs({ story, paraIds: [selection.focus.paraId], view: 'accepted' });
  const paragraph = fast ? { paraId: selection.focus.paraId, text: fast.text, atoms: fast.hasAtoms ? [0] : [] }
    : read?.ok ? read.paragraphs[0] : undefined;
  // paragraph offsets count the block embeds the paragraph starts with; its text does not
  const lead = fast ? fast.lead : leadingBlockUnits(engine, story, selection.focus.paraId);
  const anchorOffset = selection.anchor.offset - (lead ?? 0);
  const focusOffset = selection.focus.offset - (lead ?? 0);
  if (lead === undefined || anchorOffset < 0 || focusOffset < 0) return undefined;
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
        starts.add(primitive.docStart + local - focusOffset);
      }
      const next = primitive.glyphs[index + 1];
      const end = next ? byteToUtf16.get(next.cluster) : primitive.text.length;
      if (end !== undefined && glyph.advance !== undefined && Math.abs(glyph.x + glyph.advance - caret.x) < 0.01) {
        starts.add(primitive.docStart + end - focusOffset);
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
        run.wordSpacing || run.paintClip || !sameClip(run.clipGroup, runs[0]!.clipGroup) ||
        run.clipGroup?.opacity !== undefined && run.clipGroup.opacity !== 1 || run.modernEffects ||
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
    const anchorGrapheme = offsets.indexOf(anchorOffset - lineOffset);
    const focusGrapheme = offsets.indexOf(focusOffset - lineOffset);
    if (anchorGrapheme < 0 || focusGrapheme < 0) return undefined;
    const size = Math.max(...styleRuns.map((run) => run.fontSizePx));
    const x = positions.get(0)!;
    const y = Math.max(0, baseline - size);
    const height = Math.min(retained.page.height - y, size * 1.5);
    const contentEdge = retained.page.contentBounds
      ? retained.page.contentBounds.x + retained.page.contentBounds.width : retained.page.width;
    let edge = contentEdge;
    const active = new Set(runs);
    const cell = runs[0]!.clipGroup;
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
      // an empty text primitive only marks a blank line's position for hit testing and paints nothing;
      // another table cell's text keeps its place, and whatever paints over the band declines below
      if ((primitive.kind === 'text' && primitive.text !== '' || primitive.kind === 'glyphRun') && !omitted.has(index) &&
        !(cell?.id !== undefined && primitive.clipGroup?.id !== undefined && primitive.clipGroup.id !== cell.id) &&
        (primitive.kind === 'text' ? Math.abs(primitive.baselineY - baseline) < size :
          primitive.glyphs.some((glyph) => Math.abs(glyph.y - baseline) < size))) supported = false;
    }
    if (!supported || edge < right || edge <= x || height <= 0) continue;
    const page = { pageId: retained.pageId.toString(), pageIndex: retained.pageIndex };
    let damageRect: DocxEchoRect = { x, y, width: edge - x, height };
    let glyphLine: DocxEchoGlyphLine | undefined;
    if (input.glyphs) {
      const facesJson = engine.echoFacesJson?.bind(engine);
      // a table cell clips its text; the echo paints inside the clip the engine paints with
      const finite = (value: number | undefined) => typeof value === 'number' && Number.isFinite(value) ? value : 0;
      const clip = cell?.clip && { x: finite(cell.clip.x), y: finite(cell.clip.y),
        w: Math.max(0, finite(cell.clip.w)), h: Math.max(0, finite(cell.clip.h)) };
      const projected = projectGlyphLine({ provider: input.glyphs, retained, runs, omitted, text, baseline, edge,
        obstructed: edge < contentEdge, caret: focusOffset - lineOffset, ...(clip ? { clip } : {}),
        ...(facesJson ? { faces: (position: number, alphabet: string, lineStart: number, lineEnd: number) => {
          const json = facesJson(position, alphabet, lineStart, lineEnd);
          return json === undefined ? undefined : JSON.parse(json) as EchoFacesAnswer;
        } } : {}),
        paragraphText: paragraph.text, lineOffset, start, end, paragraphStart, paragraphEnd });
      if (!projected) return undefined;
      ({ line: glyphLine, rect: damageRect } = projected);
    }
    const context: DocxCommandEchoContext = {
      page, frameEpoch: layout.frame.frameEpoch, layoutRevision: layout.layoutRevision, serializedBytes: 0,
      text, styleRuns: mergeStyleRuns(styleRuns), selection: { anchorGrapheme, focusGrapheme },
      boundaries: offsets.map((utf16, grapheme) => ({ utf16, grapheme,
        caret: { x: positions.get(utf16)!, y: caret.y, width: 1, height: caret.height } })),
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

/** Whether two runs paint under the same clip group, or neither under one. */
function sameClip(a: GlyphRunPrimitive['clipGroup'], b: GlyphRunPrimitive['clipGroup']): boolean {
  if (!a || !b) return a === b;
  return a.id === b.id && a.opacity === b.opacity && a.clip?.x === b.clip?.x && a.clip?.y === b.clip?.y &&
    a.clip?.w === b.clip?.w && a.clip?.h === b.clip?.h;
}

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

/** The engine's `echoFacesJson` answer: `[display, measured, metrics, eastAsian]` per character. */
interface EchoFacesAnswer {
  faces: (readonly (number | null)[])[];
  line: ((readonly (number | null)[]) | null)[];
  latin: string | null;
  eastAsia: string | null;
  /** Display positions of the text run typing goes into. */
  run?: [number, number];
}

interface GlyphLineInput {
  provider: EchoGlyphProvider;
  /** The engine's font routing for the line, for a line the display list does not draw with its primary font. */
  faces?(position: number, alphabet: string, lineStart: number, lineEnd: number): EchoFacesAnswer | undefined;
  /** UTF-16 offset of the caret in the line text. */
  caret: number;
  retained: RetainedFrame['pages'][number];
  runs: readonly GlyphRunPrimitive[];
  omitted: ReadonlySet<number>;
  text: string;
  baseline: number;
  edge: number;
  /** Whether an image or shape beside the line, not the page's content box, sets `edge`. */
  obstructed: boolean;
  /** The table cell clip the engine paints the line inside, in page px. */
  clip?: { x: number; y: number; w: number; h: number };
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
 * A line the display list draws through more than one face, or not with its primary font, places
 * each character with its own face's glyph set and types only what the engine's font routing
 * draws and measures with one face that keeps the line's height.
 */
function projectGlyphLine(input: GlyphLineInput): { line: DocxEchoGlyphLine; rect: DocxEchoRect } | undefined {
  const { provider, retained, runs, text, baseline } = input;
  const first = runs[0]!;
  const facts = first.echo;
  const kern = facts?.kern ?? false;
  if (!facts || runs.some(run => !run.echo ||
    run.echo.x !== facts.x || run.echo.right !== facts.right || (run.echo.kern ?? false) !== kern ||
    (run.echo.kerning ?? 0) !== (facts.kerning ?? 0) || (run.echo.letterSpacing ?? 0) !== (facts.letterSpacing ?? 0) ||
    (run.echo.autoSpaceOff ?? 0) !== (facts.autoSpaceOff ?? 0) ||
    run.echo.position || run.echo.scale !== undefined && run.echo.scale !== 100 ||
    run.echo.stretch !== undefined || run.echo.centerWidth !== facts.centerWidth)) return undefined;
  // the engine's font routing places a line that mixes faces, or writes outside the base alphabet
  const routed = runs.some((run) => run.echo!.primaryFont === false || run.fontId !== first.fontId ||
    run.lang !== first.lang) || [...input.paragraphText].some((char) => !DOCX_ECHO_ALPHABET.includes(char));
  if (!routed && facts.letterSpacing) return undefined;
  // a centered line moves as a whole when its width changes, so only a paragraph of one line is placed
  const center = facts.centerWidth;
  if (center !== undefined && (routed || input.lineOffset !== 0 || text.length !== input.paragraphText.length ||
    text.endsWith(' '))) return undefined;
  if (runs.some((run) => run.size !== first.size || run.color !== first.color ||
    !routed && (run.fontId !== first.fontId || run.fallbackFont !== first.fallbackFont)) ||
    [...input.omitted].some((index) => retained.page.primitives[index]!.kind !== 'glyphRun')) return undefined;
  const font = CSS_FONT.exec(first.fallbackFont ?? '');
  const halfPoints = Math.round(first.size * 1.5);
  if (!font || echoRound3(halfPoints / 2 * 96 / 72) !== first.size) return undefined;
  const shapingSize = Math.fround(Math.fround(halfPoints / 2 * 96) / 72);
  const glyphIds: number[] = [];
  const xs: number[] = [];
  const charRuns: GlyphRunPrimitive[] = [];
  for (const run of runs) {
    if (routed && !CSS_FONT.test(run.fallbackFont ?? '')) return undefined;
    let byte = 0;
    let glyph = 0;
    for (const char of run.text) {
      const placed = run.glyphs[glyph];
      if (!placed || placed.cluster !== byte || run.glyphs[glyph + 1]?.cluster === byte) return undefined;
      glyphIds.push(placed.id);
      xs.push(placed.x);
      charRuns.push(run);
      byte += utf8CodePointBytes(char);
      glyph += 1;
    }
    if (glyph !== run.glyphs.length) return undefined;
  }
  const after = input.paragraphText.codePointAt(input.lineOffset + text.length);
  const next = after === undefined ? '' : String.fromCodePoint(after);
  const page = retained.page;
  let placement: ReturnType<typeof placeEchoLine>;
  let inkSources: EchoGlyphSource[];
  let spaceAdvance: (() => number | undefined) | undefined;
  let route: Pick<DocxEchoGlyphLine, 'glyphSet' | 'fontId' | 'faces' | 'charFaces' | 'nextFace' | 'typed' | 'typedRun' |
    'letterSpacing' | 'shapingSize' | 'autoSpaceOff'>;
  if (!routed) {
    const source = provider.glyphSet(first.fontId, shapingSize, kern);
    if (!source) return undefined;
    let previous: number | undefined;
    const count = source.glyphs.set.glyphIds.length;
    for (const character of input.paragraphText) {
      const index = source.glyphs.chars.get(character);
      if (index === undefined || previous !== undefined && source.glyphs.unsafe.has(previous * count + index)) return undefined;
      previous = index;
    }
    placement = placeEchoLine(source.glyphs, text, next, facts.x, center);
    inkSources = [source];
    spaceAdvance = () => {
      const space = source.glyphs.chars.get(' ');
      return space === undefined ? undefined : source.glyphs.advance(source.glyphs.set.advances[space]!);
    };
    route = { glyphSet: source.key, fontId: first.fontId };
  } else {
    // pairs kern within one face only where the line does not mix faces
    if (kern || !input.faces) return undefined;
    const lineChars = [...text];
    const alphabet = [...new Set(DOCX_ECHO_ALPHABET + DOCX_ECHO_JAMO + input.paragraphText)].join('');
    const caretPosition = input.start + Math.max(0, input.caret - 1);
    const answer = input.faces(caretPosition, alphabet, input.start, input.end + next.length);
    if (!answer || answer.faces.length !== [...alphabet].length) return undefined;
    const languageOf = (eastAsian: number | null | undefined) =>
      (eastAsian ? answer.eastAsia ?? answer.latin : answer.latin) ?? undefined;
    const faces: { fontId: number; language: string | undefined; metrics: number; chars: Set<string>; typed: string }[] = [];
    const faceOf = (fontId: number, language: string | undefined, metrics: number) => {
      let index = faces.findIndex((face) => face.fontId === fontId && face.language === language && face.metrics === metrics);
      if (index < 0) index = faces.push({ fontId, language, metrics, chars: new Set(), typed: '' }) - 1;
      return index;
    };
    const charFaces: number[] = [];
    const lineMetrics = new Set<number>();
    let utf16 = 0;
    for (let index = 0; index < lineChars.length; index += 1) {
      const entry = answer.line[utf16];
      const run = charRuns[index]!;
      // the display list draws it with the face measurement placed it with
      if (!entry || entry[0] !== run.fontId || entry[1] !== run.fontId || typeof entry[2] !== 'number') return undefined;
      lineMetrics.add(entry[2]);
      const face = faceOf(run.fontId, run.lang, entry[2]);
      faces[face]!.chars.add(lineChars[index]!);
      charFaces.push(face);
      utf16 += lineChars[index]!.length;
    }
    // the character after the line shapes with the last one where both take one face
    const following = next ? answer.line[text.length] : undefined;
    const nextFace = following && typeof following[0] === 'number'
      ? faces.findIndex((face) => face.fontId === following[0]) : -1;
    if (nextFace >= 0) faces[nextFace]!.chars.add(next);
    // typing goes into the run left of the caret; a character types where both paths take one face
    // whose line metrics the line already holds
    const run = answer.run;
    const typedRun: [number, number] | undefined = run && input.caret > 0
      ? [Math.max(0, run[0] - input.start), Math.min(text.length, run[1] - input.start)] : undefined;
    if (typedRun && typedRun[0] < input.caret && typedRun[1] >= input.caret) {
      [...alphabet].forEach((char, index) => {
        const [display, measured, metrics, eastAsian] = answer.faces[index]!;
        if (typeof display !== 'number' || display !== measured || typeof metrics !== 'number' ||
          !lineMetrics.has(metrics)) return;
        const face = faces[faceOf(display, languageOf(eastAsian), metrics)]!;
        face.chars.add(char);
        face.typed += char;
      });
    }
    // faces of one font and language share a glyph set, asked for once with all their characters
    const wanted = new Map<string, { fontId: number; language: string | undefined; chars: Set<string> }>();
    for (const face of faces) {
      const key = `${face.fontId}\u0000${face.language ?? ''}`;
      const entry = wanted.get(key) ?? { fontId: face.fontId, language: face.language, chars: new Set<string>() };
      for (const char of face.chars) entry.chars.add(char);
      wanted.set(key, entry);
    }
    const shaped = new Map<string, EchoGlyphSource>();
    for (const [key, entry] of wanted) {
      const source = provider.glyphSet(entry.fontId, shapingSize, false, entry.language, [...entry.chars].join(''));
      if (!source) return undefined;
      shaped.set(key, source);
    }
    const sources = faces.map((face) => shaped.get(`${face.fontId}\u0000${face.language ?? ''}`)!);
    // a following character of the last one's face that its set cannot echo may shape with it
    if (nextFace >= 0 && !sources[nextFace]!.glyphs.chars.has(next)) return undefined;
    placement = placeEchoFaceGlyphs({ sets: sources.map((source) => source.glyphs), charFaces,
      ...(nextFace >= 0 ? { nextFace } : {}), letterSpacing: facts.letterSpacing,
      autoSpaceOff: facts.autoSpaceOff, size: shapingSize }, text, next, facts.x);
    inkSources = sources;
    const lineFaces: DocxEchoGlyphFace[] = faces.map((face, index) =>
      ({ glyphSet: sources[index]!.key, fontId: face.fontId, metrics: face.metrics }));
    const typed = faces.map((face, index) => [...face.typed].filter((char) => sources[index]!.glyphs.chars.has(char)).join(''));
    route = { glyphSet: lineFaces[0]!.glyphSet, fontId: lineFaces[0]!.fontId, faces: lineFaces, charFaces,
      ...(nextFace >= 0 ? { nextFace } : {}),
      ...(typed.some(Boolean) && typedRun ? { typed, typedRun } : {}),
      ...(facts.letterSpacing ? { letterSpacing: Math.fround(facts.letterSpacing) } : {}),
      shapingSize, ...(facts.autoSpaceOff ? { autoSpaceOff: facts.autoSpaceOff } : {}) };
  }
  if (!placement || placement.glyphIds.some((id, index) => id !== glyphIds[index]) ||
    placement.xs.some((value, index) => value !== xs[index]) || !Number.isFinite(facts.x + facts.right) ||
    // the content box is serialized to 0.001 px, so its edge may round just short of the line's
    (input.obstructed ? input.edge < facts.right : facts.right - input.edge > 0.01) || facts.right <= facts.x) return undefined;
  const chosen = { x: facts.x };
  const right = facts.right;
  let inkEnd = chosen.x;
  [...text].forEach((char, index) => {
    if (char !== ' ') inkEnd = placement!.xs[index]! + placement!.advances[index]!;
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
    if (last.echo.stretch === undefined) previousSlack = last.echo.right - (glyph.x + (glyph.advance ?? 0));
    else {
      // a justified line takes a word that fits once a quarter of its space width is squeezed out
      const line = previous.filter((run) => run.glyphs[0]!.y === last.glyphs[0]!.y);
      const space = spaceAdvance?.();
      if (space === undefined || line.some((run) => run.fontId !== first.fontId || run.size !== first.size)) return undefined;
      const spaces = line.reduce((count, run) => count + [...run.text].filter((char) => char === ' ').length, 0);
      previousSlack = last.echo.right - (last.echo.x + last.echo.stretch) + (spaces / 4 + 1) * space;
    }
  }
  let nextWord: number | undefined;
  if (input.lineOffset + text.length < input.paragraphText.length) {
    const following = lineRuns(input.end, input.paragraphEnd, false);
    nextWord = 0;
    for (const run of following) {
      if (run.glyphs[0]!.y !== following[0]!.glyphs[0]!.y) break;
      if (!/^[A-Za-z0-9]+$/.test(run.text)) {
        // the least a break can move up is the first character, where it is not a letter or digit
        if (run === following[0] && !/^[A-Za-z0-9]/.test(run.text)) nextWord = run.glyphs[0]!.advance ?? 0;
        break;
      }
      nextWord += run.glyphs.reduce((width, glyph) => width + (glyph.advance ?? 0), 0);
    }
  }
  const ink = { xMin: 0, yMin: 0, yMax: 0, overhang: 0 };
  for (const source of inkSources) {
    ink.xMin = Math.min(ink.xMin, source.ink.xMin / source.ink.upem);
    ink.yMin = Math.min(ink.yMin, source.ink.yMin / source.ink.upem);
    ink.yMax = Math.max(ink.yMax, source.ink.yMax / source.ink.upem);
    ink.overhang = Math.max(ink.overhang, source.ink.overhang / source.ink.upem);
  }
  const ratio = provider.pixelRatio;
  const left = Math.floor(((center === undefined ? xs[0]! : chosen.x) + ink.xMin * first.size) * ratio) - 1;
  const top = Math.floor((baseline - ink.yMax * first.size) * ratio) - 1;
  const inkRight = Math.ceil((Math.max(right, placement.end) + ink.overhang * first.size) * ratio) + 1;
  const bottom = Math.ceil((baseline - ink.yMin * first.size) * ratio) + 1;
  // the line paints nothing outside its cell's clip, so the band stops at the pixels the clip touches
  const clip = input.clip;
  const x0 = Math.max(0, left, clip ? Math.floor(clip.x * ratio) : 0);
  const y0 = Math.max(0, top, clip ? Math.floor(clip.y * ratio) : 0);
  const x1 = Math.min(Math.ceil(page.width * ratio), inkRight, clip ? Math.ceil((clip.x + clip.w) * ratio) : Infinity);
  const y1 = Math.min(Math.ceil(page.height * ratio), bottom, clip ? Math.ceil((clip.y + clip.h) * ratio) : Infinity);
  const device = { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
  if (device.width <= 0 || device.height <= 0) return undefined;
  const band = { x: device.x / ratio, y: device.y / ratio, w: device.width / ratio, h: device.height / ratio };
  if (paintedOver(provider, page, input.omitted, band)) return undefined;
  return {
    line: { ...route, size: first.size,
      color: first.color, x: chosen.x, y: baseline, next, right, ...(center === undefined ? {} : { center }),
      ...(facts.canEmpty === undefined ? {} : { canEmpty: facts.canEmpty }),
      ...(previousSlack === undefined ? {} : { previousSlack }), ...(nextWord === undefined ? {} : { nextWord }),
      device, pixelRatio: ratio, ...(clip ? { clip } : {}) },
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
    if (primitive.kind === 'text' && primitive.text === '') return false;
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
