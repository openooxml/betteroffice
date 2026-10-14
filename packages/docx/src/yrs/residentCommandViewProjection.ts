import type {
  DocxCommandEchoContext,
  DocxCommandSelection,
  ResidentDocxCommandEngine,
} from './residentCommandProtocol';
import type { DocxCommandSelectionLayout } from './residentCommandSelection';
import type { GlyphRunPrimitive } from '../layout/render/displayList';

/** @internal */
export const DOCX_ECHO_MAX_CONTEXT_BYTES = 16 * 1024;

/** @internal */
export interface ProjectCommandEchoContextInput {
  engine: ResidentDocxCommandEngine;
  selection: DocxCommandSelection | null;
  layout: DocxCommandSelectionLayout;
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
  if (!activePage || activePage.pageId.toString() !== caret.pageId || activePage.page.unbuilt ||
    activePage.page.primitives.length > DOCX_ECHO_MAX_CONTEXT_BYTES / 16) return undefined;
  let length: number;
  try {
    const span = engine.geometryReader.locateParagraph('body', selection.focus.paraId);
    length = span.end - span.start;
  } catch { return undefined; }
  if (!Number.isSafeInteger(length) || length <= 0 || length > DOCX_ECHO_MAX_CONTEXT_BYTES) return undefined;
  const read = engine.proposalEngine.readParagraphs({ story: 'body', paraIds: [selection.focus.paraId], view: 'accepted' });
  const paragraph = read.ok ? read.paragraphs[0] : undefined;
  if (!paragraph || paragraph.paraId !== selection.focus.paraId || paragraph.atoms.length || !paragraph.text ||
    new TextEncoder().encode(paragraph.text).byteLength > DOCX_ECHO_MAX_CONTEXT_BYTES ||
    /[\u0000-\u001f\u007f\u0590-\u109f\u200e-\u200f\u202a-\u202e\u2066-\u2069]/u.test(paragraph.text)) return undefined;
  const starts = new Set<number>();
  let retainedBytes = 0;
  let retainedGlyphs = 0;
  for (const primitive of activePage.page.primitives) {
    if (primitive.kind !== 'glyphRun' || primitive.docStart === undefined || primitive.listMarker) continue;
    retainedBytes += primitive.text.length;
    retainedGlyphs += primitive.glyphs.length;
    if (retainedBytes > DOCX_ECHO_MAX_CONTEXT_BYTES || retainedGlyphs > DOCX_ECHO_MAX_CONTEXT_BYTES / 16) return undefined;
    const byteToUtf16 = new Map<number, number>();
    let byte = 0;
    let utf16 = 0;
    for (const character of primitive.text) {
      byteToUtf16.set(byte, utf16);
      byte += new TextEncoder().encode(character).byteLength;
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
  const start = starts.values().next().value!;
  const end = start + paragraph.text.length;
  for (const retained of [activePage]) {
    if (retained.page.unbuilt) continue;
    const runs = retained.page.primitives.filter((primitive): primitive is GlyphRunPrimitive =>
      primitive.kind === 'glyphRun' && primitive.docStart !== undefined && primitive.docEnd !== undefined &&
      primitive.docStart >= start && primitive.docEnd <= end && !primitive.listMarker);
    runs.sort((a, b) => a.docStart! - b.docStart!);
    if (!runs.length || runs[0]!.docStart !== start || runs.at(-1)!.docEnd !== end ||
      runs.map((run) => run.text).join('') !== paragraph.text) continue;
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
        byte += new TextEncoder().encode(character).byteLength;
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
    const offsets = [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(paragraph.text)]
      .map((segment) => segment.index).concat(paragraph.text.length);
    if (offsets.some((offset) => !positions.has(offset))) continue;
    const anchorGrapheme = offsets.indexOf(selection.anchor.offset);
    const focusGrapheme = offsets.indexOf(selection.focus.offset);
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
    const context: DocxCommandEchoContext = {
      page, frameEpoch: layout.frame.frameEpoch, layoutRevision: layout.layoutRevision, serializedBytes: 0,
      text: paragraph.text, styleRuns, selection: { anchorGrapheme, focusGrapheme },
      boundaries: offsets.map((utf16, grapheme) => ({ utf16, grapheme,
        caret: { x: positions.get(utf16)!, y, width: 1, height } })),
      transform: [1, 0, 0, 1, 0, 0], shaping: 'simple-ltr',
      damage: [{ page, rect: { x, y, width: edge - x, height },
        backgroundId: `${layout.frame.frameEpoch}:${page.pageId}:echo`,
        primitiveIds: Array.from(retained.primitiveIds, (id, index) => omitted.has(index) ? null : id.toString())
          .filter((id): id is string => id !== null) }],
    };
    let bytes = new TextEncoder().encode(JSON.stringify(context)).byteLength;
    while (bytes !== context.serializedBytes) {
      context.serializedBytes = bytes;
      bytes = new TextEncoder().encode(JSON.stringify(context)).byteLength;
    }
    return bytes <= DOCX_ECHO_MAX_CONTEXT_BYTES ? context : undefined;
  }
  return undefined;
}
