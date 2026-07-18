/**
 * The Rust measurement source — the SOLE text/layout measurement path. Every
 * block kind is measured here: paragraphs through the wasm `docx-text` engine
 * (Word metrics), tables/text boxes/shapes by pure recursion through the same
 * measurer, and image/chart/break extents synthesized from declared
 * dimensions. There is no browser `measureText` anywhere in layout.
 *
 * Division of labor:
 *  - {@link TextMeasureFontRegistry} (owned here) turns Word font requests
 *    into ordered engine font-id chains from embedded + bundled bytes, and
 *    resolves the shared per-script coverage fallbacks (CJK/RTL Noto faces).
 *    Resolution is deterministic: embedded faces → bundled metric-compatible
 *    face → always-available last-resort base face; never OS/local fonts.
 *  - {@link RustMeasureSource.prepareFonts} is the ONLY async step: it warms
 *    the registry for every (family, bold, italic) combo a block list uses,
 *    plus the script fallbacks for every script detected in the blocks'
 *    text — and for every request an earlier measure pass found unresolved
 *    (header/footer and footnote paragraphs reach measurement without
 *    passing through the adapters' block notifications; recording their
 *    misses here closes that gap). Call sites re-layout after it resolves.
 *  - {@link RustMeasureSource.createMeasureBlock} builds the sync
 *    `MeasureBlockFn` the float pipeline consumes. Paragraph blocks whose
 *    font chains are cached go to the engine — float-zoned blocks included,
 *    with the zones and cumulative Y passed through in the envelope.
 *
 * READINESS GATE (how measurement never races an unloaded face): a paragraph
 * whose font chain (or script fallback) has not settled yet measures as a
 * deterministic SYNTHETIC extent and records the miss;
 * {@link RustMeasureSource.hasPendingFonts} then reports true, the adapter
 * discards that provisional pass instead of committing it, and re-runs the
 * pipeline when `prepareFonts` settles. Once settled, `font-unready` can
 * never recur for those chains (the sync registry view is populated), so
 * committed layouts are always measured against real, loaded font bytes.
 *
 * SELF-SUFFICIENT FONTS: every family chain the registry builds terminates in
 * an always-available last-resort base face (broad-coverage Liberation
 * Sans/Serif per serif-ness, supplied by the bundled provider's
 * `resolveLastResort`), and the per-script Noto fallback is appended whenever
 * a covered script is detected. A run whose family has no embedded/bundled
 * match — or a character outside the family chain's coverage — still has real
 * font bytes to measure with; the engine substitutes the terminal face's
 * `.notdef` for genuinely uncovered characters. For a truly-unknown font the
 * measured metrics are the base font's, not the requested family's — the
 * accepted, deterministic substitution tradeoff.
 *
 * SYNTHETIC extents (deterministic pure arithmetic, no font tables, no
 * platform state) stand in for exactly two residual cases: (a) an engine
 * refusal — degenerate/attacker-crafted input tripping the security clamps,
 * or a malformed payload; (b) a chain that SETTLED EMPTY, which can only
 * happen when the host injects no bundled provider (`resolveLastResort`
 * guarantees non-empty chains otherwise) and the document embeds nothing for
 * the family. Both are counted in {@link RustMeasureStats}.
 *
 * Script fallbacks: runs are scanned for Unicode script ranges — kana → JP,
 * Hangul → KR, Arabic → arabic, Hebrew → hebrew — and Han is region-resolved:
 * kana in the same block wins (Japanese text), else the document's
 * `w:themeFontLang` eastAsia hint ({@link RustMeasureSource.setScriptHints}),
 * else Hangul in the block (hanja), else SC. The resolved script ids are
 * appended AFTER the family faces of every chain the block uses, so the
 * family face keeps supplying metrics/Latin glyphs and the script face only
 * covers what the family face cannot. RTL text measures natively (the engine
 * splits UBA level runs before shaping).
 *
 * @packageDocumentation
 * @public
 */

import type { CompatibilityFlags } from '../../docx/settingsParser';
import {
  assertExhaustiveLayoutBlock,
  DEFAULT_TEXTBOX_MARGINS,
  DEFAULT_TEXTBOX_WIDTH,
  type BlockExtent,
  type ImageBlock,
  type LayoutBlock,
  type ParagraphBlock,
  type ParagraphExtent,
  type ShapeBlock,
  type TableBlock,
  type TextBoxBlock,
  type TextRun,
} from '../pagination/types';
import type { MeasureBlockFn } from './measureBlocksPipeline';
import { hashParagraphBlock } from './paragraphHash';
import {
  TextMeasureFontRegistry,
  type BundledFontProvider,
  type EmbeddedFaceInput,
  type FontScript,
} from './fontRegistry';
import { measureTableBlock } from './measureTable';

/**
 * Font styling properties for a short-text measurement request (the
 * `measureTextWidthWithActiveRustSource` latch and tab-leader glyph
 * sizing).
 *
 * @public
 */
export interface FontStyle {
  fontFamily?: string;
  /** In points. */
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  /** In pixels. */
  letterSpacing?: number;
}

/**
 * Thin sync facade over the wasm measure exports (`register_measure_font`,
 * `clear_measure_fonts`, `measure_paragraph_json`). Injectable so unit tests
 * can script responses without loading wasm.
 *
 * @public
 */
export interface RustTextEngine {
  /** Register raw sfnt bytes; returns the engine font id. Throws on bad bytes. */
  registerFont(bytes: Uint8Array): number;
  /** Drop every registered font (ids restart at 0). */
  clearFonts(): void;
  /**
   * Measurement input JSON in, `ParagraphExtent` JSON out. Throws with a
   * message starting `UNSUPPORTED:` for blocks the engine cannot measure.
   */
  measureParagraphJson(input: string): string;
}

let enginePromise: Promise<RustTextEngine> | null = null;

/**
 * Lazily load the embedded wasm module and expose its measure surface as a
 * {@link RustTextEngine}. The dynamic `import()` keeps the ~800KB base64
 * module out of the default bundle — only callers that actually enable Rust
 * measurement pay for it (same pattern as the display-list builder's
 * `loadEngine`).
 *
 * @public
 */
export function getRustTextEngine(): Promise<RustTextEngine> {
  enginePromise ??= import('../wasm/index').then(async (m) => {
    await m.preloadLayoutWasm();
    return {
    registerFont: m.registerMeasureFont,
    clearFonts: m.clearMeasureFonts,
    measureParagraphJson: m.measureParagraphJson,
  };
  });
  return enginePromise;
}

/**
 * Fallbacks for runs/paragraphs with no explicit formatting (Word 2007+
 * document defaults). The Rust engine hardcodes no font names, so the host
 * passes these in as the `defaults` envelope field.
 */
const DEFAULT_FONT_SIZE = 11; // pt (Word 2007+ default)
const DEFAULT_FONT_FAMILY = 'Calibri';

/**
 * Bound on the internal Rust-measure memo; eviction is insert-order
 * (oldest first).
 */
const RUST_MEASURE_CACHE_LIMIT = 5000;

/**
 * Why a paragraph measured as a synthetic extent, for the debug counters.
 *  - `non-font`: a degenerate content width or an engine error unrelated to
 *    glyph coverage (malformed/attacker-crafted payload tripping the engine's
 *    security clamps). Never a font-availability problem, so warming fonts
 *    can never remove it.
 *  - `font-unready`: a font chain or the block's script fallback has not
 *    resolved yet — transient; the miss is recorded, `hasPendingFonts`
 *    reports it, and the adapter discards the pass and re-runs once
 *    `prepareFonts` settles (the readiness gate in the module doc).
 *  - `uncovered`: the engine refused for glyph-coverage reasons — dormant
 *    now that chains terminate in a last-resort face and the shaper
 *    substitutes its `.notdef`, but kept so a regression is categorized.
 */
type FallbackReason = 'non-font' | 'font-unready' | 'uncovered';

/**
 * Memoized engine refusal — carries the categorized reason so a later memo hit
 * re-counts it in the same bucket. Its `refused` key discriminates it from a
 * {@link ParagraphExtent} (whose discriminant is `kind: 'paragraph'`).
 */
interface RefusedEntry {
  readonly refused: FallbackReason;
}

type RustParagraphMeasureResult =
  | {
      readonly ok: true;
      readonly extent: ParagraphExtent;
    }
  | {
      readonly ok: false;
      readonly reason: FallbackReason;
    };

type ActiveRustTextWidthResult =
  | { readonly active: false }
  | { readonly active: true; readonly width: number | undefined };

let activeRustTextWidth: ((text: string, style: FontStyle) => number | undefined) | undefined;

/**
 * Reset the per-layout-pass Rust text-width latch. `computeLayout` calls this
 * before measurement; the measure function sets it when a pass runs, so
 * passes that never measured (e.g. engine still loading) stay inactive.
 *
 * @internal
 */
export function beginRustMeasureLayoutPass(): void {
  activeRustTextWidth = undefined;
}

/**
 * Measure a short run through the Rust source that was active during the
 * current layout pass. Inactive means no measure pass has run yet (callers
 * skip their optional geometry); active + `undefined` means the required
 * font chain is not ready or the engine refused the text — callers must
 * skip, never estimate by other means.
 *
 * @internal
 */
export function measureTextWidthWithActiveRustSource(
  text: string,
  style: FontStyle
): ActiveRustTextWidthResult {
  if (!activeRustTextWidth) return { active: false };
  return { active: true, width: activeRustTextWidth(text, style) };
}

/**
 * Whether an engine throw is a glyph-coverage refusal (`UNSUPPORTED: no font
 * in chain covers U+XXXX`) rather than another failure. The wasm boundary
 * throws a plain string (`JsValue::from_str`); the mock engine in tests throws
 * an `Error` — handle both.
 */
function isCoverageError(err: unknown): boolean {
  const msg = err instanceof Error ? err.message : typeof err === 'string' ? err : String(err);
  return msg.includes('no font in chain covers');
}

/** One (family, bold, italic) request; `family` is the raw block string. */
interface FontRequest {
  family: string;
  bold: boolean;
  italic: boolean;
}

/**
 * `fontChains` key exactly as the Rust `chain_for` lookup builds it:
 * lowercased family (no trim — the engine lowercases the raw run string),
 * then bold/italic as 0|1.
 */
function rustChainKey(req: FontRequest): string {
  return `${req.family.toLowerCase()}|${req.bold ? 1 : 0}|${req.italic ? 1 : 0}`;
}

/**
 * Collect every (family, bold, italic) combo one paragraph's text, tab and
 * field runs use, PLUS the regular (`|0|0`) chain for each text-run family —
 * the engine's empty / whitespace-only paragraph path measures with the
 * regular face. Keyed by {@link rustChainKey} so the map doubles as the
 * `fontChains` skeleton.
 */
function collectParagraphFontRequests(block: ParagraphBlock, out: Map<string, FontRequest>): void {
  const add = (family: string, bold: boolean, italic: boolean) => {
    const req = { family, bold, italic };
    const key = rustChainKey(req);
    if (!out.has(key)) out.set(key, req);
  };
  const paragraphDefault = block.attrs?.defaultFontFamily ?? DEFAULT_FONT_FAMILY;
  const addFormatting = (
    run: Extract<ParagraphBlock['runs'][number], { kind: 'text' | 'tab' | 'field' }>,
    includeRegular: boolean
  ) => {
    const baseFamily = run.fontFamily ?? paragraphDefault;
    add(baseFamily, !!run.bold, !!run.italic);
    if (includeRegular) add(baseFamily, false, false);
    const slots = run.fontSlots;
    for (const family of [slots?.ascii, slots?.hAnsi, slots?.eastAsia]) {
      if (!family) continue;
      add(family, !!run.bold, !!run.italic);
      if (includeRegular) add(family, false, false);
    }
    if (slots?.cs) {
      add(slots.cs, run.boldCs ?? !!run.bold, run.italicCs ?? !!run.italic);
      if (includeRegular) add(slots.cs, false, false);
    }
  };
  // Empty-paragraph path resolves attrs.defaultFontFamily ?? defaults.
  add(paragraphDefault, false, false);
  for (const run of block.runs) {
    if (run.kind === 'text') {
      addFormatting(run, true);
    } else if (run.kind === 'tab' || run.kind === 'field') {
      // The engine resolves a tab run's font for line metrics (the TS
      // `updateMaxFont(runToFontStyle(run))` equivalent) and measures a
      // field run's fallback text with the run's face. Both default to the
      // document family directly, not the paragraph default.
      addFormatting(run, false);
    }
  }
  // A visible list marker is measured with the marker family's regular face
  // (resolveListMarkerFont order: level rPr → first text run → paragraph
  // default). Collected whenever the marker is visible — for hanging-indent
  // paragraphs the engine never resolves it, but warming the chain is cheap
  // and keeps the collection independent of indent details.
  const attrs = block.attrs;
  if (attrs?.listMarker && !attrs.listMarkerHidden) {
    const firstTextRun = block.runs.find((r): r is TextRun => r.kind === 'text');
    add(attrs.listMarkerFontFamily ?? firstTextRun?.fontFamily ?? paragraphDefault, false, false);
  }
}

/** Deterministic order for appended script-fallback ids. */
const SCRIPT_ORDER: FontScript[] = ['cjk-sc', 'cjk-tc', 'cjk-jp', 'cjk-kr', 'arabic', 'hebrew'];

/** Region a Han-only run resolves to (kana/Hangul/hint decide — see below). */
type HanRegion = 'cjk-sc' | 'cjk-tc' | 'cjk-jp' | 'cjk-kr';

/** Raw per-range hits of one scan; Han region is resolved afterwards. */
interface ScriptScan {
  han: boolean;
  kana: boolean;
  hangul: boolean;
  arabic: boolean;
  hebrew: boolean;
}

/**
 * Flag the script ranges present in `text`. Iterates code points (surrogate
 * pairs stay whole). CJK punctuation / fullwidth forms count as Han: every
 * CJK face covers them, and the Han bucket's region resolution picks which
 * one. Deliberately coarse — this feeds font *coverage*, not shaping.
 */
function scanTextScripts(text: string, flags: ScriptScan): void {
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    if (cp < 0x0590) continue; // Latin, Greek, Cyrillic, … — family faces cover these
    if (cp <= 0x05ff)
      flags.hebrew = true; // Hebrew
    else if (cp <= 0x06ff)
      flags.arabic = true; // Arabic
    else if (cp >= 0x0750 && cp <= 0x077f)
      flags.arabic = true; // Arabic Supplement
    else if (cp >= 0x0870 && cp <= 0x08ff)
      flags.arabic = true; // Arabic Ext-B / Ext-A
    else if (cp >= 0x1100 && cp <= 0x11ff)
      flags.hangul = true; // Hangul Jamo
    else if (cp >= 0x3000 && cp <= 0x303f)
      flags.han = true; // CJK symbols & punctuation
    else if (cp >= 0x3040 && cp <= 0x30ff)
      flags.kana = true; // Hiragana + Katakana
    else if (cp >= 0x3130 && cp <= 0x318f)
      flags.hangul = true; // Hangul compat jamo
    else if (cp >= 0x31f0 && cp <= 0x31ff)
      flags.kana = true; // Katakana phonetic ext
    else if (cp >= 0x3400 && cp <= 0x4dbf)
      flags.han = true; // CJK ext A
    else if (cp >= 0x4e00 && cp <= 0x9fff)
      flags.han = true; // CJK unified
    else if (cp >= 0xa960 && cp <= 0xa97f)
      flags.hangul = true; // Hangul Jamo ext A
    else if (cp >= 0xac00 && cp <= 0xd7ff)
      flags.hangul = true; // Hangul syllables + Jamo ext B
    else if (cp >= 0xf900 && cp <= 0xfaff)
      flags.han = true; // CJK compatibility
    else if (cp >= 0xfb1d && cp <= 0xfb4f)
      flags.hebrew = true; // Hebrew presentation forms
    else if (cp >= 0xfb50 && cp <= 0xfdff)
      flags.arabic = true; // Arabic presentation forms A
    else if (cp >= 0xfe30 && cp <= 0xfe4f)
      flags.han = true; // CJK vertical forms
    else if (cp >= 0xfe70 && cp <= 0xfeff)
      flags.arabic = true; // Arabic presentation forms B
    else if (cp >= 0xff00 && cp <= 0xff65)
      flags.han = true; // fullwidth forms
    else if (cp >= 0xff66 && cp <= 0xff9f)
      flags.kana = true; // halfwidth katakana
    else if (cp >= 0xffa0 && cp <= 0xffdc)
      flags.hangul = true; // halfwidth jamo
    else if (cp >= 0x20000 && cp <= 0x3ffff) flags.han = true; // CJK ext B..I + compat suppl
  }
}

/**
 * Map a `w:themeFontLang` eastAsia BCP-47 tag to the Han region it implies;
 * `undefined` when the tag is absent or not an East-Asian language.
 */
function hanRegionFromLang(lang: string | undefined): HanRegion | undefined {
  if (!lang) return undefined;
  const tag = lang.trim().toLowerCase();
  if (tag.startsWith('ja')) return 'cjk-jp';
  if (tag.startsWith('ko')) return 'cjk-kr';
  if (tag.startsWith('zh')) {
    return /(^|-)(hant|tw|hk|mo)(-|$)/.test(tag) ? 'cjk-tc' : 'cjk-sc';
  }
  return undefined;
}

/**
 * Scripts one paragraph's text engages, in {@link SCRIPT_ORDER}. Han region
 * resolution: kana in the block → JP (Japanese text mixes both); else the
 * document hint; else Hangul in the block → KR (hanja); else SC — the
 * documented default (Simplified Chinese is by far the most common Han-only
 * case in the wild).
 */
function collectBlockScripts(block: ParagraphBlock, hint: HanRegion | undefined): FontScript[] {
  const flags: ScriptScan = {
    han: false,
    kana: false,
    hangul: false,
    arabic: false,
    hebrew: false,
  };
  for (const run of block.runs) {
    if (run.kind === 'text') scanTextScripts(run.text, flags);
  }
  const set = new Set<FontScript>();
  if (flags.kana) set.add('cjk-jp');
  if (flags.hangul) set.add('cjk-kr');
  if (flags.han) {
    if (flags.kana) set.add('cjk-jp');
    else if (hint) set.add(hint);
    else if (flags.hangul) set.add('cjk-kr');
    else set.add('cjk-sc');
  }
  if (flags.arabic) set.add('arabic');
  if (flags.hebrew) set.add('hebrew');
  if (set.size === 0) return [];
  return SCRIPT_ORDER.filter((s) => set.has(s));
}

function forEachParagraphBlock(
  blocks: readonly LayoutBlock[],
  visit: (block: ParagraphBlock) => void
): void {
  const walk = (block: LayoutBlock): void => {
    switch (block.kind) {
      case 'paragraph':
        visit(block as ParagraphBlock);
        return;
      case 'table':
        for (const row of (block as TableBlock).rows) {
          for (const cell of row.cells) {
            forEachParagraphBlock(cell.blocks, visit);
          }
        }
        return;
      case 'textBox':
        for (const paragraph of (block as TextBoxBlock).content) {
          visit(paragraph);
        }
        return;
      case 'shape': {
        const shape = block as ShapeBlock;
        for (const paragraph of shape.innerText ?? []) {
          visit(paragraph);
        }
        for (const child of shape.children ?? []) {
          walk(child);
        }
        return;
      }
      case 'image':
      case 'chart':
      case 'pageBreak':
      case 'columnBreak':
      case 'sectionBreak':
        return;
    }
  };
  for (const block of blocks) walk(block);
}

function measureTextBoxBlock(
  tb: TextBoxBlock,
  paragraphMeasurer: (block: ParagraphBlock, width: number) => ParagraphExtent
): BlockExtent {
  const margins = tb.margins ?? DEFAULT_TEXTBOX_MARGINS;
  const width = tb.width ?? DEFAULT_TEXTBOX_WIDTH;
  const innerWidth = Math.max(1, width - margins.left - margins.right);
  const innerMeasures = tb.content.map((p) => paragraphMeasurer(p, innerWidth));
  const contentHeight = innerMeasures.reduce((sum, m) => sum + m.totalHeight, 0);
  const totalHeight = tb.height ?? contentHeight + margins.top + margins.bottom;
  return {
    kind: 'textBox',
    width,
    height: totalHeight,
    innerMeasures,
  };
}

/**
 * Deterministic stand-in extent for a paragraph the engine cannot measure:
 * degenerate/attacker-crafted input it refused, or a font chain that settled
 * empty (no bundled provider injected AND nothing embedded — real hosts wire
 * `resolveLastResort`, so chains are non-empty by construction). Pure
 * arithmetic on declared values — no font tables, no platform state — so it
 * is byte-identical on every machine. Passes measured while chains were
 * still RESOLVING also produce these, but those passes are provisional by
 * contract (`hasPendingFonts`) and are never committed.
 */
function synthesizeParagraphExtent(block: ParagraphBlock, maxWidth: number): ParagraphExtent {
  const attrs = block.attrs;
  let fontSizePt = attrs?.defaultFontSize ?? DEFAULT_FONT_SIZE;
  let charCount = 0;
  for (const run of block.runs) {
    if (run.kind !== 'text') continue;
    charCount += run.text.length;
    if (typeof run.fontSize === 'number' && Number.isFinite(run.fontSize)) {
      fontSizePt = Math.max(fontSizePt, run.fontSize);
    }
  }
  if (!Number.isFinite(fontSizePt) || fontSizePt <= 0) fontSizePt = DEFAULT_FONT_SIZE;
  const fontSizePx = (fontSizePt * 96) / 72;
  // Conventional em partition (4:1 across the baseline) and the Word
  // single-spacing floor — same constants the engine's defaults derive from.
  const ascent = fontSizePx * 0.8;
  const descent = fontSizePx * 0.2;
  const lineHeight = fontSizePx * 1.15;
  // Crude half-em-per-character advance so downstream geometry has a finite,
  // deterministic box; capped at the content width.
  const cap = Number.isFinite(maxWidth) && maxWidth > 0 ? maxWidth : 0;
  const width = Math.min(cap, charCount * fontSizePx * 0.5);
  const tailRun = Math.max(0, block.runs.length - 1);
  const lastRun = block.runs[tailRun];
  const tailChar = lastRun?.kind === 'text' ? lastRun.text.length : 0;
  const spacing = attrs?.spacing;
  return {
    kind: 'paragraph',
    lines: [{ headRun: 0, headChar: 0, tailRun, tailChar, width, ascent, descent, lineHeight }],
    totalHeight: (spacing?.before ?? 0) + lineHeight + (spacing?.after ?? 0),
  };
}

function measureShapeBlock(
  shape: ShapeBlock,
  paragraphMeasurer: (block: ParagraphBlock, width: number) => ParagraphExtent
): BlockExtent {
  const width = shape.width ?? 100;
  const height = shape.height ?? 80;
  const innerMeasures = (shape.innerText ?? []).map((p) => paragraphMeasurer(p, width));
  shape.innerMeasures = innerMeasures;
  for (const child of shape.children ?? []) {
    measureShapeBlock(child, paragraphMeasurer);
  }
  return { kind: 'shape', width, height, innerMeasures };
}

/** Counters for the debug log line (rust vs synthetic per layout pass). */
export interface RustMeasureStats {
  /** Paragraph blocks measured by the Rust engine (including memo hits). */
  rustMeasured: number;
  /** Paragraphs that measured as a synthetic extent — sum of the reasons below. */
  syntheticFallback: number;
  /**
   * Synthetic extents caused by a font chain (or script fallback) not having
   * resolved yet. Transient by construction: the miss is recorded, the
   * adapter discards the pass (`hasPendingFonts`), and the fonts-settled
   * re-layout measures natively — so no COMMITTED layout ever contains one,
   * and this MUST hold constant once `prepareFonts` settles.
   */
  syntheticFontUnready: number;
  /**
   * Synthetic extents caused by an engine glyph-coverage refusal. Dormant —
   * chains terminate in a last-resort face whose `.notdef` the shaper
   * substitutes — but kept so a coverage regression is categorized rather
   * than lumped in with `non-font` refusals.
   */
  syntheticUncovered: number;
  /**
   * Synthetic-extent counts by `LayoutBlock.kind`. The audit counter proving
   * text-bearing block kinds (`paragraph`, `table`, `textBox`) measure
   * natively once fonts are ready.
   */
  syntheticBlockKinds: Record<string, number>;
}

/**
 * The selectable measurement source. Create with
 * {@link createRustMeasureSource}; see the module doc for the contract.
 *
 * @public
 */
export interface RustMeasureSource {
  /** Feed the open document's embedded faces (invalidates the measure memo). */
  setEmbeddedFaces(faces: EmbeddedFaceInput[]): void;
  /**
   * Thread the document's compat flags (`w:compat`) into every subsequent
   * engine call. Pass `undefined` to reset to all-off defaults. The memo key
   * includes the flags, so switching documents cannot serve stale extents.
   */
  setCompat(flags: CompatibilityFlags | undefined): void;
  /**
   * Thread document-level script hints into Han region resolution: pass the
   * `w:themeFontLang` eastAsia tag (`settingsParser` → `themeFontLang`).
   * Without a hint (or with a non-East-Asian tag) Han-only text defaults to
   * Simplified Chinese. Invalidates the measure memo — region choice changes
   * the fallback face and therefore the measured extents.
   */
  setScriptHints(hints: { eastAsia?: string } | undefined): void;
  /**
   * Warm the font registry for every (family, bold, italic) combo — plus the
   * per-family regulars and the per-script coverage fallbacks — that
   * `blocks`' paragraphs use, AND for every request an earlier measure pass
   * recorded as unresolved (header/footer and footnote paragraphs reach the
   * measurer without appearing in the adapters' block lists). The only async
   * step; per-face failures are tolerated (the chain settles without that
   * face). Resolves `true` when at least one previously-unsettled chain or
   * script fallback settled — even settled-empty — i.e. a re-layout would
   * measure differently (or the deferred pass can now proceed).
   */
  prepareFonts(blocks: LayoutBlock[]): Promise<boolean>;
  /**
   * True while a font chain (or script fallback) some measured paragraph
   * needed is still unresolved. The adapter's readiness gate: a pass measured
   * with pending fonts is provisional (those paragraphs carry synthetic
   * extents) and must be discarded, not committed; `prepareFonts` settles the
   * pending set and the fonts-ready re-layout measures natively. Self-clears
   * once every recorded request has settled.
   */
  hasPendingFonts(): boolean;
  /**
   * The sync measure function for the float pipeline — the sole measurement
   * path, exhaustive over every `LayoutBlock` kind. Paragraph blocks whose
   * chains (and detected script fallbacks) are cached go to the engine —
   * float-zoned blocks pass their zones and cumulative Y through in the
   * envelope (measured fresh, never memoized). Tables, text boxes and shapes
   * recurse through the same measurer; images/charts/breaks synthesize
   * extents from declared dimensions. Unresolved chains and engine refusals
   * produce deterministic synthetic extents (see the module doc).
   */
  createMeasureBlock(): MeasureBlockFn;
  /**
   * The merged, doc-wide font-id chain map for `blocks` — every
   * `"<family lowercase>|<bold 0|1>|<italic 0|1>"` combo the blocks' paragraphs
   * use, mapped to its cached engine font-id chain (with the per-block script
   * fallbacks appended after the family faces, exactly as
   * {@link RustMeasureSource.createMeasureBlock} assembles them). Only chains that
   * are already resolved (non-empty in the sync registry cache) are included; a
   * combo whose chain has not settled yet is simply omitted (its runs stay on
   * the char-distributed draw path until `prepareFonts` settles and a
   * re-layout retries). Returns an empty object when nothing is resolved.
   *
   * This is the map the canvas display-list builder consumes to gate GlyphRun
   * emission: the ids reference the same measurement `FontStore` the engine
   * registered, so the outline provider can rasterize them.
   */
  getDocumentFontChains(blocks: LayoutBlock[]): Record<string, number[]>;
  /** Drop every cached chain, registration, and memoized measure. */
  clear(): void;
  /** Snapshot of the rust-vs-fallback counters (never reset implicitly). */
  getStats(): RustMeasureStats;
}

/**
 * Build a {@link RustMeasureSource} over a loaded engine. The source owns a
 * {@link TextMeasureFontRegistry} whose sink is `engine.registerFont`;
 * `bundled` is the optional metric-compatible font provider (injected — core
 * never imports a fonts package).
 *
 * @public
 */
export function createRustMeasureSource(opts: {
  engine: RustTextEngine;
  bundled?: BundledFontProvider;
}): RustMeasureSource {
  const { engine } = opts;
  const registry = new TextMeasureFontRegistry(
    { registerFont: (bytes) => engine.registerFont(bytes) },
    // Deterministic by construction: only embedded document faces and the
    // injected bundled provider participate — never OS/local fonts — so the
    // same document + provider measure identically on every machine.
    { bundled: opts.bundled }
  );

  let compat: CompatibilityFlags | undefined;
  let hanHint: HanRegion | undefined;
  const stats: RustMeasureStats = {
    rustMeasured: 0,
    syntheticFallback: 0,
    syntheticFontUnready: 0,
    syntheticUncovered: 0,
    syntheticBlockKinds: {},
  };

  // Font requests (and scripts) a measure pass found unresolved — the
  // readiness gate's ledger. `prepareFonts` warms these alongside the given
  // blocks' requests; `hasPendingFonts` self-clears entries whose chains have
  // settled since. HF/footnote paragraphs enter measurement without passing
  // through the adapters' block notifications, so recording misses here is
  // what guarantees their fonts get warmed too.
  const pendingRequests = new Map<string, FontRequest>();
  const pendingScripts = new Set<FontScript>();

  // Internal memo for engine results, keyed by the paragraph content hash
  // extended with the fields that affect widths (letterSpacing, allCaps,
  // smallCaps, horizontalScale), plus maxWidth and the compat bits. A
  // RefusedEntry memoizes an engine throw — deterministic given block + fonts,
  // and the memo is dropped whenever the font set (or the Han region hint,
  // which swaps the fallback face) changes. The refusal carries its
  // categorized reason so a later memo hit re-counts it in the same bucket.
  const memo = new Map<string, ParagraphExtent | RefusedEntry>();

  function memoKey(block: ParagraphBlock, maxWidth: number): string {
    const rustOnly: unknown[] = [];
    for (const run of block.runs) {
      if (run.kind !== 'text') continue;
      rustOnly.push([
        run.letterSpacing ?? 0,
        run.allCaps ? 1 : 0,
        run.smallCaps ? 1 : 0,
        run.horizontalScale ?? 100,
        run.kerningMinPt ?? null,
        run.fontSlots ?? null,
        run.fontSizeCs ?? null,
        run.boldCs ?? null,
        run.italicCs ?? null,
        run.complexScript ? 1 : 0,
        run.language ?? null,
        run.hidden ? 1 : 0,
      ]);
    }
    return JSON.stringify([
      hashParagraphBlock(block),
      maxWidth,
      compat?.noLeading ? 1 : 0,
      compat?.doNotExpandShiftReturn ? 1 : 0,
      rustOnly,
    ]);
  }

  function memoWrite(key: string, value: ParagraphExtent | RefusedEntry): void {
    if (memo.has(key)) memo.delete(key);
    memo.set(key, value);
    while (memo.size > RUST_MEASURE_CACHE_LIMIT) {
      const oldest = memo.keys().next();
      if (oldest.done) break;
      memo.delete(oldest.value);
    }
  }

  /** Parse + shape-check the engine's JSON so a malformed payload falls back. */
  function parseExtent(json: string): ParagraphExtent {
    const parsed = JSON.parse(json) as ParagraphExtent;
    if (
      parsed === null ||
      typeof parsed !== 'object' ||
      parsed.kind !== 'paragraph' ||
      !Array.isArray(parsed.lines) ||
      !Number.isFinite(parsed.totalHeight)
    ) {
      throw new Error('rust measure returned a malformed ParagraphExtent');
    }
    return parsed;
  }

  function populateTabLeaderMetrics(block: ParagraphBlock, extent: ParagraphExtent): void {
    const leader = block.attrs?.tabs?.find((stop) => stop.leader && stop.leader !== 'none')?.leader;
    const glyph =
      leader === 'dot'
        ? '.'
        : leader === 'middleDot'
          ? '·'
          : leader === 'hyphen'
            ? '-'
            : leader === 'underscore'
              ? '_'
              : leader === 'heavy'
                ? '━'
                : undefined;
    for (let runIndex = 0; runIndex < block.runs.length; runIndex++) {
      const run = block.runs[runIndex];
      if (run?.kind !== 'tab') continue;
      const width = extent.lines.reduce(
        (sum, line) =>
          sum +
          (line.runAdvances ?? [])
            .filter((advance) => advance.runIndex === runIndex)
            .reduce((part, advance) => part + (advance.advance ?? 0), 0),
        0
      );
      run.width = width;
      if (!glyph) continue;
      const font = run.fontFamily ?? DEFAULT_FONT_FAMILY;
      const fontSize = run.fontSize ?? DEFAULT_FONT_SIZE;
      const advance = measureRustTextWidth(glyph, {
        fontFamily: font,
        fontSize,
        bold: run.bold,
        italic: run.italic,
      });
      if (advance === undefined || advance <= 0) continue;
      run.leaderGlyphs = {
        glyph,
        count: Math.max(1, Math.floor(width / advance)),
        advance,
        font,
        fontSize,
        color: run.color,
      };
    }
  }

  function tryMeasureParagraphRust(
    pBlock: ParagraphBlock,
    contentWidth: number,
    floatingZones: Parameters<MeasureBlockFn>[2],
    cumulativeY: Parameters<MeasureBlockFn>[3],
    countStats: boolean
  ): RustParagraphMeasureResult {
    // The engine requires a finite pre-indent width.
    if (!Number.isFinite(contentWidth) || contentWidth <= 0) {
      return { ok: false, reason: 'non-font' };
    }

    // Float-zoned blocks skip the memo entirely (no read, no write, not even
    // a refusal entry) — their extents depend on inter-block layout context
    // (cumulative Y, neighboring floats) the memo key cannot capture.
    const zoned = floatingZones !== undefined && floatingZones.length > 0;
    const key = zoned ? undefined : memoKey(pBlock, contentWidth);
    if (key !== undefined) {
      const memoized = memo.get(key);
      if (memoized !== undefined) {
        if ('refused' in memoized) return { ok: false, reason: memoized.refused };
        populateTabLeaderMetrics(pBlock, memoized);
        if (countStats) stats.rustMeasured++;
        return { ok: true, extent: memoized };
      }
    }

    // Assemble fontChains from the sync registry view. An UNRESOLVED chain
    // means this block cannot be measured yet: record the miss for
    // `prepareFonts`/`hasPendingFonts` (the readiness gate) and fall back to
    // a synthetic extent for this provisional pass. A chain that SETTLED
    // EMPTY (no provider, nothing embedded) can never improve — synthetic,
    // without recording or consulting the engine.
    const requests = new Map<string, FontRequest>();
    collectParagraphFontRequests(pBlock, requests);
    const blockScripts = collectBlockScripts(pBlock, hanHint);
    const fontChains: Record<string, number[]> = {};
    let unready = false;
    let settledEmpty = false;
    for (const [chainKey, req] of requests) {
      const chain = registry.getCachedFontIdChain(req.family, req.bold, req.italic);
      if (chain === undefined) {
        // Record EVERY unresolved combo (not just the first) so one
        // prepareFonts round-trip settles the whole block.
        pendingRequests.set(chainKey, req);
        unready = true;
      } else if (chain.length === 0) {
        settledEmpty = true;
      } else {
        fontChains[chainKey] = Array.from(chain);
      }
    }
    const scriptIds =
      blockScripts.length > 0 ? registry.getCachedScriptFallbackIds(blockScripts) : [];
    if (scriptIds === undefined) {
      for (const script of blockScripts) pendingScripts.add(script);
      unready = true;
    }
    if (unready) return { ok: false, reason: 'font-unready' };
    if (settledEmpty) return { ok: false, reason: 'non-font' };

    // Append shared script fallback ids after the family faces.
    if (scriptIds !== undefined && scriptIds.length > 0) {
      for (const chain of Object.values(fontChains)) {
        for (const id of scriptIds) {
          if (!chain.includes(id)) chain.push(id);
        }
      }
    }

    try {
      const input = JSON.stringify({
        // LayoutBlock passthrough — the engine's serde ignores unknown fields.
        block: pBlock,
        maxWidth: contentWidth,
        fontChains,
        defaults: { fontSize: DEFAULT_FONT_SIZE, fontFamily: DEFAULT_FONT_FAMILY },
        compat: {
          noLeading: compat?.noLeading ?? false,
          doNotExpandShiftReturn: compat?.doNotExpandShiftReturn ?? false,
        },
        authoritativeShaping: true,
        // Float context passthrough. Unzoned envelopes stay byte-identical to
        // the pre-float contract.
        ...(zoned ? { floatingZones, paragraphYOffset: cumulativeY ?? 0 } : {}),
      });
      const extent = parseExtent(engine.measureParagraphJson(input));
      populateTabLeaderMetrics(pBlock, extent);
      if (key !== undefined) memoWrite(key, extent);
      if (countStats) stats.rustMeasured++;
      return { ok: true, extent };
    } catch (err) {
      const reason: FallbackReason = isCoverageError(err) ? 'uncovered' : 'non-font';
      if (key !== undefined) memoWrite(key, { refused: reason });
      return { ok: false, reason };
    }
  }

  function measureRustTextWidth(text: string, style: FontStyle): number | undefined {
    if (text.length === 0) return 0;
    const block: ParagraphBlock = {
      kind: 'paragraph',
      id: '__rust-text-width',
      runs: [
        {
          kind: 'text',
          text,
          fontFamily: style.fontFamily ?? DEFAULT_FONT_FAMILY,
          fontSize: style.fontSize ?? DEFAULT_FONT_SIZE,
          bold: style.bold,
          italic: style.italic,
          letterSpacing: style.letterSpacing,
        },
      ],
      attrs: {
        defaultFontFamily: style.fontFamily ?? DEFAULT_FONT_FAMILY,
        defaultFontSize: style.fontSize ?? DEFAULT_FONT_SIZE,
      },
    };
    const result = tryMeasureParagraphRust(block, 10000, undefined, undefined, false);
    if (!result.ok) return undefined;
    return result.extent.lines[0]?.width ?? 0;
  }

  function recordFallback(block: LayoutBlock, reason: FallbackReason): void {
    stats.syntheticFallback++;
    stats.syntheticBlockKinds[block.kind] = (stats.syntheticBlockKinds[block.kind] ?? 0) + 1;
    if (reason === 'font-unready') stats.syntheticFontUnready++;
    else if (reason === 'uncovered') stats.syntheticUncovered++;
  }

  return {
    setEmbeddedFaces(faces: EmbeddedFaceInput[]): void {
      registry.setEmbeddedFaces(faces);
      memo.clear();
    },

    setCompat(flags: CompatibilityFlags | undefined): void {
      compat = flags;
    },

    setScriptHints(hints: { eastAsia?: string } | undefined): void {
      const next = hanRegionFromLang(hints?.eastAsia);
      if (next === hanHint) return;
      hanHint = next;
      memo.clear();
    },

    async prepareFonts(blocks: LayoutBlock[]): Promise<boolean> {
      const requests = new Map<string, FontRequest>();
      const scripts = new Set<FontScript>();
      forEachParagraphBlock(blocks, (block) => {
        collectParagraphFontRequests(block, requests);
        for (const script of collectBlockScripts(block, hanHint)) {
          scripts.add(script);
        }
      });
      // Fold in the misses earlier measure passes recorded (HF/footnote
      // paragraphs — see the ledger comment above).
      for (const [key, req] of pendingRequests) {
        if (!requests.has(key)) requests.set(key, req);
      }
      for (const script of pendingScripts) scripts.add(script);
      const settled = await Promise.all([
        ...Array.from(requests.values(), async (req) => {
          try {
            const before = registry.getCachedFontIdChain(req.family, req.bold, req.italic);
            if (before !== undefined) return false; // already settled
            await registry.getFontIdChain(req.family, req.bold, req.italic);
            // Settling is the news — even an empty chain unblocks a deferred
            // pass (those blocks proceed with synthetic extents instead of
            // waiting forever).
            return true;
          } catch {
            // Failure-tolerant by contract: a failed resolution behaves like
            // a settled miss.
            return false;
          }
        }),
        ...Array.from(scripts, async (script) => {
          try {
            if (registry.getCachedScriptFallbackIds([script]) !== undefined) return false;
            await registry.getScriptFallbackIds([script]);
            return true;
          } catch {
            return false;
          }
        }),
      ]);
      return settled.some(Boolean);
    },

    hasPendingFonts(): boolean {
      for (const [key, req] of pendingRequests) {
        if (registry.getCachedFontIdChain(req.family, req.bold, req.italic) === undefined) {
          return true;
        }
        pendingRequests.delete(key); // settled since — self-clear
      }
      for (const script of pendingScripts) {
        if (registry.getCachedScriptFallbackIds([script]) === undefined) return true;
        pendingScripts.delete(script);
      }
      return false;
    },

    createMeasureBlock(): MeasureBlockFn {
      const measure: MeasureBlockFn = (block, contentWidth, floatingZones, cumulativeY) => {
        activeRustTextWidth = measureRustTextWidth;

        switch (block.kind) {
          case 'paragraph': {
            const result = tryMeasureParagraphRust(
              block as ParagraphBlock,
              contentWidth,
              floatingZones,
              cumulativeY,
              true
            );
            if (result.ok) return result.extent;
            recordFallback(block, result.reason);
            return synthesizeParagraphExtent(block as ParagraphBlock, contentWidth);
          }

          case 'table':
            return measureTableBlock(block as TableBlock, contentWidth, measure);

          case 'textBox':
            return measureTextBoxBlock(block as TextBoxBlock, (paragraph, width) => {
              const result = measure(paragraph, width);
              if (result.kind !== 'paragraph') {
                throw new Error(
                  'Rust textBox paragraph measurement returned a non-paragraph extent'
                );
              }
              return result;
            });

          case 'image': {
            const imageBlock = block as ImageBlock;
            return {
              kind: 'image',
              width: imageBlock.rotationBounds?.width ?? imageBlock.width ?? 100,
              height: imageBlock.rotationBounds?.height ?? imageBlock.height ?? 100,
            };
          }

          case 'chart':
            return {
              kind: 'chart',
              width: block.width ?? 320,
              height: block.height ?? 220,
            };

          case 'pageBreak':
            return { kind: 'pageBreak' };

          case 'columnBreak':
            return { kind: 'columnBreak' };

          case 'sectionBreak':
            return { kind: 'sectionBreak' };

          case 'shape':
            return measureShapeBlock(block as ShapeBlock, (paragraph, width) => {
              const result = measure(paragraph, width);
              if (result.kind !== 'paragraph') {
                throw new Error('Rust shape paragraph measurement returned a non-paragraph extent');
              }
              return result;
            });

          default:
            // Exhaustiveness guard — this switch is the sole per-block
            // measurer; see LayoutBlock in core/layout/pagination/types.ts.
            assertExhaustiveLayoutBlock(block, 'rustMeasureSource createMeasureBlock');
        }
      };
      return measure;
    },

    getDocumentFontChains(blocks: LayoutBlock[]): Record<string, number[]> {
      const merged: Record<string, number[]> = {};
      forEachParagraphBlock(blocks, (pBlock) => {
        // Family faces first — the same requests createMeasureBlock assembles.
        // Skip a combo whose chain has not resolved yet (those runs stay on
        // the char-distributed draw path until prepareFonts settles); include
        // what IS cached.
        const requests = new Map<string, FontRequest>();
        collectParagraphFontRequests(pBlock, requests);
        const resolvedKeys: string[] = [];
        for (const [chainKey, req] of requests) {
          const chain = registry.getCachedFontIdChain(req.family, req.bold, req.italic);
          if (chain === undefined || chain.length === 0) continue;
          const target = merged[chainKey] ?? (merged[chainKey] = []);
          for (const id of chain) if (!target.includes(id)) target.push(id);
          resolvedKeys.push(chainKey);
        }

        // Append the shared script-fallback ids AFTER the family faces of every
        // chain this block uses (family faces keep supplying metrics/Latin
        // glyphs; script faces only cover what they cannot). Unresolved script
        // set → leave the family chains as-is (a later pass appends them).
        const blockScripts = collectBlockScripts(pBlock, hanHint);
        if (blockScripts.length > 0) {
          const scriptIds = registry.getCachedScriptFallbackIds(blockScripts);
          if (scriptIds !== undefined && scriptIds.length > 0) {
            for (const chainKey of resolvedKeys) {
              const target = merged[chainKey];
              for (const id of scriptIds) if (!target.includes(id)) target.push(id);
            }
          }
        }
      });
      return merged;
    },

    clear(): void {
      registry.clear();
      memo.clear();
      pendingRequests.clear();
      pendingScripts.clear();
    },

    getStats(): RustMeasureStats {
      return { ...stats, syntheticBlockKinds: { ...stats.syntheticBlockKinds } };
    },
  };
}
