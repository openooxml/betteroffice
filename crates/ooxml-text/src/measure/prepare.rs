//! Run preparation: validate each run, resolve its font chain, segment it
//! into same-font subranges, shape each subrange, and fold the shaped
//! advances back onto the run's *original* characters with UTF-16 offsets.
//!
//! Everything downstream (the line filler) works on the per-character
//! advance tables built here, so all UTF-8 ↔ UTF-16 and shaped-cluster ↔
//! character conversion happens in exactly one place.

use crate::font_store::{FontId, FontStore};
use crate::line_break::break_opportunities;
use crate::shape::{ShapeDirection, ShapeFeature, shape_with_direction, shape_with_properties};
use crate::word_metrics::{kern_enabled, kern_features};

use super::input::{MeasureInput, RunIn, validate_pt_size};
use super::{MAX_RUN_TEXT_BYTES, MeasureError, pt_to_px};

/// One indivisible shaped cluster of a run. A cluster may cover several
/// source characters (ligature or combining sequence), so every downstream
/// wrap/hit boundary is cluster-safe by construction.
#[derive(Debug, Clone, Copy)]
pub(super) struct CharAdv {
    /// UTF-16 code-unit offset of this cluster in the original text.
    pub utf16_offset: u32,
    /// UTF-16 code units covered by the complete shaped cluster.
    pub utf16_len: u32,
    pub advance: f32,
    pub level: u8,
    pub logical_order: u32,
    pub font_size_pt: f32,
    pub metrics_font: FontId,
    pub baseline_shift_px: f32,
}

/// A text run ready for line filling.
#[derive(Debug, Clone)]
pub(super) struct PreparedText {
    pub chars: Vec<CharAdv>,
    /// Total UTF-16 length of the original text.
    pub utf16_len: u32,
    /// Character indices (into `chars`) where a line may break — i.e. the
    /// next line would start at that character. UAX-14 opportunities,
    /// excluding the trivial end-of-text one.
    pub breaks: Vec<usize>,
    pub letter_spacing: f32,
    /// Points, for line-metrics bookkeeping (`updateMaxFont`).
    pub font_size_pt: f32,
    /// Font whose tables drive this run's contribution to line metrics —
    /// the head of its fallback chain, like TS's `getFontMetrics` on the
    /// primary family.
    pub metrics_font: FontId,
    pub baseline_shift_px: f32,
}

/// A tab run: width is resolved against the tab-stop grid at fill time
/// (depends on the line's current x), but its font already contributes to
/// line metrics like TS's `updateMaxFont(runToFontStyle(run))`.
#[derive(Debug, Clone, Copy)]
pub(super) struct PreparedTab {
    /// Points, for line-metrics bookkeeping.
    pub font_size_pt: f32,
    /// Head of the run's fallback chain (metrics source).
    pub metrics_font: FontId,
    pub bidi_level: u8,
}

/// A field run, measured at its cached display (fallback) text with the
/// run's formatting — the live value is swapped in at paint time (TS field
/// branch of `measureParagraph`).
#[derive(Debug, Clone, Copy)]
pub(super) struct PreparedField {
    pub width: f32,
    /// Points, for line-metrics bookkeeping.
    pub font_size_pt: f32,
    /// Head of the run's fallback chain (metrics source).
    pub metrics_font: FontId,
    pub baseline_shift_px: f32,
    pub bidi_level: u8,
}

/// An image footprint. For an inline image the width joins the line advance
/// and the (column-fitted) height plus wrap distances grows the line box in
/// `finalize_line`; for a block / `topAndBottom` (own-line) image the height
/// plus wrap distances is the whole line box (no column fit) and the width is
/// only consulted by `measureInlineWidthAfterTab` (TS parity).
#[derive(Debug, Clone, Copy)]
pub(super) struct PreparedImage {
    pub width: f32,
    pub height: f32,
    pub dist_top: f32,
    pub dist_bottom: f32,
    pub bidi_level: u8,
}

#[derive(Debug, Clone)]
pub(super) enum PreparedRun {
    Text(PreparedText),
    LineBreak,
    Tab(PreparedTab),
    Field(PreparedField),
    InlineImage(PreparedImage),
    /// A block / `topAndBottom` image (`wrapType === 'topAndBottom'` or
    /// `displayMode === 'block'`): it takes its own line, its declared height
    /// plus wrap distances IS the line box (no `max-width` column fit — TS
    /// assigns `maxImageHeightPx` directly), and it contributes no width to
    /// the line advance. Its declared width still counts toward
    /// `measureInlineWidthAfterTab` like every other image run.
    OwnLineImage(PreparedImage),
    /// A truly floating image (anchored + float/wrap mode): absolutely
    /// positioned, so it only advances the tail span — except its declared
    /// width still counts toward `measureInlineWidthAfterTab` (TS sums
    /// `next.width || 0` for *any* image run after a tab).
    SkippedImage {
        width: f32,
        bidi_level: u8,
    },
    /// Hidden text retains its logical span but has no layout footprint.
    Hidden {
        utf16_len: u32,
    },
}

/// Advance scale for synthesized small caps (w:smallCaps): lowercase
/// characters shape as their uppercase glyph with advances scaled by this
/// factor.
///
/// The painter styles smallCaps runs with CSS `font-variant: small-caps`
/// (`layout/paint/renderParagraph/runs.ts`); with no `smcp` feature in the
/// font, browsers synthesize scaled-down uppercase. The factor is
/// **0.7** — Chromium's and WebKit's synthesized-small-caps font size
/// multiplier — chosen because this engine measures what the painter
/// paints and the paint targets are Blink/WebKit first. Known divergences,
/// pinned by test: Gecko synthesizes at 0.8, Word renders small caps at
/// roughly 0.8 of the full size, and the TS measurer ignores smallCaps
/// entirely (it measures the original lowercase glyphs at full size).
const LEGACY_SMALL_CAPS_ADVANCE_SCALE: f32 = 0.7;
const WORD_SMALL_CAPS_ADVANCE_SCALE: f32 = 0.8;

/// Characters that UAX-14 treats as mandatory breaks (plus tab, which DOCX
/// represents as a `TabRun`). Their appearance inside a text run means the
/// block needs the reference engine.
fn is_disallowed_control(c: char) -> bool {
    matches!(
        c,
        '\n' | '\r' | '\t' | '\u{000b}' | '\u{000c}' | '\u{0085}' | '\u{2028}' | '\u{2029}'
    )
}

/// Per-char UBA embedding level for `text` under `base`, in logical order.
/// Levels drive shaping segmentation and direction (a shaped segment must be
/// a single directional run); line
/// breaking, spans, and widths all stay in logical order.
fn char_levels(text: &str, base: crate::bidi::BaseDirection) -> Vec<u8> {
    if text.is_empty() {
        return Vec::new();
    }
    let paras = crate::bidi::bidi_paragraphs(text, base);
    let mut runs = paras.iter().flat_map(|p| p.runs.iter());
    let mut cur = runs.next();
    let mut levels = Vec::new();
    for (byte, _) in text.char_indices() {
        while let Some(r) = cur {
            if byte < r.end {
                break;
            }
            cur = runs.next();
        }
        levels.push(cur.map_or(0, |r| r.level));
    }
    levels
}

fn shape_direction(level: u8) -> ShapeDirection {
    if crate::bidi::level_is_rtl(level) {
        ShapeDirection::Rtl
    } else {
        ShapeDirection::Ltr
    }
}

/// UBA base direction for a run's text: Word's `w:bidi` paragraph property
/// and the run-level `w:rtl` flag both force an RTL base; otherwise LTR
/// (a Word paragraph is LTR unless `bidi` is set — never first-strong).
/// The base only affects neutral characters' levels, i.e. segmentation,
/// never advance sums.
fn base_direction(run_rtl: bool, input: &MeasureInput) -> crate::bidi::BaseDirection {
    if run_rtl || input.block.attrs.as_ref().is_some_and(|a| a.bidi) {
        crate::bidi::BaseDirection::Rtl
    } else {
        crate::bidi::BaseDirection::Ltr
    }
}

pub(super) fn validate_chain(store: &FontStore, chain: &[FontId]) -> Result<(), MeasureError> {
    for &id in chain {
        store.metrics(id).map_err(|_| {
            MeasureError::Invalid(format!("unknown font id {} in chain", id.to_u32()))
        })?;
    }
    Ok(())
}

pub(super) fn prepare_runs(
    store: &FontStore,
    input: &MeasureInput,
) -> Result<Vec<PreparedRun>, MeasureError> {
    let bidi_levels = paragraph_run_levels(input);
    let mut prepared = Vec::with_capacity(input.block.runs.len());
    for (run_index, run) in input.block.runs.iter().enumerate() {
        let levels = bidi_levels.get(run_index).map(Vec::as_slice).unwrap_or(&[]);
        let object_level = levels.first().copied().unwrap_or(0);
        match run.kind.as_str() {
            "lineBreak" => prepared.push(PreparedRun::LineBreak),
            "text" if run.hidden => prepared.push(PreparedRun::Hidden {
                utf16_len: run
                    .text
                    .as_deref()
                    .unwrap_or("")
                    .chars()
                    .map(|c| c.len_utf16() as u32)
                    .sum(),
            }),
            "text" => prepared.push(PreparedRun::Text(prepare_text_run(
                store, input, run, levels,
            )?)),
            "tab" => prepared.push(PreparedRun::Tab(prepare_tab_run(
                store,
                input,
                run,
                object_level,
            )?)),
            "image" => prepared.push(prepare_image_run(run, object_level)?),
            "field" => prepared.push(PreparedRun::Field(prepare_field_run(
                store,
                input,
                run,
                object_level,
            )?)),
            other => {
                return Err(MeasureError::Unsupported(format!("run kind {other:?}")));
            }
        }
    }
    Ok(prepared)
}

fn paragraph_run_levels(input: &MeasureInput) -> Vec<Vec<u8>> {
    let mut combined = String::new();
    let mut spans = Vec::with_capacity(input.block.runs.len());
    let mut char_cursor = 0usize;
    for run in &input.block.runs {
        let start = char_cursor;
        if run.kind == "text" {
            let text = run.text.as_deref().unwrap_or("");
            combined.push_str(text);
            char_cursor += text.chars().count();
        } else {
            // UBA object replacement preserves neutral inline objects in the
            // surrounding paragraph context instead of defaulting level 0.
            combined.push('\u{fffc}');
            char_cursor += 1;
        }
        spans.push((start, char_cursor));
    }
    let base = if input.block.attrs.as_ref().is_some_and(|attrs| attrs.bidi) {
        crate::bidi::BaseDirection::Rtl
    } else {
        crate::bidi::BaseDirection::Ltr
    };
    let combined_levels = char_levels(&combined, base);
    input
        .block
        .runs
        .iter()
        .zip(spans)
        .map(|(run, (start, end))| {
            if run.kind == "text" && run.rtl {
                char_levels(
                    run.text.as_deref().unwrap_or(""),
                    crate::bidi::BaseDirection::Rtl,
                )
            } else {
                combined_levels.get(start..end).unwrap_or(&[]).to_vec()
            }
        })
        .collect()
}

/// Resolve a tab run's font for line metrics, mirroring the TS tab branch
/// (`runToFontStyle` → `updateMaxFont`): family/size default straight to the
/// document defaults, bold/italic pick the chain.
fn prepare_tab_run(
    store: &FontStore,
    input: &MeasureInput,
    run: &RunIn,
    bidi_level: u8,
) -> Result<PreparedTab, MeasureError> {
    let font_size_pt = run.font_size.unwrap_or(input.defaults.font_size);
    validate_pt_size(font_size_pt, "run.fontSize")?;
    let family = run
        .font_family
        .as_deref()
        .unwrap_or(&input.defaults.font_family);
    let chain = input.chain_for(family, run.bold, run.italic)?;
    validate_chain(store, &chain)?;
    Ok(PreparedTab {
        font_size_pt,
        metrics_font: chain[0],
        bidi_level,
    })
}

/// Security clamp for file-derived image dimensions: finite, non-negative,
/// bounded (never fed raw into line math).
fn validate_image_dim(v: f32, name: &str) -> Result<f32, MeasureError> {
    if v.is_finite() && (0.0..=100_000.0).contains(&v) {
        Ok(v)
    } else {
        Err(MeasureError::Unsupported(format!("{name} out of range")))
    }
}

/// Wrap distances may be authored negative in theory; clamp magnitude only.
fn validate_image_dist(v: f32, name: &str) -> Result<f32, MeasureError> {
    if v.is_finite() && v.abs() <= 100_000.0 {
        Ok(v)
    } else {
        Err(MeasureError::Unsupported(format!("{name} out of range")))
    }
}

/// Mirror of the TS image branch order in `measureParagraph`:
/// 1. anchored + floating (`position` set and `displayMode == "float"` or a
///    text-wrapping wrap type) → skipped (absolutely positioned);
/// 2. `topAndBottom` / block display → own-line image (its own line, height =
///    declared height + wrap distances, TS default 6px; no width advance);
/// 3. everything else measures inline.
///
/// Missing dimensions are never a bail: a dimensionless image is treated as
/// zero-size (`run.width || 0` / `run.height || 0`), matching TS's use of the
/// required-by-type `run.width`/`run.height` when they resolve to 0 — a
/// zero-size image contributes no width and no line growth but stays
/// representable.
fn prepare_image_run(run: &RunIn, bidi_level: u8) -> Result<PreparedRun, MeasureError> {
    let wrap = run.wrap_type.as_deref();
    let display = run.display_mode.as_deref();
    // TS `wrapsAroundText`: square / tight / through.
    let is_floating =
        display == Some("float") || matches!(wrap, Some("square" | "tight" | "through"));
    if run.position.is_some() && is_floating {
        // TS sums `next.width || 0` after tabs — validate when present, else 0.
        let width = validate_image_dim(
            run.rotation_bounds
                .as_ref()
                .and_then(|b| b.width)
                .or(run.width)
                .unwrap_or(0.0),
            "run.rotationBounds.width",
        )?;
        return Ok(PreparedRun::SkippedImage { width, bidi_level });
    }

    // Block / topAndBottom own-line image. TS defaults its wrap distances to
    // 6px (`run.distTop ?? 6`), unlike the inline path's 0px default.
    if wrap == Some("topAndBottom") || display == Some("block") {
        return Ok(PreparedRun::OwnLineImage(PreparedImage {
            width: validate_image_dim(
                run.rotation_bounds
                    .as_ref()
                    .and_then(|b| b.width)
                    .or(run.width)
                    .unwrap_or(0.0),
                "run.rotationBounds.width",
            )?,
            height: validate_image_dim(
                run.rotation_bounds
                    .as_ref()
                    .and_then(|b| b.height)
                    .or(run.height)
                    .unwrap_or(0.0),
                "run.rotationBounds.height",
            )?,
            dist_top: validate_image_dist(run.dist_top.unwrap_or(6.0), "run.distTop")?,
            dist_bottom: validate_image_dist(run.dist_bottom.unwrap_or(6.0), "run.distBottom")?,
            bidi_level,
        }));
    }

    Ok(PreparedRun::InlineImage(PreparedImage {
        width: validate_image_dim(
            run.rotation_bounds
                .as_ref()
                .and_then(|b| b.width)
                .or(run.width)
                .unwrap_or(0.0),
            "run.rotationBounds.width",
        )?,
        height: validate_image_dim(
            run.rotation_bounds
                .as_ref()
                .and_then(|b| b.height)
                .or(run.height)
                .unwrap_or(0.0),
            "run.rotationBounds.height",
        )?,
        dist_top: validate_image_dist(run.dist_top.unwrap_or(0.0), "run.distTop")?,
        dist_bottom: validate_image_dist(run.dist_bottom.unwrap_or(0.0), "run.distBottom")?,
        bidi_level,
    }))
}

/// Mirror of the TS field branch (`measureParagraph`): measure
/// `fallback || "1"` with the run's family/size/bold/italic. The TS field
/// style carries neither `letterSpacing` nor caps flags, so neither applies
/// here.
fn prepare_field_run(
    store: &FontStore,
    input: &MeasureInput,
    run: &RunIn,
    bidi_level: u8,
) -> Result<PreparedField, MeasureError> {
    let base_size_pt = run.font_size.unwrap_or(input.defaults.font_size);
    validate_pt_size(base_size_pt, "run.fontSize")?;
    let (font_size_pt, baseline_shift_px) = script_metrics(base_size_pt, run);
    let family = run
        .font_family
        .as_deref()
        .unwrap_or(&input.defaults.font_family);
    let chain = input.chain_for(family, run.bold, run.italic)?;
    validate_chain(store, &chain)?;
    // TS `run.fallback || '1'`: absent and empty both measure as "1"
    let fallback = match run.fallback.as_deref() {
        Some(t) if !t.is_empty() => t,
        _ => "1",
    };
    let width = measure_plain_text(
        store,
        &chain,
        fallback,
        pt_to_px(font_size_pt),
        base_direction(run.rtl, input),
    )?;
    Ok(PreparedField {
        width,
        font_size_pt,
        metrics_font: chain[0],
        baseline_shift_px,
        bidi_level,
    })
}

fn script_metrics(base_size_pt: f32, run: &RunIn) -> (f32, f32) {
    let base_px = pt_to_px(base_size_pt);
    if run.superscript {
        (base_size_pt * 0.75, base_px * 0.4)
    } else if run.subscript {
        (base_size_pt * 0.75, -base_px * 0.2)
    } else {
        (base_size_pt, 0.0)
    }
}

/// Width of `text` shaped plain — no caps, no letter spacing, no horizontal
/// scale — through `chain`: the Rust analogue of the TS canvas
/// `measureTextWidth(text, style)` calls whose style carries only
/// family/size/bold/italic (field fallbacks, list markers). Bidi-aware:
/// segments split at UBA level-run boundaries under `base`.
/// Resolve `ch` to the first covering font in `chain`; if no font covers it,
/// substitute the chain's terminal font. The host (`TextMeasureFontRegistry`)
/// guarantees every chain ends in an always-available broad-coverage
/// last-resort face, so a truly-uncovered char shapes as that face's `.notdef`
/// box glyph — a real advance width — instead of routing the whole block to
/// browser measurement. Returns `None` only for an empty chain, which
/// `validate_chain` rejects upstream.
fn resolve_with_fallback(store: &FontStore, chain: &[FontId], ch: char) -> Option<FontId> {
    store.resolve(chain, ch).or_else(|| chain.last().copied())
}

pub(super) fn measure_plain_text(
    store: &FontStore,
    chain: &[FontId],
    text: &str,
    size_px: f32,
    base: crate::bidi::BaseDirection,
) -> Result<f32, MeasureError> {
    if text.len() > MAX_RUN_TEXT_BYTES {
        return Err(MeasureError::Unsupported(format!(
            "text too long ({} bytes)",
            text.len()
        )));
    }
    if let Some(c) = text.chars().find(|&c| is_disallowed_control(c)) {
        return Err(MeasureError::Unsupported(format!(
            "control character U+{:04X} in text",
            c as u32
        )));
    }

    let levels = char_levels(text, base);
    let mut plan: Vec<(char, FontId, u8)> = Vec::new();
    for (i, ch) in text.chars().enumerate() {
        let Some(font) = resolve_with_fallback(store, chain, ch) else {
            return Err(MeasureError::Unsupported("empty font chain".to_string()));
        };
        plan.push((ch, font, levels[i]));
    }

    // shape maximal same-font single-level subranges and sum advances
    let mut width = 0.0f32;
    let mut seg_start = 0;
    while seg_start < plan.len() {
        let (_, font, level) = plan[seg_start];
        let mut seg_end = seg_start + 1;
        while seg_end < plan.len() && plan[seg_end].1 == font && plan[seg_end].2 == level {
            seg_end += 1;
        }
        let seg_text: String = plan[seg_start..seg_end].iter().map(|&(c, ..)| c).collect();
        let glyphs =
            shape_with_direction(store, font, &seg_text, size_px, &[], shape_direction(level))
                .map_err(|e| MeasureError::Invalid(format!("shaping failed: {e}")))?;
        width += glyphs.iter().map(|g| g.x_advance).sum::<f32>();
        seg_start = seg_end;
    }
    Ok(width)
}

fn prepare_text_run(
    store: &FontStore,
    input: &MeasureInput,
    run: &RunIn,
    resolved_levels: &[u8],
) -> Result<PreparedText, MeasureError> {
    let text = run.text.as_deref().unwrap_or("");
    if text.len() > MAX_RUN_TEXT_BYTES {
        return Err(MeasureError::Unsupported(format!(
            "run text too long ({} bytes)",
            text.len()
        )));
    }
    if let Some(c) = text.chars().find(|&c| is_disallowed_control(c)) {
        return Err(MeasureError::Unsupported(format!(
            "control character U+{:04X} in run text",
            c as u32
        )));
    }

    let default_size_pt = run.font_size.unwrap_or(input.defaults.font_size);
    validate_pt_size(default_size_pt, "run.fontSize")?;
    if let Some(size) = run.font_size_cs {
        validate_pt_size(size, "run.fontSizeCs")?;
    }

    let letter_spacing = run.letter_spacing.unwrap_or(0.0);
    if !(letter_spacing.is_finite() && letter_spacing.abs() <= 1000.0) {
        return Err(MeasureError::Unsupported(
            "letterSpacing out of range".to_string(),
        ));
    }

    // w:w is 1–600 percent; refuse anything outside rather than scale by it.
    let scale = match run.horizontal_scale {
        None => 1.0,
        Some(pct) => {
            if !(pct.is_finite() && pct > 0.0 && pct <= 600.0) {
                return Err(MeasureError::Unsupported(
                    "horizontalScale out of range".to_string(),
                ));
            }
            pct / 100.0
        }
    };

    if let Some(threshold) = run.kerning_min_pt
        && !(threshold.is_finite() && (0.0..=1638.0).contains(&threshold))
    {
        return Err(MeasureError::Unsupported(
            "kerningMinPt out of range".to_string(),
        ));
    }

    // Per-character plan. Font slot, complex-script weight/size and language
    // are resolved before shaping; this lets one OOXML run legitimately use
    // Latin, East Asian and complex-script faces without splitting its source
    // positions or losing cross-character shaping inside a same-style slice.
    struct PlanChar {
        source_index: usize,
        utf16_offset: u32,
        utf16_len: u32,
        shaped: ShapedChars,
        font: FontId,
        metrics_font: FontId,
        font_size_pt: f32,
        baseline_shift_px: f32,
        /// 1.0, or `SMALL_CAPS_ADVANCE_SCALE` for a synthesized small cap.
        advance_scale: f32,
        /// UBA embedding level (odd = RTL) — segmentation only.
        level: u8,
        language: Option<String>,
        features: Vec<ShapeFeature>,
    }
    enum ShapedChars {
        One(char),
        // w:caps expansions like ß → SS
        Many(Vec<char>),
    }

    #[derive(Clone, Copy, PartialEq, Eq)]
    enum FontSlot {
        Ascii,
        HAnsi,
        EastAsia,
        Cs,
    }

    fn is_complex(ch: char) -> bool {
        matches!(ch as u32,
            0x0590..=0x08ff | 0x0900..=0x0dff | 0xfb1d..=0xfdff | 0xfe70..=0xfeff)
    }
    fn is_east_asian(ch: char) -> bool {
        matches!(ch as u32,
            0x1100..=0x11ff | 0x2e80..=0x30ff | 0x3130..=0x318f |
            0x31a0..=0x31ff | 0x3400..=0x4dbf | 0x4e00..=0x9fff |
            0xa960..=0xa97f | 0xac00..=0xd7ff | 0xf900..=0xfaff |
            0xff00..=0xffef)
    }
    fn is_combining(ch: char) -> bool {
        matches!(ch as u32, 0x0300..=0x036f | 0x1ab0..=0x1aff | 0x1dc0..=0x1dff | 0x20d0..=0x20ff | 0xfe20..=0xfe2f)
    }
    fn family_for_slot<'a>(run: &'a RunIn, slot: FontSlot, fallback: &'a str) -> &'a str {
        let slots = run.font_slots.as_ref();
        let selected = match slot {
            FontSlot::Ascii => slots.and_then(|s| s.ascii.as_deref()),
            FontSlot::HAnsi => slots.and_then(|s| s.h_ansi.as_deref()),
            FontSlot::EastAsia => slots.and_then(|s| s.east_asia.as_deref()),
            FontSlot::Cs => slots.and_then(|s| s.cs.as_deref()),
        };
        selected
            .or_else(|| slots.and_then(|s| s.h_ansi.as_deref()))
            .or_else(|| slots.and_then(|s| s.ascii.as_deref()))
            .or(run.font_family.as_deref())
            .unwrap_or(fallback)
    }
    fn language_for_slot(run: &RunIn, slot: FontSlot) -> Option<&str> {
        let language = run.language.as_ref()?;
        match slot {
            FontSlot::EastAsia => language.east_asia.as_deref().or(language.latin.as_deref()),
            FontSlot::Cs => language.bidi.as_deref().or(language.latin.as_deref()),
            FontSlot::Ascii | FontSlot::HAnsi => language.latin.as_deref(),
        }
    }
    fn uppercase_for_language(ch: char, language: Option<&str>) -> Vec<char> {
        let lang = language.unwrap_or("").to_ascii_lowercase();
        if lang.starts_with("tr") || lang.starts_with("az") {
            match ch {
                'i' => return vec!['İ'],
                'ı' => return vec!['I'],
                _ => {}
            }
        }
        ch.to_uppercase().collect()
    }
    fn supports_smcp(store: &FontStore, font: FontId, ch: char, size_px: f32, level: u8) -> bool {
        let text = ch.to_string();
        let plain = shape_with_direction(store, font, &text, size_px, &[], shape_direction(level));
        let featured = shape_with_direction(
            store,
            font,
            &text,
            size_px,
            &[ShapeFeature {
                tag: *b"smcp",
                value: 1,
            }],
            shape_direction(level),
        );
        matches!((plain, featured), (Ok(a), Ok(b)) if a.iter().map(|g| g.glyph_id).collect::<Vec<_>>() != b.iter().map(|g| g.glyph_id).collect::<Vec<_>>())
    }

    let owned_levels;
    let levels = if resolved_levels.len() == text.chars().count() {
        resolved_levels
    } else {
        owned_levels = char_levels(text, base_direction(run.rtl, input));
        &owned_levels
    };
    let mut plan: Vec<PlanChar> = Vec::new();
    let mut utf16_offset: u32 = 0;
    let mut previous_slot = FontSlot::HAnsi;
    for (char_index, ch) in text.chars().enumerate() {
        let slot = if run.complex_script || is_complex(ch) {
            FontSlot::Cs
        } else if is_east_asian(ch) {
            FontSlot::EastAsia
        } else if is_combining(ch) {
            previous_slot
        } else if ch.is_ascii() {
            FontSlot::Ascii
        } else if run.font_slots.as_ref().and_then(|s| s.hint.as_deref()) == Some("eastAsia") {
            FontSlot::EastAsia
        } else {
            FontSlot::HAnsi
        };
        previous_slot = slot;
        let language = language_for_slot(run, slot);
        let (mut shaped, mut advance_scale): (Vec<char>, f32) = if run.all_caps {
            (uppercase_for_language(ch, language), 1.0)
        } else if run.small_caps && ch.is_lowercase() {
            (
                uppercase_for_language(ch, language),
                if input.authoritative_shaping {
                    WORD_SMALL_CAPS_ADVANCE_SCALE
                } else {
                    LEGACY_SMALL_CAPS_ADVANCE_SCALE
                },
            )
        } else {
            (vec![ch], 1.0)
        };
        let complex = slot == FontSlot::Cs;
        let bold = if complex {
            run.bold_cs.unwrap_or(run.bold)
        } else {
            run.bold
        };
        let italic = if complex {
            run.italic_cs.unwrap_or(run.italic)
        } else {
            run.italic
        };
        let base_size_pt = if complex {
            run.font_size_cs.unwrap_or(default_size_pt)
        } else {
            default_size_pt
        };
        let (font_size_pt, baseline_shift_px) = script_metrics(base_size_pt, run);
        let family = family_for_slot(run, slot, &input.defaults.font_family);
        let chain = input.chain_for(family, bold, italic)?;
        validate_chain(store, &chain)?;
        let first = shaped[0];
        let Some(mut font) = resolve_with_fallback(store, &chain, first) else {
            return Err(MeasureError::Unsupported("empty font chain".to_string()));
        };
        let mut features = run.kerning_min_pt.map_or_else(Vec::new, |threshold| {
            kern_features(kern_enabled(
                (font_size_pt * 2.0).round() as u32,
                (threshold * 2.0).round() as u32,
            ))
        });
        if input.authoritative_shaping
            && run.small_caps
            && !run.all_caps
            && ch.is_lowercase()
            && let Some(original_font) = resolve_with_fallback(store, &chain, ch)
            && supports_smcp(
                store,
                original_font,
                ch,
                pt_to_px(font_size_pt),
                levels[char_index],
            )
        {
            // A real small-cap glyph is already designed at the correct
            // advance/ink size. Only synthesize uppercase at 0.8 when the
            // selected face has no smcp substitution.
            shaped = vec![ch];
            advance_scale = 1.0;
            font = original_font;
            features.push(ShapeFeature {
                tag: *b"smcp",
                value: 1,
            });
        }
        let first = shaped[0];
        // Chars a resolved font doesn't cover (e.g. a rare char in a multi-char
        // uppercase expansion) shape as that font's `.notdef` — a real box-glyph
        // advance — rather than routing the block to browser measurement.
        let utf16_len = ch.len_utf16() as u32;
        plan.push(PlanChar {
            source_index: char_index,
            utf16_offset,
            utf16_len,
            shaped: if shaped.len() == 1 {
                ShapedChars::One(first)
            } else {
                ShapedChars::Many(shaped)
            },
            font,
            metrics_font: chain[0],
            font_size_pt,
            baseline_shift_px,
            advance_scale,
            level: levels[char_index],
            language: language.map(str::to_owned),
            features,
        });
        utf16_offset += utf16_len;
    }
    let utf16_total = utf16_offset;

    // Shape maximal same-font/style/language single-level subranges so
    // kerning/ligatures inside a subrange survive (TS's substring-
    // measurement rationale), then fold glyph advances back onto original
    // characters via the shaped-byte → original-char map. rustybuzz clusters
    // are byte offsets into the text handed to `shape`. Splitting at a
    // small-caps scale boundary mirrors browser synthesis (separate font
    // run, no kerning across it); splitting at a UBA level boundary keeps
    // every shaped segment a single directional run so the resolved UBA
    // direction can be passed to rustybuzz explicitly.
    let mut chars: Vec<CharAdv> = Vec::new();
    let mut seg_start = 0;
    while seg_start < plan.len() {
        let first_plan = &plan[seg_start];
        let font = first_plan.font;
        let seg_scale = first_plan.advance_scale;
        let seg_level = first_plan.level;
        let mut seg_end = seg_start + 1;
        while seg_end < plan.len()
            && plan[seg_end].font == font
            && plan[seg_end].advance_scale == seg_scale
            && plan[seg_end].level == seg_level
            && plan[seg_end].font_size_pt == first_plan.font_size_pt
            && plan[seg_end].baseline_shift_px == first_plan.baseline_shift_px
            && plan[seg_end].language == first_plan.language
            && plan[seg_end].features == first_plan.features
        {
            seg_end += 1;
        }

        let mut shaped_text = String::new();
        // (byte offset in shaped_text, original char index) per shaped char
        let mut byte_to_char: Vec<(usize, usize)> = Vec::new();
        for (char_index, pc) in plan.iter().enumerate().take(seg_end).skip(seg_start) {
            match &pc.shaped {
                ShapedChars::One(c) => {
                    byte_to_char.push((shaped_text.len(), char_index));
                    shaped_text.push(*c);
                }
                ShapedChars::Many(cs) => {
                    for &c in cs {
                        byte_to_char.push((shaped_text.len(), char_index));
                        shaped_text.push(c);
                    }
                }
            }
        }

        let glyphs = shape_with_properties(
            store,
            font,
            &shaped_text,
            pt_to_px(first_plan.font_size_pt),
            &first_plan.features,
            shape_direction(seg_level),
            first_plan.language.as_deref(),
        )
        .map_err(|e| MeasureError::Invalid(format!("shaping failed: {e}")))?;
        let mut glyph_clusters: Vec<(usize, f32)> = Vec::new();
        for g in &glyphs {
            let cluster = g.cluster as usize;
            // last mapped shaped-char at or before this cluster byte
            let slot = byte_to_char.partition_point(|&(byte, _)| byte <= cluster);
            let Some(&(_, char_index)) = slot.checked_sub(1).and_then(|i| byte_to_char.get(i))
            else {
                continue;
            };
            if let Some((_, advance)) = glyph_clusters
                .iter_mut()
                .find(|(start, _)| *start == char_index)
            {
                *advance += g.x_advance;
            } else {
                glyph_clusters.push((char_index, g.x_advance));
            }
        }
        glyph_clusters.sort_by_key(|&(start, _)| start);
        if glyph_clusters
            .first()
            .is_none_or(|&(start, _)| start != seg_start)
        {
            glyph_clusters.insert(0, (seg_start, 0.0));
        }
        for cluster_index in 0..glyph_clusters.len() {
            let (start, advance) = glyph_clusters[cluster_index];
            let end = glyph_clusters
                .get(cluster_index + 1)
                .map_or(seg_end, |&(next, _)| next);
            let pc = &plan[start];
            let utf16_len = plan[start..end].iter().map(|c| c.utf16_len).sum();
            chars.push(CharAdv {
                utf16_offset: pc.utf16_offset,
                utf16_len,
                advance: advance * pc.advance_scale * scale,
                level: pc.level,
                logical_order: pc.source_index as u32,
                font_size_pt: pc.font_size_pt,
                metrics_font: pc.metrics_font,
                baseline_shift_px: pc.baseline_shift_px,
            });
        }

        seg_start = seg_end;
    }

    // UAX-14 opportunities on the original text, translated from byte
    // offsets to char indices. The final mandatory opportunity at
    // `text.len()` is the trivial end-of-text one — the filler already ends
    // words at the run boundary, so it is dropped here. Interior mandatory
    // breaks cannot occur (control characters were refused above).
    let byte_offsets: Vec<usize> = text.char_indices().map(|(b, _)| b).collect();
    let mut breaks: Vec<usize> = Vec::new();
    if !text.is_empty() {
        for bo in break_opportunities(text) {
            if bo.byte_index == 0 || bo.byte_index >= text.len() {
                continue;
            }
            // opportunities are always char boundaries (line_break.rs)
            if let Ok(char_index) = byte_offsets.binary_search(&bo.byte_index)
                && let Some(cluster_index) = chars
                    .iter()
                    .position(|cluster| cluster.logical_order as usize == char_index)
            {
                breaks.push(cluster_index);
            }
        }
    }

    let first_font_size = chars.first().map_or(default_size_pt, |c| c.font_size_pt);
    let first_metrics_font = chars.first().map_or_else(
        || {
            let family = run
                .font_family
                .as_deref()
                .unwrap_or(&input.defaults.font_family);
            input
                .chain_for(family, run.bold, run.italic)
                .map(|chain| chain[0])
                .unwrap_or(FontId(0))
        },
        |c| c.metrics_font,
    );
    let first_baseline_shift = chars.first().map_or(0.0, |c| c.baseline_shift_px);
    Ok(PreparedText {
        chars,
        utf16_len: utf16_total,
        breaks,
        letter_spacing,
        font_size_pt: first_font_size,
        metrics_font: first_metrics_font,
        baseline_shift_px: first_baseline_shift,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::font_store::FontStore;

    const FIXTURE: &[u8] = include_bytes!("../../tests/fonts/LiberationSans-Regular.ttf");

    fn input_with(runs: serde_json::Value) -> MeasureInput {
        serde_json::from_value(serde_json::json!({
            "block": { "kind": "paragraph", "runs": runs },
            "maxWidth": 1000.0,
            "fontChains": { "liberation sans|0|0": [0] },
            "defaults": { "fontSize": 12.0, "fontFamily": "Liberation Sans" }
        }))
        .unwrap()
    }

    /// Offsets are UTF-16 code units, not UTF-8 bytes: 'é' is 2 bytes but
    /// 1 unit. (Non-BMP offsets are covered by the line-filler tests — the
    /// fixture font is BMP-only, so an emoji cannot reach this builder.)
    #[test]
    fn utf16_offsets_count_code_units_not_bytes() {
        let mut store = FontStore::new();
        store.register(FIXTURE.to_vec()).unwrap();
        // 'é' is 2 UTF-8 bytes but 1 UTF-16 unit: byte offsets [0, 2, 3],
        // UTF-16 offsets must be [0, 1, 2].
        let input = input_with(serde_json::json!([{ "kind": "text", "text": "éab" }]));
        let prepared = prepare_runs(&store, &input).unwrap();
        let PreparedRun::Text(t) = &prepared[0] else {
            panic!("expected text run");
        };
        let offsets: Vec<u32> = t.chars.iter().map(|c| c.utf16_offset).collect();
        assert_eq!(offsets, vec![0, 1, 2]);
        assert_eq!(t.utf16_len, 3);
    }

    /// An uncovered non-BMP char (fixture cmap is format 4/6, BMP-only)
    /// refuses measurement instead of guessing a width.
    #[test]
    fn uncovered_char_falls_back_to_the_terminal_font_notdef() {
        // A char no font in the chain covers (emoji, with a BMP-only fixture)
        // must NOT bail to browser measurement — it shapes as the chain's
        // terminal font's `.notdef`, so preparation succeeds with a real width.
        let mut store = FontStore::new();
        store.register(FIXTURE.to_vec()).unwrap();
        let input = input_with(serde_json::json!([{ "kind": "text", "text": "a😀b" }]));
        let prepared = prepare_runs(&store, &input).expect("uncovered char no longer bails");
        assert!(!prepared.is_empty());
    }

    #[test]
    fn authoritative_clusters_do_not_split_combining_sequences_or_ligatures() {
        let mut store = FontStore::new();
        store.register(FIXTURE.to_vec()).unwrap();
        let mut input = input_with(serde_json::json!([{
            "kind": "text",
            "text": "e\u{301} ffi",
            "letterSpacing": 2.0
        }]));
        input.authoritative_shaping = true;
        let prepared = prepare_runs(&store, &input).unwrap();
        let PreparedRun::Text(text) = &prepared[0] else {
            panic!("expected text run");
        };
        assert_eq!(text.chars[0].utf16_offset, 0);
        assert_eq!(text.chars[0].utf16_len, 2, "combining mark stays with base");
        assert!(
            text.chars.iter().any(|cluster| cluster.utf16_len > 1),
            "fixture exposes at least one multi-character shaped cluster"
        );
    }
}
