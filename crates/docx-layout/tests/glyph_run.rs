//! GlyphRun emission gates: with a populated measurement `FontStore` the
//! display-list builder shapes text runs into `GlyphRunPrimitive`s (real glyph
//! ids + pen positions from the same bytes that measured); with no resolvable
//! fonts it falls back to `TextRunPrimitive`, byte-identical to the pre-shaping
//! path. Also covers font-fallback splitting, justification stretch in the
//! glyph x, hit-testing / range-rects over GlyphRuns, and determinism.
//!
//! Fixture: the same vendored Liberation Sans Regular (SIL OFL 1.1) the
//! ooxml-text suites use.

use docx_layout::display_list::{
    DisplayList, GlyphRunPrimitive, Primitive, build_display_list_json,
    build_display_list_json_with_fonts,
};
use docx_layout::hit::{hit_test, range_rects};
use ooxml_text::{FontStore, shape};

const LIBERATION: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const FRAG_X: f64 = 96.0;
const FRAG_WIDTH: f64 = 624.0;

fn store_with_liberation() -> FontStore {
    let mut s = FontStore::new();
    s.register(LIBERATION.to_vec()).expect("fixture registers");
    s
}

/// Build a one-paragraph, one-line display-list input. `chain` is the
/// `fontChains` value for `"liberation sans|0|0"` (None ⇒ no fontChains at all,
/// the browser-measured shape). `trailing_break` appends a `<w:br>` run so a
/// last line still justifies. Text runs are 1 PM position per char starting at
/// pm 1.
fn build_input(
    text: &str,
    line_width: f64,
    alignment: Option<&str>,
    trailing_break: bool,
    chain: Option<&[u32]>,
) -> String {
    let char_count = text.chars().count() as i64;
    let pm_start = 1i64;
    let pm_end_run = pm_start + char_count;

    let mut runs = vec![serde_json::json!({
        "kind": "text",
        "text": text,
        "pmStart": pm_start,
        "pmEnd": pm_end_run,
        "fontFamily": "Liberation Sans",
        "fontSize": 12.0
    })];
    if trailing_break {
        runs.push(serde_json::json!({ "kind": "lineBreak", "pmStart": pm_end_run }));
    }

    let mut block = serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": runs,
        "pmStart": pm_start,
        "pmEnd": pm_end_run + 1
    });
    if let Some(a) = alignment {
        block["attrs"] = serde_json::json!({ "alignment": a });
    }

    let mut input = serde_json::json!({
        "measured": [{
            "block": block,
            "measure": {
                "kind": "paragraph",
                "totalHeight": 20.0,
                "lines": [{
                    "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": char_count,
                    "width": line_width, "ascent": 16.0, "descent": 4.0, "lineHeight": 20.0
                }]
            }
        }],
        "options": {},
        "layout": {
            "pages": [{
                "size": {"w": 816.0, "h": 1056.0},
                "margins": {"top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0},
                "number": 1,
                "fragments": [{
                    "kind": "paragraph",
                    "blockId": 0,
                    "x": FRAG_X, "y": 96.0, "width": FRAG_WIDTH, "height": 20.0,
                    "fromLine": 0, "toLine": 1,
                    "pmStart": pm_start, "pmEnd": pm_end_run + 1
                }]
            }]
        }
    });
    if let Some(ids) = chain {
        input["fontChains"] = serde_json::json!({ "liberation sans|0|0": ids });
    }
    input.to_string()
}

fn build_rtl_mixed_input() -> String {
    let pm_start = 1i64;
    let hebrew = "אבג";
    let latin = " abc";
    let hebrew_chars = hebrew.chars().count() as i64;
    let latin_chars = latin.chars().count() as i64;
    let hebrew_end = pm_start + hebrew_chars;
    let pm_end = hebrew_end + latin_chars;

    serde_json::json!({
        "measured": [{
            "block": {
                "kind": "paragraph",
                "id": 0,
                "runs": [
                    {
                        "kind": "text",
                        "text": hebrew,
                        "pmStart": pm_start,
                        "pmEnd": hebrew_end,
                        "fontFamily": "Liberation Sans",
                        "fontSize": 12.0
                    },
                    {
                        "kind": "text",
                        "text": latin,
                        "pmStart": hebrew_end,
                        "pmEnd": pm_end,
                        "fontFamily": "Liberation Sans",
                        "fontSize": 12.0
                    }
                ],
                "attrs": { "bidi": true },
                "pmStart": pm_start,
                "pmEnd": pm_end
            },
            "measure": {
                "kind": "paragraph",
                "totalHeight": 20.0,
                "lines": [{
                    "headRun": 0, "headChar": 0, "tailRun": 1, "tailChar": latin_chars,
                    "width": 160.0, "ascent": 16.0, "descent": 4.0, "lineHeight": 20.0
                }]
            }
        }],
        "fontChains": { "liberation sans|0|0": [0] },
        "options": {},
        "layout": {
            "pages": [{
                "size": {"w": 816.0, "h": 1056.0},
                "margins": {"top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0},
                "number": 1,
                "fragments": [{
                    "kind": "paragraph",
                    "blockId": 0,
                    "x": FRAG_X, "y": 96.0, "width": FRAG_WIDTH, "height": 20.0,
                    "fromLine": 0, "toLine": 1,
                    "pmStart": pm_start, "pmEnd": pm_end
                }]
            }]
        }
    })
    .to_string()
}

fn glyph_runs(dl: &DisplayList) -> Vec<&GlyphRunPrimitive> {
    dl.pages[0]
        .primitives
        .iter()
        .filter_map(|p| match p {
            Primitive::GlyphRun(g) => Some(g),
            _ => None,
        })
        .collect()
}

fn max_glyph_x(g: &GlyphRunPrimitive) -> f64 {
    g.glyphs.iter().map(|gl| gl.x).fold(f64::MIN, f64::max)
}

fn min_glyph_x(g: &GlyphRunPrimitive) -> f64 {
    g.glyphs.iter().map(|gl| gl.x).fold(f64::MAX, f64::min)
}

#[test]
fn emits_glyph_runs_when_fonts_resolve() {
    let store = store_with_liberation();
    let input = build_input("Hello world", 60.0, None, false, Some(&[0]));
    let json = build_display_list_json_with_fonts(&input, &store).expect("builds");

    // wire shape: camelCase keys + the `glyphRun` kind tag the canvas agent's
    // TS type expects
    assert!(json.contains(r#""kind":"glyphRun""#), "kind tag: {json}");
    assert!(json.contains(r#""fontId":0"#), "camelCase fontId: {json}");
    assert!(
        !json.contains(r#""wordSpacing""#),
        "non-justified line omits wordSpacing: {json}"
    );

    let dl: DisplayList = serde_json::from_str(&json).unwrap();
    let runs = glyph_runs(&dl);
    assert_eq!(runs.len(), 1, "single-font run ⇒ one GlyphRun");
    let g = runs[0];

    assert_eq!(g.text, "Hello world", "source text preserved");
    assert_eq!(g.font_id, 0);
    assert_eq!(g.size, 16.0, "12pt ⇒ 16px");
    assert_eq!(g.attrs.doc_start, Some(1), "doc positions carried");
    assert_eq!(g.attrs.doc_end, Some(12));
    assert!(!g.glyphs.is_empty(), "non-empty glyphs");

    // LTR glyph pen origins strictly increase
    for w in g.glyphs.windows(2) {
        assert!(w[1].x > w[0].x, "glyph x must increase: {:?}", g.glyphs);
    }
    // clusters are byte indices into the run text
    for gl in &g.glyphs {
        assert!(
            (gl.cluster as usize) < g.text.len(),
            "cluster {} out of range for {:?}",
            gl.cluster,
            g.text
        );
    }
    // no TextRunPrimitive left behind for the shaped run
    assert!(
        !json.contains(r#""kind":"text""#),
        "shaped run must not also emit a text primitive: {json}"
    );
}

#[test]
fn falls_back_to_text_run_without_fonts() {
    // fontChains present, but the store is empty ⇒ every run fails to shape and
    // takes the v0 TextRunPrimitive path
    let empty = FontStore::new();
    let with_chains = build_input("Hello world", 60.0, None, false, Some(&[0]));
    let out_fallback = build_display_list_json_with_fonts(&with_chains, &empty).expect("builds");
    assert!(
        !out_fallback.contains("glyphRun"),
        "no GlyphRun without fonts: {out_fallback}"
    );
    assert!(
        out_fallback.contains(r#""kind":"text""#),
        "TextRunPrimitive emitted"
    );

    // byte-identical to the pre-GlyphRun path (no fontChains at all)
    let no_chains = build_input("Hello world", 60.0, None, false, None);
    let out_today = build_display_list_json(&no_chains).expect("builds");
    assert_eq!(
        out_fallback, out_today,
        "fallback must be byte-identical to the browser-measured path"
    );
}

#[test]
fn unresolved_family_falls_back_per_run() {
    // fonts registered, but the run's family has no chain entry ⇒ that run falls
    // back to TextRunPrimitive (gate is per-run, not all-or-nothing)
    let store = store_with_liberation();
    // chain keyed for a DIFFERENT family than the run's "Liberation Sans"
    let mut input: serde_json::Value =
        serde_json::from_str(&build_input("Hello", 40.0, None, false, None)).unwrap();
    input["fontChains"] = serde_json::json!({ "some other font|0|0": [0] });
    let json = build_display_list_json_with_fonts(&input.to_string(), &store).expect("builds");
    assert!(
        !json.contains("glyphRun"),
        "unresolved family ⇒ no GlyphRun: {json}"
    );
    assert!(json.contains(r#""kind":"text""#));
}

#[test]
fn justified_line_glyph_x_reaches_usable_width() {
    let mut store = FontStore::new();
    let id = store.register(LIBERATION.to_vec()).unwrap();

    let text = "Hello world foo bar";
    // shape the exact run to learn its natural advance, then pin the measured
    // line width to it so the justification slack is exact
    let glyphs = shape(&store, id, text, 16.0, &[]).unwrap();
    let shaped_width: f64 = glyphs.iter().map(|g| g.x_advance as f64).sum();
    let last_adv = glyphs.last().unwrap().x_advance as f64;
    let usable_width = FRAG_WIDTH; // no indents

    // trailing <w:br> makes this (last) line justify
    let input = build_input(text, shaped_width, Some("justify"), true, Some(&[0]));
    let json = build_display_list_json_with_fonts(&input, &store).expect("builds");
    assert!(
        json.contains(r#""wordSpacing""#),
        "justified line carries wordSpacing: {json}"
    );

    let dl: DisplayList = serde_json::from_str(&json).unwrap();
    let g = glyph_runs(&dl)[0];
    let max_x = max_glyph_x(g);

    // the rightmost glyph origin lands one trailing advance short of the usable
    // right edge — i.e. the stretched glyphs reach the margin
    let right_edge = FRAG_X + usable_width;
    assert!(
        (max_x - (right_edge - last_adv)).abs() < 1.0,
        "justified rightmost glyph x {max_x} should be ~{} (edge {right_edge} - last adv {last_adv})",
        right_edge - last_adv
    );

    // and it is far past where the same run sits unjustified
    let natural = build_input(text, shaped_width, None, false, Some(&[0]));
    let ndl: DisplayList =
        serde_json::from_str(&build_display_list_json_with_fonts(&natural, &store).unwrap())
            .unwrap();
    let n_max = max_glyph_x(glyph_runs(&ndl)[0]);
    assert!(
        max_x > n_max + 100.0,
        "justification must push glyphs right: {max_x} vs natural {n_max}"
    );
}

#[test]
fn splits_a_run_across_fallback_fonts() {
    let mut store = FontStore::new();
    store.register(LIBERATION.to_vec()).unwrap(); // id 0
    store.register(LIBERATION.to_vec()).unwrap(); // id 1 — stands in for a fallback face

    // 'あ' is uncovered by Liberation, so the chain's terminal font (id 1) takes
    // it while the Latin chars resolve to id 0 ⇒ three same-font subranges
    let input = build_input("a\u{3042}b", 30.0, None, false, Some(&[0, 1]));
    let dl: DisplayList =
        serde_json::from_str(&build_display_list_json_with_fonts(&input, &store).unwrap()).unwrap();
    let runs = glyph_runs(&dl);
    assert_eq!(runs.len(), 3, "one GlyphRun per same-font subrange");

    assert_eq!((runs[0].text.as_str(), runs[0].font_id), ("a", 0));
    assert_eq!((runs[1].text.as_str(), runs[1].font_id), ("\u{3042}", 1));
    assert_eq!((runs[2].text.as_str(), runs[2].font_id), ("b", 0));

    // per-subrange doc positions split by char offset; the last closes on the
    // segment's pm_end
    assert_eq!(
        (runs[0].attrs.doc_start, runs[0].attrs.doc_end),
        (Some(1), Some(2))
    );
    assert_eq!(
        (runs[1].attrs.doc_start, runs[1].attrs.doc_end),
        (Some(2), Some(3))
    );
    assert_eq!(
        (runs[2].attrs.doc_start, runs[2].attrs.doc_end),
        (Some(3), Some(4))
    );

    // the pen keeps flowing across subranges (each cluster is byte 0 of its own
    // single-char text, so positions come from the accumulated advance)
    assert!(runs[1].glyphs[0].x > runs[0].glyphs[0].x);
    assert!(runs[2].glyphs[0].x > runs[1].glyphs[0].x);
    for g in &runs {
        assert_eq!(g.glyphs[0].cluster, 0, "single-char subrange cluster is 0");
    }
}

#[test]
fn rtl_paragraph_reorders_mixed_direction_glyph_runs_visually() {
    let store = store_with_liberation();
    let input = build_rtl_mixed_input();
    let json_a = build_display_list_json_with_fonts(&input, &store).expect("builds");
    let json_b = build_display_list_json_with_fonts(&input, &store).expect("builds");
    assert_eq!(json_a, json_b, "RTL display-list output is deterministic");

    let dl: DisplayList = serde_json::from_str(&json_a).unwrap();
    let runs = glyph_runs(&dl);
    assert_eq!(
        runs.len(),
        3,
        "Hebrew, neutral space, and Latin should split into UBA level runs"
    );

    let hebrew = runs.iter().find(|g| g.text == "אבג").expect("Hebrew run");
    let space = runs.iter().find(|g| g.text == " ").expect("space run");
    let latin = runs.iter().find(|g| g.text == "abc").expect("Latin run");

    assert_eq!(hebrew.rtl, Some(true), "Hebrew level run is RTL");
    assert_eq!(space.rtl, Some(true), "neutral separator follows RTL base");
    assert_eq!(latin.rtl, None, "Latin level run stays LTR");
    assert!(
        min_glyph_x(latin) < min_glyph_x(hebrew),
        "RTL paragraph visual order should place the embedded Latin run left of Hebrew: latin {:?}, hebrew {:?}",
        latin.glyphs,
        hebrew.glyphs
    );

    assert_eq!(
        (hebrew.attrs.doc_start, hebrew.attrs.doc_end),
        (Some(1), Some(4)),
        "Hebrew doc span remains logical"
    );
    assert_eq!(
        (latin.attrs.doc_start, latin.attrs.doc_end),
        (Some(5), Some(8)),
        "Latin doc span remains logical"
    );
    assert_eq!(
        (space.attrs.doc_start, space.attrs.doc_end),
        (Some(4), Some(5)),
        "separator doc span remains logical"
    );
    assert!(
        hebrew.glyphs.first().unwrap().cluster > hebrew.glyphs.last().unwrap().cluster,
        "RTL GlyphRun clusters should be visual-order descending: {:?}",
        hebrew.glyphs
    );
}

#[test]
fn hit_test_and_range_rects_resolve_over_glyph_runs() {
    let store = store_with_liberation();
    let input = build_input("Hello world", 80.0, None, false, Some(&[0]));
    let dl: DisplayList =
        serde_json::from_str(&build_display_list_json_with_fonts(&input, &store).unwrap()).unwrap();
    let g = glyph_runs(&dl)[0];
    let baseline = g.glyphs[0].y;

    // click near the run's left edge lands at/near doc_start
    let left_pos = hit_test(&dl, 0, g.glyphs[0].x + 1.0, baseline).expect("left hit");
    assert!(
        (1..=12).contains(&left_pos),
        "left pos {left_pos} inside the run span"
    );

    // click past the run's right edge resolves further along than the left
    let right_pos = hit_test(&dl, 0, max_glyph_x(g) + 2.0, baseline).expect("right hit");
    assert!(
        right_pos > left_pos,
        "right click resolves further: {right_pos} vs {left_pos}"
    );

    // whole-run range yields one rect covering the run's x-extent
    let rects = range_rects(&dl, 1, 12);
    assert_eq!(rects.len(), 1);
    let r = &rects[0];
    assert!(r.width > 0.0, "range rect has width");
    assert!(
        (r.x - g.glyphs[0].x).abs() < 2.0,
        "range rect starts at the run's left edge: {} vs {}",
        r.x,
        g.glyphs[0].x
    );
}

#[test]
fn glyph_run_extent_uses_real_trailing_advance() {
    // F3: the run's right edge is the trailing glyph's `x + advance` (the real
    // shaped extent), NOT the old uniform estimate `span * n/(n-1)` that drifted
    // ~3px on mixed-font lines. Pin both the per-glyph advance contract and the
    // range-rect width the hit geometry derives from it.
    let store = store_with_liberation();
    let text = "Hello world"; // 11 chars ⇒ doc span [1, 12)
    let input = build_input(text, 80.0, None, false, Some(&[0]));
    let dl: DisplayList =
        serde_json::from_str(&build_display_list_json_with_fonts(&input, &store).unwrap()).unwrap();
    let g = glyph_runs(&dl)[0];

    // every glyph carries a positive pen advance, and `x + advance` chains to the
    // next glyph's origin (the pen flows continuously across the run)
    for w in g.glyphs.windows(2) {
        assert!(w[0].advance > 0.0, "glyph advance is positive: {:?}", w[0]);
        assert!(
            (w[0].x + w[0].advance - w[1].x).abs() < 0.01,
            "x+advance chains to the next origin: {} + {} vs {}",
            w[0].x,
            w[0].advance,
            w[1].x
        );
    }

    // expected extent = the shaped natural advance sum (what the DOM painter lays
    // out); reshape the same run in an isolated store to learn it
    let mut probe = FontStore::new();
    let id = probe.register(LIBERATION.to_vec()).unwrap();
    let shaped = shape(&probe, id, text, 16.0, &[]).unwrap();
    let shaped_width: f64 = shaped.iter().map(|gl| gl.x_advance as f64).sum();

    let min_x = g.glyphs.iter().map(|gl| gl.x).fold(f64::INFINITY, f64::min);
    let right = g
        .glyphs
        .iter()
        .map(|gl| gl.x + gl.advance)
        .fold(f64::MIN, f64::max);
    let extent = right - min_x;
    assert!(
        (extent - shaped_width).abs() < 0.5,
        "real glyph extent {extent} matches the shaped width {shaped_width}"
    );

    // a whole-run range rect exposes the same extent through the hit geometry
    let rects = range_rects(&dl, 1, 12);
    assert_eq!(rects.len(), 1);
    let r = &rects[0];
    assert!(
        (r.x - min_x).abs() < 0.01,
        "range rect left {} at the run's left edge {}",
        r.x,
        min_x
    );
    assert!(
        (r.width - shaped_width).abs() < 0.5,
        "range-rect width {} matches the shaped extent {} (F3: real trailing advance)",
        r.width,
        shaped_width
    );

    // and the discarded uniform estimate (span·n/(n-1)) was measurably off — the
    // real extent is strictly closer to the shaped truth, which is what closed
    // the ~3px right-edge drift
    let max_origin = g.glyphs.iter().map(|gl| gl.x).fold(f64::MIN, f64::max);
    let n = g.glyphs.len() as f64;
    let old_estimate = (max_origin - min_x) * n / (n - 1.0);
    assert!(
        (r.width - shaped_width).abs() < (old_estimate - shaped_width).abs(),
        "real extent {} is closer to shaped {} than the old estimate {}",
        r.width,
        shaped_width,
        old_estimate
    );
}

#[test]
fn glyph_run_build_is_deterministic() {
    let store = store_with_liberation();
    let input = build_input("Hello world foo", 90.0, Some("justify"), true, Some(&[0]));
    let a = build_display_list_json_with_fonts(&input, &store).unwrap();
    let b = build_display_list_json_with_fonts(&input, &store).unwrap();
    assert_eq!(a, b, "same input ⇒ identical JSON");
}

#[test]
fn wasm_entry_threads_measure_fonts() {
    // the production wasm export shapes from the module-global measurement store
    docx_layout::clear_measure_fonts();
    let id = docx_layout::register_measure_font(LIBERATION).expect("registers");
    assert_eq!(id, 0);

    let input = build_input("Hi", 20.0, None, false, Some(&[0]));
    let json = docx_layout::build_display_list_json(&input).expect("builds");
    assert!(
        json.contains(r#""kind":"glyphRun""#),
        "wasm entry shapes via MEASURE_FONTS: {json}"
    );

    docx_layout::clear_measure_fonts();
}

/// The glyph run carries the resolved CSS face for the canvas fillText safety
/// net — the same shorthand the browser-measured TextRunPrimitive would use —
/// so a glyph-outline failure degrades to the measured family/weight/style
/// instead of generic sans-serif.
#[test]
fn glyph_runs_carry_the_resolved_fallback_font() {
    let store = store_with_liberation();
    let input = build_input("Hello world", 60.0, None, false, Some(&[0]));
    let json = build_display_list_json_with_fonts(&input, &store).expect("builds");

    let dl: DisplayList = serde_json::from_str(&json).unwrap();
    let runs = glyph_runs(&dl);
    assert_eq!(runs.len(), 1);
    assert_eq!(
        runs[0].attrs.fallback_font.as_deref(),
        Some("400 16px Liberation Sans, sans-serif"),
        "fallbackFont must carry the resolved css shorthand: {json}"
    );
    assert!(json.contains(r#""fallbackFont":"400 16px Liberation Sans, sans-serif""#));
}
