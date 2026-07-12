use ooxml_text::{FontStore, measure_paragraph_json};
use serde_json::{Value, json};

const FIXTURE: &[u8] = include_bytes!("fonts/LiberationSans-Regular.ttf");
const NOTO_NASKH_ARABIC: &[u8] = include_bytes!("fonts/NotoNaskhArabic-Regular.ttf");

const W0: f64 = 1139.0 / 128.0;
const SP: f64 = 569.0 / 128.0;
const WA: f64 = 1366.0 / 128.0;
const ASC: f64 = 14.484375;
const DESC: f64 = 3.390625;
const LEAD: f64 = 0.5234375;
const LH: f64 = 17.875 + LEAD;

fn store() -> FontStore {
    let mut s = FontStore::new();
    s.register(FIXTURE.to_vec()).expect("fixture registers");
    s
}

fn arabic_store() -> FontStore {
    let mut s = FontStore::new();
    s.register(NOTO_NASKH_ARABIC.to_vec())
        .expect("Arabic fixture registers");
    s
}

/// Measure runs at `max_width` with the standard single-font chain.
fn measure(runs: Value, max_width: f64) -> Result<Value, String> {
    measure_with(json!({ "kind": "paragraph", "runs": runs }), max_width)
}

fn measure_with(block: Value, max_width: f64) -> Result<Value, String> {
    let input = json!({
        "block": block,
        "maxWidth": max_width,
        "fontChains": { "liberation sans|0|0": [0] },
        "defaults": { "fontSize": 12.0, "fontFamily": "Liberation Sans" }
    });
    let out = measure_paragraph_json(&store(), &input.to_string())?;
    Ok(serde_json::from_str(&out).expect("output is valid JSON"))
}

fn measure_arabic(runs: Value, max_width: f64) -> Result<Value, String> {
    let input = json!({
        "block": { "kind": "paragraph", "runs": runs, "attrs": { "bidi": true } },
        "maxWidth": max_width,
        "fontChains": { "noto naskh arabic|0|0": [0] },
        "defaults": { "fontSize": 12.0, "fontFamily": "Noto Naskh Arabic" }
    });
    let out = measure_paragraph_json(&arabic_store(), &input.to_string())?;
    Ok(serde_json::from_str(&out).expect("output is valid JSON"))
}

fn approx(actual: f64, expected: f64, what: &str) {
    assert!(
        (actual - expected).abs() < 1e-3,
        "{what}: expected {expected}, got {actual}"
    );
}

fn spans(v: &Value) -> Vec<(u64, u64, u64, u64)> {
    v["lines"]
        .as_array()
        .unwrap()
        .iter()
        .map(|l| {
            (
                l["headRun"].as_u64().unwrap(),
                l["headChar"].as_u64().unwrap(),
                l["tailRun"].as_u64().unwrap(),
                l["tailChar"].as_u64().unwrap(),
            )
        })
        .collect()
}

// 1. single line fits: one TypesetRow covering the whole run
#[test]
fn single_line_fits() {
    let v = measure(json!([{ "kind": "text", "text": "0 0 0" }]), 200.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 5)]);
    let line = &v["lines"][0];
    approx(
        line["width"].as_f64().unwrap(),
        3.0 * W0 + 2.0 * SP,
        "width",
    );
    approx(line["ascent"].as_f64().unwrap(), ASC, "ascent");
    approx(line["descent"].as_f64().unwrap(), DESC, "descent");
    approx(line["lineHeight"].as_f64().unwrap(), LH, "lineHeight");
    approx(v["totalHeight"].as_f64().unwrap(), LH, "totalHeight");
}

#[test]
fn wrap_at_space_keeps_trailing_space_in_line_width() {
    let v = measure(json!([{ "kind": "text", "text": "00 00" }]), 30.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 3), (0, 3, 0, 5)]);
    let lines = v["lines"].as_array().unwrap();
    approx(
        lines[0]["width"].as_f64().unwrap(),
        2.0 * W0 + SP,
        "line 1 width includes the trailing space",
    );
    approx(
        lines[1]["width"].as_f64().unwrap(),
        2.0 * W0,
        "line 2 width",
    );
    approx(v["totalHeight"].as_f64().unwrap(), 2.0 * LH, "totalHeight");
}

// 3. overlong unbreakable word hard-breaks mid-word, minimum 1 char/line
#[test]
fn overlong_word_hard_breaks() {
    // 3 zeros (26.7px) fit in 30px; 4 (35.6px) do not
    let v = measure(json!([{ "kind": "text", "text": "0000000000" }]), 30.0).unwrap();
    assert_eq!(
        spans(&v),
        vec![(0, 0, 0, 3), (0, 3, 0, 6), (0, 6, 0, 9), (0, 9, 0, 10)]
    );
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        3.0 * W0,
        "chunk width",
    );

    // nothing fits: one forced char per line
    let v = measure(json!([{ "kind": "text", "text": "000" }]), 5.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 1), (0, 1, 0, 2), (0, 2, 0, 3)]);
}

// 4. soft return (LineBreakRun) forces a new line
#[test]
fn soft_return_forces_new_line() {
    let v = measure(
        json!([
            { "kind": "text", "text": "0" },
            { "kind": "lineBreak" },
            { "kind": "text", "text": "0" }
        ]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 1, 0), (2, 0, 2, 1)]);

    let v = measure(
        json!([{ "kind": "text", "text": "0" }, { "kind": "lineBreak" }]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 1, 0), (2, 0, 2, 0)]);
    let last = &v["lines"][1];
    approx(
        last["ascent"].as_f64().unwrap(),
        16.0 * 0.8,
        "fallback ascent",
    );
    approx(
        last["descent"].as_f64().unwrap(),
        16.0 * 0.2,
        "fallback descent",
    );
    approx(
        last["lineHeight"].as_f64().unwrap(),
        16.0 * 1.15,
        "fallback lineHeight",
    );
}

// 5. multi-run line: metrics follow the largest font on the line
#[test]
fn multi_run_line_takes_max_font_basis() {
    let mut s = FontStore::new();
    s.register(FIXTURE.to_vec()).unwrap();
    s.register(FIXTURE.to_vec()).unwrap();
    let input = json!({
        "block": { "kind": "paragraph", "runs": [
            { "kind": "text", "text": "0" },
            { "kind": "text", "text": "0", "bold": true, "fontSize": 24.0 }
        ]},
        "maxWidth": 200.0,
        "fontChains": {
            "liberation sans|0|0": [0],
            "liberation sans|1|0": [1]
        },
        "defaults": { "fontSize": 12.0, "fontFamily": "Liberation Sans" }
    });
    let out = measure_paragraph_json(&s, &input.to_string()).unwrap();
    let v: Value = serde_json::from_str(&out).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 1, 1)]);
    let line = &v["lines"][0];
    approx(line["width"].as_f64().unwrap(), W0 + 2.0 * W0, "width");
    approx(line["ascent"].as_f64().unwrap(), 2.0 * ASC, "24pt ascent");
    approx(
        line["descent"].as_f64().unwrap(),
        2.0 * DESC,
        "24pt descent",
    );
    approx(line["lineHeight"].as_f64().unwrap(), 2.0 * LH, "24pt line");
}

#[test]
fn line_rules_match_typography_semantics() {
    let with_spacing = |spacing: Value| {
        measure_with(
            json!({
                "kind": "paragraph",
                "runs": [{ "kind": "text", "text": "0" }],
                "attrs": { "spacing": spacing }
            }),
            200.0,
        )
        .unwrap()["lines"][0]["lineHeight"]
            .as_f64()
            .unwrap()
    };

    approx(
        with_spacing(json!({ "line": 20.0, "lineRule": "exact" })),
        20.0,
        "exact",
    );
    approx(
        with_spacing(json!({ "line": 10.0, "lineRule": "atLeast" })),
        LH,
        "atLeast below natural height keeps the natural height",
    );
    approx(
        with_spacing(json!({ "line": 50.0, "lineRule": "atLeast" })),
        50.0,
        "atLeast above natural height wins",
    );
    approx(
        with_spacing(json!({ "line": 2.0, "lineUnit": "multiplier" })),
        2.0 * LH,
        "multiplier scales the single-line basis",
    );
    approx(
        with_spacing(json!({ "line": 30.0, "lineUnit": "px" })),
        30.0,
        "px",
    );

    // no spacing at all: single spacing off the OS/2 win metrics
    let v = measure(json!([{ "kind": "text", "text": "0" }]), 200.0).unwrap();
    approx(v["lines"][0]["lineHeight"].as_f64().unwrap(), LH, "default");

    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "text", "text": "0" }],
            "attrs": { "spacing": { "before": 10.0, "after": 5.0 } }
        }),
        200.0,
    )
    .unwrap();
    approx(
        v["totalHeight"].as_f64().unwrap(),
        LH + 15.0,
        "before/after",
    );
}

// 7. empty paragraph: one line with the single-line floor
#[test]
fn empty_paragraph_floor_behavior() {
    // Liberation's leading-inclusive single line is 18.3984375 < 16 × 1.15 = 18.4,
    // so the floor must (barely) win:
    let v = measure(json!([]), 200.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 0)]);
    let line = &v["lines"][0];
    assert_eq!(line["width"].as_f64().unwrap(), 0.0);
    approx(line["ascent"].as_f64().unwrap(), ASC, "ascent");
    approx(line["descent"].as_f64().unwrap(), DESC, "descent");
    approx(line["lineHeight"].as_f64().unwrap(), 16.0 * 1.15, "floored");
    approx(
        v["totalHeight"].as_f64().unwrap(),
        16.0 * 1.15,
        "totalHeight",
    );

    // no floor under an exact rule
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [],
            "attrs": { "spacing": { "line": 10.0, "lineRule": "exact" } }
        }),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["lineHeight"].as_f64().unwrap(),
        10.0,
        "exact empty",
    );

    // a single whitespace-only run measures like an empty paragraph
    let v = measure(json!([{ "kind": "text", "text": "   " }]), 200.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 0)]);
    assert_eq!(v["lines"][0]["width"].as_f64().unwrap(), 0.0);
    approx(
        v["lines"][0]["lineHeight"].as_f64().unwrap(),
        16.0 * 1.15,
        "ws",
    );

    // spacing before/after still applies to empty paragraphs
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [],
            "attrs": { "spacing": { "before": 10.0, "after": 5.0 } }
        }),
        200.0,
    )
    .unwrap();
    approx(
        v["totalHeight"].as_f64().unwrap(),
        16.0 * 1.15 + 15.0,
        "empty +sp",
    );

    // the zero-height anchor variant
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [],
            "attrs": { "suppressEmptyParagraphHeight": true }
        }),
        200.0,
    )
    .unwrap();
    assert_eq!(v["totalHeight"].as_f64().unwrap(), 0.0);
    assert_eq!(v["lines"][0]["lineHeight"].as_f64().unwrap(), 0.0);
}

// 8. output char indices are UTF-16 code units, not UTF-8 bytes: 'é' is two
// UTF-8 bytes but one UTF-16 unit. (Non-BMP/surrogate safety is proven on
// the line filler directly — the fixture font is BMP-only, so a covered
// emoji cannot flow end-to-end; see measure::line_filler tests and test 9.)
#[test]
fn char_indices_are_utf16_not_bytes() {
    // width of "éé" from the pipeline itself (wide measurement)
    let wide = measure(json!([{ "kind": "text", "text": "éé" }]), 1000.0).unwrap();
    let w2 = wide["lines"][0]["width"].as_f64().unwrap();
    assert!(w2 > 10.0, "sanity: éé has real width, got {w2}");

    // fits "éé " but not both words → wrap at the space opportunity.
    // "éé éé" is 5 UTF-16 units (8 UTF-8 bytes); the first line's tail must
    // be 3 — a byte-counting implementation would emit 5.
    let max_width = w2 + SP + w2 / 2.0;
    let v = measure(json!([{ "kind": "text", "text": "éé éé" }]), max_width).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 3), (0, 3, 0, 5)]);
}

// 9. UNSUPPORTED escape hatches
#[test]
fn unsupported_inputs_bail_with_reason() {
    let cases: Vec<(Value, &str)> = vec![
        (
            json!([{ "kind": "tab", "bold": true }]),
            "tab run with no chain for its bold face",
        ),
        (
            json!([{ "kind": "field", "italic": true }]),
            "field run with no chain for its italic face",
        ),
        (
            json!([{ "kind": "text", "text": "a\tb" }]),
            "mandatory-break control char in text run",
        ),
        (json!([{ "kind": "somethingNew" }]), "unknown run kind"),
    ];
    for (runs, what) in cases {
        let err = measure(runs, 200.0).unwrap_err();
        assert!(
            err.starts_with("UNSUPPORTED"),
            "{what}: expected UNSUPPORTED, got {err:?}"
        );
    }

    // uncovered chars (emoji, CJK with a BMP-only chain) no longer bail — they
    // shape as the chain's terminal font's .notdef, so measurement succeeds.
    for runs in [
        json!([{ "kind": "text", "text": "a😀b" }]),
        json!([{ "kind": "text", "text": "中文" }]),
    ] {
        assert!(
            measure(runs, 200.0).is_ok(),
            "uncovered char should fall back to .notdef, not bail"
        );
    }

    // a visible marker resolves its font like the body: an unresolvable
    // marker family refuses instead of guessing a width
    let err = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "text", "text": "abc" }],
            "attrs": { "listMarker": "1.", "listMarkerFontFamily": "Nope" }
        }),
        200.0,
    )
    .unwrap_err();
    assert!(err.starts_with("UNSUPPORTED"), "marker chain: {err:?}");

    // no chain registered for the run's family
    let err = measure(
        json!([{ "kind": "text", "text": "abc", "fontFamily": "Nope" }]),
        200.0,
    )
    .unwrap_err();
    assert!(err.starts_with("UNSUPPORTED"), "missing chain: {err:?}");
}

#[test]
fn json_round_trip_preserves_wire_field_names() {
    let v = measure(json!([{ "kind": "text", "text": "0 0" }]), 200.0).unwrap();

    let mut top: Vec<&str> = v.as_object().unwrap().keys().map(String::as_str).collect();
    top.sort_unstable();
    assert_eq!(top, vec!["kind", "lines", "totalHeight"]);
    assert_eq!(v["kind"], "paragraph");

    let line = v["lines"][0].as_object().unwrap();
    let mut keys: Vec<&str> = line.keys().map(String::as_str).collect();
    keys.sort_unstable();
    assert_eq!(
        keys,
        vec![
            "ascent",
            "descent",
            "headChar",
            "headRun",
            "lineHeight",
            "tailChar",
            "tailRun",
            "width"
        ]
    );
}

#[test]
fn formatting_effects_on_widths() {
    // allCaps: 'a' measures as 'A'
    let v = measure(
        json!([{ "kind": "text", "text": "a", "allCaps": true }]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        WA,
        "allCaps width",
    );

    // horizontalScale 200% doubles the advance
    let v = measure(
        json!([{ "kind": "text", "text": "0", "horizontalScale": 200.0 }]),
        200.0,
    )
    .unwrap();
    approx(v["lines"][0]["width"].as_f64().unwrap(), 2.0 * W0, "scaled");

    // letterSpacing: n-1 gaps within the word
    let v = measure(
        json!([{ "kind": "text", "text": "00", "letterSpacing": 2.0 }]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        2.0 * W0 + 2.0,
        "letterSpacing",
    );
}

// 11. default 48px grid with no custom stops; a mid-line tab spans to the
// next grid line (96px), not a full stride
#[test]
fn tab_advances_to_default_grid_stops() {
    let v = measure(
        json!([{ "kind": "tab" }, { "kind": "text", "text": "0" }]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 1, 1)]);
    approx(v["lines"][0]["width"].as_f64().unwrap(), 48.0 + W0, "tab+0");

    let v = measure(
        json!([
            { "kind": "tab" },
            { "kind": "text", "text": "0" },
            { "kind": "tab" },
            { "kind": "text", "text": "0" }
        ]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 3, 1)]);
    // second tab starts at 48 + W0 = 56.898 and lands on the 96px grid line
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        96.0 + W0,
        "2 tabs",
    );
}

#[test]
fn tab_stop_alignment_semantics() {
    let with_tabs = |val: &str, text: &str| {
        measure_with(
            json!({
                "kind": "paragraph",
                "runs": [{ "kind": "tab" }, { "kind": "text", "text": text }],
                "attrs": { "tabs": [{ "val": val, "pos": 1500.0 }] }
            }),
            300.0,
        )
        .unwrap()["lines"][0]["width"]
            .as_f64()
            .unwrap()
    };

    approx(with_tabs("start", "00"), 100.0 + 2.0 * W0, "start");
    approx(
        with_tabs("end", "00"),
        100.0,
        "end: text right edge on stop",
    );
    approx(
        with_tabs("center", "00"),
        100.0 + W0,
        "center: text centered",
    );
    approx(
        with_tabs("decimal", "00"),
        100.0 + 2.0 * W0,
        "decimal≡start",
    );

    let bar = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "tab" }, { "kind": "text", "text": "0" }],
            "attrs": { "tabs": [{ "val": "bar", "pos": 720.0 }] }
        }),
        300.0,
    )
    .unwrap();
    approx(
        bar["lines"][0]["width"].as_f64().unwrap(),
        W0,
        "bar: width 0",
    );
}

// 13. degenerate stops fall back to the default grid: following text wider
// than an end stop's span, and a cleared grid position is skipped
#[test]
fn tab_falls_back_to_default_grid() {
    // end stop at 48px but following text is ~89px wide → span < 1 → grid
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "tab" }, { "kind": "text", "text": "0000000000" }],
            "attrs": { "tabs": [{ "val": "end", "pos": 720.0 }] }
        }),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        48.0 + 10.0 * W0,
        "give up on stop",
    );

    // val=clear knocks the 720tw grid line out; the tab lands on 1440tw
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "tab" }],
            "attrs": { "tabs": [{ "val": "clear", "pos": 720.0 }] }
        }),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 1)]);
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        96.0,
        "cleared 720",
    );
}

#[test]
fn tab_clamps_to_line_edge_and_wraps_when_full() {
    // end stop at 200px on a 100px line: clamp to 100 − W0
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "tab" }, { "kind": "text", "text": "0" }],
            "attrs": { "tabs": [{ "val": "end", "pos": 3000.0 }] }
        }),
        100.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 1, 1)]);
    approx(v["lines"][0]["width"].as_f64().unwrap(), 100.0, "clamped");

    // 22 zeros fill 195.77px of a 200px line; the tab's 44.23px to the 240px
    // grid line cannot fit or clamp (no room for the following zero), so the
    // tab wraps carrying that width
    let v = measure(
        json!([
            { "kind": "text", "text": "0000000000000000000000" },
            { "kind": "tab" },
            { "kind": "text", "text": "0" }
        ]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 22), (1, 0, 2, 1)]);
    approx(
        v["lines"][1]["width"].as_f64().unwrap(),
        (240.0 - 22.0 * W0) + W0,
        "wrapped tab keeps pre-wrap width",
    );
}

// 15. content-area coordinates: a hanging-indent first line starts left of
// the indent, and the implicit stop at the indent catches the tab
#[test]
fn tab_in_hanging_indent_lands_on_the_body_edge() {
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "tab" }, { "kind": "text", "text": "0" }],
            "attrs": { "indent": { "left": 48.0, "hanging": 24.0 } }
        }),
        200.0,
    )
    .unwrap();
    // first line starts at 24px content-x; the implicit 48px indent stop is
    // 24px away
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        24.0 + W0,
        "tab to indent stop",
    );
}

#[test]
fn tab_font_size_drives_line_metrics() {
    let v = measure(
        json!([{ "kind": "tab", "fontSize": 24.0 }, { "kind": "text", "text": "0" }]),
        200.0,
    )
    .unwrap();
    let line = &v["lines"][0];
    approx(line["ascent"].as_f64().unwrap(), 2.0 * ASC, "24pt ascent");
    approx(line["lineHeight"].as_f64().unwrap(), 2.0 * LH, "24pt line");
}

// ---- field runs ---------------------------------------------------------

#[test]
fn field_measures_at_fallback_text() {
    // '1' and '0' share the 1139-unit digit advance
    let v = measure(json!([{ "kind": "field", "fallback": "00" }]), 200.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 1)]);
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        2.0 * W0,
        "fallback",
    );

    // absent and empty fallback both measure as "1"
    for runs in [
        json!([{ "kind": "field" }]),
        json!([{ "kind": "field", "fallback": "" }]),
    ] {
        let v = measure(runs, 200.0).unwrap();
        approx(v["lines"][0]["width"].as_f64().unwrap(), W0, "default '1'");
    }

    // field font size drives line metrics like any run (updateMaxFont)
    let v = measure(
        json!([{ "kind": "field", "fontSize": 24.0 }, { "kind": "text", "text": "0" }]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["lineHeight"].as_f64().unwrap(),
        2.0 * LH,
        "24pt field line",
    );
}

// 18. a field that doesn't fit a non-empty line wraps whole (one unbreakable
// glyph), and a field after a tab anchors on end stops via followingWidth
#[test]
fn field_wraps_whole_and_anchors_after_tabs() {
    // 22 zeros fill 195.77px of a 200px line; the 2-digit field wraps
    let v = measure(
        json!([
            { "kind": "text", "text": "0000000000000000000000" },
            { "kind": "field", "fallback": "00" }
        ]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 22), (1, 0, 1, 1)]);
    approx(
        v["lines"][1]["width"].as_f64().unwrap(),
        2.0 * W0,
        "wrapped field",
    );

    // TOC pattern: tab to an end stop at 100px, page-number field after —
    // the field's width anchors the tab, closing the line at exactly 100px
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "tab" }, { "kind": "field", "fallback": "00" }],
            "attrs": { "tabs": [{ "val": "end", "pos": 1500.0 }] }
        }),
        300.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        100.0,
        "field anchored on end stop",
    );
}

/// First-line availability probe: lines produced for `"00 00"` (40.0390625px)
/// against `max_width` with the given attrs.
fn marker_lines(attrs: Value, max_width: f64) -> usize {
    measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "text", "text": "00 00" }],
            "attrs": attrs
        }),
        max_width,
    )
    .unwrap()["lines"]
        .as_array()
        .unwrap()
        .len()
}

const TEXT_00_00: f64 = 4.0 * W0 + SP;

// 19. suffix semantics: nothing = natural width, space = + one space glyph,
// tab (default) = grow to the next default-grid stop
#[test]
fn list_marker_suffix_footprints() {
    // nothing: marker "0" costs exactly W0
    let attrs = |suffix: &str| json!({ "listMarker": "0", "listMarkerSuffix": suffix });
    assert_eq!(marker_lines(attrs("nothing"), W0 + TEXT_00_00), 1);
    assert_eq!(marker_lines(attrs("nothing"), W0 + TEXT_00_00 - 1.0), 2);

    // space: + one space advance
    assert_eq!(marker_lines(attrs("space"), W0 + SP + TEXT_00_00), 1);
    assert_eq!(marker_lines(attrs("space"), W0 + SP + TEXT_00_00 - 1.0), 2);

    // default tab suffix: "1." (13.34px) grows to the 48px grid line
    let tab_attrs = json!({ "listMarker": "1." });
    assert_eq!(marker_lines(tab_attrs.clone(), 48.0 + TEXT_00_00), 1);
    assert_eq!(marker_lines(tab_attrs, 48.0 + TEXT_00_00 - 1.0), 2);
}

// 20. tab-suffix stop resolution: a closer custom stop beats the grid, the
// document defaultTabStopTwips drives the grid, and no grid at all falls
// back to natural + half an em
#[test]
fn list_marker_tab_stop_resolution() {
    // custom start stop at 300tw = 20px beats the 48px grid line
    let custom = json!({
        "listMarker": "1.",
        "tabs": [{ "val": "start", "pos": 300.0 }]
    });
    assert_eq!(marker_lines(custom.clone(), 20.0 + TEXT_00_00), 1);
    assert_eq!(marker_lines(custom, 20.0 + TEXT_00_00 - 1.0), 2);

    // A 300tw default interval yields 20px grid stops.
    let grid = json!({ "listMarker": "1.", "defaultTabStopTwips": 300.0 });
    assert_eq!(marker_lines(grid.clone(), 20.0 + TEXT_00_00), 1);
    assert_eq!(marker_lines(grid, 20.0 + TEXT_00_00 - 1.0), 2);

    // defaultTabStop 0 and no custom stops: natural + 0.5em = 13.34375 + 8
    let bare = json!({ "listMarker": "1.", "defaultTabStopTwips": 0.0 });
    let footprint = 13.34375 + 8.0;
    assert_eq!(marker_lines(bare.clone(), footprint + TEXT_00_00), 1);
    assert_eq!(marker_lines(bare, footprint + TEXT_00_00 - 1.0), 2);
}

// 21. marker font size from the numbering level rPr scales the footprint;
// hidden markers and hanging-indent markers cost nothing
#[test]
fn list_marker_font_and_zero_width_paths() {
    // listMarkerFontSize 24pt doubles the marker "0" to 2×W0
    let big = json!({
        "listMarker": "0",
        "listMarkerSuffix": "nothing",
        "listMarkerFontSize": 24.0
    });
    assert_eq!(marker_lines(big.clone(), 2.0 * W0 + TEXT_00_00), 1);
    assert_eq!(marker_lines(big, 2.0 * W0 + TEXT_00_00 - 1.0), 2);

    // Hidden markers have no footprint.
    let hidden = json!({ "listMarker": "00000000", "listMarkerHidden": true });
    assert_eq!(marker_lines(hidden, TEXT_00_00), 1);

    let hanging = json!({
        "listMarker": "1.",
        "listMarkerFontFamily": "Nope",
        "indent": { "hanging": 12.0 }
    });
    measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "text", "text": "abc" }],
            "attrs": hanging
        }),
        200.0,
    )
    .expect("hanging-indent marker skips marker-font resolution");

    // the empty-paragraph path returns before marker resolution too
    measure_with(
        json!({
            "kind": "paragraph",
            "runs": [],
            "attrs": { "listMarker": "1.", "listMarkerFontFamily": "Nope" }
        }),
        200.0,
    )
    .expect("empty paragraph never measures its marker");
}

// 22. an image alone on the line grows it to the image height plus the
// descent buffer on BOTH sides; with text, the image seats on the baseline
// (full height above, text descent below)
#[test]
fn inline_image_grows_the_line_box() {
    // image alone: fallback descent 3.2 buffers both sides
    let v = measure(
        json!([{ "kind": "image", "width": 50.0, "height": 100.0 }]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 1)]);
    let line = &v["lines"][0];
    approx(line["width"].as_f64().unwrap(), 50.0, "image width");
    approx(
        line["lineHeight"].as_f64().unwrap(),
        106.4,
        "alone: h + 2×3.2",
    );
    approx(line["ascent"].as_f64().unwrap(), 103.2, "alone ascent");
    approx(line["descent"].as_f64().unwrap(), 3.2, "alone descent");

    // image flowing with text: baseline-seated, text descent below only
    let v = measure(
        json!([
            { "kind": "text", "text": "0" },
            { "kind": "image", "width": 50.0, "height": 100.0 }
        ]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 1, 1)]);
    let line = &v["lines"][0];
    approx(
        line["width"].as_f64().unwrap(),
        W0 + 50.0,
        "text+image width",
    );
    approx(
        line["lineHeight"].as_f64().unwrap(),
        100.0 + DESC,
        "with text: h + text descent",
    );
    approx(line["ascent"].as_f64().unwrap(), 100.0, "with text ascent");
    approx(line["descent"].as_f64().unwrap(), DESC, "text descent kept");

    // an image shorter than the text line changes nothing
    let v = measure(
        json!([
            { "kind": "text", "text": "0" },
            { "kind": "image", "width": 10.0, "height": 10.0 }
        ]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["lineHeight"].as_f64().unwrap(),
        LH,
        "no growth",
    );

    // wrap distances join the footprint (wp:inline distT/distB)
    let v = measure(
        json!([
            { "kind": "text", "text": "0" },
            { "kind": "image", "width": 20.0, "height": 100.0,
              "distTop": 5.0, "distBottom": 7.0 }
        ]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["lineHeight"].as_f64().unwrap(),
        112.0 + DESC,
        "footprint includes dist",
    );
}

#[test]
fn inline_image_wrapping_and_column_fit() {
    // 22 zeros fill 195.77px; the 50px image wraps to its own line
    let v = measure(
        json!([
            { "kind": "text", "text": "0000000000000000000000" },
            { "kind": "image", "width": 50.0, "height": 30.0 }
        ]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 22), (1, 0, 1, 1)]);
    approx(
        v["lines"][1]["lineHeight"].as_f64().unwrap(),
        30.0 + 2.0 * 3.2,
        "wrapped image line",
    );

    let v = measure(
        json!([{ "kind": "image", "width": 400.0, "height": 100.0 }]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 0), (0, 0, 0, 1)]);
    approx(
        v["lines"][0]["lineHeight"].as_f64().unwrap(),
        16.0 * 1.15,
        "empty leading row",
    );
    approx(
        v["lines"][1]["lineHeight"].as_f64().unwrap(),
        50.0 + 2.0 * 3.2,
        "rendered (fitted) height reserved",
    );
    approx(
        v["lines"][1]["width"].as_f64().unwrap(),
        400.0,
        "declared width kept",
    );
}

#[test]
fn floating_images_skip_but_count_after_tabs() {
    let v = measure(
        json!([
            { "kind": "text", "text": "0" },
            { "kind": "image", "width": 50.0, "height": 500.0,
              "wrapType": "square", "displayMode": "float",
              "position": { "horizontal": { "align": "right" } } }
        ]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 1, 1)]);
    approx(v["lines"][0]["width"].as_f64().unwrap(), W0, "no advance");
    approx(
        v["lines"][0]["lineHeight"].as_f64().unwrap(),
        LH,
        "no growth",
    );

    // end stop at 100px: the floating image's 20px width joins the
    // following-runs width, pulling the tab back with it
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [
                { "kind": "tab" },
                { "kind": "image", "width": 20.0, "height": 20.0,
                  "wrapType": "square", "displayMode": "float",
                  "position": { "horizontal": { "posOffset": 0 } } },
                { "kind": "text", "text": "0" }
            ],
            "attrs": { "tabs": [{ "val": "end", "pos": 1500.0 }] }
        }),
        300.0,
    )
    .unwrap();
    // tab = 100 − (20 + W0); line advance adds only the text W0
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        100.0 - 20.0,
        "floating width anchored the stop",
    );
}

// 24a. a block/topAndBottom image alone gets its own line at declared
// height + distances + descent buffer both sides, then a trailing empty line
#[test]
fn own_line_image_takes_its_own_line() {
    // `displayMode: block` and `wrapType: topAndBottom` share the path
    let variants = [
        json!([{ "kind": "image", "width": 50.0, "height": 100.0, "displayMode": "block" }]),
        json!([{ "kind": "image", "width": 50.0, "height": 100.0, "wrapType": "topAndBottom" }]),
    ];
    for runs in variants {
        let v = measure(runs, 200.0).unwrap();
        // the image's own line, then the empty line opened after it
        assert_eq!(spans(&v), vec![(0, 0, 0, 1), (1, 0, 1, 0)]);
        let img = &v["lines"][0];
        approx(
            img["width"].as_f64().unwrap(),
            0.0,
            "own-line image adds no width",
        );
        // maxImageHeightPx = 100 + 6 + 6 = 112; alone → + 2 × 3.2 descent
        approx(
            img["lineHeight"].as_f64().unwrap(),
            112.0 + 2.0 * 3.2,
            "own-line height",
        );
        approx(
            img["ascent"].as_f64().unwrap(),
            112.0 + 3.2,
            "own-line ascent",
        );
        approx(img["descent"].as_f64().unwrap(), 3.2, "fallback descent");
        // trailing empty line at the metrics-less fallback height
        approx(
            v["lines"][1]["lineHeight"].as_f64().unwrap(),
            16.0 * 1.15,
            "trailing empty line",
        );
        approx(
            v["totalHeight"].as_f64().unwrap(),
            112.0 + 2.0 * 3.2 + 16.0 * 1.15,
            "total height",
        );
    }
}

// 24b. an own-line image finishes the current (text) line first, then takes
// its line, then opens a trailing empty one. Explicit zero wrap distances
// isolate the box to the declared image height.
#[test]
fn own_line_image_finishes_the_current_line_first() {
    let v = measure(
        json!([
            { "kind": "text", "text": "0" },
            { "kind": "image", "width": 40.0, "height": 80.0,
              "displayMode": "block", "distTop": 0.0, "distBottom": 0.0 }
        ]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 1), (1, 0, 1, 1), (2, 0, 2, 0)]);
    approx(v["lines"][0]["width"].as_f64().unwrap(), W0, "text width");
    approx(
        v["lines"][0]["lineHeight"].as_f64().unwrap(),
        LH,
        "text line kept",
    );
    // image alone: 80 + 2 × 3.2 (no distances), no width advance
    approx(
        v["lines"][1]["lineHeight"].as_f64().unwrap(),
        80.0 + 2.0 * 3.2,
        "image line height",
    );
    approx(
        v["lines"][1]["width"].as_f64().unwrap(),
        0.0,
        "no width advance",
    );
    approx(
        v["lines"][2]["lineHeight"].as_f64().unwrap(),
        16.0 * 1.15,
        "trailing empty line",
    );
}

#[test]
fn own_line_image_width_counts_after_a_tab() {
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [
                { "kind": "tab" },
                { "kind": "image", "width": 20.0, "height": 20.0, "displayMode": "block" }
            ],
            "attrs": { "tabs": [{ "val": "end", "pos": 1500.0 }] }
        }),
        300.0,
    )
    .unwrap();
    // end stop at 100px, following width = the block image's 20px, so the
    // tab on the first line measures 100 − 20 = 80px (the image then takes
    // its own line, adding no width there).
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        100.0 - 20.0,
        "block image width anchored the end stop",
    );
}

#[test]
fn dimensionless_image_is_zero_size() {
    // lone image with no dims: one line, zero width, no growth
    let v = measure(json!([{ "kind": "image" }]), 200.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 1)]);
    approx(v["lines"][0]["width"].as_f64().unwrap(), 0.0, "zero width");
    approx(
        v["lines"][0]["lineHeight"].as_f64().unwrap(),
        16.0 * 1.15,
        "no growth (fallback height)",
    );

    // inline after text: contributes nothing to the line width or height
    let v = measure(
        json!([
            { "kind": "text", "text": "0" },
            { "kind": "image" }
        ]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 1, 1)]);
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        W0,
        "text width only",
    );
    approx(
        v["lines"][0]["lineHeight"].as_f64().unwrap(),
        LH,
        "text height only",
    );
}

// ---- smallCaps ----------------------------------------------------------

#[test]
fn small_caps_scales_uppercased_lowercase() {
    // 'a' → 'A' at 0.7: WA × 0.7
    let v = measure(
        json!([{ "kind": "text", "text": "a", "smallCaps": true }]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        WA * 0.7,
        "lowercase scaled",
    );

    // uppercase and uncased chars are untouched
    let v = measure(
        json!([{ "kind": "text", "text": "A0", "smallCaps": true }]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        WA + W0,
        "uppercase/digits full size",
    );

    let v = measure(
        json!([{ "kind": "text", "text": "aA", "smallCaps": true }]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        WA * 0.7 + WA,
        "mixed case",
    );

    // allCaps wins over smallCaps: full-size uppercase (CSS text-transform
    // runs before font-variant finds any lowercase)
    let v = measure(
        json!([{ "kind": "text", "text": "a", "smallCaps": true, "allCaps": true }]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        WA,
        "allCaps beats smallCaps",
    );

    // Small caps compose with horizontal scaling.
    let v = measure(
        json!([{ "kind": "text", "text": "a", "smallCaps": true, "horizontalScale": 200.0 }]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        WA * 0.7 * 2.0,
        "smallCaps × horizontalScale",
    );
}

const W_SHALOM: f64 = 4501.0 / 128.0;
const W_ABG: f64 = 3377.0 / 128.0;

// 26. a Hebrew word: shaped RTL, width is the logical advance sum, spans
// count UTF-16 units; the rtl run flag and the bidi paragraph attr only
// pick the UBA base direction and change nothing about the sums
#[test]
fn hebrew_word_width_and_utf16_spans() {
    let v = measure(json!([{ "kind": "text", "text": "שלום" }]), 200.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 4)]);
    approx(v["lines"][0]["width"].as_f64().unwrap(), W_SHALOM, "shalom");

    let v = measure(
        json!([{ "kind": "text", "text": "שלום", "rtl": true }]),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        W_SHALOM,
        "rtl flag",
    );

    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": [{ "kind": "text", "text": "שלום" }],
            "attrs": { "bidi": true }
        }),
        200.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        W_SHALOM,
        "bidi attr",
    );

    // fields measure bidi text too (measure_plain_text path)
    let v = measure(json!([{ "kind": "field", "fallback": "שלום" }]), 200.0).unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        W_SHALOM,
        "rtl field",
    );
}

// 27. mixed LTR/RTL on one line: per-level segments shaped separately,
// width = sum of segment advances, span stays logical
#[test]
fn mixed_ltr_rtl_line_sums_segment_advances() {
    let v = measure(json!([{ "kind": "text", "text": "0 אבג" }]), 200.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 5)]);
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        W0 + SP + W_ABG,
        "mixed line",
    );
}

#[test]
fn wrap_between_ltr_and_rtl_keeps_logical_spans() {
    // "00 " = 22.24px fits 40px; "שלום" (35.16) wraps whole to line 2
    let v = measure(json!([{ "kind": "text", "text": "00 שלום" }]), 40.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 3), (0, 3, 0, 7)]);
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        2.0 * W0 + SP,
        "ltr line keeps trailing space",
    );
    approx(
        v["lines"][1]["width"].as_f64().unwrap(),
        W_SHALOM,
        "rtl line",
    );
}

#[test]
fn arabic_word_measures_without_fallback_and_keeps_logical_spans() {
    let v = measure_arabic(
        json!([{ "kind": "text", "text": "سلام", "rtl": true }]),
        200.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 4)]);
    let width = v["lines"][0]["width"].as_f64().unwrap();
    assert!(
        width > 10.0 && width < 200.0,
        "Arabic word should measure to a plausible positive width, got {width}"
    );
}

// indents narrow the affected lines (first-line offset vs body width)
#[test]
fn first_line_indent_narrows_only_the_first_line() {
    // "0 0" is 22.24px; fits 25px unindented on one line
    let runs = json!([{ "kind": "text", "text": "0 0" }]);
    let v = measure(runs.clone(), 25.0).unwrap();
    assert_eq!(spans(&v).len(), 1, "no indent: single line");

    // firstLine indent 8 → first line available = 17 → wrap after "0 "
    let v = measure_with(
        json!({
            "kind": "paragraph",
            "runs": runs,
            "attrs": { "indent": { "firstLine": 8.0 } }
        }),
        25.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 2), (0, 2, 0, 3)]);
}

fn measure_block_floats(
    block: Value,
    max_width: f64,
    zones: Value,
    paragraph_y_offset: f64,
) -> Result<Value, String> {
    let input = json!({
        "block": block,
        "maxWidth": max_width,
        "fontChains": { "liberation sans|0|0": [0] },
        "defaults": { "fontSize": 12.0, "fontFamily": "Liberation Sans" },
        "floatingZones": zones,
        "paragraphYOffset": paragraph_y_offset
    });
    let out = measure_paragraph_json(&store(), &input.to_string())?;
    Ok(serde_json::from_str(&out).expect("output is valid JSON"))
}

fn measure_floats(runs: Value, max_width: f64, zones: Value) -> Result<Value, String> {
    measure_block_floats(
        json!({ "kind": "paragraph", "runs": runs }),
        max_width,
        zones,
        0.0,
    )
}

fn line_keys(v: &Value, line: usize) -> Vec<String> {
    let mut keys: Vec<String> = v["lines"][line]
        .as_object()
        .unwrap()
        .keys()
        .cloned()
        .collect();
    keys.sort_unstable();
    keys
}

const BASE_LINE_KEYS: [&str; 8] = [
    "ascent",
    "descent",
    "headChar",
    "headRun",
    "lineHeight",
    "tailChar",
    "tailRun",
    "width",
];

// 29. a left zone covering lines 1–2 of a 4-line wrap: the covered lines
// narrow (breaks shift vs the zone-free baseline) and emit leftOffset; once
// the zone ends mid-paragraph the later lines regain full width and carry
// no float keys at all
#[test]
fn left_zone_narrows_covered_lines_then_releases() {
    let runs = json!([{ "kind": "text", "text": "000 000 000 000 000 000" }]);

    // baseline: no zone → two 100px lines
    let v = measure(runs.clone(), 100.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 12), (0, 12, 0, 23)]);

    // zone bottom 35 sits between line 2's probe top (LH = 18.3984) and
    // line 3's (2·LH = 36.7969): lines 1–2 intersect, line 3 doesn't
    let v = measure_floats(
        runs,
        100.0,
        json!([{ "leftMargin": 40.0, "rightMargin": 0.0, "topY": 0.0, "bottomY": 35.0 }]),
    )
    .unwrap();
    assert_eq!(
        spans(&v),
        vec![(0, 0, 0, 4), (0, 4, 0, 8), (0, 8, 0, 20), (0, 20, 0, 23)]
    );
    for i in [0, 1] {
        approx(
            v["lines"][i]["leftOffset"].as_f64().unwrap(),
            40.0,
            "covered line leftOffset",
        );
        approx(
            v["lines"][i]["width"].as_f64().unwrap(),
            3.0 * W0 + SP,
            "narrowed line width",
        );
        let mut expected: Vec<String> = BASE_LINE_KEYS.iter().map(|s| s.to_string()).collect();
        expected.push("leftOffset".to_string());
        expected.sort_unstable();
        assert_eq!(line_keys(&v, i), expected, "only leftOffset added");
    }
    for i in [2, 3] {
        assert_eq!(line_keys(&v, i), BASE_LINE_KEYS.to_vec(), "full-width line");
    }
    approx(
        v["lines"][2]["width"].as_f64().unwrap(),
        3.0 * (3.0 * W0 + SP),
        "line 3 regains full width",
    );
    approx(v["totalHeight"].as_f64().unwrap(), 4.0 * LH, "no skips");
}

// 30. right zone → rightOffset; zones on both sides → both offsets, width
// shrunk by their sum
#[test]
fn right_and_both_side_zones_emit_offsets() {
    let runs = json!([{ "kind": "text", "text": "000 000 000" }]);

    // zone bottom 17 < line 2's probe top 18.3984 → first line only
    let v = measure_floats(
        runs.clone(),
        100.0,
        json!([{ "leftMargin": 0.0, "rightMargin": 40.0, "topY": 0.0, "bottomY": 17.0 }]),
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 4), (0, 4, 0, 11)]);
    approx(
        v["lines"][0]["rightOffset"].as_f64().unwrap(),
        40.0,
        "rightOffset",
    );
    assert!(
        v["lines"][0].get("leftOffset").is_none(),
        "no leftOffset for a right-side zone"
    );

    // one zone per side: margins max independently, both fields emitted
    let v = measure_floats(
        runs,
        100.0,
        json!([
            { "leftMargin": 30.0, "rightMargin": 0.0, "topY": 0.0, "bottomY": 17.0 },
            { "leftMargin": 0.0, "rightMargin": 20.0, "topY": 0.0, "bottomY": 17.0 }
        ]),
    )
    .unwrap();
    assert_eq!(spans(&v)[0], (0, 0, 0, 4), "50px strip fits one word");
    approx(v["lines"][0]["leftOffset"].as_f64().unwrap(), 30.0, "left");
    approx(
        v["lines"][0]["rightOffset"].as_f64().unwrap(),
        20.0,
        "right",
    );
}

// 31. obstructed lines hop below the float: under MIN_WRAP_SEGMENT_WIDTH
// (24px) of room — a near-full-width margin, a margin wider than the whole
// line, or a fullWidthBlock band — the skip lands on the next line as
// floatSkipBefore, the line measures at full width below the zone, and
// totalHeight includes the gap
#[test]
fn obstructed_lines_skip_below_floats() {
    let runs = json!([{ "kind": "text", "text": "000" }]);

    // 100 − 80 = 20px < 24px → skip = zone bottom − 0 = 50
    for left_margin in [80.0, 150.0] {
        let v = measure_floats(
            runs.clone(),
            100.0,
            json!([{ "leftMargin": left_margin, "rightMargin": 0.0, "topY": 0.0, "bottomY": 50.0 }]),
        )
        .unwrap();
        approx(
            v["lines"][0]["floatSkipBefore"].as_f64().unwrap(),
            50.0,
            "skip to the zone bottom",
        );
        assert!(
            v["lines"][0].get("leftOffset").is_none(),
            "below the zone: full width (y = bottomY is exclusive)"
        );
        approx(
            v["lines"][0]["width"].as_f64().unwrap(),
            3.0 * W0,
            "full-width line below the zone",
        );
        approx(
            v["totalHeight"].as_f64().unwrap(),
            LH + 50.0,
            "totalHeight includes the skip",
        );
    }

    // topAndBottom band: full-width block → zero usable width → same hop
    let v = measure_floats(
        runs,
        100.0,
        json!([{ "leftMargin": 0.0, "rightMargin": 0.0, "topY": 0.0, "bottomY": 40.0,
                 "fullWidthBlock": true }]),
    )
    .unwrap();
    approx(
        v["lines"][0]["floatSkipBefore"].as_f64().unwrap(),
        40.0,
        "band skip",
    );
    assert!(
        v["lines"][0].get("segments").is_none(),
        "below the band no synthetic segment leaks out"
    );
}

#[test]
fn tab_content_x_includes_float_left_offset() {
    // indent.left 48px (720tw → implicit stop at the indent), hanging 24px
    // → first-line grid x starts at 48 − 24 = 24
    let block = json!({
        "kind": "paragraph",
        "runs": [{ "kind": "tab" }, { "kind": "text", "text": "0" }],
        "attrs": { "indent": { "left": 48.0, "hanging": 24.0 } }
    });

    // baseline: contentX = 24 → tab spans to the 48px indent stop = 24px
    let v = measure_with(block.clone(), 200.0).unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        24.0 + W0,
        "no float: tab lands on the body edge",
    );

    // zone leftMargin 10 → contentX = 24 + 10 = 34 → tab shrinks to 14px
    let v = measure_block_floats(
        block,
        200.0,
        json!([{ "leftMargin": 10.0, "rightMargin": 0.0, "topY": 0.0, "bottomY": 17.0 }]),
        0.0,
    )
    .unwrap();
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        14.0 + W0,
        "leftOffset participates in the tab grid x",
    );
    approx(
        v["lines"][0]["leftOffset"].as_f64().unwrap(),
        10.0,
        "offset",
    );
}

// 33. first-line indent + list-marker inline width + zone compose: all
// three subtract from the first line's width (marker footprint = tab stop
// at 48px − markerStart 12px = 36px; see list-marker tests)
#[test]
fn zone_composes_with_marker_and_first_line_indent() {
    let block = json!({
        "kind": "paragraph",
        "runs": [{ "kind": "text", "text": "000 000 000" }],
        "attrs": {
            "listMarker": "1.",
            "indent": { "firstLine": 12.0 }
        }
    });

    // baseline: first line = 150 − 12 (firstLine) − 36 (marker) = 102 →
    // all three words fit (88.98px)
    let v = measure_with(block.clone(), 150.0).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 11)]);

    // zone leftMargin 40 → 62px: exactly two words fit (62.28 ≤ 62.5 slack),
    // the third wraps to a full-width second line
    let v = measure_block_floats(
        block,
        150.0,
        json!([{ "leftMargin": 40.0, "rightMargin": 0.0, "topY": 0.0, "bottomY": 17.0 }]),
        0.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 8), (0, 8, 0, 11)]);
    approx(
        v["lines"][0]["leftOffset"].as_f64().unwrap(),
        40.0,
        "marker line still reports the float offset",
    );
    approx(
        v["lines"][0]["width"].as_f64().unwrap(),
        2.0 * (3.0 * W0 + SP),
        "narrowed marker first line",
    );
}

#[test]
fn centered_zone_splits_line_into_segments() {
    let zones = json!([{
        "leftMargin": 0.0, "rightMargin": 0.0, "topY": 0.0, "bottomY": 17.0,
        "segments": [
            { "leftOffset": 0.0, "availableWidth": 30.0 },
            { "leftOffset": 70.0, "availableWidth": 130.0 }
        ]
    }]);

    // "00000 00000" = 93.43px ≤ strip sum 160 → one line, split at the
    // 3-char prefix (26.70 ≤ 30 < 35.59)
    let v = measure_floats(
        json!([{ "kind": "text", "text": "00000 00000" }]),
        200.0,
        zones.clone(),
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 11)]);
    let segments = v["lines"][0]["segments"].as_array().unwrap();
    assert_eq!(segments.len(), 2);
    let seg_spans: Vec<(u64, u64, u64, u64)> = segments
        .iter()
        .map(|s| {
            (
                s["headRun"].as_u64().unwrap(),
                s["headChar"].as_u64().unwrap(),
                s["tailRun"].as_u64().unwrap(),
                s["tailChar"].as_u64().unwrap(),
            )
        })
        .collect();
    assert_eq!(seg_spans, vec![(0, 0, 0, 3), (0, 3, 0, 11)]);
    approx(
        segments[0]["leftOffset"].as_f64().unwrap(),
        0.0,
        "strip 1 x",
    );
    approx(
        segments[0]["availableWidth"].as_f64().unwrap(),
        30.0,
        "strip 1 room",
    );
    approx(
        segments[0]["width"].as_f64().unwrap(),
        3.0 * W0,
        "strip 1 text",
    );
    approx(
        segments[1]["leftOffset"].as_f64().unwrap(),
        70.0,
        "strip 2 x",
    );
    approx(
        segments[1]["width"].as_f64().unwrap(),
        7.0 * W0 + SP,
        "strip 2 text",
    );
    let mut keys: Vec<&str> = segments[0]
        .as_object()
        .unwrap()
        .keys()
        .map(String::as_str)
        .collect();
    keys.sort_unstable();
    assert_eq!(
        keys,
        vec![
            "availableWidth",
            "headChar",
            "headRun",
            "leftOffset",
            "tailChar",
            "tailRun",
            "width"
        ]
    );

    // a line fitting the first strip: one segment covering the whole line
    let v = measure_floats(
        json!([{ "kind": "text", "text": "00" }]),
        200.0,
        zones.clone(),
    )
    .unwrap();
    let segments = v["lines"][0]["segments"].as_array().unwrap();
    assert_eq!(segments.len(), 1);
    approx(
        segments[0]["width"].as_f64().unwrap(),
        2.0 * W0,
        "whole line in strip 1",
    );
    approx(
        segments[0]["availableWidth"].as_f64().unwrap(),
        30.0,
        "strip 1 room",
    );

    let v = measure_floats(
        json!([
            { "kind": "text", "text": "00000" },
            { "kind": "text", "text": "0" }
        ]),
        200.0,
        zones,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 1, 1)]);
    assert_eq!(
        line_keys(&v, 0),
        BASE_LINE_KEYS.to_vec(),
        "bail emits nothing"
    );
}

// 35. paragraphYOffset shifts the paragraph within the zones' space: the
// same zone misses the paragraph at offset 0 and covers its first line at
// offset 30
#[test]
fn paragraph_y_offset_shifts_zone_intersection() {
    let runs = json!([{ "kind": "text", "text": "000 000" }]);
    let zones = json!([{ "leftMargin": 50.0, "rightMargin": 0.0, "topY": 30.0, "bottomY": 47.0 }]);

    // offset 0: line probe [0, 16) misses [30, 47) → single full line
    let v = measure_floats(runs.clone(), 100.0, zones.clone()).unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 7)]);
    assert!(v["lines"][0].get("leftOffset").is_none());

    // offset 30: probe [30, 46) intersects → 50px strip fits only the first
    // word (57.84 > 50.5); line 2's probe top 30 + LH = 48.4 clears the zone
    let v = measure_block_floats(
        json!({ "kind": "paragraph", "runs": runs }),
        100.0,
        zones,
        30.0,
    )
    .unwrap();
    assert_eq!(spans(&v), vec![(0, 0, 0, 4), (0, 4, 0, 7)]);
    approx(
        v["lines"][0]["leftOffset"].as_f64().unwrap(),
        50.0,
        "offset hit",
    );
    assert!(v["lines"][1].get("leftOffset").is_none(), "line 2 clears");
}

// 36. security clamps on the float context: bounded zone/segment counts and
// sane finite ranges, refused as UNSUPPORTED (host falls back per block)
#[test]
fn float_zone_input_validation() {
    let runs = json!([{ "kind": "text", "text": "0" }]);
    let zone = |left: f64, top: f64, bottom: f64| json!({ "leftMargin": left, "rightMargin": 0.0, "topY": top, "bottomY": bottom });

    // > 200 zones
    let many: Vec<Value> = (0..201).map(|_| zone(10.0, 0.0, 10.0)).collect();
    let err = measure_floats(runs.clone(), 100.0, json!(many)).unwrap_err();
    assert!(err.starts_with("UNSUPPORTED"), "zone count: {err:?}");

    // absurd margin / Y magnitude
    for bad in [
        json!([zone(200_000.0, 0.0, 10.0)]),
        json!([zone(10.0, 0.0, 1.0e10)]),
    ] {
        let err = measure_floats(runs.clone(), 100.0, bad).unwrap_err();
        assert!(err.starts_with("UNSUPPORTED"), "range: {err:?}");
    }

    // absurd paragraphYOffset
    let err = measure_block_floats(
        json!({ "kind": "paragraph", "runs": runs.clone() }),
        100.0,
        json!([zone(10.0, 0.0, 10.0)]),
        1.0e10,
    )
    .unwrap_err();
    assert!(err.starts_with("UNSUPPORTED"), "offset: {err:?}");

    // > 100 segments in one zone
    let segments: Vec<Value> = (0..101)
        .map(|i| json!({ "leftOffset": i as f64, "availableWidth": 1.0 }))
        .collect();
    let err = measure_floats(
        runs,
        100.0,
        json!([{ "leftMargin": 0.0, "rightMargin": 0.0, "topY": 0.0, "bottomY": 10.0,
                 "segments": segments }]),
    )
    .unwrap_err();
    assert!(err.starts_with("UNSUPPORTED"), "segment count: {err:?}");
}
