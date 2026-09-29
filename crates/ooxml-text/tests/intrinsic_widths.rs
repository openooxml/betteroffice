use ooxml_text::FontStore;
use ooxml_text::measure::{MeasureInput, measure_intrinsic_widths};
use serde_json::json;

fn widths(runs: serde_json::Value, attrs: serde_json::Value) -> (f32, f32) {
    let mut store = FontStore::new();
    let font = store
        .register(include_bytes!("fonts/LiberationSans-Regular.ttf").to_vec())
        .unwrap();
    let input: MeasureInput = serde_json::from_value(json!({
        "block": {"kind": "paragraph", "runs": runs, "attrs": attrs},
        "maxWidth": 10000,
        "fontChains": {"arial|0|0": [font.to_u32()]},
        "defaults": {"fontFamily": "Arial", "fontSize": 12},
        "authoritativeShaping": true
    }))
    .unwrap();
    let request = ooxml_text::measure::MeasureRequest {
        block: &input.block,
        max_width: input.max_width,
        font_chains: ooxml_text::measure::FontChains::Hash(&input.font_chains),
        defaults: &input.defaults,
        compat: input.compat,
        floating_zones: None,
        paragraph_y_offset: None,
        authoritative_shaping: input.authoritative_shaping,
    };
    measure_intrinsic_widths(&store, &request).unwrap()
}

#[test]
fn a_word_across_run_boundaries_stays_unbreakable() {
    let (minimum, maximum) = widths(
        json!([
            {"kind": "text", "text": "long", "allCaps": true},
            {"kind": "text", "text": "word", "horizontalScale": 80}
        ]),
        json!({}),
    );
    let a = widths(
        json!([{"kind": "text", "text": "long", "allCaps": true}]),
        json!({}),
    )
    .1;
    let b = widths(
        json!([{"kind": "text", "text": "word", "horizontalScale": 80}]),
        json!({}),
    )
    .1;
    assert!((minimum - a - b).abs() < 0.001);
    assert_eq!(minimum, maximum);
}

#[test]
fn break_opportunities_cross_runs_and_trailing_spaces_do_not_set_minimums() {
    let expected = widths(json!([{"kind": "text", "text": "longest"}]), json!({})).1;
    let (minimum, maximum) = widths(
        json!([
            {"kind": "text", "text": "longest "},
            {"kind": "text", "text": "x   "}
        ]),
        json!({}),
    );
    assert!((minimum - expected).abs() < 0.001);
    assert!(maximum > minimum);
    assert!(
        (maximum - widths(json!([{"kind": "text", "text": "longest x"}]), json!({})).1).abs()
            < 0.001
    );
}

#[test]
fn explicit_breaks_and_hidden_text_do_not_inflate_natural_width() {
    let expected = widths(json!([{"kind": "text", "text": "longest"}]), json!({})).1;
    let (minimum, maximum) = widths(
        json!([
            {"kind": "text", "text": "longest"},
            {"kind": "lineBreak"},
            {"kind": "text", "text": "invisible very wide text", "hidden": true},
            {"kind": "text", "text": "x"}
        ]),
        json!({}),
    );
    assert_eq!(minimum, expected);
    assert_eq!(maximum, expected);
}

#[test]
fn inline_images_and_paragraph_indents_constrain_both_widths() {
    let (minimum, maximum) = widths(
        json!([
            {"kind": "image", "width": 80, "height": 10, "displayMode": "inline"}
        ]),
        json!({"indent": {"left": 12, "right": 8, "firstLine": 4}}),
    );
    assert_eq!(minimum, 104.0);
    assert_eq!(maximum, 104.0);
}
