//! Native smoke test for the measurement surface exported from this crate:
//! the thread_local font registry wiring, not the measurement math (that is
//! covered in `ooxml-text`'s own suites). Only success paths run here — the
//! error paths construct `JsValue`s, which are wasm-only at runtime.

use docx_layout::{clear_measure_fonts, measure_paragraph_json, register_measure_font};

const FIXTURE: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[test]
fn register_measure_clear_round_trip() {
    let id = register_measure_font(FIXTURE).expect("fixture font registers");
    assert_eq!(id, 0);

    let input = serde_json::json!({
        "block": {
            "kind": "paragraph",
            "runs": [{ "kind": "text", "text": "0 0" }]
        },
        "maxWidth": 200.0,
        "fontChains": { "liberation sans|0|0": [id] },
        "defaults": { "fontSize": 12.0, "fontFamily": "Liberation Sans" }
    })
    .to_string();

    let out = measure_paragraph_json(&input).expect("simple paragraph measures");
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(v["kind"], "paragraph");
    assert_eq!(v["lines"].as_array().unwrap().len(), 1);
    assert!(v["totalHeight"].as_f64().unwrap() > 0.0);

    // ids restart after clearing
    clear_measure_fonts();
    let id = register_measure_font(FIXTURE).expect("re-registers after clear");
    assert_eq!(id, 0);
}
