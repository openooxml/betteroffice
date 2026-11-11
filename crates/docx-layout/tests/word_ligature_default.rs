//! Word draws Calibri's optional ligatures only where `w14:ligatures` asks:
//! a docx run without it measures one glyph per letter of "fi ti".

use docx_layout::{clear_measure_fonts, measure_paragraph_json_str, register_measure_font};

const CARLITO: &[u8] = include_bytes!("../../../packages/fonts/assets/Carlito-Regular.ttf");

fn glyph_count(run: serde_json::Value) -> usize {
    clear_measure_fonts();
    let id = register_measure_font(CARLITO).expect("Carlito registers");
    let input = serde_json::json!({
        "block": { "kind": "paragraph", "runs": [run] },
        "maxWidth": 1000.0,
        "fontChains": { "calibri|0|0": [id] },
        "defaults": { "fontSize": 12.0, "fontFamily": "Calibri" },
        "authoritativeShaping": true
    })
    .to_string();
    let out: serde_json::Value =
        serde_json::from_str(&measure_paragraph_json_str(&input).expect("measures")).unwrap();
    let packed = out["lines"][0]["glyphRuns"].as_str().expect("glyph runs");
    ooxml_text::unpack_glyph_runs(packed)
        .unwrap()
        .iter()
        .map(|run| run.glyphs.len())
        .sum()
}

#[test]
fn docx_run_without_ligature_request_keeps_separate_letters() {
    assert_eq!(
        glyph_count(serde_json::json!({ "kind": "text", "text": "fi ti" })),
        5
    );
}

#[test]
fn docx_run_asking_for_standard_ligatures_joins_them() {
    let joined = glyph_count(serde_json::json!({
        "kind": "text", "text": "fi ti", "ligatures": "standardContextual"
    }));
    assert!(joined < 5, "{joined}");
}
