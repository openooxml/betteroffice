//! A classic `w:shadow` run draws with the classic text shadow and carries no
//! Word 2010 effects for the canvas to take instead.

#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EngineSession, seed_from_docx};
use fixture::{FONT, p, region_request, with_body};
use serde_json::Value;

fn drawn_runs<'a>(value: &'a Value, found: &mut Vec<&'a Value>) {
    match value {
        Value::Array(items) => items.iter().for_each(|item| drawn_runs(item, found)),
        Value::Object(entries) => {
            if matches!(
                entries.get("kind").and_then(Value::as_str),
                Some("text" | "glyphRun")
            ) {
                found.push(value);
            }
            entries.values().for_each(|item| drawn_runs(item, found));
        }
        _ => {}
    }
}

#[test]
fn a_classic_shadow_run_draws_the_classic_shadow_without_modern_effects() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let bytes = with_body(&p(
        "00000001",
        r#"<w:r><w:rPr><w:shadow/></w:rPr><w:t>Shadowed</w:t></w:r>"#,
    ));
    let engine = EngineSession::new(75420);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let request = region_request(&engine, &bytes, font);
    let output = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    let mut runs = Vec::new();
    drawn_runs(&display, &mut runs);
    let shadowed: Vec<&Value> = runs
        .into_iter()
        .filter(|run| run["text"].as_str().is_some_and(|text| !text.is_empty()))
        .collect();
    let text: String = shadowed
        .iter()
        .filter_map(|run| run["text"].as_str())
        .collect();
    assert_eq!(text, "Shadowed", "{shadowed:?}");
    for run in shadowed {
        assert_eq!(run["textShadow"], "shadow", "{run}");
        assert!(run.get("modernEffects").is_none(), "{run}");
    }
}
