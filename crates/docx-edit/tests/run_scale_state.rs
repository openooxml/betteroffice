//! A collaborative state can hold a run scale the parser no longer lets in,
//! because it was seeded before the parser refused it or came from a peer.

#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use std::collections::{BTreeMap, HashMap};
use std::sync::Arc;

use docx_edit::{
    EditCtx, EditingDoc, EngineSession, InlineFormatDelta, StoryRange, seed_from_docx,
};
use fixture::{FONT, p, r, region_request, with_body};
use serde_json::Value;
use yrs::Any;

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

/// The drawn runs of `bytes` laid out from a state seeded from it, with a
/// `characterSpacing` of scale 0 written over the run when `scaled` is set.
fn drawn(bytes: &[u8], font: u32, scaled: bool) -> Vec<Value> {
    let seed = EditingDoc::new(75410);
    seed_from_docx(&seed, bytes).unwrap();
    if scaled {
        let spacing = Any::Map(Arc::new(HashMap::from([(
            "scale".to_owned(),
            Any::Number(0.0),
        )])));
        seed.format_range(
            &EditCtx::local("Ada", "2026-11-09T00:00:00Z"),
            StoryRange::new("body", 0, 6),
            &InlineFormatDelta {
                other: BTreeMap::from([("characterSpacing".to_owned(), Some(spacing))]),
                ..InlineFormatDelta::default()
            },
        )
        .unwrap();
    }
    let engine = EngineSession::new(75411);
    engine
        .doc()
        .apply_host_update_v1(&seed.encode_state_as_update_v1())
        .unwrap();
    let request = region_request(&engine, bytes, font);
    let output = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    let mut runs = Vec::new();
    drawn_runs(&display, &mut runs);
    runs.into_iter().cloned().collect()
}

#[test]
fn a_state_holding_a_run_scale_of_0_draws_as_the_same_run_with_no_scale() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let bytes = with_body(&p("00000001", &r("Scaled")));
    let scaled = drawn(&bytes, font, true);
    let plain = drawn(&bytes, font, false);
    let text: String = scaled
        .iter()
        .filter_map(|run| run["text"].as_str())
        .collect();
    assert_eq!(text, "Scaled", "{scaled:?}");
    assert_eq!(scaled, plain);
    let width: f64 = plain.iter().filter_map(|run| run["width"].as_f64()).sum();
    assert!(width > 0.0, "{plain:?}");
}
