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

#[test]
fn a_state_holding_a_run_scale_of_0_lays_out_at_full_width() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let bytes = with_body(&p("00000001", &r("Scaled")));
    let seed = EditingDoc::new(75410);
    seed_from_docx(&seed, &bytes).unwrap();
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

    let engine = EngineSession::new(75411);
    engine
        .doc()
        .apply_host_update_v1(&seed.encode_state_as_update_v1())
        .unwrap();
    let request = region_request(&engine, &bytes, font);
    let output = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    let mut runs = Vec::new();
    drawn_runs(&display, &mut runs);
    assert!(!runs.is_empty(), "the run is drawn");
    for run in runs {
        assert_eq!(run.get("horizontalScale"), None, "{run}");
    }
}
