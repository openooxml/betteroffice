use super::lowering_pages as fixture;
use super::*;
use serde_json::{Value, json};

pub(super) struct OpenSwitch(Option<bool>);

impl OpenSwitch {
    pub(super) fn new(enabled: Option<bool>) -> Self {
        Self(OPEN_REGION_PATH_OVERRIDE.with(|value| value.replace(enabled)))
    }
}

impl Drop for OpenSwitch {
    fn drop(&mut self) {
        OPEN_REGION_PATH_OVERRIDE.with(|enabled| enabled.set(self.0));
    }
}

pub(super) struct PreviewSwitch(Option<bool>);

impl PreviewSwitch {
    pub(super) fn new(enabled: Option<bool>) -> Self {
        Self(PREVIEW_REGION_PATH_OVERRIDE.with(|value| value.replace(enabled)))
    }
}

impl Drop for PreviewSwitch {
    fn drop(&mut self) {
        PREVIEW_REGION_PATH_OVERRIDE.with(|enabled| enabled.set(self.0));
    }
}

#[test]
fn default_preview_and_open_use_plain_layout_while_bulk_uses_regions() {
    let _preview = PreviewSwitch::new(None);
    let _open = OpenSwitch::new(None);
    for (trigger, region) in [
        (RelayoutTrigger::Interactive, false),
        (RelayoutTrigger::Preview, false),
        (RelayoutTrigger::Open, false),
        (RelayoutTrigger::Bulk, true),
    ] {
        assert_eq!(trigger.uses_region_path(), region);
        let engine = EngineSession::new(75310);
        engine
            .layout_regions_for_trigger("{}", None, trigger)
            .unwrap();
        assert_eq!(engine.region_retention_valid.get(), region);
        assert_eq!(engine.pagination.borrow().display_uses_region_path, region);
        assert_eq!(
            engine
                .regions
                .borrow()
                .as_ref()
                .unwrap()
                .region_request_fingerprint
                .is_some(),
            region
        );
    }
    {
        let _preview = PreviewSwitch::new(Some(true));
        let _open = OpenSwitch::new(Some(true));
        assert!(RelayoutTrigger::Preview.uses_region_path());
        assert!(RelayoutTrigger::Open.uses_region_path());
        {
            let _preview = PreviewSwitch::new(Some(false));
            let _open = OpenSwitch::new(Some(false));
            assert!(!RelayoutTrigger::Preview.uses_region_path());
            assert!(!RelayoutTrigger::Open.uses_region_path());
        }
        assert!(RelayoutTrigger::Preview.uses_region_path());
        assert!(RelayoutTrigger::Open.uses_region_path());
    }
    assert!(!RelayoutTrigger::Preview.uses_region_path());
    assert!(!RelayoutTrigger::Open.uses_region_path());
}

#[test]
fn preview_font_preflight_survives_plain_layout_and_region_entry() {
    let _preview = PreviewSwitch::new(None);
    let _open = OpenSwitch::new(None);
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(fixture::FONT).unwrap();
    let bytes = split_document(40, 4, 20, true);
    let engine = EngineSession::new(75313);
    crate::seed_from_docx(engine.doc(), &bytes).unwrap();
    let mut request = split_request(&engine, &bytes, font);
    let revision = engine.doc().list_revisions().unwrap()[0]
        .change
        .revision_id
        .clone();
    request["renderEnv"]["revisionPreview"] = json!({});
    request["renderEnv"]["revisionPreview"][&revision] = json!("accepted");
    let requirements = engine
        .layout_font_requirements_json(&request.to_string())
        .unwrap();
    for trigger in [
        RelayoutTrigger::Open,
        RelayoutTrigger::Preview,
        RelayoutTrigger::Interactive,
        RelayoutTrigger::Bulk,
    ] {
        engine
            .layout_regions_for_trigger(&request.to_string(), None, trigger)
            .unwrap();
        assert!(engine.preview_font_requirements.borrow().is_some());
        let cold = EngineSession::new(75314);
        cold.doc()
            .apply_update_v1(&engine.doc().encode_state_as_update_v1())
            .unwrap();
        cold.layout_document_with_regions_retained(&request.to_string())
            .unwrap();
        assert_eq!(cold.stats().incremental_pagination_calls, 0);
        assert_eq!(
            engine.retained_layout_json().unwrap(),
            cold.retained_layout_json().unwrap()
        );
        assert_eq!(
            engine.retained_kernel_inputs_json().unwrap(),
            cold.retained_kernel_inputs_json().unwrap()
        );
        let before = engine.stats();
        let decision = if request["renderEnv"]["revisionPreview"][&revision] == "accepted" {
            "rejected"
        } else {
            "accepted"
        };
        request["renderEnv"]["revisionPreview"][&revision] = json!(decision);
        assert_eq!(
            engine
                .layout_font_requirements_json(&request.to_string())
                .unwrap(),
            requirements
        );
        assert_eq!(engine.stats().lower_cache_hits, before.lower_cache_hits);
        assert_eq!(engine.stats().lower_cache_misses, before.lower_cache_misses);
    }
}

#[test]
fn first_lowering_records_preview_units_only_for_region_paths() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(fixture::FONT).unwrap();
    let bytes = split_document(40, 4, 20, true);
    for (trigger, region) in [
        (RelayoutTrigger::Open, false),
        (RelayoutTrigger::Preview, false),
        (RelayoutTrigger::Interactive, false),
        (RelayoutTrigger::Bulk, true),
        (RelayoutTrigger::Open, true),
        (RelayoutTrigger::Preview, true),
    ] {
        let _open = OpenSwitch::new(Some(region));
        let _preview = PreviewSwitch::new(Some(region));
        let engine = EngineSession::new(75315);
        crate::seed_from_docx(engine.doc(), &bytes).unwrap();
        let mut request = json!({
            "bodyStory": "body",
            "renderEnv": {},
            "measurement": {
                "fontChains": {"calibri|0|0": [font]},
                "defaults": {"fontFamily": "Calibri", "fontSize": 11},
                "authoritativeShaping": true,
            },
        });
        for relower in [false, true] {
            if relower {
                request["renderEnv"]["showHiddenText"] = json!(true);
            }
            engine
                .layout_regions_for_trigger(&request.to_string(), None, trigger)
                .unwrap();
            assert_eq!(
                engine.render.borrow().stories["body"].preview.is_some(),
                region || relower,
                "{trigger:?}, relower={relower}"
            );
            let cold = EngineSession::new(75316);
            cold.doc()
                .apply_update_v1(&engine.doc().encode_state_as_update_v1())
                .unwrap();
            cold.layout_document_with_regions_retained(&request.to_string())
                .unwrap();
            assert_eq!(cold.stats().incremental_pagination_calls, 0);
            assert_eq!(
                engine.retained_layout_json().unwrap(),
                cold.retained_layout_json().unwrap()
            );
            assert_eq!(
                engine.retained_kernel_inputs_json().unwrap(),
                cold.retained_kernel_inputs_json().unwrap()
            );
        }
    }
}

#[test]
fn default_preview_accept_undo_reject_undo_matches_cold_full_layout() {
    let _preview = PreviewSwitch::new(None);
    let _open = OpenSwitch::new(None);
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(fixture::FONT).unwrap();
    let bytes = split_document(40, 4, 20, true);
    let engine = EngineSession::new(75311);
    crate::seed_from_docx(engine.doc(), &bytes).unwrap();
    engine.set_relayout_trigger(RelayoutTrigger::Open);
    let mut request = split_request(&engine, &bytes, font);
    engine
        .layout_document_with_regions_retained(&request.to_string())
        .unwrap();
    let revision = engine.doc().list_revisions().unwrap()[0]
        .change
        .revision_id
        .clone();
    let version = engine.doc().version();
    let epoch = engine.doc_epoch();
    for decision in [Some("accepted"), None, Some("rejected"), None] {
        request["renderEnv"]["revisionPreview"] = json!({});
        if let Some(decision) = decision {
            request["renderEnv"]["revisionPreview"][&revision] = json!(decision);
        }
        let request = request.to_string();
        assert_eq!(
            engine.region_relayout_trigger(&request).unwrap(),
            RelayoutTrigger::Preview
        );
        REGION_WORK_COUNTS.with(|counts| counts.set(RegionWorkCounts::default()));
        let actual = engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        assert_eq!(
            REGION_WORK_COUNTS.with(Cell::get),
            RegionWorkCounts::default()
        );
        assert!(!engine.region_retention_valid.get());
        assert_eq!(engine.doc_epoch(), epoch);
        assert_eq!(engine.doc().version(), version);
        let cold = EngineSession::new(75312);
        cold.doc()
            .apply_update_v1(&engine.doc().encode_state_as_update_v1())
            .unwrap();
        cold.doc()
            .set_note_separator_state(engine.doc().note_separator_state().unwrap());
        cold.set_relayout_trigger(RelayoutTrigger::Open);
        let expected = cold
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        assert_eq!(cold.stats().incremental_pagination_calls, 0);
        assert_eq!(actual.as_bytes(), expected.as_bytes());
        assert_eq!(engine.retained_layout_json().unwrap(), expected);
        assert_eq!(
            engine.retained_kernel_inputs_json().unwrap(),
            cold.retained_kernel_inputs_json().unwrap()
        );
    }
}

#[test]
fn an_early_footnote_bounds_retained_probe_payload_and_matches_fresh_with_open_regions() {
    let _open = OpenSwitch::new(Some(true));
    tests::an_early_footnote_bounds_retained_probe_payload_and_matches_fresh();
}

#[test]
fn interactive_note_layout_batches_match_main_frames_and_cold_rebuilds_with_open_regions() {
    let _open = OpenSwitch::new(Some(true));
    tests::interactive_note_layout_batches_match_main_frames_and_cold_rebuilds();
}

#[test]
fn interactive_note_sections_keystrokes_match_cold_without_region_work_with_open_regions() {
    let _open = OpenSwitch::new(Some(true));
    tests::interactive_note_sections_keystrokes_match_cold_without_region_work();
}

#[test]
fn host_proposals_before_initial_layout_keep_the_open_trigger_with_open_regions() {
    let _open = OpenSwitch::new(Some(true));
    tests::host_proposals_before_initial_layout_keep_the_open_trigger();
}

#[test]
fn open_and_bulk_region_switches_preserve_cold_note_layouts_with_open_regions() {
    let _open = OpenSwitch::new(Some(true));
    tests::open_and_bulk_region_switches_preserve_cold_note_layouts();
}

fn split_document(preceding: u32, lines: usize, height: u32, widow: bool) -> Vec<u8> {
    let paragraph = |id: u32, height: u32, content: &str| {
        fixture::p(
            &format!("{id:08X}"),
            &format!(
                r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="{}" w:lineRule="exact"/><w:widowControl w:val="{}"/></w:pPr>{content}"#,
                height * 15,
                u8::from(widow),
            ),
        )
    };
    let split = (0..lines)
        .map(|index| {
            let line = fixture::r(&format!("Split {index}"));
            if index == 0 {
                line
            } else {
                format!("<w:r><w:br/></w:r>{line}")
            }
        })
        .collect::<String>();
    let mut body = [
        paragraph(1, 100, &fixture::r("First")),
        paragraph(2, 100, &fixture::r("Second")),
        paragraph(3, preceding, &fixture::r("Preceding")),
        paragraph(4, height, &split),
        paragraph(
            5,
            10,
            &format!(
                r#"{}<w:ins w:id="1" w:author="Ann" w:date="2026-09-29T12:00:00Z">{}</w:ins>"#,
                fixture::r("After"),
                fixture::r(" changed"),
            ),
        ),
    ]
    .concat();
    for id in 6..12 {
        body.push_str(&paragraph(id, 100, &fixture::r("Tail")));
    }
    body.push_str("<w:sectPr/>");
    fixture::with_body(&body)
}

fn split_request(engine: &EngineSession, bytes: &[u8], font: u32) -> Value {
    let mut request = fixture::region_request(engine, bytes, font);
    request["regions"]["sections"] = json!([{
        "sectionId": "main",
        "pageSize": {"w": 200, "h": 120},
        "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10},
    }]);
    request["notes"]["contents"] = json!([]);
    request
}

fn prime_split(
    bytes: &[u8],
    font: u32,
    slices: &[(usize, usize, usize)],
) -> (EngineSession, Value) {
    let engine = EngineSession::new(75301);
    let seed = EditingDoc::new(75300);
    crate::seed_from_docx(&seed, bytes).unwrap();
    engine
        .doc()
        .apply_host_update_v1(&seed.encode_state_as_update_v1())
        .unwrap();
    engine
        .doc()
        .set_note_separator_state(seed.note_separator_state().unwrap());
    let request = split_request(&engine, bytes, font);
    engine.set_relayout_trigger(RelayoutTrigger::Preview);
    let output = engine
        .layout_document_with_regions_retained_json(&request.to_string())
        .unwrap();
    let layout: Value = serde_json::from_str(&output).unwrap();
    let input: Value =
        serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
    let split_id = &input["measured"][3]["block"]["id"];
    let actual: Vec<_> = layout["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .enumerate()
        .flat_map(|(page, value)| {
            value["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .filter(move |fragment| &fragment["blockId"] == split_id)
                .map(move |fragment| {
                    (
                        page,
                        fragment["fromLine"].as_u64().unwrap() as usize,
                        fragment["toLine"].as_u64().unwrap() as usize,
                    )
                })
        })
        .collect();
    assert_eq!(actual, slices);
    (engine, request)
}

fn assert_resumed_matches_fresh(engine: &EngineSession, request: &Value) {
    let before = engine.stats();
    let output = engine
        .layout_document_with_regions_retained_json(&request.to_string())
        .unwrap();
    let after = engine.stats();
    assert_eq!(
        after.incremental_pagination_calls,
        before.incremental_pagination_calls + 1
    );
    assert!(after.pagination_blocks_placed - before.pagination_blocks_placed <= 4);
    assert!(after.rebuilt_pages < after.retained_pages);
    let fresh = EngineSession::new(75302);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    let expected = fresh
        .layout_document_with_regions_retained_json(&request.to_string())
        .unwrap();
    assert_eq!(output.as_bytes(), expected.as_bytes());
    assert_eq!(engine.retained_layout_json().unwrap(), expected);
}

#[test]
fn a_preview_decision_after_a_split_paragraph_resumes_and_matches_fresh() {
    let _preview = PreviewSwitch::new(Some(true));
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    for (preceding, lines, height, widow, slices) in [
        (95, 2, 10, false, vec![(3, 0, 1), (3, 1, 2)]),
        (90, 2, 10, false, vec![(2, 0, 1), (3, 1, 2)]),
        (40, 4, 20, true, vec![(2, 0, 2), (3, 2, 4)]),
    ] {
        let bytes = split_document(preceding, lines, height, widow);
        for decision in ["accepted", "rejected"] {
            let (engine, mut request) = prime_split(&bytes, font, &slices);
            let revision = engine.doc().list_revisions().unwrap()[0]
                .change
                .revision_id
                .clone();
            request["renderEnv"]["revisionPreview"] = json!({});
            request["renderEnv"]["revisionPreview"][&revision] = json!(decision);
            assert_resumed_matches_fresh(&engine, &request);
        }
    }
}
