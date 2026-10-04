use super::*;

fn assert_preview_honoured(engine: &EngineSession) {
    assert!(!engine.interactive_pending.get());
    engine.set_relayout_trigger(RelayoutTrigger::Preview);
    assert_eq!(engine.pending_relayout_trigger(), RelayoutTrigger::Preview);
}

fn apply_resident(engine: &EngineSession, profiled: bool) -> Result<(), String> {
    let epoch = engine.display.borrow().binary_frame_epoch;
    if profiled {
        let mut time = 0.0;
        engine
            .apply_and_layout_profiled("body", epoch, &mut || {
                time += 1.0;
                time
            })
            .map(|_| ())
    } else {
        engine.apply_and_layout("body", epoch).map(|_| ())
    }
}

#[test]
fn resident_errors_consume_interactive_before_the_next_preview() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for profiled in [false, true] {
        for failure in ["render", "measurement", "regions"] {
            let engine = paragraphs_engine(9390, 3);
            engine.set_relayout_trigger(RelayoutTrigger::Interactive);
            let request = small_page_request(font);
            engine
                .layout_document_with_regions_retained(&request)
                .unwrap();
            engine
                .edit_resident_text(crate::StoryRange::new("body", 3, 3), Some("x"), true)
                .unwrap();
            assert!(engine.interactive_pending.get());
            let expected = match failure {
                "render" => {
                    engine.regions.borrow_mut().take();
                    engine.render.borrow_mut().stories.remove("body");
                    "resident render environment missing"
                }
                "measurement" => {
                    engine.regions.borrow_mut().take();
                    engine.measurement.borrow_mut().templates.clear();
                    "resident measurement template missing"
                }
                "regions" => {
                    let mut request: serde_json::Value = serde_json::from_str(&request).unwrap();
                    request["renderEnv"] = json!({"revisionPreview": []});
                    let mut regions = engine.regions.borrow_mut();
                    let state = regions.as_mut().unwrap();
                    state.fast_path = None;
                    state.request_json = request.to_string();
                    "parse render environment"
                }
                _ => unreachable!(),
            };
            let error = apply_resident(&engine, profiled).unwrap_err();
            assert!(
                error.contains(expected),
                "{failure}, profiled={profiled}: {error}"
            );
            assert_preview_honoured(&engine);
        }
    }
}

#[test]
fn resident_successes_consume_interactive_before_the_next_preview() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for profiled in [false, true] {
        let engine = paragraphs_engine(9391, 3);
        engine.set_relayout_trigger(RelayoutTrigger::Interactive);
        let request = small_page_request(font);
        let request_value: serde_json::Value = serde_json::from_str(&request).unwrap();
        let extras = json!({"fontChains": request_value["measurement"]["fontChains"]}).to_string();
        engine
            .layout_document_with_regions_retained(&request)
            .unwrap();
        engine.build_display_list_frame(&extras, 0).unwrap();
        engine
            .edit_resident_text(crate::StoryRange::new("body", 3, 3), Some("x"), true)
            .unwrap();
        apply_resident(&engine, profiled).unwrap();
        assert_preview_honoured(&engine);
    }
}

#[test]
fn sliced_errors_consume_interactive_before_the_next_preview() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for begin in [false, true] {
        let engine = paragraphs_engine(9392, 3);
        if begin {
            let progress = engine
                .begin_region_layout(&small_page_request(font))
                .unwrap();
            assert!(progress.layout_json.is_none());
        }
        engine
            .edit_resident_text(crate::StoryRange::new("body", 3, 3), Some("x"), true)
            .unwrap();
        let error = engine.resume_region_layout(1).unwrap_err();
        assert!(error.contains(if begin {
            "the document or its fonts changed"
        } else {
            "no region layout to resume"
        }));
        assert_preview_honoured(&engine);
    }
}

#[test]
fn display_reset_preserves_plain_damage_and_clears_region_damage() {
    for region in [false, true] {
        let mut pagination = PaginationState {
            display_uses_region_path: region,
            display_rebuilt_pages: BTreeSet::from([2]),
            display_layout_pending: true,
            display_full_rebuild: true,
            position_deltas: HashMap::from([("body".to_owned(), 1)]),
            note_changed_pages: vec![3],
            restamped_pages: Some(BTreeSet::from([4])),
            ..Default::default()
        };
        pagination.clear_display_damage();
        assert_eq!(pagination.display_rebuilt_pages.is_empty(), region);
        assert_eq!(pagination.display_layout_pending, !region);
        assert_eq!(pagination.display_full_rebuild, !region);
        assert_eq!(pagination.position_deltas.is_empty(), region);
        assert_eq!(pagination.note_changed_pages.is_empty(), region);
        assert_eq!(pagination.restamped_pages, Some(BTreeSet::new()));
    }
}

#[test]
fn repeated_interactive_frames_match_main_with_and_without_notes() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for notes in [false, true] {
        let (engine, request) = if notes {
            interactive_note_sections_engine(font)
        } else {
            let engine = paragraphs_engine(9393, 25);
            let request: serde_json::Value = serde_json::from_str(&small_page_request(font)).unwrap();
            (engine, request)
        };
        engine.set_relayout_trigger(RelayoutTrigger::Interactive);
        let extras = json!({"fontChains": request["measurement"]["fontChains"]}).to_string();
        engine
            .layout_document_with_regions_retained(&request.to_string())
            .unwrap();
        engine.build_display_list_frame(&extras, 0).unwrap();
        let mut reference_display = main_frame_baseline(&engine, &extras);
        engine
            .edit_resident_text(crate::StoryRange::new("body", 3, 3), Some("x"), true)
            .unwrap();
        engine
            .layout_document_with_regions_retained(&request.to_string())
            .unwrap();
        let (deltas, note_pages, rebuilt_ranges, display_pending) = {
            let pagination = engine.pagination.borrow();
            assert!(pagination.last_incremental);
            assert!(!pagination.display_uses_region_path);
            assert!(!pagination.position_deltas.is_empty());
            (
                pagination.position_deltas.clone(),
                pagination.note_changed_pages.clone(),
                pagination.rebuilt_page_ranges.clone(),
                pagination.display_layout_pending,
            )
        };
        for frame in 0..2 {
            let (expected_bytes, expected_list, _) = main_frame_oracle(
                &engine,
                &extras,
                std::mem::take(&mut reference_display),
            );
            let epoch = engine.display.borrow().binary_frame_epoch;
            assert_eq!(
                engine.build_display_list_frame(&extras, epoch).unwrap(),
                expected_bytes,
                "notes={notes}, frame={frame}"
            );
            assert_eq!(
                engine.with_display_list(Clone::clone).unwrap(),
                expected_list
            );
            {
                let pagination = engine.pagination.borrow();
                assert_eq!(pagination.position_deltas, deltas);
                assert_eq!(pagination.note_changed_pages, note_pages);
                assert_eq!(pagination.rebuilt_page_ranges, rebuilt_ranges);
                assert_eq!(pagination.display_layout_pending, display_pending);
                assert_eq!(pagination.restamped_pages, Some(BTreeSet::new()));
            }
            if frame == 0 {
                reference_display = main_frame_baseline(&engine, &extras);
            }
        }
    }
}
