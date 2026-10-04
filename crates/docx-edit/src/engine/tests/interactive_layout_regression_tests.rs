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
            let request_value: serde_json::Value = serde_json::from_str(&request).unwrap();
            let extras =
                json!({"fontChains": request_value["measurement"]["fontChains"]}).to_string();
            engine.build_display_list_frame(&extras, 0).unwrap();
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
                    request["renderEnv"] = json!({"showHiddenText": []});
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
            let request: serde_json::Value =
                serde_json::from_str(&small_page_request(font)).unwrap();
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
            let (expected_bytes, expected_list, _) =
                main_frame_oracle(&engine, &extras, std::mem::take(&mut reference_display));
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

fn interactive_damage_engine(
    font: u32,
    page_count: usize,
    notes: bool,
    window: Option<std::ops::Range<usize>>,
) -> (EngineSession, String, String) {
    let body: String = (1..=page_count)
        .map(|page| {
            let mut content = if page == 1 {
                String::new()
            } else {
                "<w:pPr><w:pageBreakBefore/></w:pPr>".to_owned()
            };
            let suffix = if notes { "" } else { " A" };
            content += &lowering_pages::r(&format!("Page {page}{suffix}"));
            if notes && page == 4 {
                content += r#"<w:r><w:footnoteReference w:id="1"/></w:r>"#;
            }
            lowering_pages::p(&format!("{:08X}", 0x7300_0000 + page), &content)
        })
        .collect();
    let body = format!(
        concat!(
            "{}<w:sectPr><w:pgSz w:w=\"7200\" w:h=\"5760\"/><w:pgMar w:top=\"720\" ",
            "w:right=\"720\" w:bottom=\"720\" w:left=\"720\" w:header=\"300\" ",
            "w:footer=\"300\" w:gutter=\"0\"/></w:sectPr>"
        ),
        body
    );
    let bytes = lowering_pages::with_body_and_note(
        &body,
        &lowering_pages::p("73010000", &lowering_pages::r("Note")),
    );
    let engine = EngineSession::new(9394);
    crate::seed_from_docx(engine.doc(), &bytes).unwrap();
    engine.set_relayout_trigger(RelayoutTrigger::Interactive);
    let request = lowering_pages::region_request(&engine, &bytes, font);
    let extras = json!({"fontChains": request["measurement"]["fontChains"]}).to_string();
    let request = request.to_string();
    engine
        .layout_document_with_regions_retained(&request)
        .unwrap();
    engine.set_display_window(window);
    engine.build_display_list_frame(&extras, 0).unwrap();
    assert_eq!(
        engine.with_display_list(|list| list.pages.len()),
        Some(page_count)
    );
    (engine, request, extras)
}

#[test]
fn separated_interactive_layouts_match_main_note_anchors_and_frame_bytes() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let (engine, request, extras) = interactive_damage_engine(font, 10, true, None);
    let reference_display = main_frame_baseline(&engine, &extras);
    let initial_notes = reference_display.list.as_ref().unwrap().pages[3]
        .note_areas
        .clone();
    for (para_id, text) in [("73000002", "x"), ("73000008", "y")] {
        engine
            .doc()
            .insert_text(
                &crate::EditCtx::local("", ""),
                engine.doc().paragraph_mark_position(para_id).unwrap(),
                text,
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
        engine
            .layout_document_with_regions_retained(&request)
            .unwrap();
        let pagination = engine.pagination.borrow();
        assert!(pagination.last_incremental);
        assert!(!pagination.display_uses_region_path);
    }
    assert!(engine.pagination.borrow().note_changed_pages.is_empty());
    let (expected_bytes, expected_list, expected_rebuilt) =
        main_frame_oracle(&engine, &extras, reference_display);
    let before = engine.stats();
    let epoch = engine.display.borrow().binary_frame_epoch;
    assert_eq!(
        engine.build_display_list_frame(&extras, epoch).unwrap(),
        expected_bytes
    );
    assert_eq!(
        engine.with_display_list(Clone::clone).unwrap(),
        expected_list
    );
    assert_eq!(expected_list.pages[3].note_areas, initial_notes);
    assert_eq!(
        engine.stats().rebuilt_display_pages - before.rebuilt_display_pages,
        expected_rebuilt as u64
    );
}

#[test]
fn interactive_page_build_before_undo_matches_main_frame_bytes() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let (engine, request, extras) = interactive_damage_engine(font, 7, true, Some(0..6));
    let (resident_input, list) = {
        let pagination = engine.pagination.borrow();
        docx_layout::build_resident_display_list_partial_observed(
            pagination.input.as_ref().unwrap(),
            pagination.layout.as_ref().unwrap(),
            &extras,
            &|index| index < 6,
            &mut || {},
        )
        .unwrap()
    };
    let mut reference_display = {
        let display = engine.display.borrow();
        assert_eq!(display.list.as_ref().unwrap(), &list);
        assert!(list.pages[6].unbuilt);
        DisplayState {
            list: Some(list),
            resident_input: Some(resident_input),
            pages: display.pages.clone(),
            frame_epoch: display.frame_epoch,
            binary_frame_epoch: display.binary_frame_epoch,
            next_page_id: display.next_page_id,
            ..Default::default()
        }
    };
    let undo = crate::UndoSession::new();
    undo.track(engine.doc());
    engine
        .doc()
        .insert_text(
            &crate::EditCtx::local("", ""),
            engine.doc().paragraph_mark_position("73000002").unwrap(),
            "x",
            crate::FormatPolicy::Inherit,
        )
        .unwrap();
    engine
        .layout_document_with_regions_retained(&request)
        .unwrap();
    let built = {
        let pagination = engine.pagination.borrow();
        docx_layout::build_resident_display_pages(
            pagination.input.as_ref().unwrap(),
            pagination.layout.as_ref().unwrap(),
            reference_display.resident_input.as_mut().unwrap(),
            reference_display.list.as_mut().unwrap(),
            &[6],
        )
        .unwrap()
    };
    let epoch = reference_display.binary_frame_epoch;
    let epochs = FrameEpochs {
        doc_epoch: engine.doc_epoch(),
        layout_epoch: engine.pagination.borrow().layout_epoch,
        frame_epoch: reference_display.frame_epoch + 1,
        base_frame_epoch: epoch,
    };
    let expected_page_bytes = encode_frame_delta_changes(
        reference_display.list.as_ref().unwrap(),
        &mut reference_display.pages,
        epochs,
        DisplayChanges {
            rebuilt: &built,
            repositioned: &[],
            shifts: &[],
        },
    )
    .unwrap();
    assert_eq!(
        engine.build_display_pages_frame(&[6], epoch).unwrap(),
        expected_page_bytes
    );
    reference_display.frame_epoch = epochs.frame_epoch;
    reference_display.binary_frame_epoch = epochs.frame_epoch;
    assert!(undo.undo());
    assert_eq!(
        engine.region_relayout_trigger(&request).unwrap(),
        RelayoutTrigger::Interactive
    );
    engine
        .layout_document_with_regions_retained(&request)
        .unwrap();
    assert!(engine.pagination.borrow().display_rebuilt_pages.is_empty());
    let (expected_bytes, expected_list, _) = main_frame_oracle(&engine, &extras, reference_display);
    assert_eq!(
        engine
            .build_display_list_frame(&extras, epochs.frame_epoch)
            .unwrap(),
        expected_bytes
    );
    assert_eq!(
        engine.with_display_list(Clone::clone).unwrap(),
        expected_list
    );
}

#[test]
fn interactive_layouts_past_the_region_damage_limit_match_main_frames() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let page_count = MAX_RETAINED_DISPLAY_REBUILT_PAGES + 2;
    let (engine, request, extras) = interactive_damage_engine(font, page_count, false, None);
    let mut reference_display = main_frame_baseline(&engine, &extras);
    for (text, each_layout) in [("B", false), ("C", true)] {
        let before = engine.stats();
        for page in 1..=page_count {
            let mark = engine
                .doc()
                .paragraph_mark_position(&format!("{:08X}", 0x7300_0000 + page))
                .unwrap();
            engine
                .doc()
                .replace_range(
                    &crate::EditCtx::local("", ""),
                    crate::StoryRange::new("body", mark.index - 1, mark.index),
                    text,
                )
                .unwrap();
            if each_layout {
                engine
                    .layout_document_with_regions_retained(&request)
                    .unwrap();
            }
        }
        if !each_layout {
            engine
                .layout_document_with_regions_retained(&request)
                .unwrap();
        }
        {
            let pagination = engine.pagination.borrow();
            assert!(pagination.last_incremental);
            assert!(!pagination.display_full_rebuild);
            assert!(!pagination.display_uses_region_path);
            assert!(pagination.display_rebuilt_pages.is_empty());
        }
        assert_eq!(engine.stats().display_builds, before.display_builds);
        let (expected_bytes, expected_list, expected_rebuilt) =
            main_frame_oracle(&engine, &extras, std::mem::take(&mut reference_display));
        let epoch = engine.display.borrow().binary_frame_epoch;
        assert_eq!(
            engine.build_display_list_frame(&extras, epoch).unwrap(),
            expected_bytes
        );
        assert_eq!(
            engine.with_display_list(Clone::clone).unwrap(),
            expected_list
        );
        assert_eq!(
            engine.stats().incremental_display_builds,
            before.incremental_display_builds + 1
        );
        assert_eq!(
            engine.stats().rebuilt_display_pages - before.rebuilt_display_pages,
            expected_rebuilt as u64
        );
        if !each_layout {
            reference_display = main_frame_baseline(&engine, &extras);
        }
    }
}
