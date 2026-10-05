use super::*;

fn assert_preview_refresh_outputs(
    engine: &EngineSession,
    oracle: &EngineSession,
    request: &serde_json::Value,
    trigger: RelayoutTrigger,
) {
    let env: RenderEnv = serde_json::from_value(request["renderEnv"].clone()).unwrap();
    for session in [engine, oracle] {
        preview_mapped_oracle(session, &env);
        session
            .layout_regions_for_trigger(&request.to_string(), None, trigger)
            .unwrap();
    }
    let lowered = |session: &EngineSession| {
        let render = session.render.borrow();
        let story = &render.stories["body"];
        preview_mapped_snapshot(&story.blocks, &story.map, &story.revealable_blocks)
    };
    assert_eq!(lowered(engine), lowered(oracle));
    assert_eq!(
        engine.retained_layout_json().unwrap(),
        oracle.retained_layout_json().unwrap()
    );
    assert_eq!(
        engine.retained_kernel_inputs_json().unwrap(),
        oracle.retained_kernel_inputs_json().unwrap()
    );
    let extras = json!({
        "fontChains": request["measurement"]["fontChains"],
        "headersFooters": engine.regions.borrow().as_ref().unwrap().headers_footers,
    })
    .to_string();
    let frame = |session: &EngineSession| {
        let epoch = session.display.borrow().binary_frame_epoch;
        session.build_display_list_frame(&extras, epoch).unwrap()
    };
    assert_eq!(frame(engine), frame(oracle));
    assert_eq!(
        engine.with_display_list(Clone::clone).unwrap(),
        oracle.with_display_list(Clone::clone).unwrap()
    );
}

fn preview_refresh_edit_oracle(bytes: &[u8], font: u32) {
    let source = EngineSession::new(75270);
    crate::seed_from_docx(source.doc(), bytes).unwrap();
    let state = source.doc().encode_state_as_update_v1();
    let engine = EngineSession::new(75271);
    let oracle = EngineSession::new(75271);
    oracle.preview_refresh.set(false);
    for session in [&engine, &oracle] {
        session.doc().apply_update_v1(&state).unwrap();
        *session.doc().metadata.lock().unwrap() = source.doc().source_metadata();
        session
            .lower_story_json(
                "body",
                &RenderEnv {
                    show_hidden_text: true,
                    ..RenderEnv::default()
                },
            )
            .unwrap();
    }
    let mut request = lowering_pages::region_request(&engine, bytes, font);
    assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Open);
    let positions = {
        let render = engine.render.borrow();
        let story = &render.stories["body"];
        story.map.spans.iter()
            .filter(|span| {
                !span.atom && span.raw_end > span.raw_start
                    && story.map.paragraphs[span.paragraph as usize].0 == 0
            })
            .map(|span| span.raw_start)
            .collect::<Vec<_>>()
    };
    assert!(!positions.is_empty());
    for raw in [
        positions[0],
        positions[positions.len() / 2],
        *positions.last().unwrap(),
    ] {
        for text in [Some("x"), Some("y"), None, None] {
            for session in [&engine, &oracle] {
                session
                    .edit_resident_text(
                        crate::StoryRange::new("body", raw, raw + u32::from(text.is_none())),
                        text,
                        false,
                    )
                    .unwrap();
            }
            assert_preview_refresh_outputs(
                &engine, &oracle, &request, RelayoutTrigger::Interactive,
            );
        }
        for id in preview_fixture::ids(&engine) {
            for decision in [RevisionPreview::Accepted, RevisionPreview::Rejected] {
                request["renderEnv"] = serde_json::to_value(
                    RenderEnv::default().with_revision_preview(&id, decision),
                )
                .unwrap();
                let before = engine.stats();
                let expected_before = oracle.stats();
                assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Preview);
                assert_eq!(
                    engine.stats().lower_preview_patches - before.lower_preview_patches,
                    oracle.stats().lower_preview_patches - expected_before.lower_preview_patches
                );
            }
        }
        request["renderEnv"] = json!({});
        assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Preview);
    }
}

#[test]
fn preview_refresh_plain_edits_match_recorded_outputs() {
    let _open = OpenSwitch::new(Some(false));
    let _preview = PreviewSwitch::new(Some(false));
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font(LIBERATION).unwrap();
    for bytes in [
        preview_fixture::document(&preview_fixture::paragraph(1, &preview_fixture::run("Plain"))),
        preview_fixture::plain(),
        preview_fixture::nested(),
        preview_fixture::breaks(),
        preview_fixture::drawings(),
        preview_fixture::fields(),
        preview_fixture::hidden_fields(),
        preview_fixture::sequence(),
    ] {
        preview_refresh_edit_oracle(&bytes, font);
    }
}

#[test]
fn preview_refresh_plain_edits_match_recorded_corpus_outputs() {
    let _open = OpenSwitch::new(Some(false));
    let _preview = PreviewSwitch::new(Some(false));
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font(LIBERATION).unwrap();
    for (_, bytes) in preview_fixture::corpus() {
        preview_refresh_edit_oracle(bytes, font);
    }
}

#[test]
fn preview_refresh_plain_edits_match_recorded_region_outputs() {
    let _open = OpenSwitch::new(Some(true));
    let _preview = PreviewSwitch::new(Some(true));
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font(LIBERATION).unwrap();
    preview_refresh_edit_oracle(&preview_fixture::plain(), font);
}

#[test]
fn preview_refresh_reuses_unedited_units_without_copying_chunks() {
    let body = (0..100)
        .map(|index| {
            let content = if index == 50 {
                preview_fixture::revision("ins", "1", &preview_fixture::run("changed"))
            } else {
                preview_fixture::run("Ordinary paragraph")
            };
            preview_fixture::paragraph(index + 1, &content)
        })
        .collect::<String>();
    let engine = preview_seeded(&preview_fixture::document(&body));
    preview_mapped_oracle(&engine, &RenderEnv::default());
    for text in [Some("x"), None] {
        engine
            .edit_resident_text(
                crate::StoryRange::new("body", 1, 1 + u32::from(text.is_none())),
                text,
                false,
            )
            .unwrap();
        preview_mapped_oracle(&engine, &RenderEnv::default());
        let work = recording_work(&engine);
        assert_eq!(work.copied_chunks, 0);
        assert_eq!(work.reused_units, 1);
        assert_eq!(work.chunks, 3);
        let (_, _, _, recorded) = crate::bridge::preview::lower_recorded(
            engine.doc(),
            "body",
            &RenderEnv::default(),
            &mut crate::bridge::local::LocalLowering::new(false),
            true,
        )
        .unwrap();
        let recorded = recorded.unwrap().work;
        assert!(recorded.chunks >= 200);
        assert!(recorded.copied_chunks > 0);
    }
    let raw = {
        let render = engine.render.borrow();
        render.stories["body"].map.spans.iter()
            .find(|span| span.paragraph == 50)
            .unwrap().raw_start + 1
    };
    for text in [Some("x"), None] {
        engine
            .edit_resident_text(
                crate::StoryRange::new("body", raw, raw + u32::from(text.is_none())),
                text,
                false,
            )
            .unwrap();
        preview_mapped_oracle(&engine, &RenderEnv::default());
        let work = recording_work(&engine);
        assert_eq!(work.reused_units, 0);
        assert!(work.copied_chunks > 0);
        assert!(work.chunks < 10);
    }
    let before = engine.stats();
    preview_mapped_oracle(
        &engine,
        &RenderEnv::default().with_revision_preview("1", RevisionPreview::Rejected),
    );
    assert_eq!(engine.stats().lower_cache_misses, before.lower_cache_misses);
    assert_eq!(
        engine.stats().lower_preview_patches,
        before.lower_preview_patches + 1
    );
}

#[test]
fn preview_refresh_multiple_pending_edits_use_full_recording() {
    let engine = preview_seeded(&preview_fixture::plain());
    preview_mapped_oracle(&engine, &RenderEnv::default());
    for _ in 0..2 {
        engine
            .edit_resident_text(crate::StoryRange::new("body", 1, 1), Some("x"), false)
            .unwrap();
    }
    preview_mapped_oracle(&engine, &RenderEnv::default());
    let work = recording_work(&engine);
    assert_eq!(work.reused_units, 0);
    assert!(work.copied_chunks > 0);
    for decision in [RevisionPreview::Accepted, RevisionPreview::Rejected] {
        preview_mapped_oracle(
            &engine,
            &RenderEnv::default().with_revision_preview("1", decision),
        );
    }
}

#[test]
fn preview_refresh_without_revisions_skips_recorder_callbacks() {
    let bytes = preview_fixture::document(&preview_fixture::paragraph(
        1,
        &preview_fixture::run("Ordinary text"),
    ));
    let engine = preview_seeded(&bytes);
    preview_mapped_oracle(&engine, &RenderEnv::default());
    for text in [Some("x"), None] {
        engine
            .edit_resident_text(
                crate::StoryRange::new("body", 1, 1 + u32::from(text.is_none())),
                text,
                false,
            )
            .unwrap();
        preview_mapped_oracle(&engine, &RenderEnv::default());
        let work = recording_work(&engine);
        assert_eq!(work.chunks, 0);
        assert_eq!(work.copied_chunks, 0);
        assert_eq!(work.reused_units, 0);
    }
}

fn recording_work(engine: &EngineSession) -> crate::bridge::preview::RecordingWork {
    engine.render.borrow().stories["body"]
        .preview
        .as_ref()
        .unwrap()
        .work
        .clone()
}
