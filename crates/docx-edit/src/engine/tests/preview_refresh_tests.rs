use super::*;

fn assert_preview_refresh_record(
    session: &EngineSession,
    doc: &crate::EditingDoc,
    env: &RenderEnv,
) {
    let (blocks, map, revealable, recorded) = crate::bridge::preview::lower_recorded(
        doc,
        "body",
        env,
        &mut crate::bridge::local::LocalLowering::new(false),
        true,
    )
    .unwrap();
    let render = session.render.borrow();
    let story = &render.stories["body"];
    let recorded = recorded.unwrap();
    assert_eq!(
        story.preview.as_ref().unwrap().snapshot(session.doc()),
        recorded.snapshot(doc)
    );
    assert_eq!(
        preview_mapped_snapshot(&story.blocks, &story.map, &story.revealable_blocks),
        preview_mapped_snapshot(&blocks, &map, &revealable)
    );
}

fn assert_preview_refresh_outputs(
    engine: &EngineSession,
    oracle: &EngineSession,
    request: &serde_json::Value,
    trigger: RelayoutTrigger,
) {
    let env: RenderEnv = serde_json::from_value(request["renderEnv"].clone()).unwrap();
    for session in [engine, oracle] {
        preview_mapped_oracle(session, &env);
        assert_preview_refresh_record(session, session.doc(), &env);
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

fn preview_refresh_sessions(
    bytes: &[u8],
    font: u32,
) -> (EngineSession, EngineSession, serde_json::Value) {
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
    let request = lowering_pages::region_request(&engine, bytes, font);
    assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Open);
    (engine, oracle, request)
}

fn assert_preview_refresh_decisions(
    engine: &EngineSession,
    oracle: &EngineSession,
    request: &mut serde_json::Value,
    ids: &[String],
) {
    let env = request["renderEnv"].clone();
    for id in ids {
        for decision in [RevisionPreview::Accepted, RevisionPreview::Rejected] {
            let mut preview: RenderEnv = serde_json::from_value(env.clone()).unwrap();
            preview.revision_preview.insert(id.clone(), decision);
            request["renderEnv"] = serde_json::to_value(preview).unwrap();
            let before = engine.stats();
            let expected_before = oracle.stats();
            assert_preview_refresh_outputs(engine, oracle, request, RelayoutTrigger::Preview);
            assert_eq!(
                engine.stats().lower_preview_patches - before.lower_preview_patches,
                oracle.stats().lower_preview_patches - expected_before.lower_preview_patches
            );
        }
    }
    request["renderEnv"] = env;
    assert_preview_refresh_outputs(engine, oracle, request, RelayoutTrigger::Preview);
}

fn preview_refresh_edit_oracle(bytes: &[u8], font: u32) {
    let (engine, oracle, mut request) = preview_refresh_sessions(bytes, font);
    let positions = {
        let render = engine.render.borrow();
        let story = &render.stories["body"];
        story
            .map
            .spans
            .iter()
            .filter(|span| {
                !span.atom
                    && span.raw_end > span.raw_start
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
                &engine,
                &oracle,
                &request,
                RelayoutTrigger::Interactive,
            );
            assert_preview_refresh_decisions(
                &engine,
                &oracle,
                &mut request,
                &preview_fixture::ids(&engine),
            );
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
        preview_fixture::document(&preview_fixture::paragraph(
            1,
            &preview_fixture::run("Plain"),
        )),
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

fn preview_refresh_revision_document(kind: &str) -> Vec<u8> {
    preview_fixture::document(&format!(
        "{}{}{}{}",
        preview_fixture::paragraph(1, &preview_fixture::run("Before")),
        preview_fixture::paragraph(
            2,
            &preview_fixture::revision(kind, "1", &preview_fixture::run("changed")),
        ),
        preview_fixture::paragraph(
            3,
            &preview_fixture::revision("ins", "2", &preview_fixture::run("later")),
        ),
        preview_fixture::paragraph(4, &preview_fixture::run("After")),
    ))
}

fn assert_live_revision_previews(
    engine: &EngineSession,
    oracle: &EngineSession,
    request: &mut serde_json::Value,
    ids: &[String],
) {
    let before = engine.stats();
    assert_preview_refresh_decisions(engine, oracle, request, ids);
    assert_eq!(engine.stats().lower_cache_misses, before.lower_cache_misses);
    assert_eq!(
        engine.stats().lower_preview_patches - before.lower_preview_patches,
        2 * ids.len() as u64 + 1
    );
}

#[test]
fn preview_refresh_revision_unit_boundary_edits_match_recording() {
    let _open = OpenSwitch::new(Some(false));
    let _preview = PreviewSwitch::new(Some(false));
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font(LIBERATION).unwrap();
    for (case, text) in [
        ("start", "x"),
        ("end", "x"),
        ("inside", "x"),
        ("start", "😀"),
        ("end", "😀"),
    ] {
        let bytes = preview_refresh_revision_document("ins");
        let (engine, oracle, mut request) = preview_refresh_sessions(&bytes, font);
        let ranges = engine.render.borrow().stories["body"]
            .preview
            .as_ref()
            .unwrap()
            .raw_ranges();
        assert_eq!(ranges.len(), 2);
        assert_eq!(ranges[0].end, ranges[1].start);
        let raw = match case {
            "start" => ranges[0].start,
            "end" => ranges[0].end,
            "inside" => ranges[0].start + 1,
            _ => unreachable!(),
        };
        let length = engine.doc().story_len("body").unwrap();
        for session in [&engine, &oracle] {
            session
                .edit_resident_text(crate::StoryRange::new("body", raw, raw), Some(text), false)
                .unwrap();
            assert_eq!(
                session.doc().story_len("body").unwrap(),
                length + text.encode_utf16().count() as u32
            );
        }
        assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Interactive);
        let work = recording_work(&engine);
        assert_eq!(work.reused_units, 1, "{case}: {text}");
        assert!(work.copied_chunks > 0, "{case}: {text}");
        assert!(
            work.chunks < recording_work(&oracle).chunks,
            "{case}: {text}"
        );
        assert_live_revision_previews(
            &engine,
            &oracle,
            &mut request,
            &preview_fixture::ids(&engine),
        );
        assert_eq!(
            engine.doc().story_len("body").unwrap(),
            length + text.encode_utf16().count() as u32
        );
    }
}

#[test]
fn preview_refresh_empty_revision_text_matches_recording() {
    let _open = OpenSwitch::new(Some(false));
    let _preview = PreviewSwitch::new(Some(false));
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font(LIBERATION).unwrap();
    for kind in ["ins", "del"] {
        let bytes = preview_refresh_revision_document(kind);
        let (engine, oracle, mut request) = preview_refresh_sessions(&bytes, font);
        let ids = preview_fixture::ids(&engine);
        assert_eq!(ids, ["1", "2"]);
        let range = engine.render.borrow().stories["body"]
            .preview
            .as_ref()
            .unwrap()
            .raw_ranges()
            .remove(0);
        assert_eq!(range.end - range.start, 8);
        let length = engine.doc().story_len("body").unwrap();
        for session in [&engine, &oracle] {
            session
                .edit_resident_text(
                    crate::StoryRange::new("body", range.start, range.end - 1),
                    None,
                    false,
                )
                .unwrap();
            assert_eq!(session.doc().story_len("body").unwrap(), length - 7);
            assert!(
                session
                    .doc()
                    .list_revisions()
                    .unwrap()
                    .iter()
                    .all(|revision| revision.change.revision_id != "1")
            );
        }
        assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Interactive);
        let work = recording_work(&engine);
        assert_eq!(work.reused_units, 1);
        assert!(work.chunks > 0);
        assert!(work.chunks < recording_work(&oracle).chunks);
        assert_live_revision_previews(&engine, &oracle, &mut request, &ids);
    }
}

#[test]
fn preview_refresh_reuses_unedited_units_without_copying_chunks() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font(LIBERATION).unwrap();
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
    let bytes = preview_fixture::document(&body);
    let (engine, oracle, mut request) = preview_refresh_sessions(&bytes, font);
    for text in [Some("x"), None] {
        for session in [&engine, &oracle] {
            session
                .edit_resident_text(
                    crate::StoryRange::new("body", 1, 1 + u32::from(text.is_none())),
                    text,
                    false,
                )
                .unwrap();
        }
        assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Interactive);
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
        assert_live_revision_previews(
            &engine,
            &oracle,
            &mut request,
            &preview_fixture::ids(&engine),
        );
    }
    let raw = {
        let render = engine.render.borrow();
        render.stories["body"]
            .map
            .spans
            .iter()
            .find(|span| span.paragraph == 50)
            .unwrap()
            .raw_start
            + 1
    };
    for text in [Some("x"), None] {
        for session in [&engine, &oracle] {
            session
                .edit_resident_text(
                    crate::StoryRange::new("body", raw, raw + u32::from(text.is_none())),
                    text,
                    false,
                )
                .unwrap();
        }
        assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Interactive);
        let work = recording_work(&engine);
        assert_eq!(work.reused_units, 0);
        assert!(work.copied_chunks > 0);
        assert!(work.chunks < 10);
        assert_live_revision_previews(
            &engine,
            &oracle,
            &mut request,
            &preview_fixture::ids(&engine),
        );
    }
    let before = engine.stats();
    request["renderEnv"] = serde_json::to_value(
        RenderEnv::default().with_revision_preview("1", RevisionPreview::Rejected),
    )
    .unwrap();
    assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Preview);
    assert_eq!(engine.stats().lower_cache_misses, before.lower_cache_misses);
    assert_eq!(
        engine.stats().lower_preview_patches,
        before.lower_preview_patches + 1
    );
}

#[test]
fn preview_refresh_multiple_pending_edits_use_full_recording() {
    let _open = OpenSwitch::new(Some(false));
    let _preview = PreviewSwitch::new(Some(false));
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font(LIBERATION).unwrap();
    for edit_count in [2, 4] {
        let bytes = preview_fixture::plain();
        let (engine, oracle, mut request) = preview_refresh_sessions(&bytes, font);
        let (checkpoint, checkpoint_oracle, mut checkpoint_request) =
            preview_refresh_sessions(&bytes, font);
        let ranges = engine.render.borrow().stories["body"]
            .preview
            .as_ref()
            .unwrap()
            .raw_ranges();
        assert_eq!(ranges.len(), 2);
        let epoch = engine.render.borrow().stories["body"].doc_epoch;
        for (raw, text) in [
            (1, "x"),
            (1, "x"),
            (ranges[0].start + 5, "y"),
            (ranges[0].end + 3, "😀"),
        ]
        .into_iter()
        .take(edit_count)
        {
            for session in [&engine, &oracle, &checkpoint, &checkpoint_oracle] {
                session
                    .edit_resident_text(crate::StoryRange::new("body", raw, raw), Some(text), false)
                    .unwrap();
            }
            assert_preview_refresh_outputs(
                &checkpoint,
                &checkpoint_oracle,
                &checkpoint_request,
                RelayoutTrigger::Interactive,
            );
            assert_preview_refresh_record(&checkpoint, engine.doc(), &RenderEnv::default());
            assert!(recording_work(&checkpoint).reused_units > 0);
            assert_live_revision_previews(
                &checkpoint,
                &checkpoint_oracle,
                &mut checkpoint_request,
                &preview_fixture::ids(&checkpoint),
            );
            assert_eq!(engine.render.borrow().stories["body"].doc_epoch, epoch);
        }
        assert!(
            engine.render.borrow().stories["body"]
                .preview_edit
                .is_none()
        );
        assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Interactive);
        let work = recording_work(&engine);
        assert_eq!(work.reused_units, 0);
        assert!(work.copied_chunks > 0);
        for decision in [RevisionPreview::Accepted, RevisionPreview::Rejected] {
            request["renderEnv"] =
                serde_json::to_value(RenderEnv::default().with_revision_preview("1", decision))
                    .unwrap();
            assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Preview);
        }
        request["renderEnv"] = json!({});
        assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Preview);
        assert_live_revision_previews(
            &engine,
            &oracle,
            &mut request,
            &preview_fixture::ids(&engine),
        );
    }
}

#[test]
fn preview_refresh_without_revisions_skips_recorder_callbacks() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font(LIBERATION).unwrap();
    let bytes = preview_fixture::document(&preview_fixture::paragraph(
        1,
        &preview_fixture::run("Ordinary text"),
    ));
    let (engine, oracle, request) = preview_refresh_sessions(&bytes, font);
    for text in [Some("x"), None] {
        for session in [&engine, &oracle] {
            session
                .edit_resident_text(
                    crate::StoryRange::new("body", 1, 1 + u32::from(text.is_none())),
                    text,
                    false,
                )
                .unwrap();
        }
        assert_preview_refresh_outputs(&engine, &oracle, &request, RelayoutTrigger::Interactive);
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
