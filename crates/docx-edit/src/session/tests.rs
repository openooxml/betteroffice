use std::sync::Arc;

use serde_json::Value;

use super::{EditSessionCore, EngineErrorKind, Host};
use crate::{EditTextView, ReadParagraphsRequest, StoryRange, TriState};

const BODY: &str = r#"[{"storyId":"body","paragraphs":[{"text":"First"},{"text":"Second"}]}]"#;

fn core(client_id: u64) -> EditSessionCore {
    EditSessionCore::new(client_id, Host::deterministic())
}

#[test]
fn load_json_then_load_converges() {
    let source = core(1);
    source.load_json(BODY).unwrap();
    let loaded = core(2);
    loaded.load(&source.encode_state()).unwrap();
    assert_eq!(source.encode_state_vector(), loaded.encode_state_vector());
    assert_eq!(loaded.story_ids(), vec!["body".to_owned()]);
}

#[test]
fn diff_chunks_round_trip_converges() {
    let source = core(1);
    source.load_json(BODY).unwrap();
    let target = core(2);
    let chunks = source
        .encode_diff_chunks(&target.encode_state_vector(), 64)
        .unwrap();
    assert!(!chunks.is_empty());
    target.apply_update_chunks(&chunks).unwrap();
    assert_eq!(source.encode_state_vector(), target.encode_state_vector());
}

#[test]
fn update_events_carry_origin_byte() {
    let source = core(1);
    source.load_json(BODY).unwrap();
    let state = source.encode_state();

    let mut remote = core(2);
    remote.start_update_event_observation().unwrap();
    remote.apply_update(&state).unwrap();
    let event = remote.drain_update_event();
    assert_eq!(event.first(), Some(&1));
    assert!(event.len() > 1);
    assert!(remote.drain_update_event().is_empty());

    let mut local = core(3);
    local.start_update_event_observation().unwrap();
    local.apply_local_update(&state).unwrap();
    let event = local.drain_update_event();
    assert_eq!(event.first(), Some(&0));
    assert!(event.len() > 1);
}

#[test]
fn malformed_update_error_matches_shim_text() {
    let session = core(1);
    let error = session.apply_update(&[255, 255, 255]).unwrap_err();
    assert_eq!(error.kind, EngineErrorKind::Message);
    assert!(
        error.message.starts_with("invalid yrs update: "),
        "{}",
        error.message
    );
}

#[test]
fn version_nonce_depends_on_host_entropy() {
    let one = EditSessionCore::new(7, Host::with_entropy(Arc::new(|| 1)));
    let two = EditSessionCore::new(7, Host::with_entropy(Arc::new(|| 2)));
    assert_ne!(one.engine.doc().version(), two.engine.doc().version());
}

fn paragraph_texts(session: &EditSessionCore) -> Vec<String> {
    session
        .engine
        .doc()
        .read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: None,
            view: EditTextView::Accepted,
        })
        .unwrap()
        .paragraphs
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect()
}

#[test]
fn an_empty_read_answers_with_the_version_of_a_full_read() {
    fn read(
        session: &EditSessionCore,
        story: Option<&str>,
        para_ids: Option<Vec<String>>,
    ) -> Result<crate::ReadParagraphsResponse, crate::EditRefusal> {
        session
            .engine
            .doc()
            .read_paragraphs(&ReadParagraphsRequest {
                story: story.map(str::to_owned),
                para_ids,
                view: EditTextView::Accepted,
            })
    }
    let same = |session: &EditSessionCore| {
        let empty = read(session, None, Some(Vec::new())).unwrap();
        let full = read(session, None, None).unwrap();
        assert_eq!(empty.version, full.version);
        assert!(empty.paragraphs.is_empty());
        assert!(!full.paragraphs.is_empty());
        full.version
    };
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap().to_owned();
    session.track_undo();
    let opened = same(&session);
    session
        .insert_text("body", &para_id, 5, "!", None, None)
        .unwrap();
    let edited = same(&session);
    assert_ne!(edited, opened);
    assert!(session.undo());
    let undone = same(&session);
    assert_ne!(undone, edited);
    let remote = core(2);
    let doc = session.engine.doc();
    remote
        .engine
        .doc()
        .apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    remote
        .insert_text("body", &para_id, 0, "r", None, None)
        .unwrap();
    doc.apply_update_v1(&remote.engine.doc().encode_state_as_update_v1())
        .unwrap();
    assert_ne!(same(&session), undone);
    assert!(read(&session, Some("missing"), Some(Vec::new())).is_err());
    assert!(read(&session, Some("missing"), None).is_err());
}

#[test]
fn text_edits_marks_and_raw_ops_run_without_wasm() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    let inserted: Value = serde_json::from_str(
        &session
            .insert_text("body", para_id, 5, "!", None, None)
            .unwrap(),
    )
    .unwrap();
    assert!(inserted["revisionId"].is_null());
    assert_eq!(inserted["range"]["story"], "body");
    assert_eq!(paragraph_texts(&session), ["First!", "Second"]);
    let deleted: Value = serde_json::from_str(
        &session
            .delete_range("body", para_id, 5, para_id, 6, None, None)
            .unwrap(),
    )
    .unwrap();
    assert!(deleted["revisionId"].is_null());
    assert_eq!(deleted["range"]["story"], "body");
    assert_eq!(paragraph_texts(&session), ["First", "Second"]);

    session
        .toggle_mark(
            "body",
            para_id,
            0,
            para_id,
            5,
            r#"{"type":"bold"}"#,
            None,
            None,
        )
        .unwrap();
    assert_eq!(
        session
            .engine
            .doc()
            .selection_context(&StoryRange::new("body", 0, 5))
            .unwrap()
            .bold,
        TriState::On
    );
    session
        .apply_raw_ops("body", r#"[{"op":"insert","index":0,"text":"+"}]"#)
        .unwrap();
    assert_eq!(paragraph_texts(&session), ["+First", "Second"]);
}

#[test]
fn table_edits_run_without_wasm() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    let inserted: Value = serde_json::from_str(
        &session
            .insert_table("body", para_id, 5, 1, 2, None, None)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(inserted["rows"], 1);
    assert_eq!(inserted["columns"], 2);
    let row: Value = serde_json::from_str(
        &session
            .insert_row(
                r#"{"story":"body","tableIndex":0,"row":0,"column":0}"#,
                true,
                None,
                None,
            )
            .unwrap(),
    )
    .unwrap();
    assert_eq!(row["rows"], 2);
    assert_eq!(row["columns"], 2);
    let deleted: Value = serde_json::from_str(
        &session
            .delete_table(&inserted["table"].to_string())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(deleted["deletedTable"], true);
    assert_eq!(session.story_ids(), ["body"]);
}

#[test]
fn incomplete_suggesting_author_matches_shim_error() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    let before = session.encode_state();
    let error = session
        .insert_text("body", para_id, 0, "!", Some("author".to_owned()), None)
        .unwrap_err();
    assert_eq!(error.kind, EngineErrorKind::Message);
    assert_eq!(
        error.message,
        "suggesting requires both an author name and an ISO date"
    );
    assert_eq!(session.encode_state(), before);
}

#[test]
fn selection_and_undo_run_without_wasm() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    session
        .set_selection("body", para_id, 1, para_id, 4)
        .unwrap();
    let selection: Value = serde_json::from_str(&session.selection().unwrap()).unwrap();
    assert_eq!(
        selection,
        serde_json::json!({
            "anchor": { "story": "body", "paraId": para_id, "offset": 1 },
            "head": { "story": "body", "paraId": para_id, "offset": 4 },
        })
    );

    session.track_undo();
    session
        .insert_text("body", para_id, 5, "!", None, None)
        .unwrap();
    assert_eq!(paragraph_texts(&session), ["First!", "Second"]);
    assert!(session.undo());
    assert_eq!(paragraph_texts(&session), ["First", "Second"]);
    assert!(session.can_redo());
}

#[test]
fn comments_reads_and_identities_run_without_wasm() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    let ranges = serde_json::json!([{
        "story": "body", "startPara": para_id, "startOffset": 0,
        "endPara": para_id, "endOffset": 5,
    }]);
    let added: Value = serde_json::from_str(
        &session
            .add_comment(&ranges.to_string(), "Reviewer", "2026-01-01", r#""Note""#)
            .unwrap(),
    )
    .unwrap();
    let comments: Value = serde_json::from_str(&session.list_comments().unwrap()).unwrap();
    assert_eq!(comments.as_array().unwrap().len(), 1);
    assert_eq!(comments[0]["id"], added["commentId"]);
    assert_eq!(comments[0]["body"], "Note");

    let read: Value = serde_json::from_str(
        &session
            .read_paragraphs_json(r#"{"story":"body","view":"accepted"}"#)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(read["ok"], true);
    let paragraphs = read["paragraphs"].as_array().unwrap();
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(paragraphs[0]["text"], "First");
    assert_eq!(paragraphs[1]["text"], "Second");
    let identities: Value = serde_json::from_str(&session.paragraph_identities().unwrap()).unwrap();
    let identities = identities["paragraphs"].as_array().unwrap();
    assert_eq!(identities.len(), paragraphs.len());
    for (identity, paragraph) in identities.iter().zip(paragraphs) {
        assert_eq!(identity["session"]["paraId"], paragraph["paraId"]);
    }
}

#[test]
fn unknown_selection_story_matches_shim_error() {
    let error = core(1)
        .set_selection("missing", "p1", 0, "p1", 0)
        .unwrap_err();
    assert_eq!(error.kind, EngineErrorKind::Message);
    assert_eq!(error.message, "story \"missing\" was not found");
}

#[test]
fn a_resident_frame_paints_an_arabic_word_from_its_measured_glyphs() {
    let session = core(1);
    session
        .load_json(r#"[{"storyId":"body","paragraphs":[{"text":"مرحبا بالعالم"}]}]"#)
        .unwrap();
    let font = session
        .register_measure_font(include_bytes!(
            "../../../../packages/fonts/assets/NotoNaskhArabic-Regular.ttf"
        ))
        .unwrap();
    let request = serde_json::json!({
        "bodyStory": "body",
        "regions": {"sections": [{
            "sectionId": "main",
            "properties": {
                "pageWidth": 4320,
                "pageHeight": 2880,
                "marginTop": 300,
                "marginRight": 300,
                "marginBottom": 300,
                "marginLeft": 300
            }
        }]},
        "measurement": {
            "fontChains": {"calibri|0|0": [font]},
            "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    });
    session
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    session
        .build_display_list_frame(
            &serde_json::json!({"fontChains": {"calibri|0|0": [font]}}).to_string(),
            0.0,
        )
        .unwrap();
    let texts = session
        .engine
        .with_display_list(|list| {
            list.pages[0]
                .primitives
                .iter()
                .filter_map(|primitive| match primitive {
                    docx_layout::display_list::Primitive::GlyphRun(run) => Some(run.text.clone()),
                    _ => None,
                })
                .collect::<Vec<_>>()
        })
        .unwrap();
    assert_eq!(texts, ["مرحبا بالعالم"]);
}

#[test]
fn fonts_layout_frames_and_resident_input_run_without_wasm() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    let ordinary = core(2);
    ordinary.load(&session.encode_state()).unwrap();
    let font = session
        .register_measure_font(include_bytes!(
            "../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
    let request = serde_json::json!({
        "bodyStory": "body",
        "regions": {"sections": [{
            "sectionId": "main",
            "properties": {
                "pageWidth": 4320,
                "pageHeight": 2880,
                "marginTop": 300,
                "marginRight": 300,
                "marginBottom": 300,
                "marginLeft": 300
            }
        }]},
        "measurement": {
            "fontChains": {"calibri|0|0": [font]},
            "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    });
    let layout: Value = serde_json::from_str(
        &session
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    assert!(!layout["layout"]["pages"].as_array().unwrap().is_empty());
    let frame = session
        .build_display_list_frame(
            &serde_json::json!({"fontChains": {"calibri|0|0": [font]}}).to_string(),
            0.0,
        )
        .unwrap();
    assert!(!frame.is_empty());
    session
        .set_selection("body", para_id, 2, para_id, 2)
        .unwrap();
    let caret: Value =
        serde_json::from_str(&session.resident_caret_snapshot_json().unwrap()).unwrap();
    let rect = &caret["caretRect"];
    let hit: Value = serde_json::from_str(
        &session
            .display_hit_test_regions_json(
                rect["pageIndex"].as_u64().unwrap() as u32,
                rect["x"].as_f64().unwrap(),
                rect["y"].as_f64().unwrap() + rect["height"].as_f64().unwrap() / 2.0,
            )
            .unwrap(),
    )
    .unwrap();
    assert_eq!(hit["region"], "body");
    assert_eq!(hit["target"], "text");
    let hit_pos = hit["pos"].as_i64().unwrap();
    let hit_caret = session
        .engine
        .with_display_list(|list| docx_layout::hit::caret_rect(list, hit_pos))
        .unwrap()
        .unwrap();
    let hit_caret = serde_json::to_value(hit_caret).unwrap();
    for key in ["pageIndex", "x", "y", "height"] {
        assert_eq!(hit_caret[key], rect[key], "{key}");
    }

    let receipt: Value = serde_json::from_str(
        &ordinary
            .insert_text("body", para_id, 2, "!", None, None)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(
        receipt,
        serde_json::json!({
            "revisionId": null,
            "range": {
                "story": "body",
                "start": {"paraId": para_id, "offset": 2},
                "end": {"paraId": para_id, "offset": 3}
            }
        })
    );
    let frame = session
        .apply_input("!", caret["frameEpoch"].as_f64().unwrap())
        .unwrap();
    assert!(!frame.is_empty());
    assert_eq!(paragraph_texts(&session), ["Fi!rst", "Second"]);
    assert_eq!(paragraph_texts(&session), paragraph_texts(&ordinary));
    let selection: Value = serde_json::from_str(&session.selection().unwrap()).unwrap();
    assert_eq!(selection["head"]["story"], receipt["range"]["story"]);
    assert_eq!(
        selection["head"]["paraId"],
        receipt["range"]["end"]["paraId"]
    );
    assert_eq!(
        selection["head"]["offset"],
        receipt["range"]["end"]["offset"]
    );
    assert_eq!(selection["anchor"], selection["head"]);
}

#[test]
fn invalid_resident_frame_epoch_matches_shim_error_before_mutation() {
    let session = core(1);
    session.load_json(BODY).unwrap();
    let before = session.encode_state();
    let error = session.apply_input("!", -1.0).unwrap_err();
    assert_eq!(error.kind, EngineErrorKind::Message);
    assert_eq!(
        error.message,
        "expected_frame_epoch must be a non-negative safe integer"
    );
    assert_eq!(session.encode_state(), before);
}

/// A session whose body table names cell (0, 2) while its story is missing: one replica split
/// a cell, another inserted a row, the first undid its split, after the row reached it when
/// `synced`. With `select_early` the session placed its caret in that cell while it still had
/// its story.
fn missing_cell_session(synced: bool, select_early: bool) -> (EditSessionCore, String) {
    let ctx = crate::EditCtx::local("Ada", "2026-11-10T00:00:00Z");
    let origin = crate::EditingDoc::new(100);
    origin
        .create_story("body", "Hello", "Normal", "left")
        .unwrap();
    origin
        .insert_table(&ctx, crate::Position::new("body", 0), 1, 2)
        .unwrap();
    let state = origin.encode_state_as_update_v1();
    let open = |client| {
        let doc = crate::EditingDoc::new(client);
        doc.apply_update_v1(&state).unwrap();
        doc
    };
    let (splitter, peer) = (open(91), open(92));
    let exchange = || {
        let from_splitter = splitter.encode_state_as_update_v1();
        let from_peer = peer.encode_state_as_update_v1();
        peer.apply_update_v1(&from_splitter).unwrap();
        splitter.apply_update_v1(&from_peer).unwrap();
    };
    let mut undo = splitter.undo_manager();
    splitter
        .split_cell_grid(
            &ctx,
            &crate::CellLoc::new("body", 0, 0, 1),
            Some(1),
            Some(2),
        )
        .unwrap();
    exchange();
    let session = core(93);
    if select_early {
        session
            .apply_update(&splitter.encode_state_as_update_v1())
            .unwrap();
        let payload = session
            .engine
            .doc()
            .table_payload(&crate::TableLocator::new("body", 0))
            .unwrap();
        let rows = serde_json::to_value(&payload["rows"]).unwrap();
        let cell = rows[0]["cells"][2]["story"].as_str().unwrap().to_owned();
        let paragraphs = session.engine.doc().paragraphs(&cell).unwrap();
        let para = paragraphs[0].para_id.clone();
        session.set_selection(&cell, &para, 0, &para, 0).unwrap();
    }
    peer.insert_row(&ctx, &crate::CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    if synced {
        exchange();
    }
    assert!(undo.undo());
    exchange();
    session
        .apply_update(&splitter.encode_state_as_update_v1())
        .unwrap();
    let payload = session
        .engine
        .doc()
        .table_payload(&crate::TableLocator::new("body", 0))
        .unwrap();
    let rows = serde_json::to_value(&payload["rows"]).unwrap();
    let cell = rows[0]["cells"][2]["story"].as_str().unwrap().to_owned();
    assert!(!session.story_ids().contains(&cell), "{cell}");
    (session, cell)
}

#[test]
fn every_write_aimed_at_a_cell_without_its_story_is_refused_before_its_argument_checks() {
    for synced in [true, false] {
        refuses_every_write(synced);
    }
}

fn refuses_every_write(synced: bool) {
    let (session, cell) = missing_cell_session(synced, false);
    let key = format!("{cell}#p0");
    let collapsed = serde_json::json!([{
        "story": cell, "startPara": key, "startOffset": 0, "endPara": key, "endOffset": 0,
    }])
    .to_string();
    let before = (session.encode_state_vector(), session.encode_state());
    let (c, k) = (cell.as_str(), key.as_str());
    let results: Vec<(&str, Result<(), super::EngineError>)> = vec![
        (
            "insert_text with a break",
            session.insert_text(c, k, 0, "a\nb", None, None).map(drop),
        ),
        (
            "edit_text_resident",
            session
                .edit_text_resident(c, k, 0, k, 0, Some("x".to_owned()))
                .map(drop),
        ),
        (
            "edit_text_resident over a range",
            session.edit_text_resident(c, k, 0, k, 0, None).map(drop),
        ),
        (
            "collapsed delete_range",
            session.delete_range(c, k, 0, k, 0, None, None).map(drop),
        ),
        (
            "collapsed replace_range with no text",
            session
                .replace_range(c, k, 0, k, 0, "", None, None)
                .map(drop),
        ),
        (
            "collapsed toggle_mark",
            session
                .toggle_mark(c, k, 0, k, 0, r#"{"type":"bold"}"#, None, None)
                .map(drop),
        ),
        (
            "collapsed format_range",
            session
                .format_range(c, k, 0, k, 0, r#"{"bold":true}"#, None, None)
                .map(drop),
        ),
        (
            "collapsed set_hyperlink",
            session
                .set_hyperlink(c, k, 0, k, 0, r#"{"url":"x"}"#, None, None)
                .map(drop),
        ),
        (
            "collapsed clear_formatting",
            session
                .clear_formatting(c, k, 0, k, 0, None, None)
                .map(drop),
        ),
        (
            "merge_paragraphs on the last paragraph",
            session.merge_paragraphs(c, k, None, None).map(drop),
        ),
        (
            "insert_image with a malformed payload",
            session
                .insert_image(c, k, 0, "not json", None, None)
                .map(drop),
        ),
        (
            "insert_section_break",
            session.insert_section_break(c, k, 0, "nextPage"),
        ),
        (
            "insert_watermark",
            session.insert_watermark(c, k, 0, r#"{"text":"DRAFT"}"#),
        ),
        ("insert_page_break", session.insert_page_break(c, k, 0)),
        (
            "collapsed add_comment",
            session
                .add_comment(&collapsed, "Ada", "2026-11-10T00:00:00Z", "\"note\"")
                .map(drop),
        ),
        (
            "collapsed set_comment_ranges",
            session.set_comment_ranges("missing", &collapsed),
        ),
        ("apply_raw_ops", session.apply_raw_ops(c, "not json")),
    ];
    for (what, result) in results {
        let error = result.expect_err(what);
        assert!(error.message.contains("has no story"), "{what}: {error}");
        assert_eq!(error.kind, EngineErrorKind::CellHasNoStory, "{what}");
    }
    assert_eq!(
        (session.encode_state_vector(), session.encode_state()),
        before
    );
}

#[test]
fn a_caret_placed_before_the_cell_lost_its_story_reads_at_its_empty_paragraph() {
    for synced in [true, false] {
        let (session, cell) = missing_cell_session(synced, true);
        let before = (session.encode_state_vector(), session.encode_state());
        let selection: Value = serde_json::from_str(&session.selection().unwrap()).unwrap();
        for end in ["anchor", "head"] {
            assert_eq!(selection[end]["story"], cell.as_str(), "synced={synced}");
            assert_eq!(selection[end]["paraId"], format!("{cell}#p0"));
            assert_eq!(selection[end]["offset"], 0);
        }
        assert_eq!(
            (session.encode_state_vector(), session.encode_state()),
            before
        );
    }
}

#[test]
fn loading_stories_that_include_a_cell_without_its_story_seeds_none_of_them() {
    for synced in [true, false] {
        let (session, cell) = missing_cell_session(synced, false);
        let before = (session.encode_state_vector(), session.encode_state());
        let request = serde_json::json!([
            {"storyId": "fresh", "paragraphs": [{"text": "first"}]},
            {"storyId": cell, "paragraphs": [{"text": "second"}]},
        ])
        .to_string();
        let error = session.load_json(&request).unwrap_err();
        assert_eq!(error.kind, EngineErrorKind::CellHasNoStory, "{error}");
        assert!(!session.story_ids().contains(&"fresh".to_owned()));
        assert_eq!(
            (session.encode_state_vector(), session.encode_state()),
            before
        );
    }
}

#[test]
fn resident_input_and_delete_in_a_cell_without_its_story_refuse_with_its_kind() {
    for synced in [true, false] {
        let (session, cell) = missing_cell_session(synced, true);
        let key = format!("{cell}#p0");
        let before = (session.encode_state_vector(), session.encode_state());
        let results: Vec<(&str, Result<(), super::EngineError>)> = vec![
            ("apply_input", session.apply_input("x", 0.0).map(drop)),
            (
                "apply_input_profiled",
                session.apply_input_profiled("x", 0.0).map(drop),
            ),
            (
                "apply_delete",
                session.apply_delete("backward", 0.0, 1).map(drop),
            ),
            (
                "apply_delete_profiled",
                session.apply_delete_profiled("forward", 0.0, 1).map(drop),
            ),
            (
                "edit_text_resident",
                session
                    .edit_text_resident(&cell, &key, 0, &key, 0, Some("x".to_owned()))
                    .map(drop),
            ),
        ];
        for (what, result) in results {
            let error = result.expect_err(what);
            assert_eq!(error.kind, EngineErrorKind::CellHasNoStory, "{what}");
        }
        assert_eq!(
            (session.encode_state_vector(), session.encode_state()),
            before
        );
    }
}

#[test]
fn a_cell_refusal_keeps_its_kind_through_a_plain_message() {
    let refusal = crate::OpError::CellHasNoStory("body:t0:r0c2~7".to_owned());
    assert_eq!(
        super::EngineError::message(refusal).kind,
        EngineErrorKind::CellHasNoStory
    );
    let missing = crate::EditError::StoryNotFound("table cell \"x\" has no story".to_owned());
    assert_eq!(
        super::EngineError::message(missing).kind,
        EngineErrorKind::Message
    );
}
