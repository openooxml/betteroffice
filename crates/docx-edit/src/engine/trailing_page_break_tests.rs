use super::{EngineSession, lowering_pages};
use crate::{StoryRange, bridge::RenderEnv, seed_from_docx};
use serde_json::Value;

fn document(text: &str) -> Vec<u8> {
    let body = format!(
        "{}{}<w:sectPr/>",
        lowering_pages::p("71000000", &lowering_pages::r("A")),
        lowering_pages::p(
            "71000001",
            &format!(
                r#"{}<w:r><w:br w:type="page"/></w:r>"#,
                if text.is_empty() {
                    String::new()
                } else {
                    lowering_pages::r(text)
                }
            ),
        ),
    );
    let mut parts = ooxml_opc::unzip_parts(&lowering_pages::with_body(&body)).unwrap();
    for (name, bytes) in &mut parts {
        let root = match name.as_str() {
            "word/footnotes.xml" => "footnotes",
            "word/endnotes.xml" => "endnotes",
            _ => continue,
        };
        *bytes = format!(
            r#"<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
        )
        .into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_edit_matches_fresh(before: &str, after: &str, text: Option<&str>, pages: usize) {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let bytes = document(before);
    let engine = EngineSession::new(9410);
    engine.set_local_lowering(true);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let request = lowering_pages::region_request(&engine, &bytes, font).to_string();
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    assert_eq!(
        engine.stats().retained_pages,
        if before.is_empty() { 1 } else { 2 }
    );
    engine
        .edit_resident_text(
            StoryRange::new("body", 2, 2 + before.len() as u32),
            text,
            true,
        )
        .unwrap();
    engine
        .layout_document_with_regions_after_edit(&request, Some("body"))
        .unwrap();
    let actual: Value = serde_json::from_str(&engine.retained_layout_json().unwrap()).unwrap();
    assert_eq!(actual["layout"]["pages"].as_array().unwrap().len(), pages);

    let fresh = EngineSession::new(9411);
    seed_from_docx(fresh.doc(), &document(after)).unwrap();
    let expected: Value =
        serde_json::from_str(&fresh.layout_document_with_regions_json(&request).unwrap()).unwrap();
    assert_eq!(actual["layout"], expected["layout"]);
    assert_eq!(
        engine
            .lower_story_json("body", &RenderEnv::default())
            .unwrap(),
        fresh
            .lower_story_json("body", &RenderEnv::default())
            .unwrap()
    );
}

#[test]
fn typing_into_a_terminal_break_only_paragraph_matches_fresh_open() {
    assert_edit_matches_fresh("", "x", Some("x"), 2);
}

#[test]
fn emptying_a_terminal_text_and_break_paragraph_matches_fresh_open() {
    assert_edit_matches_fresh("x", "", None, 1);
}

const PAIR_TEXT: &str = "text on its own page";
const PAIRS: u32 = 3;

/// Which notes the package carries: none, the fixture's own notes (an endnote and page-foot
/// footnotes), or those notes with the footnotes placed at the document's end.
#[derive(Clone, Copy, Debug, PartialEq)]
enum Notes {
    None,
    Fixture,
    DocumentEndFootnotes,
}

fn break_pairs(notes: Notes) -> Vec<u8> {
    let body: String = (0..PAIRS)
        .map(|index| {
            format!(
                "{}{}",
                lowering_pages::p(
                    &format!("{:08X}", 0x7000_0000 + index * 2),
                    &lowering_pages::r(PAIR_TEXT)
                ),
                lowering_pages::p(
                    &format!("{:08X}", 0x7000_0001 + index * 2),
                    r#"<w:r><w:br w:type="page"/></w:r>"#
                ),
            )
        })
        .collect();
    let section = if notes == Notes::DocumentEndFootnotes {
        r#"<w:sectPr><w:footnotePr><w:pos w:val="docEnd"/></w:footnotePr></w:sectPr>"#
    } else {
        "<w:sectPr/>"
    };
    let bytes = lowering_pages::with_body(&format!("{body}{section}"));
    if notes != Notes::None {
        return bytes;
    }
    let mut parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    for (name, bytes) in &mut parts {
        let root = match name.as_str() {
            "word/footnotes.xml" => "footnotes",
            "word/endnotes.xml" => "endnotes",
            _ => continue,
        };
        *bytes = format!(
            r#"<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
        )
        .into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// A second replica joining `engine`'s current state the way collaboration peers do, with the
/// source package and its metadata, laid out from scratch.
fn fresh_peer(engine: &EngineSession, bytes: &[u8]) -> EngineSession {
    let peer = EngineSession::new(9413);
    let state = engine.doc().encode_state_as_update_v1();
    let metadata = engine.doc().encode_peer_metadata().unwrap();
    let prepared = peer
        .doc()
        .prepare_peer_bootstrap(&state, &metadata, Some(bytes.to_vec().into()))
        .unwrap();
    peer.doc().install_peer_bootstrap(prepared, 17).unwrap();
    peer.doc().apply_update_v1(&state).unwrap();
    peer
}

#[test]
fn edits_around_a_terminal_break_only_paragraph_match_a_fresh_peer() {
    let pair = PAIR_TEXT.len() as u32 + 3;
    let last = pair * (PAIRS - 1);
    let edits = [
        ("start", 0, false),
        ("middle", pair + 5, false),
        ("inside the last text paragraph", last + 10, false),
        (
            "end of the last text paragraph",
            last + PAIR_TEXT.len() as u32,
            false,
        ),
        (
            "final break-only paragraph",
            last + PAIR_TEXT.len() as u32 + 1,
            true,
        ),
    ];
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for notes in [Notes::None, Notes::Fixture, Notes::DocumentEndFootnotes] {
        let notes_keep_break = notes == Notes::DocumentEndFootnotes;
        let bytes = break_pairs(notes);
        for (name, index, keeps_break) in edits {
            for resident in [false, true] {
                let case = format!("notes={notes:?} resident={resident} edit at {name}");
                let engine = EngineSession::new(9412);
                engine.set_local_lowering(resident);
                seed_from_docx(engine.doc(), &bytes).unwrap();
                let request = lowering_pages::region_request(&engine, &bytes, font).to_string();
                engine
                    .layout_document_with_regions_retained_json(&request)
                    .unwrap();
                let pages = PAIRS as usize + usize::from(notes_keep_break);
                assert_eq!(engine.stats().retained_pages, pages, "{case}");
                if resident {
                    engine
                        .edit_resident_text(StoryRange::new("body", index, index), Some("x"), true)
                        .unwrap();
                    engine
                        .layout_document_with_regions_after_edit(&request, Some("body"))
                        .unwrap();
                } else {
                    engine
                        .doc()
                        .insert_text(
                            &crate::EditCtx::local("", ""),
                            crate::Position::new("body", index),
                            "x",
                            crate::FormatPolicy::Inherit,
                        )
                        .unwrap();
                    engine
                        .layout_document_with_regions_retained_json(&request)
                        .unwrap();
                }
                let actual: Value =
                    serde_json::from_str(&engine.retained_layout_json().unwrap()).unwrap();
                assert_eq!(
                    actual["layout"]["pages"].as_array().unwrap().len(),
                    pages + usize::from(keeps_break && !notes_keep_break),
                    "{case}"
                );
                let peer = fresh_peer(&engine, &bytes);
                let expected: Value = serde_json::from_str(
                    &peer.layout_document_with_regions_json(&request).unwrap(),
                )
                .unwrap();
                assert_eq!(actual["layout"], expected["layout"], "{case}");
                assert_eq!(
                    engine
                        .lower_story_json("body", &RenderEnv::default())
                        .unwrap(),
                    peer.lower_story_json("body", &RenderEnv::default())
                        .unwrap(),
                    "{case}"
                );
            }
        }
    }
}
