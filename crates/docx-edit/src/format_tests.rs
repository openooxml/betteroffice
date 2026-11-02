use super::*;
use crate::{ChangeKind, ChangeTarget, EngineSession, seed_from_docx};
use serde_json::json;

fn document(body: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
        ("word/document.xml".to_owned(), format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#).into_bytes()),
    ]).unwrap()
}

fn engine() -> EngineSession {
    let engine = EngineSession::new(9301);
    seed_from_docx(
        engine.doc(),
        &document("<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>abcdef</w:t></w:r></w:p>"),
    )
    .unwrap();
    engine
}

fn suggesting() -> EditCtx {
    EditCtx::local("Ada", "2026-10-01T12:00:00Z").suggesting()
}

fn attrs(engine: &EngineSession) -> Vec<(String, Value)> {
    engine
        .doc()
        .story_segments("body")
        .unwrap()
        .into_iter()
        .filter_map(|segment| {
            let crate::SegmentContent::Text(text) = segment.content else {
                return None;
            };
            Some((
                text.to_string(),
                serde_json::to_value(segment.attributes).unwrap(),
            ))
        })
        .collect()
}

#[test]
fn suggested_format_captures_previous_run_properties() {
    let engine = engine();
    let receipt = engine
        .doc()
        .toggle_format(
            &suggesting(),
            StoryRange::new("body", 0, 6),
            SimpleFormat::Bold,
        )
        .unwrap();
    assert_eq!(receipt.revision_ids.len(), 1);
    let runs = attrs(&engine);
    assert_eq!(runs[0].1["bold"], true);
    let change = &runs[0].1["rPrChange"][0];
    assert_eq!(change["previousFormatting"]["italic"], true);
    assert!(change["previousFormatting"]["bold"].is_null());
    assert_eq!(change["info"]["author"], "Ada");
    assert_eq!(change["info"]["date"], "2026-10-01T12:00:00Z");
    assert_eq!(
        engine.doc().list_revisions().unwrap()[0].change.kind,
        ChangeKind::RunPropertiesChanged
    );
}

#[test]
fn partial_and_repeated_format_keep_original_properties_on_both_sides() {
    let engine = engine();
    engine
        .doc()
        .toggle_format(
            &suggesting(),
            StoryRange::new("body", 1, 5),
            SimpleFormat::Bold,
        )
        .unwrap();
    engine
        .doc()
        .clear_formatting(&suggesting(), StoryRange::new("body", 2, 4))
        .unwrap();
    let runs = attrs(&engine);
    assert_eq!(
        runs.iter().map(|run| run.0.as_str()).collect::<Vec<_>>(),
        ["a", "b", "cd", "e", "f"]
    );
    assert!(runs[0].1["rPrChange"].is_null());
    assert!(runs[4].1["rPrChange"].is_null());
    for run in &runs[1..4] {
        assert_eq!(
            run.1["rPrChange"][0]["previousFormatting"],
            json!({"italic": true})
        );
    }
}

#[test]
fn accept_and_reject_run_properties_by_id_and_range() {
    for accept in [true, false] {
        for by_id in [true, false] {
            let engine = engine();
            let receipt = engine
                .doc()
                .toggle_format(
                    &suggesting(),
                    StoryRange::new("body", 0, 6),
                    SimpleFormat::Bold,
                )
                .unwrap();
            let target = if by_id {
                ChangeTarget::Revision(receipt.revision_ids[0].clone())
            } else {
                ChangeTarget::Range(StoryRange::new("body", 0, 6))
            };
            if accept {
                engine.doc().accept_change(&suggesting(), &target).unwrap();
            } else {
                engine.doc().reject_change(&suggesting(), &target).unwrap();
            }
            assert!(engine.doc().list_revisions().unwrap().is_empty());
            let runs = attrs(&engine);
            assert_eq!(runs[0].1["bold"].as_bool().unwrap_or(false), accept);
            assert_eq!(runs[0].1["italic"], true);
            assert!(runs[0].1["rPrChange"].is_null());
        }
    }
}

#[test]
fn suggested_hyperlink_can_be_rejected() {
    let engine = engine();
    let receipt = engine
        .doc()
        .set_hyperlink(
            &suggesting(),
            StoryRange::new("body", 1, 5),
            Some(Any::from_json(r##"{"href":"#target"}"##).unwrap()),
        )
        .unwrap();
    assert!(!receipt.revision_ids.is_empty());
    engine
        .doc()
        .reject_change(
            &suggesting(),
            &ChangeTarget::Revision(receipt.revision_ids[0].clone()),
        )
        .unwrap();
    assert!(
        attrs(&engine)
            .iter()
            .all(|run| run.1["hyperlink"].is_null())
    );
}

#[test]
fn generated_run_change_serializes_and_reopens() {
    let engine = engine();
    engine
        .doc()
        .toggle_format(
            &suggesting(),
            StoryRange::new("body", 0, 6),
            SimpleFormat::Bold,
        )
        .unwrap();
    let runs = attrs(&engine);
    let mut change = runs[0].1["rPrChange"][0].clone();
    change["info"]["id"] = json!(42);
    let run = serde_json::from_value(json!({"type":"run", "formatting":{"bold":true,"italic":true}, "propertyChanges":[change], "content":[{"type":"text","text":"abcdef"}]})).unwrap();
    let mut context = docx_parse::serializer::SerializerContext::new(
        &docx_parse::serializer::SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2026-10-01T12:00:00Z".to_owned(),
        },
    )
    .unwrap();
    let xml = docx_parse::serializer::serialize_run(&run, &mut context).unwrap();
    assert!(xml.contains("w:rPrChange"));
    let reopened = EngineSession::new(9302);
    seed_from_docx(reopened.doc(), &document(&format!("<w:p>{xml}</w:p>"))).unwrap();
    let revisions = reopened.doc().list_revisions().unwrap();
    assert_eq!(revisions[0].change.kind, ChangeKind::RunPropertiesChanged);
    reopened
        .doc()
        .reject_change(&suggesting(), &ChangeTarget::Revision("42".to_owned()))
        .unwrap();
    assert!(attrs(&reopened)[0].1["bold"].is_null());
    assert_eq!(attrs(&reopened)[0].1["italic"], true);
}

#[test]
fn run_property_decision_preview_uses_old_or_new_properties_without_mutating() {
    use crate::bridge::{RenderEnv, RevisionPreview};
    let engine = engine();
    let receipt = engine
        .doc()
        .toggle_format(
            &suggesting(),
            StoryRange::new("body", 1, 5),
            SimpleFormat::Bold,
        )
        .unwrap();
    let before = engine.doc().encode_state_as_update_v1();
    for (decision, bold) in [
        (RevisionPreview::Accepted, true),
        (RevisionPreview::Rejected, false),
    ] {
        let env = RenderEnv::default().with_revision_preview(&receipt.revision_ids[0], decision);
        let blocks: Value =
            serde_json::from_str(&engine.lower_story_json("body", &env).unwrap()).unwrap();
        for run in blocks[0]["runs"].as_array().unwrap() {
            if run["text"].as_str().unwrap_or_default().contains("bcde") {
                assert_eq!(run["bold"].as_bool().unwrap_or(false), bold);
                assert_eq!(run["italic"], true);
                assert!(run["formatChange"].is_null());
            }
        }
    }
    assert_eq!(engine.doc().encode_state_as_update_v1(), before);
}

#[test]
fn formatted_change_has_a_dotted_display_cue_and_decisions_remove_it() {
    use crate::bridge::{RenderEnv, RevisionPreview};
    use docx_layout::display_list::{Primitive, RevisionKind};
    let font = docx_layout::register_measure_font(include_bytes!(
        "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
    ))
    .unwrap();
    let engine = engine();
    let receipt = engine
        .doc()
        .toggle_format(
            &suggesting(),
            StoryRange::new("body", 0, 6),
            SimpleFormat::Bold,
        )
        .unwrap();
    for decision in [
        None,
        Some(RevisionPreview::Accepted),
        Some(RevisionPreview::Rejected),
        None,
    ] {
        let env = decision.map_or_else(RenderEnv::default, |decision| {
            RenderEnv::default().with_revision_preview(&receipt.revision_ids[0], decision)
        });
        engine.layout_document_with_regions_json(&json!({
            "bodyStory":"body", "renderEnv":env, "options":{},
            "measurement":{"fontChains":{"calibri|0|0":[font]},"defaults":{"fontFamily":"Calibri","fontSize":12}}
        }).to_string()).unwrap();
        engine
            .build_display_list_frame("{}", engine.stats().frame_epoch)
            .unwrap();
        let cues = engine.with_display_list(|list| list.pages.iter().flat_map(|page| &page.primitives).filter(|primitive| {
            matches!(primitive, Primitive::Decoration(cue) if cue.dotted && cue.attrs.revision.as_ref().is_some_and(|revision| revision.kind == RevisionKind::Format && revision.author == "Ada"))
        }).count()).unwrap();
        assert_eq!(cues > 0, decision.is_none());
    }
}

#[test]
fn no_op_format_does_not_create_a_revision_and_typing_does_not_inherit_one() {
    let engine = engine();
    engine
        .doc()
        .format_range(
            &suggesting(),
            StoryRange::new("body", 0, 6),
            &InlineFormatDelta {
                italic: Patch::Set(true),
                ..Default::default()
            },
        )
        .unwrap();
    assert!(engine.doc().list_revisions().unwrap().is_empty());
    engine
        .doc()
        .toggle_format(
            &suggesting(),
            StoryRange::new("body", 0, 6),
            SimpleFormat::Bold,
        )
        .unwrap();
    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            crate::Position::new("body", 3),
            "X",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let runs = attrs(&engine);
    let inserted = runs.iter().find(|run| run.0 == "X").unwrap();
    assert!(inserted.1["rPrChange"].is_null());
    assert_eq!(inserted.1["bold"], true);
}

#[test]
fn imported_change_keeps_its_original_properties_after_a_partial_format() {
    let engine = EngineSession::new(9304);
    seed_from_docx(engine.doc(), &document(r#"<w:p><w:r><w:rPr><w:i/><w:rPrChange w:id="8" w:author="Prior" w:date="2026-09-01T00:00:00Z"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>abcdef</w:t></w:r></w:p>"#)).unwrap();
    let receipt = engine
        .doc()
        .toggle_format(
            &suggesting(),
            StoryRange::new("body", 2, 4),
            SimpleFormat::Underline,
        )
        .unwrap();
    assert_eq!(receipt.revision_ids, ["8"]);
    let runs = attrs(&engine);
    assert_eq!(runs.len(), 3);
    for run in &runs {
        assert_eq!(
            run.1["rPrChange"][0]["previousFormatting"],
            json!({"bold":true})
        );
        assert_eq!(run.1["rPrChange"][0]["info"]["author"], "Prior");
    }
    engine
        .doc()
        .reject_change(&suggesting(), &ChangeTarget::Revision("8".to_owned()))
        .unwrap();
    for run in attrs(&engine) {
        assert_eq!(run.1["bold"], true);
        assert!(run.1["italic"].is_null());
        assert!(run.1["underline"].is_null());
    }
}

#[test]
fn resolving_all_includes_run_properties_alongside_text_changes() {
    for accept in [true, false] {
        let engine = engine();
        engine
            .doc()
            .toggle_format(
                &suggesting(),
                StoryRange::new("body", 0, 6),
                SimpleFormat::Bold,
            )
            .unwrap();
        engine
            .doc()
            .insert_text(
                &suggesting(),
                crate::Position::new("body", 6),
                "X",
                FormatPolicy::Inherit,
            )
            .unwrap();
        let ids: Vec<_> = engine
            .doc()
            .list_revisions()
            .unwrap()
            .into_iter()
            .map(|revision| revision.change.revision_id)
            .collect();
        assert_eq!(ids.len(), 2);
        engine
            .doc()
            .settle_revisions(
                if accept { &ids } else { &[] },
                if accept { &[] } else { &ids },
                &crate::UndoSession::new(),
            )
            .unwrap()
            .unwrap();
        assert!(engine.doc().list_revisions().unwrap().is_empty());
        assert_eq!(
            attrs(&engine)[0].1["bold"].as_bool().unwrap_or(false),
            accept
        );
    }
}

#[test]
fn every_character_format_can_be_suggested_and_rejected() {
    for format in [
        SimpleFormat::Bold,
        SimpleFormat::Italic,
        SimpleFormat::Underline,
        SimpleFormat::Strike,
        SimpleFormat::Superscript,
        SimpleFormat::Subscript,
    ] {
        let engine = engine();
        let receipt = engine
            .doc()
            .toggle_format(&suggesting(), StoryRange::new("body", 0, 6), format)
            .unwrap();
        assert_eq!(receipt.revision_ids.len(), 1);
        engine
            .doc()
            .reject_change(
                &suggesting(),
                &ChangeTarget::Revision(receipt.revision_ids[0].clone()),
            )
            .unwrap();
        assert_eq!(attrs(&engine)[0].1, json!({"italic":true}));
    }
    let engine = engine();
    let delta = InlineFormatDelta {
        color: Patch::Set(ColorPatch::Rgb("123456".to_owned())),
        highlight: Patch::Set("yellow".to_owned()),
        font_size: Patch::Set(18.0),
        font_family: Patch::Set(FontFamilyPatch {
            ascii: "Serif".to_owned(),
            h_ansi: None,
        }),
        ..Default::default()
    };
    engine
        .doc()
        .format_range(
            &EditCtx::local("", ""),
            StoryRange::new("body", 0, 6),
            &delta,
        )
        .unwrap();
    let previous = attrs(&engine)[0].1.clone();
    let receipt = engine
        .doc()
        .clear_formatting(&suggesting(), StoryRange::new("body", 0, 6))
        .unwrap();
    let changes = attrs(&engine);
    assert_eq!(
        changes[0].1["rPrChange"][0]["previousFormatting"]["highlight"],
        "yellow"
    );
    assert_eq!(
        changes[0].1["rPrChange"][0]["previousFormatting"]["fontSize"].as_f64(),
        Some(36.0)
    );
    engine
        .doc()
        .reject_change(
            &suggesting(),
            &ChangeTarget::Revision(receipt.revision_ids[0].clone()),
        )
        .unwrap();
    assert_eq!(attrs(&engine)[0].1, previous);
}

#[test]
fn partial_rejection_keeps_the_remaining_run_property_change() {
    let engine = engine();
    engine
        .doc()
        .toggle_format(
            &suggesting(),
            StoryRange::new("body", 0, 6),
            SimpleFormat::Bold,
        )
        .unwrap();
    engine
        .doc()
        .reject_change(
            &suggesting(),
            &ChangeTarget::Range(StoryRange::new("body", 2, 4)),
        )
        .unwrap();
    let runs = attrs(&engine);
    assert_eq!(runs.len(), 3);
    assert_eq!(runs[0].1["bold"], true);
    assert!(runs[1].1["bold"].is_null());
    assert_eq!(runs[2].1["bold"], true);
    assert!(runs[1].1["rPrChange"].is_null());
    assert!(runs[0].1["rPrChange"].is_array());
    assert!(runs[2].1["rPrChange"].is_array());
}

#[test]
fn empty_previous_and_current_properties_still_reopen_as_a_change() {
    let engine = EngineSession::new(9305);
    seed_from_docx(engine.doc(), &document(r#"<w:p><w:r><w:rPr><w:rPrChange w:id="9" w:author="Writer"><w:rPr/></w:rPrChange></w:rPr><w:t>text</w:t></w:r></w:p>"#)).unwrap();
    assert_eq!(
        engine.doc().list_revisions().unwrap()[0].change.kind,
        ChangeKind::RunPropertiesChanged
    );
    engine
        .doc()
        .accept_change(&suggesting(), &ChangeTarget::Revision("9".to_owned()))
        .unwrap();
    assert!(engine.doc().list_revisions().unwrap().is_empty());
}
