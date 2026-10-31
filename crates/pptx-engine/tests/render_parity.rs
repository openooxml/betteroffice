use pptx_edit::{
    DeckSession, EditCtx, ProposalEdit, ProposalRequest, TextParagraphDraft, TextRunDraft,
    TextStyle,
};
use pptx_engine::{Renderer, parse_pptx_json, parse_pptx_json_without_media};
use pptx_render::{HitTestResult, RenderError, SlideRenderer};
use serde_json::Value;

const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const FALLBACK: &[u8] = include_bytes!("../../../packages/fonts/assets/NotoSansArabic-Regular.ttf");

fn renderers(fonts: bool) -> (Renderer, SlideRenderer) {
    let mut engine = Renderer::new();
    let mut renderer = SlideRenderer::new();
    if fonts {
        for bold in [false, true] {
            assert_eq!(
                engine.register_font("Arial", bold, false, FONT).unwrap(),
                renderer.register_font("Arial", bold, false, FONT).unwrap()
            );
        }
    }
    (engine, renderer)
}

#[test]
fn hit_test_boundary_matches_the_shared_native_path() {
    let session = DeckSession::open(DECK, 1).unwrap();
    for fonts in [false, true] {
        let (mut engine, renderer) = renderers(fonts);
        let scope = session.slide_scope(0).unwrap();
        let reference = renderer
            .layout_scoped_slide(session.package(), &scope)
            .unwrap();
        for (x, y) in [
            (8.0, 8.0),
            (100.0, 150.0),
            (40.0, 20.0),
            (-1.0, -1.0),
            (1281.0, 721.0),
            (f32::NAN, 1.0),
            (1.0, f32::INFINITY),
        ] {
            let json = engine
                .hit_test_slide_json(&session, &scope.slide.id, x, y)
                .unwrap();
            assert_eq!(
                serde_json::from_str::<Option<HitTestResult>>(&json).unwrap(),
                reference.hit_test(x, y)
            );
        }
        assert_eq!(
            engine
                .hit_test_slide_json(&session, "unknown", 0.0, 0.0)
                .unwrap(),
            "null"
        );
    }
}

#[test]
fn fallback_boundary_matches_native_layout_after_a_cached_frame() {
    let session = DeckSession::open(DECK, 1).unwrap();
    let scope = session.slide_scope(0).unwrap();
    let story_id = &scope
        .slide
        .shapes
        .iter()
        .find(|shape| !shape.text_stories.is_empty())
        .unwrap()
        .text_stories[0]
        .id;
    session
        .set_story_paragraphs(
            &EditCtx::local("render-parity"),
            story_id,
            &[TextParagraphDraft {
                alignment: None,
                bullet: None,
                runs: vec![TextRunDraft {
                    text: "Hello مرحبا world".to_owned(),
                    style: TextStyle {
                        font_family: Some("Unregistered parity face".to_owned()),
                        font_size_pt: Some(18.0),
                        bold: Some(false),
                        italic: Some(false),
                        ..TextStyle::default()
                    },
                }],
            }],
        )
        .unwrap();
    let (mut engine, mut renderer) = renderers(true);
    let before = engine.layout_slide_json(&session, 0).unwrap();
    assert_eq!(
        engine
            .register_fallback_font("Noto Sans Arabic", false, false, FALLBACK)
            .unwrap(),
        renderer
            .register_fallback_font("Noto Sans Arabic", false, false, FALLBACK)
            .unwrap()
    );
    let scope = session.slide_scope(0).unwrap();
    let expected = renderer
        .layout_scoped_slide(session.package(), &scope)
        .unwrap();
    let after = engine.layout_slide_json(&session, 0).unwrap();
    assert_ne!(before, after);
    assert_eq!(
        after,
        serde_json::to_string(&expected.display_list).unwrap()
    );
}

#[test]
fn proposal_diff_boundary_preserves_the_original_json_contract() {
    let session = DeckSession::open(DECK, 1).unwrap();
    let scope = session.slide_scope(0).unwrap();
    let story_id = &scope
        .slide
        .shapes
        .iter()
        .find(|shape| !shape.text_stories.is_empty())
        .unwrap()
        .text_stories[0]
        .id;
    let proposal = session
        .propose(ProposalRequest {
            agent_id: "render-parity".to_owned(),
            note: Some("Title".to_owned()),
            edits: vec![ProposalEdit::ReplaceText {
                story_id: story_id.clone(),
                start: 0,
                end: 3,
                text: "Reviewed".to_owned(),
                style: None,
            }],
        })
        .unwrap();
    for fonts in [false, true] {
        let (engine, renderer) = renderers(fonts);
        for index in 0..session.slide_ids().unwrap().len() {
            let preview = session
                .preview_proposal_diff_slide(&proposal.id, index)
                .unwrap();
            let rendered = renderer
                .layout_scoped_slide(session.package(), &preview.scope)
                .unwrap();
            let expected = serde_json::json!({
                "proposal": preview.proposal,
                "snapshot": preview.snapshot,
                "textChanges": preview.text_changes,
                "frame": rendered.display_list,
            });
            assert_eq!(
                engine
                    .layout_proposal_diff_slide_json(&session, &proposal.id, index as u32)
                    .unwrap(),
                serde_json::to_string(&expected).unwrap()
            );
        }
        assert_eq!(
            engine
                .layout_proposal_diff_slide_json(&session, &proposal.id, 99)
                .unwrap_err()
                .to_string(),
            RenderError::SlideNotFound(99).to_string()
        );
        assert_eq!(
            engine
                .layout_proposal_diff_slide_json(&session, "unknown", 0)
                .unwrap_err()
                .to_string(),
            session
                .preview_proposal_diff_slide("unknown", 0)
                .unwrap_err()
                .to_string()
        );
    }
}

#[test]
fn inspection_boundary_matches_parse_json_with_and_without_media() {
    let package = pptx_parse::parse_pptx(DECK).unwrap();
    let full: Value = serde_json::from_str(&parse_pptx_json(DECK).unwrap()).unwrap();
    assert_eq!(full, serde_json::to_value(&package).unwrap());
    let mut expected = full;
    expected["media"] = serde_json::to_value(
        package
            .media
            .iter()
            .map(|part| {
                serde_json::json!({
                    "partPath": part.part_path,
                    "contentType": part.content_type,
                    "byteLength": part.bytes.len(),
                })
            })
            .collect::<Vec<_>>(),
    )
    .unwrap();
    assert_eq!(
        serde_json::from_str::<Value>(&parse_pptx_json_without_media(DECK).unwrap()).unwrap(),
        expected
    );
    for parse in [parse_pptx_json, parse_pptx_json_without_media] {
        assert_eq!(
            parse(b"not a presentation").unwrap_err().to_string(),
            pptx_parse::parse_pptx(b"not a presentation")
                .unwrap_err()
                .to_string()
        );
    }
}
