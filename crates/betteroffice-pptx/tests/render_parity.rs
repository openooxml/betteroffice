use betteroffice_pptx::{
    EditCtx, Error, HitTestResult, Presentation, Primitive, ProposalDiffSlide, ProposalEdit,
    ProposalError, ProposalRequest, RenderError, RenderedSlide, TextParagraphDraft, TextRunDraft,
    TextStyle,
};
use pptx_edit::DeckSession;
use pptx_render::SlideRenderer;

const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const FALLBACK: &[u8] = include_bytes!("../../../packages/fonts/assets/NotoSansArabic-Regular.ttf");

struct Pair {
    facade: Presentation,
    session: DeckSession,
    renderer: SlideRenderer,
}

impl Pair {
    fn new(fonts: bool) -> Self {
        let mut pair = Self {
            facade: Presentation::open_with_options(DECK, &backstop_options()).unwrap(),
            session: DeckSession::open(DECK, 1).unwrap(),
            renderer: SlideRenderer::new(),
        };
        if fonts {
            for bold in [false, true] {
                assert_eq!(
                    pair.facade
                        .register_font("Arial", bold, false, FONT)
                        .unwrap(),
                    pair.renderer
                        .register_font("Arial", bold, false, FONT)
                        .unwrap()
                );
            }
        }
        pair
    }

    fn layout(&self, index: usize) -> RenderedSlide {
        let scope = self.session.slide_scope(index).unwrap();
        self.renderer
            .layout_scoped_slide(self.session.package(), &scope)
            .unwrap()
    }

    fn propose(&self) -> String {
        let snapshot = self.session.snapshot().unwrap();
        let story = &snapshot.slides[0]
            .shapes
            .iter()
            .find(|shape| !shape.text_stories.is_empty())
            .unwrap()
            .text_stories[0];
        let request = ProposalRequest {
            agent_id: "render-parity".to_owned(),
            note: Some("Title".to_owned()),
            edits: vec![ProposalEdit::ReplaceText {
                story_id: story.id.clone(),
                start: 0,
                end: 3,
                text: "Reviewed".to_owned(),
                style: None,
            }],
        };
        let actual = self.facade.propose(request.clone()).unwrap();
        let expected = self.session.propose(request).unwrap();
        assert_eq!(actual, expected);
        actual.id
    }
}

#[test]
fn hit_tests_match_the_shared_typescript_layout_path() {
    for fonts in [false, true] {
        let pair = Pair::new(fonts);
        let rendered = pair.layout(0);
        for (x, y) in [
            (8.0, 8.0),
            (100.0, 150.0),
            (40.0, 20.0),
            (-1.0, -1.0),
            (1281.0, 721.0),
            (f32::NAN, 1.0),
            (1.0, f32::INFINITY),
        ] {
            assert_eq!(
                pair.facade.hit_test(0, x, y).unwrap(),
                rendered.hit_test(x, y),
                "point ({x}, {y}), fonts={fonts}"
            );
        }
        assert!(matches!(
            rendered.hit_test(8.0, 8.0),
            Some(HitTestResult::Shape { .. })
        ));
        assert!(matches!(
            rendered.hit_test(100.0, 150.0),
            Some(HitTestResult::Text { .. })
        ));
        assert_eq!(rendered.hit_test(40.0, 20.0), None);
        for index in 1..pair.facade.slide_ids().unwrap().len() {
            let reference = pair.layout(index);
            assert_eq!(
                pair.facade.hit_test(index, 8.0, 8.0).unwrap(),
                reference.hit_test(8.0, 8.0)
            );
        }
    }
}

#[test]
fn hit_tests_follow_edits_without_a_previous_layout() {
    let pair = Pair::new(true);
    let snapshot = pair.session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = &slide.shapes[0];
    assert!(pair.facade.hit_test(0, 8.0, 8.0).unwrap().is_some());
    let context = EditCtx::local("render-parity");
    pair.facade
        .remove_shape(&context, &slide.id, &shape.id)
        .unwrap();
    pair.session
        .remove_shape(&context, &slide.id, &shape.id)
        .unwrap();
    assert_eq!(pair.facade.hit_test(0, 8.0, 8.0).unwrap(), None);
    assert_eq!(
        pair.facade.hit_test(0, 8.0, 8.0).unwrap(),
        pair.layout(0).hit_test(8.0, 8.0)
    );
}

#[test]
fn fallback_registration_changes_missing_glyph_layout_like_typescript() {
    let mut pair = Pair::new(true);
    let snapshot = pair.session.snapshot().unwrap();
    let story_id = &snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| !shape.text_stories.is_empty())
        .unwrap()
        .text_stories[0]
        .id;
    let paragraphs = [TextParagraphDraft {
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
    }];
    let context = EditCtx::local("render-parity");
    pair.facade
        .set_story_paragraphs(&context, story_id, &paragraphs)
        .unwrap();
    pair.session
        .set_story_paragraphs(&context, story_id, &paragraphs)
        .unwrap();
    let before = pair.facade.render_slide(0).unwrap().display_list;
    assert_eq!(before, pair.layout(0).display_list);
    let id = pair
        .facade
        .register_fallback_font("Noto Sans Arabic", false, false, FALLBACK)
        .unwrap();
    assert_eq!(
        id,
        pair.renderer
            .register_fallback_font("Noto Sans Arabic", false, false, FALLBACK)
            .unwrap()
    );
    let after = pair.facade.render_slide(0).unwrap();
    let reference = pair.layout(0);
    assert_eq!(after.display_list, reference.display_list);
    assert_eq!(after.font_substitutions, reference.font_substitutions);
    assert_ne!(after.display_list, before);
    assert!(after.display_list.primitives.iter().any(|primitive| {
        matches!(primitive, Primitive::TextBox { lines, .. }
            if lines.iter().flat_map(|line| &line.runs).any(|run| run.font_id == id))
    }));
    assert_eq!(
        pair.facade.hit_test(0, 100.0, 70.0).unwrap(),
        reference.hit_test(100.0, 70.0)
    );
}

#[test]
fn proposal_diff_matches_the_typescript_preview_and_render_paths() {
    for fonts in [false, true] {
        let pair = Pair::new(fonts);
        let id = pair.propose();
        let before = pair.facade.snapshot().unwrap();
        for index in 0..before.slides.len() {
            let preview = pair
                .session
                .preview_proposal_diff_slide(&id, index)
                .unwrap();
            let frame = pair
                .renderer
                .layout_scoped_slide(pair.session.package(), &preview.scope)
                .unwrap();
            let expected = ProposalDiffSlide {
                proposal: preview.proposal,
                snapshot: preview.snapshot,
                text_changes: preview.text_changes,
                frame: frame.display_list,
            };
            assert_eq!(
                pair.facade.layout_proposal_diff_slide(&id, index).unwrap(),
                expected
            );
            if index == 0 {
                assert!(!expected.text_changes.is_empty());
            }
        }
        assert_eq!(pair.facade.snapshot().unwrap(), before);
        assert!(!pair.facade.can_undo());
    }
}

#[test]
fn invalid_targets_and_fonts_keep_engine_errors() {
    let mut pair = Pair::new(false);
    assert!(matches!(
        pair.facade.render_slide(0),
        Err(Error::Render(RenderError::NoFont))
    ));
    assert!(matches!(
        pair.facade.hit_test(99, 0.0, 0.0),
        Err(Error::Render(RenderError::SlideNotFound(99)))
    ));
    let id = pair.propose();
    assert!(matches!(
        pair.facade.layout_proposal_diff_slide(&id, 99),
        Err(Error::Render(RenderError::SlideNotFound(99)))
    ));
    let error = pair
        .facade
        .layout_proposal_diff_slide("unknown", 0)
        .unwrap_err();
    assert!(matches!(
        &error,
        Error::Proposal(ProposalError::NotFound(_))
    ));
    assert_eq!(
        error.to_string(),
        pair.session
            .preview_proposal_diff_slide("unknown", 0)
            .unwrap_err()
            .to_string()
    );
    for (family, bytes) in [("", FONT), ("Invalid", b"not a font".as_slice())] {
        let actual = pair
            .facade
            .register_fallback_font(family, false, false, bytes)
            .unwrap_err();
        let expected = pair
            .renderer
            .register_fallback_font(family, false, false, bytes)
            .unwrap_err();
        assert_eq!(actual.to_string(), expected.to_string());
        assert!(matches!(actual, Error::Render(RenderError::Font(_))));
    }
}

fn backstop_options() -> betteroffice_pptx::OpenOptions {
    let options = betteroffice_pptx::OpenOptions::default();
    #[cfg(feature = "fonts")]
    let options = betteroffice_pptx::OpenOptions {
        font_source: betteroffice_pptx::fonts::FontSource::Backstop,
        ..options
    };
    options
}
