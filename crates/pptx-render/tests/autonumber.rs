use pptx_edit::DeckSession;
use pptx_parse::{Bullet, ShapeNode};
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/autonumber-bullets.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const MONO: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationMono-Regular.ttf");

fn renderer() -> SlideRenderer {
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .register_font("Courier New", false, false, MONO)
        .unwrap();
    renderer
}

fn lines(list: &SurfaceDisplayList, id: u32) -> &[PositionedTextLine] {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id, lines, ..
            } if *object_id == id => Some(lines.as_slice()),
            _ => None,
        })
        .unwrap()
}

fn markers(list: &SurfaceDisplayList, id: u32) -> Vec<&str> {
    lines(list, id)
        .iter()
        .flat_map(|line| &line.runs)
        .filter(|run| run.start == run.end)
        .map(|run| run.text.as_str())
        .collect()
}

#[test]
fn automatic_numbers_keep_styles_and_story_geometry() {
    let session = DeckSession::open(DECK, 300).unwrap();
    let renderer = renderer();
    let snapshot = session.snapshot().unwrap();
    let rendered = renderer
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap();
    let list = &rendered.display_list;
    assert_eq!(
        markers(list, 10),
        ["1.", "–", "–", "–", "–", "–", "–", "2.", "3.", "4."]
    );
    assert_eq!(
        markers(list, 11),
        ["1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10."]
    );
    assert_eq!(markers(list, 12), ["7.", "8.", "1.", "2."]);
    assert_eq!(markers(list, 13), ["7.", "8."]);
    assert_eq!(markers(list, 14), ["Z.", "AA.", "AB."]);
    assert_eq!(markers(list, 15), ["1.", "1.", "1.", "•", "1."]);
    assert_eq!(markers(list, 16), ["IX.", "10."]);
    let mut edited = snapshot.clone();
    let story = &mut edited.slides[0]
        .shapes
        .iter_mut()
        .find(|shape| shape.name == "List 12")
        .unwrap()
        .text_stories[0];
    story.paragraphs[1].bullet_json = Some(
        r#"{"type":"autoNumber","scheme":"arabicPeriod","startAt":42,"restart":true}"#.to_owned(),
    );
    let edited = renderer
        .layout_slide(session.package(), &edited, 0)
        .unwrap();
    assert_eq!(markers(&edited.display_list, 12), ["7.", "42.", "1.", "2."]);
    for line in lines(list, 11) {
        let marker = &line.runs[0];
        assert_eq!(marker.x, 320.0);
        assert_eq!(marker.font_size_px, 18.0);
        assert_eq!(marker.color, "#D02020");
        assert_eq!(marker.font_family, "Courier New");
        assert!(
            marker
                .glyphs
                .iter()
                .all(|glyph| glyph.cluster == marker.start)
        );
        assert_eq!(line.x, 356.0);
        assert_eq!(line.runs[1].color, "#147D40");
        assert_eq!(line.runs[1].font_size_px, 24.0);
    }
    let mut plain_package = session.package().clone();
    for node in &mut plain_package.slides[0].shapes {
        if let ShapeNode::Shape(shape) = node {
            let body = shape.text.as_mut().unwrap();
            for properties in body
                .list_style
                .iter_mut()
                .chain(body.paragraphs.iter_mut().map(|p| &mut p.properties))
            {
                if matches!(properties.bullet, Some(Bullet::AutoNumber { .. })) {
                    properties.bullet = Some(Bullet::None);
                }
            }
        }
    }
    let mut plain_snapshot = snapshot.clone();
    for shape in &mut plain_snapshot.slides[0].shapes {
        for story in &mut shape.text_stories {
            for paragraph in &mut story.paragraphs {
                if paragraph
                    .bullet_json
                    .as_deref()
                    .and_then(|json| serde_json::from_str::<Bullet>(json).ok())
                    .is_some_and(|bullet| matches!(bullet, Bullet::AutoNumber { .. }))
                {
                    paragraph.bullet_json = None;
                }
            }
        }
    }
    let plain = renderer
        .layout_slide(&plain_package, &plain_snapshot, 0)
        .unwrap();
    let mut stripped = list.clone();
    for primitive in &mut stripped.primitives {
        if let Primitive::TextBox { lines, .. } = primitive {
            for line in lines {
                line.runs
                    .retain(|run| run.start != run.end || ["–", "•"].contains(&run.text.as_str()));
            }
        }
    }
    assert!(stripped == plain.display_list);
    for line in lines(list, 11) {
        for x in [320.0, line.x, line.x + line.width] {
            assert_eq!(
                rendered.hit_test(x, line.baseline),
                plain.hit_test(x, line.baseline)
            );
        }
    }
}

#[test]
fn automatic_numbers_are_stable_with_normal_and_shape_autofit() {
    let session = DeckSession::open(DECK, 301).unwrap();
    let renderer = renderer();
    let snapshot = session.snapshot().unwrap();
    let first = renderer
        .layout_slide(session.package(), &snapshot, 1)
        .unwrap();
    for id in [20, 21] {
        assert_eq!(
            markers(&first.display_list, id),
            ["1.", "2.", "3.", "4.", "5."]
        );
        let lines = lines(&first.display_list, id);
        assert!(lines.len() > 5);
        if id == 20 {
            assert!(lines[0].runs[0].font_size_px < 18.0);
        } else {
            assert_eq!(lines[0].runs[0].font_size_px, 18.0);
        }
    }
    let second = renderer
        .layout_slide(session.package(), &snapshot, 1)
        .unwrap();
    assert_eq!(first.display_list, second.display_list);
}
