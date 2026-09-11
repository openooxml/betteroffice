use pptx_edit::DeckSession;
use pptx_parse::{Bullet, BulletSize, PptxPackage, ShapeNode, TextBody};
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/list-style-bullets.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const MONO: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationMono-Regular.ttf");

fn renderer() -> SlideRenderer {
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        renderer.register_font("Arial", bold, false, FONT).unwrap();
        renderer
            .register_font("Courier New", bold, false, MONO)
            .unwrap();
    }
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

#[test]
fn list_styles_cascade_defaults_levels_and_direct_properties() {
    let session = DeckSession::open(DECK, 2941).unwrap();
    let list = renderer()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    let title = &lines(&list, 2)[0];
    assert_eq!(title.runs[0].font_size_px, 88.0);
    assert_eq!(title.runs[0].color, "#24265D");
    assert!(title.runs[0].bold);
    assert!((title.x - (64.0 + (1152.0 - title.width) / 2.0)).abs() < 0.001);
    let body = lines(&list, 3);
    assert_eq!(body[0].runs.last().unwrap().font_size_px, 24.0);
    assert_eq!(body[0].runs.last().unwrap().color, "#147D40");
    assert_eq!(body[1].runs.last().unwrap().color, "#2040B0");
    assert_eq!(body[2].runs.len(), 1);
    assert_eq!(body[2].runs[0].color, "#505050");
    assert!((body[2].runs[0].font_size_px - 56.0 / 3.0).abs() < 0.001);
    let defaults = lines(&list, 4);
    assert_eq!(defaults[0].runs.last().unwrap().font_size_px, 32.0);
    assert_eq!(defaults[0].runs.last().unwrap().color, "#9C27B0");
    assert_eq!(defaults[1].runs.last().unwrap().font_size_px, 40.0);
    assert_eq!(defaults[1].runs.last().unwrap().color, "#006699");
    let direct = lines(&list, 5);
    assert_eq!(direct[0].runs.len(), 1);
    assert_eq!(direct[0].runs[0].color, "#008080");
    assert!((direct[0].runs[0].font_size_px - 80.0 / 3.0).abs() < 0.001);
}

fn clear_bullets(body: &mut TextBody) {
    for properties in body
        .default_list_style
        .as_deref_mut()
        .into_iter()
        .chain(body.list_style.iter_mut())
        .chain(body.paragraphs.iter_mut().map(|p| &mut p.properties))
    {
        properties.bullet = Some(Bullet::None);
    }
}

fn clear_shapes(shapes: &mut [ShapeNode]) {
    for shape in shapes {
        match shape {
            ShapeNode::Shape(shape) => {
                if let Some(body) = &mut shape.text {
                    clear_bullets(body);
                }
            }
            ShapeNode::Group(group) => clear_shapes(&mut group.children),
            _ => {}
        }
    }
}

fn without_bullets(mut package: PptxPackage) -> PptxPackage {
    for slide in &mut package.slides {
        clear_shapes(&mut slide.shapes);
    }
    for layout in &mut package.layouts {
        clear_shapes(&mut layout.shapes);
    }
    for master in &mut package.masters {
        clear_shapes(&mut master.shapes);
        for properties in master
            .text_styles
            .title
            .iter_mut()
            .chain(&mut master.text_styles.body)
            .chain(&mut master.text_styles.other)
        {
            properties.bullet = Some(Bullet::None);
        }
    }
    package
}

#[test]
fn bullets_use_own_formatting_without_changing_story_positions() {
    let session = DeckSession::open(DECK, 2942).unwrap();
    let renderer = renderer();
    let snapshot = session.snapshot().unwrap();
    let rendered = renderer
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap();
    let plain_rendered = renderer
        .layout_slide(&without_bullets(session.package().clone()), &snapshot, 0)
        .unwrap();
    let list = &rendered.display_list;
    let plain = &plain_rendered.display_list;
    for (line, plain) in lines(list, 3).iter().zip(lines(plain, 3)) {
        let mut actual = line.clone();
        actual.runs.retain(|run| run.start != run.end);
        assert_eq!(&actual, plain);
        for x in [line.x - 30.0, line.x, line.x + line.width] {
            assert_eq!(
                rendered.hit_test(x, line.baseline),
                plain_rendered.hit_test(x, line.baseline)
            );
        }
    }
    let body = lines(list, 3);
    assert_eq!(body[0].runs[0].text, "•");
    assert_eq!(body[0].runs[0].x, 64.0);
    assert_eq!(body[0].runs[0].font_size_px, 12.0);
    assert_eq!(body[0].runs[0].color, "#D02020");
    assert_eq!(body[0].runs[0].font_family, "Courier New");
    assert_eq!(body[0].runs[0].start, body[0].runs[0].end);
    assert_eq!(body[1].runs[0].text, "–");
    assert_eq!(body[1].runs[0].x, 100.0);
    assert_eq!(body[1].runs[0].color, "#2040B0");
    assert_eq!(body[3].runs[0].text, "•");
    assert_eq!(body[3].runs[0].font_size_px, 24.0);
    assert_eq!(body[3].runs[0].color, "#147D40");
    assert!(body[3].runs[0].bold);
    assert_eq!(body[3].runs[0].font_family, "Arial");
    let mut package = session.package().clone();
    let ShapeNode::Shape(shape) = &mut package.layouts[0].shapes[1] else {
        panic!("body")
    };
    shape.text.as_mut().unwrap().list_style[0].bullet_size = Some(BulletSize::Points(12.0));
    let points = renderer
        .layout_slide(&package, &snapshot, 0)
        .unwrap()
        .display_list;
    assert_eq!(lines(&points, 3)[0].runs[0].font_size_px, 16.0);
}
