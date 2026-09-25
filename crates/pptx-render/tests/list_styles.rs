use pptx_edit::DeckSession;
use pptx_parse::{
    Bullet, BulletSize, ParagraphProperties, PptxPackage, RunProperties, ShapeNode, TextBody,
};
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

#[test]
fn text_outside_placeholders_takes_the_presentation_default_style() {
    let session = DeckSession::open(DECK, 2943).unwrap();
    let mut package = session.package().clone();
    let sized = |size| {
        vec![
            ParagraphProperties {
                default_run: Some(RunProperties {
                    font_size_pt: Some(size),
                    ..RunProperties::default()
                }),
                ..ParagraphProperties::default()
            };
            3
        ]
    };
    package.presentation.default_text_style = sized(27.0);
    for master in &mut package.masters {
        master.text_styles.other = sized(18.0);
    }
    for shape in &mut package.slides[0].shapes {
        if let ShapeNode::Shape(shape) = shape
            && let Some(body) = &mut shape.text
        {
            body.default_list_style = None;
            body.list_style.clear();
        }
    }
    let list = renderer()
        .layout_slide(&package, &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    assert_eq!(lines(&list, 2)[0].runs[0].font_size_px, 88.0);
    assert_eq!(lines(&list, 4)[0].runs.last().unwrap().font_size_px, 36.0);
}

#[test]
fn a_default_text_style_given_only_as_def_ppr_reaches_text_outside_placeholders() {
    let mut parts = ooxml_opc::unzip_parts(DECK).unwrap();
    for (path, bytes) in &mut parts {
        if path == "ppt/presentation.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    "</p:presentation>",
                    r#"<p:defaultTextStyle><a:defPPr><a:defRPr sz="2700"/></a:defPPr></p:defaultTextStyle></p:presentation>"#,
                )
                .into_bytes();
        }
    }
    let session = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 2944).unwrap();
    let mut package = session.package().clone();
    assert!(package.presentation.default_text_style.is_empty());
    for shape in &mut package.slides[0].shapes {
        if let ShapeNode::Shape(shape) = shape
            && let Some(body) = &mut shape.text
        {
            body.default_list_style = None;
            body.list_style.clear();
        }
    }
    let list = renderer()
        .layout_slide(&package, &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    assert_eq!(lines(&list, 2)[0].runs[0].font_size_px, 88.0);
    assert_eq!(lines(&list, 4)[0].runs.last().unwrap().font_size_px, 36.0);
}

#[test]
fn a_right_to_left_paragraph_hangs_its_marker_off_the_right_edge() {
    let session = DeckSession::open(DECK, 2945).unwrap();
    let renderer = renderer();
    let snapshot = session.snapshot().unwrap();
    let ltr = renderer
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap()
        .display_list;
    let mut package = session.package().clone();
    for shape in &mut package.slides[0].shapes {
        if let ShapeNode::Shape(shape) = shape
            && let Some(body) = &mut shape.text
        {
            for paragraph in &mut body.paragraphs {
                paragraph.properties.rtl = Some(true);
            }
        }
    }
    let rtl = renderer
        .layout_slide(&package, &snapshot, 0)
        .unwrap()
        .display_list;
    let marker = |line: &PositionedTextLine| {
        let run = line.runs.iter().find(|run| run.text == "•").unwrap();
        (run.x, run.x + run.width)
    };
    let text = |line: &PositionedTextLine| {
        line.runs
            .iter()
            .filter(|run| run.text != "•")
            .fold((f32::MAX, f32::MIN), |(left, right), run| {
                (left.min(run.x), right.max(run.x + run.width))
            })
    };
    let (ltr, rtl) = (&lines(&ltr, 3)[0], &lines(&rtl, 3)[0]);
    assert!(
        marker(ltr).1 <= text(ltr).0,
        "left to right, the marker leads on the left"
    );
    assert!(
        marker(rtl).0 >= text(rtl).1,
        "right to left, it leads on the right"
    );
    // The body is 600 px wide with no insets, and its markers hang to the edge.
    assert!((marker(rtl).1 - (marker(ltr).0 + 600.0)).abs() < 0.01);
    assert!(
        text(rtl).0 < text(ltr).0,
        "the left margin no longer holds the text"
    );
}

#[test]
fn a_right_to_left_paragraph_read_from_xml_paints_and_carets_in_reading_order() {
    let mut parts = ooxml_opc::unzip_parts(DECK).unwrap();
    for (path, bytes) in &mut parts {
        if path == "ppt/slides/slide1.xml" {
            let xml = String::from_utf8(bytes.clone()).unwrap();
            let rtl = xml.replacen(
                r#"<a:pPr lvl="0" /><a:r><a:rPr /><a:t>First level</a:t>"#,
                "<a:pPr lvl=\"0\" rtl=\"1\" /><a:r><a:rPr /><a:t>\u{645}\u{631}\u{62d}\u{628}\u{627} \u{628}\u{643}\u{645}</a:t>",
                1,
            );
            assert_ne!(rtl, xml, "the fixture paragraph is rewritten");
            *bytes = rtl.into_bytes();
        }
    }
    let session = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 2946).unwrap();
    let ShapeNode::Shape(body) = &session.package().slides[0].shapes[1] else {
        panic!("the bullet body");
    };
    assert_eq!(
        body.text.as_ref().unwrap().paragraphs[0].properties.rtl,
        Some(true)
    );
    let list = renderer()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    let line = &lines(&list, 3)[0];
    let run = line
        .runs
        .iter()
        .find(|run| run.text.contains('\u{645}'))
        .expect("the Arabic run");
    assert!(run.glyphs.len() > 2);
    assert!(
        run.glyphs.windows(2).all(|pair| pair[1].x < pair[0].x),
        "each glyph paints left of the one before it"
    );
    for glyph in &run.glyphs {
        let stop = line
            .caret_stops
            .iter()
            .find(|stop| stop.position == glyph.cluster)
            .expect("a caret before every cluster");
        assert!(
            (stop.x - (glyph.x + glyph.advance)).abs() < 0.01,
            "the caret before a character sits at its right edge"
        );
    }
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
        // The marker stands in the hanging indent, so a paragraph that loses
        // it would start its first line there instead.
        properties.indent = Some(0);
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
            properties.indent = Some(0);
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
