use pptx_edit::{DeckSession, EditCtx, TextStyle};
use pptx_parse::{ShapeNode, TextAutofit};
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer};

const DECK: &[u8] = include_bytes!("fixtures/line-spacing.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn render(family: &str, size_pt: f64, scale: f64) -> (f32, Vec<PositionedTextLine>) {
    let mut package = pptx_parse::parse_pptx(DECK).unwrap();
    let ShapeNode::Shape(shape) = &mut package.slides[0].shapes[2] else {
        panic!("text shape");
    };
    let body = shape.text.as_mut().unwrap();
    for paragraph in &mut body.paragraphs {
        for run in &mut paragraph.runs {
            run.properties.font_family = Some(family.to_owned());
            run.properties.font_size_pt = Some(size_pt);
        }
    }
    body.autofit = Some(TextAutofit::Normal {
        font_scale: Some(scale),
        line_space_reduction: None,
    });
    let session = DeckSession::from_package(package, 8_021).unwrap();
    let snapshot = session.snapshot().unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap()
        .display_list
        .primitives
        .into_iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id: 4,
                paragraphs,
                lines,
                ..
            } => Some((paragraphs[0].runs[0].font_size_pt, lines)),
            _ => None,
        })
        .unwrap()
}

#[test]
fn fractional_autofit_matches_an_authored_whole_point_size() {
    for family in ["Arial", "Trebuchet MS"] {
        let (size, lines) = render(family, 32.0, 0.85);
        let (_, authored) = render(family, 27.0, 1.0);
        assert_eq!(size, 27.0);
        assert_eq!(lines, authored);
        assert_eq!(lines[0].runs[0].font_size_px, 36.0);
    }
}

#[test]
fn unscaled_text_keeps_its_authored_fractional_size() {
    let (size, lines) = render("Arial", 27.5, 1.0);
    assert_eq!(size, 27.5);
    assert!((lines[0].runs[0].font_size_px - 27.5 * 4.0 / 3.0).abs() < 0.001);
}

fn assert_unwrapped_autofit_fits(margin: i64, indent: i64) {
    let mut package = pptx_parse::parse_pptx(DECK).unwrap();
    let source = format!(
        r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<p:cSld><p:spTree><p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="476250" y="476250"/><a:ext cx="1143000" cy="952500"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr wrap="none" lIns="95250" rIns="95250" tIns="0" bIns="0"><a:normAutofit fontScale="100000"/></a:bodyPr>
<a:lstStyle/><a:p><a:pPr algn="l" marL="{margin}" marR="0" indent="{indent}"><a:buNone/></a:pPr>
<a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>WW</a:t></a:r>
<a:endParaRPr sz="2400"/></a:p></p:txBody>
</p:sp></p:spTree></p:cSld></p:sld>"#
    );
    assert!(package.replace_part("ppt/slides/slide1.xml", source.into_bytes()));
    let bytes = pptx_parse::write_pptx(&package).unwrap();
    let session = DeckSession::open(&bytes, 8_022).unwrap();
    let story = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0].clone();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let text = |session: &DeckSession| {
        renderer
            .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
            .unwrap()
            .display_list
            .primitives
            .into_iter()
            .find_map(|primitive| match primitive {
                Primitive::TextBox {
                    object_id: 2,
                    x,
                    w,
                    lines,
                    ..
                } => Some((x + 10.0, w - 20.0, lines)),
                _ => None,
            })
            .unwrap()
    };
    let (_, _, before) = text(&session);
    assert_eq!(before[0].runs[0].font_size_px, 32.0);
    session
        .insert_text(
            &EditCtx::local("test"),
            &story.id,
            story.length - 1,
            "W",
            &TextStyle::default(),
        )
        .unwrap();
    let (x, width, lines) = text(&session);
    assert_eq!(lines.len(), 1);
    assert!(lines[0].width < width);
    assert!(
        lines[0].x >= x && lines[0].x + lines[0].width <= x + width + 0.01,
        "positioned line must fit: x={}, width={}, content=[{}, {}]",
        lines[0].x,
        lines[0].width,
        x,
        x + width
    );
    assert!(lines[0].runs[0].font_size_px < 32.0);
    let saved = renderer.save(&session).unwrap();
    let reopened = DeckSession::open(&saved, 8_023).unwrap();
    let ShapeNode::Shape(shape) = &reopened.package().slides[0].shapes[0] else {
        panic!("text shape");
    };
    assert!(matches!(
        shape.text.as_ref().unwrap().autofit,
        Some(TextAutofit::Normal { font_scale: Some(scale), .. }) if scale < 1.0
    ));
    assert_eq!(text(&reopened), (x, width, lines));
}

#[test]
fn edited_unwrapped_autofit_includes_paragraph_margin() {
    assert_unwrapped_autofit_fits(285_750, 0);
}

#[test]
fn edited_unwrapped_autofit_includes_first_line_indent() {
    assert_unwrapped_autofit_fits(0, 285_750);
}

fn right_aligned_refit(text: &str, typed: &str) -> (f32, Option<TextAutofit>) {
    let mut package = pptx_parse::parse_pptx(DECK).unwrap();
    let source = format!(
        r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<p:cSld><p:spTree><p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="476250" y="476250"/><a:ext cx="2857500" cy="1905000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr wrap="square" lIns="95250" rIns="95250" tIns="0" bIns="0"><a:normAutofit/></a:bodyPr>
<a:lstStyle/><a:p><a:pPr algn="r" marL="0" marR="0" indent="0"><a:buNone/></a:pPr>
<a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>{text}</a:t></a:r>
<a:endParaRPr sz="2400"/></a:p></p:txBody>
</p:sp></p:spTree></p:cSld></p:sld>"#
    );
    assert!(package.replace_part("ppt/slides/slide1.xml", source.into_bytes()));
    let bytes = pptx_parse::write_pptx(&package).unwrap();
    let session = DeckSession::open(&bytes, 8_024).unwrap();
    let story = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0].clone();
    session
        .insert_text(
            &EditCtx::local("test"),
            &story.id,
            story.length - 1,
            typed,
            &TextStyle::default(),
        )
        .unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let size = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
        .primitives
        .into_iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id: 2,
                lines,
                ..
            } => Some(lines[0].runs[0].font_size_px),
            _ => None,
        })
        .unwrap();
    let saved = renderer.save(&session).unwrap();
    let reopened = DeckSession::open(&saved, 8_025).unwrap();
    let ShapeNode::Shape(shape) = &reopened.package().slides[0].shapes[0] else {
        panic!("text shape");
    };
    (size, shape.text.as_ref().unwrap().autofit.clone())
}

fn assert_full_size((size, autofit): (f32, Option<TextAutofit>)) {
    assert_eq!(size, 32.0);
    assert!(
        matches!(
            autofit,
            Some(TextAutofit::Normal { font_scale, .. }) if font_scale.is_none_or(|scale| scale == 1.0)
        ),
        "{autofit:?}"
    );
}

#[test]
fn edited_right_aligned_wrapped_paragraph_keeps_its_full_size() {
    assert_full_size(right_aligned_refit(
        "Alpha beta gamma delta epsilon zeta eta theta",
        " iota",
    ));
}

#[test]
fn a_typed_trailing_space_keeps_right_aligned_text_at_full_size() {
    assert_full_size(right_aligned_refit("Right aligned", " "));
}
