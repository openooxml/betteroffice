use pptx_edit::DeckSession;
use pptx_parse::{PptxPackage, ShapeNode, TextBody};
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/paragraph-right-margin.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const BOX_WIDTH_PX: f32 = 600.0;
const UNWRAPPED_FIRST_LINE: &str = "A right margin narrows the wrap width of every line in ";

fn renderer() -> SlideRenderer {
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
}

fn lines(list: &SurfaceDisplayList, object_id: u32) -> &[PositionedTextLine] {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id: id,
                lines,
                ..
            } if *id == object_id => Some(lines.as_slice()),
            _ => None,
        })
        .unwrap()
}

fn text(line: &PositionedTextLine) -> String {
    line.runs.iter().map(|run| run.text.as_str()).collect()
}

fn package() -> PptxPackage {
    DeckSession::open(DECK, 2951).unwrap().package().clone()
}

fn display_list(package: &PptxPackage) -> SurfaceDisplayList {
    let session = DeckSession::open(DECK, 2951).unwrap();
    renderer()
        .layout_slide(package, &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
}

#[test]
fn a_right_margin_narrows_the_wrap_width_at_every_rung_of_the_cascade() {
    let list = display_list(&package());
    for (object_id, line_index, wrap_width_px, first_line) in [
        (
            2,
            0,
            BOX_WIDTH_PX - 96.0,
            "A right margin narrows the wrap width of every ",
        ),
        (
            2,
            2,
            BOX_WIDTH_PX - 48.0,
            "A right margin narrows the wrap width of every line ",
        ),
        (
            3,
            0,
            BOX_WIDTH_PX - 216.0,
            "A right margin narrows the wrap ",
        ),
    ] {
        let lines = lines(&list, object_id);
        assert_eq!(text(&lines[line_index]), first_line, "shape {object_id}");
        for line in &lines[line_index..line_index + 2] {
            assert!(
                line.width <= wrap_width_px,
                "shape {object_id} line {line_index}"
            );
        }
    }
}

#[test]
fn a_direct_paragraph_margin_overrides_the_shape_list_style() {
    let list = display_list(&package());
    let lines = lines(&list, 4);
    assert_eq!(text(&lines[0]), UNWRAPPED_FIRST_LINE);
    assert!(lines[0].width > BOX_WIDTH_PX - 48.0);
    assert!(lines[0].width <= BOX_WIDTH_PX);
    assert_eq!(text(&lines[2]), "A right margin narrows the wrap ");
}

#[test]
fn clearing_every_right_margin_restores_the_full_box_width() {
    let mut package = package();
    for shapes in package
        .slides
        .iter_mut()
        .map(|slide| &mut slide.shapes)
        .chain(package.layouts.iter_mut().map(|layout| &mut layout.shapes))
    {
        clear_shapes(shapes);
    }
    for master in &mut package.masters {
        for properties in &mut master.text_styles.body {
            properties.margin_right = None;
        }
    }
    let list = display_list(&package);
    for (object_id, line_index) in [(2, 0), (2, 2), (3, 0), (4, 0), (4, 2)] {
        assert_eq!(
            text(&lines(&list, object_id)[line_index]),
            UNWRAPPED_FIRST_LINE,
            "shape {object_id} line {line_index}"
        );
    }
}

fn clear_shapes(shapes: &mut [ShapeNode]) {
    for shape in shapes {
        match shape {
            ShapeNode::Shape(shape) => {
                if let Some(body) = &mut shape.text {
                    clear_margins(body);
                }
            }
            ShapeNode::Group(group) => clear_shapes(&mut group.children),
            _ => {}
        }
    }
}

fn clear_margins(body: &mut TextBody) {
    for properties in body
        .default_list_style
        .as_deref_mut()
        .into_iter()
        .chain(body.list_style.iter_mut())
        .chain(
            body.paragraphs
                .iter_mut()
                .map(|paragraph| &mut paragraph.properties),
        )
    {
        properties.margin_right = None;
    }
}
