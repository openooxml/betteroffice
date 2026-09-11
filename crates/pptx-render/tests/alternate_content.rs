use pptx_edit::DeckSession;
use pptx_render::{Paint, Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/alternate-content.pptx");
const SANS: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationSans-Regular.ttf");

fn render() -> SurfaceDisplayList {
    let session = DeckSession::open(DECK, 331).unwrap();
    let mut renderer = SlideRenderer::new();
    for family in ["Arial", "Liberation Sans"] {
        renderer.register_font(family, false, false, SANS).unwrap();
    }
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
}

fn shape(list: &SurfaceDisplayList, name: &str) -> (f32, f32, Option<Paint>) {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Shape {
                name: shape_name,
                x,
                y,
                fill,
                ..
            } if shape_name == name => Some((*x, *y, fill.clone())),
            _ => None,
        })
        .unwrap_or_else(|| panic!("{name} is missing from the display list"))
}

fn text(list: &SurfaceDisplayList, object_id: u32) -> String {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id: id,
                lines,
                ..
            } if *id == object_id => Some(
                lines
                    .iter()
                    .flat_map(|line| line.runs.iter().map(|run| run.text.clone()))
                    .collect::<String>(),
            ),
            _ => None,
        })
        .unwrap_or_else(|| panic!("object {object_id} has no text"))
}

#[test]
fn a_fallback_shape_in_the_tree_is_drawn_where_the_slide_places_it() {
    let list = render();

    let (x, y, fill) = shape(&list, "fallback");
    assert_eq!(x, 690.0);
    assert_eq!(y, 150.0);
    assert_eq!(
        fill,
        Some(Paint::Solid {
            color: "#315EFB".to_owned()
        })
    );
    assert_eq!(text(&list, 12), "Fallback shape");

    let (control_x, _, _) = shape(&list, "control");
    assert_eq!(control_x, 60.0);
    assert_eq!(text(&list, 11), "Control shape");
}
