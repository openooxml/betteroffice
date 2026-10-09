use ooxml_drawingml::GeometryPathCommand as C;
use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/star-inner-radius.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn star(name: &str) -> Vec<(f64, f64)> {
    let session = DeckSession::open(FIXTURE, 318).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let rendered = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    let path = rendered
        .display_list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Shape {
                name: shape, path, ..
            } if shape == name => Some(path.clone()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no shape named {name}"));
    path.into_iter()
        .filter_map(|command| match command {
            C::Move { x, y } | C::Line { x, y } => Some((x, y)),
            _ => None,
        })
        .collect()
}

fn cross(origin: (f64, f64), a: (f64, f64), b: (f64, f64)) -> f64 {
    (a.0 - origin.0) * (b.1 - origin.1) - (a.1 - origin.1) * (b.0 - origin.0)
}

#[test]
fn a_default_five_point_star_is_a_regular_pentagram() {
    for name in ["star5 default 150x150", "star5 adj 19098 150x150"] {
        let v = star(name);
        assert!(cross(v[0], v[4], v[1]).abs() < 1e-5, "{name}");
        assert!(cross(v[0], v[4], v[3]).abs() < 1e-5, "{name}");
    }
}

#[test]
fn a_five_point_star_reaches_the_bottom_of_its_frame() {
    for name in ["star5 default 150x150", "star5 default 350x150"] {
        let bottom = star(name).iter().map(|p| p.1).fold(f64::MIN, f64::max);
        assert!((bottom - 1.0).abs() < 1e-4, "{name}: {bottom}");
    }
}

#[test]
fn a_star_adjustment_past_half_draws_the_half() {
    assert_eq!(
        star("star5 adj 80000 150x150"),
        star("star5 adj 50000 150x150")
    );
}
