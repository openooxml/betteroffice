use ooxml_drawingml::GeometryPathCommand as C;
use pptx_edit::DeckSession;
use pptx_render::{Paint, Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/metafile-gradient.pptx");

fn shading(bytes: &[u8]) -> Vec<(String, [f64; 4])> {
    let session = DeckSession::open(bytes, 318).unwrap();
    let rendered = SlideRenderer::new()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    rendered
        .display_list
        .primitives
        .iter()
        .map(|primitive| {
            let Primitive::Shape {
                fill: Some(Paint::Solid { color }),
                path,
                ..
            } = primitive
            else {
                panic!("shaded metafile did not become filled vector artwork: {primitive:?}");
            };
            let mut box_rect = [f64::MAX, f64::MAX, f64::MIN, f64::MIN];
            for command in path {
                let (x, y) = match command {
                    C::Move { x, y } | C::Line { x, y } => (*x, *y),
                    _ => continue,
                };
                box_rect = [
                    box_rect[0].min(x),
                    box_rect[1].min(y),
                    box_rect[2].max(x),
                    box_rect[3].max(y),
                ];
            }
            (color.clone(), box_rect)
        })
        .collect()
}

fn green(color: &str) -> u8 {
    u8::from_str_radix(&color[3..5], 16).unwrap()
}

#[test]
fn gradient_and_pattern_blit_records_shade_the_picture() {
    let ops = shading(FIXTURE);
    assert_eq!(ops.len(), 193);

    let (triangles, rest) = ops.split_at(128);
    assert!(triangles.iter().all(|(_, box_rect)| box_rect[3] <= 0.5));
    for pair in triangles.chunks(64) {
        assert_eq!((green(&pair[0].0), green(&pair[63].0)), (0x4f, 0xa8));
        assert!(
            pair.windows(2)
                .all(|two| green(&two[0].0) < green(&two[1].0))
        );
    }

    let (bars, blit) = rest.split_at(64);
    assert_eq!(bars[0].0, "#f0e3c1");
    assert_eq!(bars[63].0, "#237b3d");
    assert_eq!(bars[0].1[0], 0.0);
    assert_eq!(bars[63].1[2], 0.5);
    assert!(bars.iter().all(|(_, box_rect)| box_rect[1] == 0.5));

    assert_eq!(blit, [("#e03a2f".to_owned(), [0.5, 0.5, 1.0, 0.8])]);
}
