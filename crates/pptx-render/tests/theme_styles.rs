use pptx_edit::{DeckSession, EditCtx, ShapeStroke};
use pptx_render::{Paint, Primitive, SlideRenderer, Stroke, StrokeEnd, SurfaceDisplayList};

const MATRIX: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/style-matrix-deck.pptx");
const MASTER: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/master-style-deck.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn render(session: &DeckSession) -> SurfaceDisplayList {
    render_index(session, 0)
}

fn render_index(session: &DeckSession, index: usize) -> SurfaceDisplayList {
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        renderer.register_font("Arial", bold, false, FONT).unwrap();
    }
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), index)
        .unwrap()
        .display_list
}

fn shape<'a>(list: &'a SurfaceDisplayList, name: &str) -> (Option<&'a Paint>, Option<&'a Stroke>) {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Shape {
                name: actual,
                shape_id: Some(id),
                fill,
                stroke,
                ..
            } if actual == name && id.starts_with("slide:") => {
                Some((fill.as_ref(), stroke.as_ref()))
            }
            _ => None,
        })
        .unwrap_or_else(|| panic!("missing {name}"))
}

fn solid(value: &str) -> Option<Paint> {
    Some(Paint::Solid {
        color: value.to_owned(),
    })
}

fn matrix_stroke(color: &str) -> Stroke {
    Stroke {
        color: color.to_owned(),
        width: 3.0,
        dashed: true,
        head_end: Some(StrokeEnd {
            kind: "triangle".to_owned(),
            width: 15.0,
            length: 6.0,
        }),
        tail_end: Some(StrokeEnd {
            kind: "oval".to_owned(),
            width: 6.0,
            length: 15.0,
        }),
    }
}

#[test]
fn theme_matrix_renders_reference_colours_indices_and_explicit_line_defaults() {
    let session = DeckSession::open(MATRIX, 8010).unwrap();
    let list = render(&session);
    for name in ["PictureLine", "MasterPictureLine"] {
        let stroke = list.primitives.iter().find_map(|p| match p {
            Primitive::Image {
                name: actual,
                stroke,
                ..
            } if actual == name => stroke.as_ref(),
            _ => None,
        });
        assert_eq!(stroke, Some(&matrix_stroke("#CDEFAB")), "{name}");
    }

    for (name, color) in [
        ("FillRef3", "#474747"),
        ("FillRef1001", "#4A4A4A"),
        ("FillRef1002", "#556677"),
        ("ExplicitFill", "#112233"),
    ] {
        assert_eq!(shape(&list, name).0, solid(color).as_ref(), "{name}");
    }
    for name in [
        "FillRef0",
        "FillRef1000",
        "ExplicitNoFill",
        "UnsupportedFill",
        "LineSlot2",
    ] {
        assert_eq!(shape(&list, name), (None, None), "{name}");
    }
    let Some(Paint::Gradient {
        stops, angle_deg, ..
    }) = shape(&list, "Gradient2").0
    else {
        panic!("missing gradient")
    };
    assert_eq!(*angle_deg, Some(90.0));
    assert_eq!(
        stops
            .iter()
            .map(|s| (s.position, s.color.as_str()))
            .collect::<Vec<_>>(),
        [(0.0, "#474747"), (0.5, "#123456"), (1.0, "#4A4A4A")]
    );
    for (name, expected) in [
        ("LnRef3", matrix_stroke("#0F9ED5")),
        (
            "WidthOnly",
            Stroke {
                color: "#474747".to_owned(),
                width: 4.0,
                dashed: false,
                head_end: None,
                tail_end: None,
            },
        ),
        ("ExplicitLineColour", matrix_stroke("#112233")),
    ] {
        assert_eq!(shape(&list, name).1, Some(&expected), "{name}");
    }
}

#[test]
fn placeholder_properties_outrank_shape_and_layout_style_references() {
    let list = render(&DeckSession::open(MATRIX, 8011).unwrap());
    assert_eq!(shape(&list, "Placeholder50").0, solid("#334455").as_ref());
    assert_eq!(
        shape(&list, "Placeholder50").1,
        Some(&Stroke {
            color: "#667788".to_owned(),
            width: 8.0,
            dashed: true,
            head_end: Some(StrokeEnd {
                kind: "triangle".to_owned(),
                width: 40.0,
                length: 16.0,
            }),
            tail_end: Some(StrokeEnd {
                kind: "oval".to_owned(),
                width: 16.0,
                length: 40.0,
            }),
        })
    );
    assert_eq!(shape(&list, "Placeholder51"), (None, None));
    assert_eq!(shape(&list, "Placeholder52").0, solid("#ABCDEF").as_ref());
    let inherited = render_index(&DeckSession::open(MATRIX, 8014).unwrap(), 1);
    for (name, color) in [
        ("PhOwnStyle", "#00FF00"),
        ("PhInheritStyle", "#4EA72E"),
        ("PhInheritFill", "#00FF00"),
    ] {
        assert_eq!(shape(&inherited, name).0, solid(color).as_ref(), "{name}");
    }
}

#[test]
fn master_shapes_render_theme_fills() {
    let list = render(&DeckSession::open(MASTER, 8012).unwrap());
    let fills: Vec<_> = list
        .primitives
        .iter()
        .filter_map(|p| match p {
            Primitive::Shape {
                shape_id: Some(id),
                fill: Some(Paint::Solid { color }),
                ..
            } => Some((id.split(':').next().unwrap(), color.as_str())),
            _ => None,
        })
        .collect();
    assert!(fills.contains(&("master", "#000000")));
    assert!(fills.contains(&("slide", "#FF0000")));
    assert_eq!(
        shape(&list, "Oval 3").1,
        Some(&Stroke {
            color: "#030E13".to_owned(),
            width: 2.0,
            dashed: false,
            head_end: None,
            tail_end: None,
        })
    );
}

#[test]
fn edited_outlines_override_the_resolved_style() {
    let session = DeckSession::open(MATRIX, 8013).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let target = slide.shapes.iter().find(|s| s.name == "WidthOnly").unwrap();
    session
        .set_shape_stroke(
            &EditCtx::local("test"),
            &slide.id,
            &target.id,
            &ShapeStroke {
                color: None,
                width_pt: Some(1.5),
            },
        )
        .unwrap();
    assert_eq!(
        shape(&render(&session), "WidthOnly").1,
        Some(&Stroke {
            color: "#000000".to_owned(),
            width: 2.0,
            dashed: false,
            head_end: None,
            tail_end: None,
        })
    );
    session
        .set_shape_stroke(
            &EditCtx::local("test"),
            &slide.id,
            &target.id,
            &ShapeStroke {
                color: Some("#FEDCBA".to_owned()),
                width_pt: Some(1.5),
            },
        )
        .unwrap();
    let list = render(&session);
    assert_eq!(
        shape(&list, "WidthOnly").1,
        Some(&Stroke {
            color: "#FEDCBA".to_owned(),
            width: 2.0,
            dashed: false,
            head_end: None,
            tail_end: None,
        })
    );
    session
        .set_shape_stroke(
            &EditCtx::local("test"),
            &slide.id,
            &target.id,
            &ShapeStroke::default(),
        )
        .unwrap();
    assert_eq!(shape(&render(&session), "WidthOnly").1, None);
}

#[test]
fn theme_reference_and_style_alpha_reach_fills_gradients_and_outlines() {
    for (style_alpha, suffix) in [(false, "80"), (true, "40")] {
        let mut parts = ooxml_opc::unzip_parts(MATRIX).unwrap();
        for (path, bytes) in &mut parts {
            if path == "ppt/slides/slide1.xml" {
                *bytes = String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace(
                        r#"<a:srgbClr val="808080">"#,
                        r#"<a:srgbClr val="808080"><a:alpha val="50000"/>"#,
                    )
                    .into_bytes();
            }
            if style_alpha && path == "ppt/theme/theme1.xml" {
                *bytes = String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace(
                        r#"<a:schemeClr val="phClr">"#,
                        r#"<a:schemeClr val="phClr"><a:alpha val="25000"/>"#,
                    )
                    .into_bytes();
            }
        }
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let list = render(&DeckSession::open(&bytes, 8015).unwrap());
        assert_eq!(
            shape(&list, "FillRef3").0,
            solid(&format!("#474747{suffix}")).as_ref()
        );
        assert_eq!(
            shape(&list, "WidthOnly").1.unwrap().color,
            format!("#474747{suffix}")
        );
        let Some(Paint::Gradient { stops, .. }) = shape(&list, "Gradient2").0 else {
            panic!("missing gradient")
        };
        assert_eq!(
            stops
                .iter()
                .map(|stop| stop.color.clone())
                .collect::<Vec<_>>(),
            [
                format!("#474747{suffix}"),
                "#123456".to_owned(),
                format!("#4A4A4A{suffix}")
            ]
        );
    }
}
