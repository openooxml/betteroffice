use super::*;

fn geometry_section(index: Option<u32>, control: Option<(&str, &str)>) -> Section {
    let mut children = vec![
        row(0, "MoveTo", vec![cell("X", "0"), cell("Y", "0")]),
        row(1, "LineTo", vec![cell("X", "1"), cell("Y", "0")]),
        row(2, "LineTo", vec![cell("X", "1"), cell("Y", "1")]),
        row(3, "LineTo", vec![cell("X", "0"), cell("Y", "1")]),
    ]
    .into_iter()
    .map(SectionChild::Row)
    .collect::<Vec<_>>();
    if let Some((name, formula)) = control {
        children.push(SectionChild::Unknown(vsdx_parse::OpaqueXml {
            name: "Cell".into(),
            attributes: vec![
                ("N".into(), name.into()),
                ("F".into(), formula.into()),
                ("V".into(), "0".into()),
            ],
            children: Vec::new(),
        }));
    }
    Section {
        name: "Geometry".into(),
        index,
        del: false,
        children,
        other_attrs: vec![],
    }
}

fn paint_control_shape(id: u32, sections: Vec<Section>) -> Shape {
    let mut shape = shape(id, 1.0, 1.0);
    shape
        .children
        .retain(|child| !matches!(child, ShapeChild::Section(_)));
    for section in sections {
        shape.children.push(ShapeChild::Section(section));
    }
    shape
}

fn connector_shape(id: u32, sections: Vec<Section>) -> Shape {
    let mut shape = paint_control_shape(id, sections);
    shape.children.extend([
        ShapeChild::Cell(cell("OneD", "1")),
        ShapeChild::Cell(cell("BeginX", "1")),
        ShapeChild::Cell(cell("BeginY", "1")),
        ShapeChild::Cell(cell("EndX", "4")),
        ShapeChild::Cell(cell("EndY", "1")),
    ]);
    shape
}

fn with_formula(shape: &mut Shape, name: &str, formula_value: &str) {
    shape.children.retain(
        |child| !matches!(child, ShapeChild::Cell(Cell { name: actual, .. }) if actual == name),
    );
    shape.children.push(ShapeChild::Cell(Cell {
        name: name.into(),
        formula: Some(formula_value.into()),
        value: None,
        unit: None,
        del: false,
        other_attrs: vec![],
    }));
}

fn shape_primitives(list: &VsdxDisplayList) -> Vec<&Primitive> {
    list.primitives
        .iter()
        .filter(|primitive| matches!(primitive, Primitive::Shape { .. }))
        .collect()
}

#[test]
fn geometry_no_fill_strokes_without_filling() {
    let list = render(vec![paint_control_shape(
        1,
        vec![geometry_section(None, Some(("NoFill", "1")))],
    )]);
    let primitives = shape_primitives(&list);
    assert_eq!(primitives.len(), 1);
    match primitives[0] {
        Primitive::Shape {
            id, fill, stroke, ..
        } => {
            assert_eq!(id, "page:1");
            assert!(fill.is_none());
            assert!(stroke.is_some());
        }
        _ => unreachable!(),
    }
    assert!(
        !list
            .primitives
            .iter()
            .any(|primitive| matches!(primitive, Primitive::Placeholder { .. }))
    );
}

#[test]
fn geometry_no_line_fills_without_stroking() {
    let list = render(vec![paint_control_shape(
        1,
        vec![geometry_section(None, Some(("NoLine", "1")))],
    )]);
    let primitives = shape_primitives(&list);
    assert_eq!(primitives.len(), 1);
    match primitives[0] {
        Primitive::Shape {
            id, fill, stroke, ..
        } => {
            assert_eq!(id, "page:1");
            assert!(fill.is_some());
            assert!(stroke.is_none());
        }
        _ => unreachable!(),
    }
}

#[test]
fn geometry_no_show_emits_no_geometry_primitive() {
    let list = render(vec![paint_control_shape(
        1,
        vec![geometry_section(None, Some(("NoShow", "1")))],
    )]);
    assert!(shape_primitives(&list).is_empty());
    assert!(
        !list
            .primitives
            .iter()
            .any(|primitive| matches!(primitive, Primitive::Placeholder { .. }))
    );
}

#[test]
fn geometry_hidden_section_does_not_hide_sibling_sections() {
    let list = render(vec![paint_control_shape(
        1,
        vec![
            geometry_section(Some(0), Some(("NoShow", "1"))),
            geometry_section(Some(1), None),
        ],
    )]);
    let primitives = shape_primitives(&list);
    assert_eq!(primitives.len(), 1);
    match primitives[0] {
        Primitive::Shape { id, path, .. } => {
            assert_eq!(id, "page:1");
            assert!(!path.is_empty());
        }
        _ => unreachable!(),
    }
}

#[test]
fn geometry_sections_carry_their_own_paint() {
    let list = render(vec![paint_control_shape(
        1,
        vec![
            geometry_section(Some(0), Some(("NoFill", "1"))),
            geometry_section(Some(1), Some(("NoLine", "1"))),
        ],
    )]);
    let primitives = shape_primitives(&list);
    assert_eq!(primitives.len(), 2);
    for primitive in &primitives {
        match primitive {
            Primitive::Shape { id, .. } => assert_eq!(*id, "page:1"),
            _ => unreachable!(),
        }
    }
    match (primitives[0], primitives[1]) {
        (
            Primitive::Shape {
                fill: first_fill,
                stroke: first_stroke,
                ..
            },
            Primitive::Shape {
                fill: second_fill,
                stroke: second_stroke,
                ..
            },
        ) => {
            assert!(first_fill.is_none());
            assert!(first_stroke.is_some());
            assert!(second_fill.is_some());
            assert!(second_stroke.is_none());
        }
        _ => unreachable!(),
    }
}

#[test]
fn geometry_inactive_controls_paint_normally() {
    for control in ["NoFill", "NoLine", "NoShow"] {
        let list = render(vec![paint_control_shape(
            1,
            vec![geometry_section(None, Some((control, "0")))],
        )]);
        let primitives = shape_primitives(&list);
        assert_eq!(primitives.len(), 1, "{control}");
        match primitives[0] {
            Primitive::Shape { fill, stroke, .. } => {
                assert!(fill.is_some(), "{control}");
                assert!(stroke.is_some(), "{control}");
            }
            _ => unreachable!(),
        }
    }
}

#[test]
fn geometry_unevaluable_controls_produce_one_placeholder() {
    for control in ["NoFill", "NoLine", "NoShow"] {
        let list = render(vec![paint_control_shape(
            1,
            vec![
                geometry_section(Some(0), None),
                geometry_section(Some(1), Some((control, "Unknown(1)"))),
            ],
        )]);
        assert!(
            matches!(list.primitives.as_slice(), [Primitive::Placeholder { id, z_order: 0, reason, .. }] if id == "page:1" && reason.contains(control))
        );
    }
}

#[test]
fn connector_no_show_emits_no_geometry_primitive() {
    for invalid_endpoint in [false, true] {
        let mut shape = connector_shape(1, vec![geometry_section(None, Some(("NoShow", "1")))]);
        if invalid_endpoint {
            with_formula(&mut shape, "EndX", "Unknown(1)");
        }
        assert!(render(vec![shape]).primitives.is_empty());
    }
}

#[test]
fn connector_no_line_fills_without_stroking() {
    let list = render(vec![connector_shape(
        1,
        vec![geometry_section(None, Some(("NoLine", "1")))],
    )]);
    let primitives = shape_primitives(&list);
    assert_eq!(primitives.len(), 1);
    match primitives[0] {
        Primitive::Shape {
            id, fill, stroke, ..
        } => {
            assert_eq!(id, "page:1");
            assert!(fill.is_some());
            assert!(stroke.is_none());
        }
        _ => unreachable!(),
    }
}

#[test]
fn unused_paint_channel_does_not_block_the_used_channel() {
    for make_shape in [paint_control_shape, connector_shape] {
        for (control, unused) in [("NoFill", "FillForegnd"), ("NoLine", "LineColor")] {
            let mut shape = make_shape(1, vec![geometry_section(None, Some((control, "1")))]);
            with_formula(&mut shape, unused, "Unknown(1)");
            let list = render(vec![shape]);
            assert!(
                matches!(list.primitives.as_slice(), [Primitive::Shape { id, fill, stroke, .. }] if id == "page:1" && fill.is_some() == (control == "NoLine") && stroke.is_some() == (control == "NoFill"))
            );
        }
    }
}

#[test]
fn unused_paint_channel_failure_is_not_reported() {
    for make_shape in [paint_control_shape, connector_shape] {
        for (control, unused) in [("NoFill", "FillForegnd"), ("NoLine", "LineColor")] {
            let mut shape = make_shape(1, vec![geometry_section(None, Some((control, "1")))]);
            with_formula(&mut shape, unused, "Unknown(1)");
            let list = render(vec![shape]);
            assert!(
                matches!(list.primitives.as_slice(), [Primitive::Shape { diagnostics, .. }] if diagnostics.is_empty())
            );
        }
    }
}

#[test]
fn connector_sections_emit_one_route_with_the_used_channels() {
    let list = render(vec![connector_shape(
        1,
        vec![
            geometry_section(Some(10), Some(("NoFill", "1"))),
            geometry_section(Some(2), Some(("NoLine", "1"))),
            geometry_section(Some(0), Some(("NoShow", "1"))),
        ],
    )]);
    assert!(
        matches!(list.primitives.as_slice(), [Primitive::Shape { id, z_order: 0, fill: Some(_), stroke: Some(_), path, .. }] if id == "page:1" && path == &connector_route(ScenePoint { x: 1.0, y: 1.0 }, ScenePoint { x: 4.0, y: 1.0 }, None))
    );
}

#[test]
fn empty_sections_do_not_request_unused_paint() {
    for (control, unused) in [("NoFill", "FillForegnd"), ("NoLine", "LineColor")] {
        let mut empty = geometry_section(Some(0), None);
        empty.children.clear();
        let mut shape = paint_control_shape(
            1,
            vec![empty, geometry_section(Some(1), Some((control, "1")))],
        );
        with_formula(&mut shape, unused, "Unknown(1)");
        let list = render(vec![shape]);
        assert!(
            matches!(list.primitives.as_slice(), [Primitive::Shape { fill, stroke, .. }] if fill.is_some() == (control == "NoLine") && stroke.is_some() == (control == "NoFill"))
        );
    }
}

#[test]
fn defaulted_paint_is_reported_once_across_sections() {
    let mut shape = paint_control_shape(
        1,
        vec![
            geometry_section(Some(0), None),
            geometry_section(Some(1), None),
        ],
    );
    with_formula(&mut shape, "FillForegnd", "Unknown(1)");
    let list = render(vec![shape]);
    let primitives = shape_primitives(&list);
    assert_eq!(primitives.len(), 2);
    let reported = primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                fill, diagnostics, ..
            } => {
                assert!(fill.is_some());
                Some(diagnostics)
            }
            _ => None,
        })
        .flatten()
        .map(|diagnostic| diagnostic.code.as_str())
        .collect::<Vec<_>>();
    assert_eq!(reported, ["unresolvable-fill-colour"]);
}

#[test]
fn geometry_sections_preserve_numeric_order_identity_and_hit_testing() {
    let mut stroked = geometry_section(Some(10), Some(("NoFill", "1")));
    for child in &mut stroked.children {
        if let SectionChild::Row(row) = child {
            for child in &mut row.children {
                if let RowChild::Cell(cell) = child
                    && cell.name == "X"
                {
                    cell.value = Some(
                        (cell.value.as_ref().unwrap().parse::<f64>().unwrap() + 2.0).to_string(),
                    );
                }
            }
        }
    }
    let shape = paint_control_shape(
        1,
        vec![stroked, geometry_section(Some(2), Some(("NoLine", "1")))],
    );
    let behind = super::shape(7, 8.0, 6.0);
    let list = render(vec![behind.clone(), shape.clone()]);
    assert!(
        matches!(list.primitives.as_slice(), [_, Primitive::Shape { id: first, z_order: 1, fill: Some(_), stroke: None, .. }, Primitive::Shape { id: second, z_order: 1, fill: None, stroke: Some(_), .. }] if first == "page:1" && second == first)
    );
    for (x, y, expected) in [
        (1.5, 1.5, Some("page:1")),
        (3.5, 1.0, Some("page:1")),
        (2.5, 1.5, None),
        (3.5, 1.5, None),
    ] {
        let hit = hit_test(&list, x * PIXELS_PER_INCH, (8.0 - y) * PIXELS_PER_INCH);
        assert_eq!(
            hit,
            expected.map(|id| HitTestResult::Shape {
                shape_id: id.into()
            })
        );
    }
    let list = render(vec![behind, shape, super::shape(9, 1.0, 1.0)]);
    assert_eq!(
        hit_test(&list, 1.5 * PIXELS_PER_INCH, 6.5 * PIXELS_PER_INCH),
        Some(HitTestResult::Shape {
            shape_id: "page:9".into()
        })
    );
}

#[test]
fn unsupported_geometry_emits_one_placeholder_without_partial_paths() {
    let mut invalid = geometry_section(Some(1), None);
    invalid
        .children
        .push(SectionChild::Row(row(4, "NURBSTo", vec![])));
    let list = render(vec![paint_control_shape(
        1,
        vec![geometry_section(Some(0), None), invalid],
    )]);
    assert!(
        matches!(list.primitives.as_slice(), [Primitive::Placeholder { id, z_order: 0, reason, .. }] if id == "page:1" && reason.contains("NURBSTo"))
    );
}

#[test]
fn geometry_editing_cells_do_not_suppress_drawing() {
    for control in ["NoSnap", "NoQuickDrag"] {
        let list = render(vec![paint_control_shape(
            1,
            vec![geometry_section(None, Some((control, "1")))],
        )]);
        assert!(
            matches!(
                list.primitives.as_slice(),
                [Primitive::Shape {
                    fill: Some(_),
                    stroke: Some(_),
                    ..
                }]
            ),
            "{control}"
        );
    }
}

#[test]
fn geometry_unknown_control_reaches_the_placeholder() {
    let list = render(vec![paint_control_shape(
        1,
        vec![geometry_section(None, Some(("NoSuchControl", "1")))],
    )]);
    assert!(
        matches!(list.primitives.as_slice(), [Primitive::Placeholder { reason, .. }] if reason.contains("NoSuchControl"))
    );
}

#[test]
fn hidden_geometry_preserves_one_text_box() {
    let mut shape = paint_control_shape(
        1,
        vec![
            geometry_section(Some(0), Some(("NoShow", "1"))),
            geometry_section(Some(1), Some(("NoShow", "1"))),
        ],
    );
    shape
        .children
        .push(ShapeChild::Text(vec![TextToken::Literal("label".into())]));
    let list = render(vec![shape]);
    assert!(
        matches!(list.primitives.as_slice(), [Primitive::TextBox { id, .. }] if id == "page:1")
    );
}
