use super::*;
use vsdx_parse::{Row, RowChild, Section, SectionChild};

fn stop_cell(name: &str, formula: Option<&str>, value: Option<&str>) -> Cell {
    Cell {
        name: name.into(),
        formula: formula.map(str::to_owned),
        value: value.map(str::to_owned),
        unit: None,
        del: false,
        other_attrs: vec![],
    }
}

fn stop_row(index: u32, color: (&str, &str), position: &str) -> SectionChild {
    stop_row_with_trans(index, color, position, "0")
}

fn stop_row_with_trans(
    index: u32,
    color: (&str, &str),
    position: &str,
    trans: &str,
) -> SectionChild {
    SectionChild::Row(Row {
        index: Some(index),
        name: None,
        local_name: None,
        row_type: None,
        del: false,
        children: vec![
            RowChild::Cell(stop_cell("GradientStopColor", Some(color.0), Some(color.1))),
            RowChild::Cell(stop_cell("GradientStopColorTrans", None, Some(trans))),
            RowChild::Cell(stop_cell("GradientStopPosition", None, Some(position))),
        ],
        other_attrs: vec![],
    })
}

fn gradient_shape(id: u32, angle: &str, rows: Vec<SectionChild>) -> Shape {
    let mut shape = shape(id, 1.0, 1.0);
    shape
        .children
        .push(ShapeChild::Cell(cell("FillGradientEnabled", "1")));
    shape
        .children
        .push(ShapeChild::Cell(cell("FillGradientAngle", angle)));
    shape.children.push(ShapeChild::Section(Section {
        name: "FillGradient".into(),
        index: None,
        del: false,
        children: rows,
        other_attrs: vec![],
    }));
    shape
}

fn shape_fill(list: &VsdxDisplayList) -> Option<Paint> {
    match list.primitives.as_slice() {
        [Primitive::Shape { fill, .. }] => fill.clone(),
        _ => unreachable!("expected one shape primitive"),
    }
}

fn fidelity_codes(list: &VsdxDisplayList) -> Vec<&str> {
    list.primitives
        .iter()
        .flat_map(|primitive| match primitive {
            Primitive::Shape { diagnostics, .. } => diagnostics.as_slice(),
            _ => &[],
        })
        .filter(|diagnostic| diagnostic.category == DiagnosticCategory::Fidelity)
        .map(|diagnostic| diagnostic.code.as_str())
        .collect()
}

#[test]
fn linear_gradient_resolves_stops_through_the_colour_path() {
    let mut package = package(vec![gradient_shape(
        1,
        &std::f64::consts::FRAC_PI_2.to_string(),
        vec![
            stop_row(0, ("THEMEVAL(\"FillStop1Color\",3)", "3"), "0"),
            stop_row(1, ("GUARD(RGB(78,142,194))", "#4E8EC2"), "1"),
        ],
    )]);
    package.colors = serde_json::from_value(serde_json::json!([
        {"name": "ColorEntry", "attributes": [["IX", "3"], ["RGB", "0A0B0C"]]}
    ]))
    .unwrap();
    let list = Renderer::default().layout_page(&package, "page").unwrap();
    match shape_fill(&list) {
        Some(Paint::Gradient { angle_deg, stops }) => {
            assert!((angle_deg.unwrap() - 90.0).abs() < 0.01);
            assert_eq!(
                stops,
                vec![
                    GradientStop {
                        position: 0.0,
                        color: "#0A0B0C".into(),
                    },
                    GradientStop {
                        position: 1.0,
                        color: "#4E8EC2".into(),
                    },
                ]
            );
        }
        other => unreachable!("expected gradient, got {other:?}"),
    }
}

#[test]
fn gradient_stop_formula_failure_falls_back_to_the_cached_value() {
    let list = render(vec![gradient_shape(
        1,
        "0",
        vec![
            stop_row(
                0,
                (
                    "THEMEGUARD(IF(ISTHEMED(),SETATREFEXPR(RGB(1,2,3)),SETATREFEXPR(RGB(255,255,255))))",
                    "#FFFFFF",
                ),
                "0",
            ),
            stop_row(1, ("GUARD(RGB(202,187,151))", "#CABB97"), "1"),
        ],
    )]);
    match shape_fill(&list) {
        Some(Paint::Gradient { stops, .. }) => {
            assert_eq!(stops[0].color, "#FFFFFF");
            assert_eq!(stops[1].color, "#CABB97");
        }
        other => unreachable!("expected gradient, got {other:?}"),
    }
}

#[test]
fn gradient_stops_sort_by_position_and_skip_deleted_rows() {
    let mut rows = vec![
        stop_row(1, ("GUARD(RGB(0,0,255))", "#0000FF"), "1"),
        stop_row(0, ("GUARD(RGB(255,0,0))", "#FF0000"), "0"),
    ];
    rows.push(SectionChild::Row(Row {
        index: Some(2),
        name: None,
        local_name: None,
        row_type: None,
        del: true,
        children: vec![],
        other_attrs: vec![],
    }));
    let list = render(vec![gradient_shape(1, "0", rows)]);
    match shape_fill(&list) {
        Some(Paint::Gradient { stops, .. }) => {
            assert_eq!(
                stops
                    .iter()
                    .map(|stop| stop.color.as_str())
                    .collect::<Vec<_>>(),
                ["#FF0000", "#0000FF"]
            );
        }
        other => unreachable!("expected gradient, got {other:?}"),
    }
}

#[test]
fn gradient_without_an_angle_falls_back_to_solid_without_a_diagnostic() {
    let mut shape = gradient_shape(
        1,
        &std::f64::consts::FRAC_PI_2.to_string(),
        vec![stop_row(0, ("GUARD(RGB(1,2,3))", "#010203"), "0")],
    );
    shape.children.retain(|child| {
        !matches!(child, ShapeChild::Cell(Cell { name, .. }) if name == "FillGradientAngle")
    });
    let list = render(vec![shape]);
    assert!(matches!(shape_fill(&list), Some(Paint::Solid { .. })));
    assert!(!fidelity_codes(&list).contains(&"unresolvable-fill-gradient"));
}

#[test]
fn gradient_with_fewer_than_two_stops_reports_and_falls_back_to_solid() {
    let list = render(vec![gradient_shape(
        1,
        "0",
        vec![stop_row(0, ("GUARD(RGB(1,2,3))", "#010203"), "0")],
    )]);
    assert!(matches!(shape_fill(&list), Some(Paint::Solid { .. })));
    assert!(fidelity_codes(&list).contains(&"unresolvable-fill-gradient"));
}

#[test]
fn a_transparent_stop_keeps_the_gradient_and_reports_the_lost_transparency() {
    let list = render(vec![gradient_shape(
        1,
        "0",
        vec![
            stop_row_with_trans(0, ("GUARD(RGB(255,0,0))", "#FF0000"), "0", "0"),
            stop_row_with_trans(1, ("GUARD(RGB(0,255,0))", "#00FF00"), "0.5", "0.5"),
            stop_row_with_trans(2, ("GUARD(RGB(0,0,255))", "#0000FF"), "1", "0"),
        ],
    )]);
    match shape_fill(&list) {
        Some(Paint::Gradient { stops, .. }) => assert_eq!(
            stops
                .iter()
                .map(|stop| stop.color.as_str())
                .collect::<Vec<_>>(),
            ["#FF0000", "#00FF00", "#0000FF"]
        ),
        other => unreachable!("expected gradient, got {other:?}"),
    }
    assert!(fidelity_codes(&list).contains(&"lossy-fill-gradient"));
}

#[test]
fn a_non_linear_gradient_direction_reports_and_falls_back_to_solid() {
    let mut shape = gradient_shape(
        1,
        "0",
        vec![
            stop_row(0, ("GUARD(RGB(255,0,0))", "#FF0000"), "0"),
            stop_row(1, ("GUARD(RGB(0,0,255))", "#0000FF"), "1"),
        ],
    );
    shape
        .children
        .push(ShapeChild::Cell(cell("FillGradientDir", "5")));
    let list = render(vec![shape]);
    assert!(matches!(shape_fill(&list), Some(Paint::Solid { .. })));
    assert!(fidelity_codes(&list).contains(&"unresolvable-fill-gradient"));
}

#[test]
fn an_unresolvable_stop_colour_reports_and_falls_back_to_solid() {
    let list = render(vec![gradient_shape(
        1,
        "0",
        vec![
            stop_row(0, ("GUARD(RGB(255,0,0))", "#FF0000"), "0"),
            SectionChild::Row(Row {
                index: Some(1),
                name: None,
                local_name: None,
                row_type: None,
                del: false,
                children: vec![
                    RowChild::Cell(stop_cell("GradientStopColor", None, Some("not a colour"))),
                    RowChild::Cell(stop_cell("GradientStopPosition", None, Some("1"))),
                ],
                other_attrs: vec![],
            }),
            stop_row(2, ("GUARD(RGB(0,0,255))", "#0000FF"), "1"),
        ],
    )]);
    assert!(matches!(shape_fill(&list), Some(Paint::Solid { .. })));
    assert!(fidelity_codes(&list).contains(&"unresolvable-fill-gradient"));
}

#[test]
fn disabled_gradient_paints_the_solid_fill() {
    let mut shape = gradient_shape(
        1,
        "0",
        vec![
            stop_row(0, ("GUARD(RGB(255,0,0))", "#FF0000"), "0"),
            stop_row(1, ("GUARD(RGB(0,0,255))", "#0000FF"), "1"),
        ],
    );
    for child in &mut shape.children {
        if let ShapeChild::Cell(cell) = child
            && cell.name == "FillGradientEnabled"
        {
            cell.value = Some("0".into());
        }
    }
    let list = render(vec![shape]);
    assert!(matches!(
        shape_fill(&list),
        Some(Paint::Solid { color }) if color == "#010203"
    ));
}
