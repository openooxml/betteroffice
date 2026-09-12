use crate::{GeometryIssue, Lookup, RealizedGeometry, ResolvedRow, ResolvedSection};
use ooxml_drawingml::GeometryPathCommand;

pub fn realize_geometry(section: &ResolvedSection) -> RealizedGeometry {
    let mut out = RealizedGeometry::default();
    let mut current = (0.0, 0.0);
    let mut rows: Vec<&ResolvedRow> = section.rows.values().collect();
    rows.sort_by(|left, right| row_order(&left.key).cmp(&row_order(&right.key)));
    for row in rows {
        let ty = row.row_type.as_deref().unwrap_or("");
        if matches!(
            ty,
            "NURBSTo" | "PolylineTo" | "SplineStart" | "SplineKnot" | "InfiniteLine"
        ) {
            out.issues
                .push(GeometryIssue::UnsupportedRowType(ty.into()));
            continue;
        }
        let value = |name: &str| {
            row.cells
                .get(name)
                .and_then(|v| match v {
                    Lookup::Found(value) => value.cell.value.as_deref(),
                    Lookup::Deleted | Lookup::Absent => None,
                })
                .and_then(|v| v.parse::<f64>().ok())
                .filter(|value| value.is_finite())
        };
        let mut required = |names: &[&str]| -> Option<Vec<f64>> {
            let mut values = Vec::with_capacity(names.len());
            for name in names {
                match value(name) {
                    Some(value) => values.push(value),
                    None if row.cells.contains_key(*name) => {
                        out.issues.push(GeometryIssue::UnevaluatedCell {
                            row_type: ty.into(),
                            cell: (*name).into(),
                        })
                    }
                    None => out.issues.push(GeometryIssue::MissingCell {
                        row_type: ty.into(),
                        cell: (*name).into(),
                    }),
                }
            }
            (values.len() == names.len()).then_some(values)
        };
        let xy = match (value("X"), value("Y")) {
            (Some(x), Some(y)) => (x, y),
            _ => {
                for n in ["X", "Y"] {
                    if row.cells.contains_key(n) && value(n).is_none() {
                        out.issues.push(GeometryIssue::UnevaluatedCell {
                            row_type: ty.into(),
                            cell: n.into(),
                        });
                    }
                }
                continue;
            }
        };
        match ty {
            "MoveTo" => {
                if push_checked(&mut out, GeometryPathCommand::Move { x: xy.0, y: xy.1 }, ty) {
                    current = xy;
                }
            }
            "LineTo" => {
                if push_checked(&mut out, GeometryPathCommand::Line { x: xy.0, y: xy.1 }, ty) {
                    current = xy;
                }
            }
            "RelMoveTo" => {
                let end = (current.0 + xy.0, current.1 + xy.1);
                if push_checked(
                    &mut out,
                    GeometryPathCommand::Move { x: end.0, y: end.1 },
                    ty,
                ) {
                    current = end;
                }
            }
            "RelLineTo" => {
                let end = (current.0 + xy.0, current.1 + xy.1);
                if push_checked(
                    &mut out,
                    GeometryPathCommand::Line { x: end.0, y: end.1 },
                    ty,
                ) {
                    current = end;
                }
            }
            "ArcTo" => {
                let Some(values) = required(&["A"]) else {
                    continue;
                };
                if emit_row(&mut out, |out| cubic_arc(out, current, xy, values[0], ty)) {
                    current = xy;
                }
            }
            "EllipticalArcTo" => {
                let Some(values) = required(&["A", "B", "C", "D"]) else {
                    continue;
                };
                if emit_row(&mut out, |out| {
                    cubic_elliptical_arc(
                        out,
                        current,
                        xy,
                        (values[0], values[1]),
                        values[2],
                        values[3],
                        ty,
                    )
                }) {
                    current = xy;
                }
            }
            "Ellipse" => {
                let Some(values) = required(&["A", "B", "C", "D"]) else {
                    continue;
                };
                if emit_row(&mut out, |out| {
                    cubic_ellipse(out, xy, values[0], values[1], values[2], values[3], ty)
                }) {
                    current = xy;
                }
            }
            _ => out
                .issues
                .push(GeometryIssue::UnsupportedRowType(ty.into())),
        }
    }
    out
}

/// Keys sort lexically, so `IX:10` would precede `IX:2` and displace relative rows.
fn row_order(key: &str) -> (u32, &str) {
    key.strip_prefix("IX:")
        .and_then(|index| index.parse().ok())
        .map_or((u32::MAX, key), |index| (index, ""))
}

fn emit_row(out: &mut RealizedGeometry, emit: impl FnOnce(&mut RealizedGeometry) -> bool) -> bool {
    let command_count = out.commands.len();
    if emit(out) {
        true
    } else {
        out.commands.truncate(command_count);
        false
    }
}

fn push_checked(out: &mut RealizedGeometry, command: GeometryPathCommand, row_type: &str) -> bool {
    let finite = match &command {
        GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
            x.is_finite() && y.is_finite()
        }
        GeometryPathCommand::Quad { cpx, cpy, x, y } => {
            cpx.is_finite() && cpy.is_finite() && x.is_finite() && y.is_finite()
        }
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => {
            cp1x.is_finite()
                && cp1y.is_finite()
                && cp2x.is_finite()
                && cp2y.is_finite()
                && x.is_finite()
                && y.is_finite()
        }
        GeometryPathCommand::Close => true,
    };
    if finite {
        out.commands.push(command);
    } else {
        out.issues.push(GeometryIssue::UnevaluatedCell {
            row_type: row_type.into(),
            cell: "geometry".into(),
        });
    }
    finite
}

fn cubic_arc(
    out: &mut RealizedGeometry,
    start: (f64, f64),
    end: (f64, f64),
    bow: f64,
    row_type: &str,
) -> bool {
    let dx = end.0 - start.0;
    let dy = end.1 - start.1;
    let chord = dx.hypot(dy);
    if chord == 0.0 || bow == 0.0 {
        return push_checked(
            out,
            GeometryPathCommand::Line { x: end.0, y: end.1 },
            row_type,
        );
    }
    let radius = chord * chord / (8.0 * bow.abs()) + bow.abs() / 2.0;
    let midpoint = ((start.0 + end.0) / 2.0, (start.1 + end.1) / 2.0);
    let normal = (-dy / chord * bow.signum(), dx / chord * bow.signum());
    let center = (
        midpoint.0 - normal.0 * (radius - bow.abs()),
        midpoint.1 - normal.1 * (radius - bow.abs()),
    );
    let a0 = (start.1 - center.1).atan2(start.0 - center.0);
    let a1 = (end.1 - center.1).atan2(end.0 - center.0);
    let bow_point = (
        midpoint.0 + normal.0 * bow.abs(),
        midpoint.1 + normal.1 * bow.abs(),
    );
    let bow_angle = (bow_point.1 - center.1).atan2(bow_point.0 - center.0);
    let sweep = [
        a1 - a0,
        a1 - a0 + std::f64::consts::TAU,
        a1 - a0 - std::f64::consts::TAU,
    ]
    .into_iter()
    .min_by(|left, right| {
        angle_distance(a0 + left / 2.0, bow_angle)
            .total_cmp(&angle_distance(a0 + right / 2.0, bow_angle))
    })
    .unwrap();
    cubic_arc_segment(
        out,
        center,
        (radius, radius),
        0.0,
        a0,
        sweep / 2.0,
        row_type,
    ) && cubic_arc_segment(
        out,
        center,
        (radius, radius),
        0.0,
        a0 + sweep / 2.0,
        sweep / 2.0,
        row_type,
    )
}

fn cubic_elliptical_arc(
    out: &mut RealizedGeometry,
    start: (f64, f64),
    end: (f64, f64),
    through: (f64, f64),
    angle: f64,
    axis_ratio: f64,
    row_type: &str,
) -> bool {
    let original_end = end;
    let ratio = axis_ratio.abs();
    if ratio <= f64::EPSILON {
        return push_checked(
            out,
            GeometryPathCommand::Line {
                x: original_end.0,
                y: original_end.1,
            },
            row_type,
        );
    }
    let rotate = |point: (f64, f64)| {
        (
            point.0 * angle.cos() + point.1 * angle.sin(),
            -point.0 * angle.sin() + point.1 * angle.cos(),
        )
    };
    let start = rotate(start);
    let end = rotate(end);
    let through = rotate(through);
    let metric = |point: (f64, f64)| (point.0, point.1 * ratio * ratio);
    let delta_end = metric((end.0 - start.0, end.1 - start.1));
    let delta_through = metric((through.0 - start.0, through.1 - start.1));
    let determinant = 2.0 * (delta_end.0 * delta_through.1 - delta_end.1 * delta_through.0);
    if determinant.abs() <= f64::EPSILON {
        return push_checked(
            out,
            GeometryPathCommand::Line {
                x: original_end.0,
                y: original_end.1,
            },
            row_type,
        );
    }
    let squared = |point: (f64, f64)| point.0 * point.0 + ratio * ratio * point.1 * point.1;
    let end_difference = squared(end) - squared(start);
    let through_difference = squared(through) - squared(start);
    let center = (
        (end_difference * delta_through.1 - delta_end.1 * through_difference) / determinant,
        (delta_end.0 * through_difference - end_difference * delta_through.0) / determinant,
    );
    let rx = squared((start.0 - center.0, start.1 - center.1)).sqrt();
    if rx <= f64::EPSILON {
        return push_checked(
            out,
            GeometryPathCommand::Line {
                x: original_end.0,
                y: original_end.1,
            },
            row_type,
        );
    }
    let ry = rx / ratio;
    let start_angle = ((start.1 - center.1) / ry).atan2((start.0 - center.0) / rx);
    let end_angle = ((end.1 - center.1) / ry).atan2((end.0 - center.0) / rx);
    let through_angle = ((through.1 - center.1) / ry).atan2((through.0 - center.0) / rx);
    let sweep = sweep_through(start_angle, end_angle, through_angle);
    let through_sweep = if angle_distance(through_angle, start_angle) <= f64::EPSILON {
        0.0
    } else if sweep >= 0.0 {
        (through_angle - start_angle).rem_euclid(std::f64::consts::TAU)
    } else {
        (through_angle - start_angle).rem_euclid(std::f64::consts::TAU) - std::f64::consts::TAU
    };
    let center = (
        center.0 * angle.cos() - center.1 * angle.sin(),
        center.0 * angle.sin() + center.1 * angle.cos(),
    );
    cubic_arc_segment(
        out,
        center,
        (rx, ry),
        angle,
        start_angle,
        through_sweep,
        row_type,
    ) && cubic_arc_segment(
        out,
        center,
        (rx, ry),
        angle,
        start_angle + through_sweep,
        sweep - through_sweep,
        row_type,
    )
}

fn cubic_ellipse(
    out: &mut RealizedGeometry,
    center: (f64, f64),
    axis_x: f64,
    axis_y: f64,
    other_axis_x: f64,
    other_axis_y: f64,
    row_type: &str,
) -> bool {
    let axis = (axis_x - center.0, axis_y - center.1);
    let other_axis = (other_axis_x - center.0, other_axis_y - center.1);
    let start = (center.0 + axis.0, center.1 + axis.1);
    if !push_checked(
        out,
        GeometryPathCommand::Move {
            x: start.0,
            y: start.1,
        },
        row_type,
    ) {
        return false;
    }
    for quarter in 0..4 {
        if !cubic_axis_arc_segment(
            out,
            center,
            axis,
            other_axis,
            quarter as f64 * std::f64::consts::FRAC_PI_2,
            std::f64::consts::FRAC_PI_2,
            row_type,
        ) {
            return false;
        }
    }
    true
}

fn angle_distance(left: f64, right: f64) -> f64 {
    ((left - right + std::f64::consts::PI).rem_euclid(std::f64::consts::TAU) - std::f64::consts::PI)
        .abs()
}

fn sweep_through(start: f64, end: f64, through: f64) -> f64 {
    let positive = (end - start).rem_euclid(std::f64::consts::TAU);
    let to_through = (through - start).rem_euclid(std::f64::consts::TAU);
    if to_through <= positive {
        positive
    } else {
        positive - std::f64::consts::TAU
    }
}

fn cubic_axis_arc_segment(
    out: &mut RealizedGeometry,
    center: (f64, f64),
    axis: (f64, f64),
    other_axis: (f64, f64),
    start: f64,
    sweep: f64,
    row_type: &str,
) -> bool {
    let k = 4.0 / 3.0 * (sweep / 4.0).tan();
    let point = |t: f64| {
        (
            center.0 + axis.0 * t.cos() + other_axis.0 * t.sin(),
            center.1 + axis.1 * t.cos() + other_axis.1 * t.sin(),
        )
    };
    let tangent = |t: f64| {
        (
            -axis.0 * t.sin() + other_axis.0 * t.cos(),
            -axis.1 * t.sin() + other_axis.1 * t.cos(),
        )
    };
    let p0 = point(start);
    let p1 = point(start + sweep);
    let d0 = tangent(start);
    let d1 = tangent(start + sweep);
    push_checked(
        out,
        GeometryPathCommand::Cubic {
            cp1x: p0.0 + k * d0.0,
            cp1y: p0.1 + k * d0.1,
            cp2x: p1.0 - k * d1.0,
            cp2y: p1.1 - k * d1.1,
            x: p1.0,
            y: p1.1,
        },
        row_type,
    )
}

fn cubic_arc_segment(
    out: &mut RealizedGeometry,
    center: (f64, f64),
    radii: (f64, f64),
    rotation: f64,
    start: f64,
    sweep: f64,
    row_type: &str,
) -> bool {
    let (rx, ry) = radii;
    let count = (sweep.abs() / std::f64::consts::FRAC_PI_2).ceil().max(1.0) as usize;
    let step = sweep / count as f64;
    for index in 0..count {
        let t0 = start + step * index as f64;
        let t1 = t0 + step;
        let k = 4.0 / 3.0 * (step / 4.0).tan();
        let point = |t: f64| {
            (
                center.0 + rx * t.cos() * rotation.cos() - ry * t.sin() * rotation.sin(),
                center.1 + rx * t.cos() * rotation.sin() + ry * t.sin() * rotation.cos(),
            )
        };
        let tangent = |t: f64| {
            (
                -rx * t.sin() * rotation.cos() - ry * t.cos() * rotation.sin(),
                -rx * t.sin() * rotation.sin() + ry * t.cos() * rotation.cos(),
            )
        };
        let p0 = point(t0);
        let p1 = point(t1);
        let d0 = tangent(t0);
        let d1 = tangent(t1);
        if !push_checked(
            out,
            GeometryPathCommand::Cubic {
                cp1x: p0.0 + k * d0.0,
                cp1y: p0.1 + k * d0.1,
                cp2x: p1.0 - k * d1.0,
                cp2y: p1.1 - k * d1.1,
                x: p1.0,
                y: p1.1,
            },
            row_type,
        ) {
            return false;
        }
    }
    true
}

#[cfg(test)]
mod tests {
    use crate::*;
    use ooxml_drawingml::GeometryPathCommand;
    use std::collections::BTreeMap;
    use vsdx_parse::Cell;

    fn cell(name: &str, value: &str) -> Cell {
        Cell {
            name: name.into(),
            formula: None,
            value: Some(value.into()),
            unit: None,
            del: false,
            other_attrs: vec![],
        }
    }
    fn resolved_row(ty: &str, cells: Vec<Cell>) -> ResolvedRow {
        ResolvedRow {
            key: "IX:0".into(),
            deleted: false,
            row_type: Some(ty.into()),
            cells: cells
                .into_iter()
                .map(|cell| {
                    (
                        cell.name.clone(),
                        Lookup::Found(ResolvedCell {
                            cell,
                            provenance: Provenance::Local,
                        }),
                    )
                })
                .collect(),
        }
    }

    #[test]
    fn geometry_uses_cached_values_and_reports_unsupported_rows() {
        let section = ResolvedSection {
            name: "Geometry".into(),
            deleted: false,

            rows: BTreeMap::from([
                (
                    "IX:0".into(),
                    resolved_row("MoveTo", vec![cell("X", "1"), cell("Y", "2")]),
                ),
                ("IX:1".into(), resolved_row("NURBSTo", vec![])),
            ]),
        };
        let geometry = realize_geometry(&section);
        assert!(matches!(
            geometry.commands[0],
            GeometryPathCommand::Move { x: 1.0, y: 2.0 }
        ));
        assert_eq!(
            geometry.issues,
            vec![GeometryIssue::UnsupportedRowType("NURBSTo".into())]
        );
    }

    #[test]
    fn geometry_rejects_non_finite_cached_values() {
        for value in ["NaN", "inf", "1e999"] {
            let section = ResolvedSection {
                name: "Geometry".into(),
                deleted: false,
                rows: BTreeMap::from([(
                    "IX:0".into(),
                    resolved_row("MoveTo", vec![cell("X", value), cell("Y", "2")]),
                )]),
            };
            let geometry = realize_geometry(&section);
            assert!(geometry.commands.is_empty(), "{value}");
            assert_eq!(
                geometry.issues,
                vec![GeometryIssue::UnevaluatedCell {
                    row_type: "MoveTo".into(),
                    cell: "X".into(),
                }],
                "{value}"
            );
        }
    }

    #[test]
    fn geometry_rejects_non_finite_realized_relative_coordinates() {
        let section = ResolvedSection {
            name: "Geometry".into(),
            deleted: false,
            rows: BTreeMap::from([
                (
                    "IX:0".into(),
                    resolved_row("MoveTo", vec![cell("X", "1e308"), cell("Y", "0")]),
                ),
                (
                    "IX:1".into(),
                    resolved_row("RelLineTo", vec![cell("X", "1e308"), cell("Y", "0")]),
                ),
            ]),
        };
        let geometry = realize_geometry(&section);
        assert_eq!(
            geometry.commands,
            vec![GeometryPathCommand::Move { x: 1e308, y: 0.0 }]
        );
        assert_eq!(
            geometry.issues,
            vec![GeometryIssue::UnevaluatedCell {
                row_type: "RelLineTo".into(),
                cell: "geometry".into(),
            }]
        );
    }

    #[test]
    fn arc_to_rejects_non_finite_derived_geometry() {
        let section = ResolvedSection {
            name: "Geometry".into(),
            deleted: false,
            rows: BTreeMap::from([(
                "IX:0".into(),
                resolved_row(
                    "ArcTo",
                    vec![cell("X", "1e308"), cell("Y", "0"), cell("A", "1")],
                ),
            )]),
        };
        let geometry = realize_geometry(&section);
        assert!(geometry.commands.is_empty());
        assert_eq!(
            geometry.issues,
            vec![GeometryIssue::UnevaluatedCell {
                row_type: "ArcTo".into(),
                cell: "geometry".into(),
            }]
        );
    }

    #[test]
    fn ellipse_rejects_non_finite_derived_geometry() {
        let section = ResolvedSection {
            name: "Geometry".into(),
            deleted: false,
            rows: BTreeMap::from([(
                "IX:0".into(),
                resolved_row(
                    "Ellipse",
                    vec![
                        cell("X", "-7e307"),
                        cell("Y", "0"),
                        cell("A", "0"),
                        cell("B", "0"),
                        cell("C", "1e308"),
                        cell("D", "0"),
                    ],
                ),
            )]),
        };
        let geometry = realize_geometry(&section);
        assert!(geometry.commands.is_empty());
        assert_eq!(
            geometry.issues,
            vec![GeometryIssue::UnevaluatedCell {
                row_type: "Ellipse".into(),
                cell: "geometry".into(),
            }]
        );
    }

    #[test]
    fn geometry_emits_finite_commands_unchanged() {
        let section = ResolvedSection {
            name: "Geometry".into(),
            deleted: false,
            rows: BTreeMap::from([
                (
                    "IX:0".into(),
                    resolved_row("MoveTo", vec![cell("X", "1"), cell("Y", "2")]),
                ),
                (
                    "IX:1".into(),
                    resolved_row("RelLineTo", vec![cell("X", "3"), cell("Y", "4")]),
                ),
            ]),
        };
        let geometry = realize_geometry(&section);
        assert_eq!(
            geometry.commands,
            vec![
                GeometryPathCommand::Move { x: 1.0, y: 2.0 },
                GeometryPathCommand::Line { x: 4.0, y: 6.0 },
            ]
        );
        assert!(geometry.issues.is_empty());
    }

    #[test]
    fn arc_to_bows_by_its_height_at_the_curve_midpoint() {
        let section = ResolvedSection {
            name: "Geometry".into(),
            deleted: false,
            rows: BTreeMap::from([
                (
                    "IX:0".into(),
                    resolved_row("MoveTo", vec![cell("X", "0"), cell("Y", "0")]),
                ),
                (
                    "IX:1".into(),
                    resolved_row(
                        "ArcTo",
                        vec![cell("X", "2"), cell("Y", "0"), cell("A", "0.5")],
                    ),
                ),
            ]),
        };
        let geometry = realize_geometry(&section);
        assert!(geometry.commands.iter().any(|command| matches!(
            command,
            GeometryPathCommand::Cubic { x, y, .. }
                if (x - 1.0).abs() < 1e-12 && (y - 0.5).abs() < 1e-12
        )));
    }

    #[test]
    fn ellipse_uses_center_and_axis_endpoints() {
        let section = ResolvedSection {
            name: "Geometry".into(),
            deleted: false,
            rows: BTreeMap::from([(
                "IX:0".into(),
                resolved_row(
                    "Ellipse",
                    vec![
                        cell("X", "3"),
                        cell("Y", "4"),
                        cell("A", "5"),
                        cell("B", "5"),
                        cell("C", "2"),
                        cell("D", "6"),
                    ],
                ),
            )]),
        };
        let geometry = realize_geometry(&section);
        assert!(matches!(
            geometry.commands[0],
            GeometryPathCommand::Move { x: 5.0, y: 5.0 }
        ));
        let endpoints = geometry
            .commands
            .iter()
            .filter_map(|command| match command {
                GeometryPathCommand::Move { x, y } => Some((*x, *y)),
                GeometryPathCommand::Cubic { x, y, .. } => Some((*x, *y)),
                GeometryPathCommand::Line { x, y } => Some((*x, *y)),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(endpoints.len(), 5);
        for (actual, expected) in
            endpoints
                .iter()
                .zip([(5.0, 5.0), (2.0, 6.0), (1.0, 3.0), (4.0, 2.0), (5.0, 5.0)])
        {
            assert!((actual.0 - expected.0).abs() < 1e-12);
            assert!((actual.1 - expected.1).abs() < 1e-12);
        }
    }

    #[test]
    fn elliptical_arc_requires_all_cached_schema_cells() {
        let section = ResolvedSection {
            name: "Geometry".into(),
            deleted: false,
            rows: BTreeMap::from([(
                "IX:0".into(),
                resolved_row(
                    "EllipticalArcTo",
                    vec![
                        cell("X", "1"),
                        cell("Y", "1"),
                        cell("A", "1"),
                        cell("B", "1"),
                        cell("C", "0"),
                    ],
                ),
            )]),
        };
        let geometry = realize_geometry(&section);
        assert!(geometry.commands.is_empty());
        assert_eq!(
            geometry.issues,
            vec![GeometryIssue::MissingCell {
                row_type: "EllipticalArcTo".into(),
                cell: "D".into()
            }]
        );
    }

    #[test]
    fn elliptical_arc_passes_through_its_control_point() {
        let section = ResolvedSection {
            name: "Geometry".into(),
            deleted: false,
            rows: BTreeMap::from([
                (
                    "IX:0".into(),
                    resolved_row("MoveTo", vec![cell("X", "2"), cell("Y", "0")]),
                ),
                (
                    "IX:1".into(),
                    resolved_row(
                        "EllipticalArcTo",
                        vec![
                            cell("X", "0"),
                            cell("Y", "1"),
                            cell("A", "1.4142135623730951"),
                            cell("B", "0.7071067811865476"),
                            cell("C", "0"),
                            cell("D", "2"),
                        ],
                    ),
                ),
            ]),
        };
        let geometry = realize_geometry(&section);
        assert!(geometry.commands.iter().any(|command| matches!(
            command,
            GeometryPathCommand::Cubic { x, y, .. }
                if (x - std::f64::consts::SQRT_2).abs() < 1e-12
                    && (y - std::f64::consts::FRAC_1_SQRT_2).abs() < 1e-12
        )));
    }

    #[test]
    fn geometry_realizes_two_digit_rows_in_numeric_order() {
        let keyed = |key: &str, ty: &str, cells: Vec<Cell>| {
            (
                key.to_owned(),
                ResolvedRow {
                    key: key.into(),
                    ..resolved_row(ty, cells)
                },
            )
        };
        let section = ResolvedSection {
            name: "Geometry".into(),
            deleted: false,
            rows: BTreeMap::from([
                keyed("IX:1", "MoveTo", vec![cell("X", "0"), cell("Y", "0")]),
                keyed("IX:2", "RelLineTo", vec![cell("X", "1"), cell("Y", "0")]),
                keyed("IX:10", "RelLineTo", vec![cell("X", "0"), cell("Y", "1")]),
            ]),
        };
        assert_eq!(
            realize_geometry(&section).commands,
            vec![
                GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                GeometryPathCommand::Line { x: 1.0, y: 0.0 },
                GeometryPathCommand::Line { x: 1.0, y: 1.0 },
            ]
        );
    }
}
