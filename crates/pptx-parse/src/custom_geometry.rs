use std::collections::BTreeMap;
use std::f64::consts::{FRAC_PI_2, TAU};

use ooxml_drawingml::GeometryPathCommand;

use crate::CustomGeometryPath;
use crate::drawing::{GuideValue, evaluate_guide_formula, guide_operand, standard_guide_values};
use crate::xml::{ParseBudget, XmlElement};

const MAX_CUSTOM_PATH_COMMANDS: usize = 16_384;
const MAX_ARC_SEGMENTS: usize = 64;

pub(crate) fn parse_custom_geometry(
    custom: &XmlElement,
    extent: Option<(f64, f64)>,
    budget: &mut ParseBudget<'_>,
) -> Option<Vec<CustomGeometryPath>> {
    let definitions = guide_definitions(custom);
    let path_list = custom.child("pathLst")?;
    let guide_extent = extent.or_else(|| {
        path_list
            .child_elements()
            .filter(|path| path.local_name() == "path")
            .find_map(|path| {
                let width = dimension(path, "w", None)?;
                let height = dimension(path, "h", None)?;
                Some((width, height))
            })
    });
    let guides = match guide_extent {
        Some((width, height)) => guide_values(&definitions, width, height),
        None => BTreeMap::new(),
    };
    let mut paths = Vec::new();
    let available = MAX_CUSTOM_PATH_COMMANDS.min(budget.geometry_commands_left());
    let mut remaining = available;
    for path in path_list.child_elements() {
        if path.local_name() != "path" {
            continue;
        }
        let (Some(width), Some(height)) = (
            dimension(path, "w", extent.map(|size| size.0)),
            dimension(path, "h", extent.map(|size| size.1)),
        ) else {
            continue;
        };
        let mut budget_exceeded = false;
        if let Some(parsed) = parse_custom_path(
            path,
            width,
            height,
            &guides,
            &mut remaining,
            &mut budget_exceeded,
        ) && !parsed.commands.is_empty()
        {
            paths.push(parsed);
        }
        if budget_exceeded && available == budget.geometry_commands_left() {
            budget.spend_geometry_commands(available);
            return None;
        }
    }
    budget.spend_geometry_commands(available - remaining);
    (!paths.is_empty()).then_some(paths)
}

fn guide_definitions(custom: &XmlElement) -> Vec<(&str, &str)> {
    ["avLst", "gdLst"]
        .into_iter()
        .filter_map(|list| custom.child(list))
        .flat_map(|list| list.child_elements())
        .filter(|guide| guide.local_name() == "gd")
        .filter_map(|guide| Some((guide.attribute("name")?, guide.attribute("fmla")?)))
        .collect()
}

fn guide_values(
    definitions: &[(&str, &str)],
    width: f64,
    height: f64,
) -> BTreeMap<String, GuideValue> {
    let mut values = standard_guide_values((width, height));
    for (name, formula) in definitions {
        if let Some(value) = evaluate_guide_formula(formula, &values) {
            values.insert((*name).to_owned(), value);
        }
    }
    values
}

fn parse_custom_path(
    path: &XmlElement,
    width: f64,
    height: f64,
    guides: &BTreeMap<String, GuideValue>,
    remaining: &mut usize,
    budget_exceeded: &mut bool,
) -> Option<CustomGeometryPath> {
    let mut commands = Vec::new();
    let mut current = None;
    let mut start = None;
    for command in path.child_elements() {
        let Some(left) = remaining.checked_sub(1) else {
            *budget_exceeded = true;
            return None;
        };
        *remaining = left;
        let points: Vec<_> = command
            .child_elements()
            .take(4)
            .map(|point| {
                (point.local_name() == "pt").then_some(())?;
                normalized_point(
                    coordinate(point, "x", guides)?,
                    coordinate(point, "y", guides)?,
                    width,
                    height,
                )
            })
            .collect::<Option<_>>()?;
        let next = match (command.local_name(), points.as_slice()) {
            ("moveTo", &[(x, y)]) => {
                start = Some((x, y));
                GeometryPathCommand::Move { x, y }
            }
            ("lnTo", &[(x, y)]) if current.is_some() => GeometryPathCommand::Line { x, y },
            ("quadBezTo", &[(cpx, cpy), (x, y)]) if current.is_some() => {
                GeometryPathCommand::Quad { cpx, cpy, x, y }
            }
            ("cubicBezTo", &[(cp1x, cp1y), (cp2x, cp2y), (x, y)]) if current.is_some() => {
                GeometryPathCommand::Cubic {
                    cp1x,
                    cp1y,
                    cp2x,
                    cp2y,
                    x,
                    y,
                }
            }
            ("arcTo", []) => {
                let arc = arc_to_cubics(
                    command,
                    current?,
                    width,
                    height,
                    *remaining + 1,
                    guides,
                    budget_exceeded,
                )?;
                *remaining = remaining.checked_sub(arc.len().saturating_sub(1))?;
                if let Some(GeometryPathCommand::Cubic { x, y, .. }) = arc.last() {
                    current = Some((*x, *y));
                }
                commands.extend(arc);
                continue;
            }
            ("close", []) => {
                current = Some(start?);
                commands.push(GeometryPathCommand::Close);
                continue;
            }
            _ => return None,
        };
        current = match next {
            GeometryPathCommand::Move { x, y }
            | GeometryPathCommand::Line { x, y }
            | GeometryPathCommand::Quad { x, y, .. }
            | GeometryPathCommand::Cubic { x, y, .. } => Some((x, y)),
            GeometryPathCommand::Close => unreachable!(),
        };
        commands.push(next);
    }
    Some(CustomGeometryPath {
        commands,
        no_fill: path.attribute("fill") == Some("none"),
        no_stroke: matches!(path.attribute("stroke"), Some("0" | "false" | "off")),
    })
}

fn coordinate(
    element: &XmlElement,
    name: &str,
    guides: &BTreeMap<String, GuideValue>,
) -> Option<f64> {
    let value = guide_operand(element.attribute(name)?, guides)?.value;
    value.is_finite().then_some(value)
}

fn dimension(path: &XmlElement, name: &str, fallback: Option<f64>) -> Option<f64> {
    if let Some(raw) = path.attribute(name) {
        let value = raw.parse::<f64>().ok()?;
        if !value.is_finite() || value < 0.0 {
            return None;
        }
        if value > 0.0 {
            return Some(value);
        }
    }
    match fallback {
        Some(value) if value.is_finite() && value > 0.0 => Some(value),
        Some(value) if value.is_finite() && value == 0.0 => Some(1.0),
        _ => None,
    }
}

fn normalized_point(x: f64, y: f64, width: f64, height: f64) -> Option<(f64, f64)> {
    let point = (x / width, y / height);
    (point.0.is_finite() && point.1.is_finite()).then_some(point)
}

fn ellipse_angle(angle: f64, rx: f64, ry: f64) -> f64 {
    let turns = (angle / TAU).floor();
    let angle = angle.rem_euclid(TAU);
    (rx * angle.sin()).atan2(ry * angle.cos()).rem_euclid(TAU) + turns * TAU
}

fn arc_to_cubics(
    arc: &XmlElement,
    current: (f64, f64),
    width: f64,
    height: f64,
    remaining: usize,
    guides: &BTreeMap<String, GuideValue>,
    budget_exceeded: &mut bool,
) -> Option<Vec<GeometryPathCommand>> {
    let rx = coordinate(arc, "wR", guides)?;
    let ry = coordinate(arc, "hR", guides)?;
    let start =
        (coordinate(arc, "stAng", guides)?.rem_euclid(21_600_000.0) / 60_000.0).to_radians();
    let sweep = (coordinate(arc, "swAng", guides)? / 60_000.0).to_radians();
    let remaining = remaining.min(MAX_ARC_SEGMENTS);
    if rx <= 0.0 || ry <= 0.0 || sweep.abs() > MAX_ARC_SEGMENTS as f64 * FRAC_PI_2 {
        return None;
    }
    if sweep.abs() > remaining as f64 * FRAC_PI_2 {
        *budget_exceeded = true;
        return None;
    }
    if sweep == 0.0 {
        return Some(Vec::new());
    }
    let end = ellipse_angle(start + sweep, rx, ry);
    let start = ellipse_angle(start, rx, ry);
    let sweep = end - start;
    let count = (sweep.abs() / FRAC_PI_2).ceil() as usize;
    if count == 0 || count > MAX_ARC_SEGMENTS {
        return None;
    }
    if count > remaining {
        *budget_exceeded = true;
        return None;
    }
    let (rx, ry) = normalized_point(rx, ry, width, height)?;
    let center = (current.0 - rx * start.cos(), current.1 - ry * start.sin());
    let step = sweep / count as f64;
    let alpha = 4.0 / 3.0 * (step / 4.0).tan();
    let mut commands = Vec::with_capacity(count);
    let mut from = current;
    for index in 0..count {
        let a = start + step * index as f64;
        let b = a + step;
        let to = (center.0 + rx * b.cos(), center.1 + ry * b.sin());
        let cp1 = (from.0 - alpha * rx * a.sin(), from.1 + alpha * ry * a.cos());
        let cp2 = (to.0 + alpha * rx * b.sin(), to.1 - alpha * ry * b.cos());
        if [cp1.0, cp1.1, cp2.0, cp2.1, to.0, to.1]
            .iter()
            .any(|v| !v.is_finite())
        {
            return None;
        }
        commands.push(GeometryPathCommand::Cubic {
            cp1x: cp1.0,
            cp1y: cp1.1,
            cp2x: cp2.0,
            cp2y: cp2.1,
            x: to.0,
            y: to.1,
        });
        from = to;
    }
    Some(commands)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ParseLimits;
    use crate::xml::{ParseBudget, parse_xml};

    fn parse(paths: &str) -> Option<Vec<CustomGeometryPath>> {
        parse_with(paths, "", Some((400.0, 200.0)))
    }

    fn parse_with(
        paths: &str,
        guides: &str,
        extent: Option<(f64, f64)>,
    ) -> Option<Vec<CustomGeometryPath>> {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        parse_in(paths, guides, extent, &mut budget)
    }

    fn parse_in(
        paths: &str,
        guides: &str,
        extent: Option<(f64, f64)>,
        budget: &mut ParseBudget<'_>,
    ) -> Option<Vec<CustomGeometryPath>> {
        let xml = parse_xml(
            format!("<a:custGeom>{guides}<a:pathLst>{paths}</a:pathLst></a:custGeom>").as_bytes(),
            "custom.xml",
            budget,
        )
        .unwrap();
        parse_custom_geometry(&xml, extent, budget)
    }

    fn assert_cubic(command: &GeometryPathCommand, expected: [f64; 6]) {
        let GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } = command
        else {
            panic!("expected cubic, got {command:?}");
        };
        for (actual, expected) in [cp1x, cp1y, cp2x, cp2y, x, y].into_iter().zip(expected) {
            assert!((actual - expected).abs() < 1e-12, "{actual} != {expected}");
        }
    }

    #[test]
    fn arcs_preserve_ellipse_angles_direction_and_current_point() {
        let paths = parse(r#"<a:path w="200" h="100"><a:moveTo><a:pt x="180" y="20"/></a:moveTo><a:arcTo wR="80" hR="30" stAng="0" swAng="5400000"/><a:arcTo wR="80" hR="30" stAng="5400000" swAng="-5400000"/><a:lnTo><a:pt x="20" y="30"/></a:lnTo><a:close/><a:arcTo wR="80" hR="30" stAng="0" swAng="5400000"/></a:path>"#).unwrap();
        let commands = &paths[0].commands;
        assert_eq!(commands.len(), 6);
        let k = 0.552_284_749_830_793_6;
        assert_cubic(
            &commands[1],
            [0.9, 0.2 + 0.3 * k, 0.5 + 0.4 * k, 0.5, 0.5, 0.5],
        );
        assert_cubic(
            &commands[2],
            [0.5 + 0.4 * k, 0.5, 0.9, 0.2 + 0.3 * k, 0.9, 0.2],
        );
        assert_cubic(
            &commands[5],
            [0.9, 0.2 + 0.3 * k, 0.5 + 0.4 * k, 0.5, 0.5, 0.5],
        );

        let diagonal = parse(r#"<a:path w="200" h="100"><a:moveTo><a:pt x="130" y="90"/></a:moveTo><a:arcTo wR="80" hR="40" stAng="2700000" swAng="10800000"/></a:path>"#).unwrap();
        let GeometryPathCommand::Cubic { x, y, .. } = diagonal[0].commands.last().unwrap() else {
            panic!()
        };
        let offset = 160.0 / 5.0_f64.sqrt();
        assert!((x - (130.0 - offset) / 200.0).abs() < 1e-12);
        assert!((y - (90.0 - offset) / 100.0).abs() < 1e-12);
        for sweep in [21_600_000, -21_600_000] {
            let full = parse(&format!(r#"<a:path w="200" h="100"><a:moveTo><a:pt x="180" y="20"/></a:moveTo><a:arcTo wR="80" hR="30" stAng="0" swAng="{sweep}"/></a:path>"#)).unwrap();
            assert_eq!(full[0].commands.len(), 5);
            let GeometryPathCommand::Cubic { x, y, .. } = full[0].commands[4] else {
                panic!()
            };
            assert!((x - 0.9).abs() < 1e-12 && (y - 0.2).abs() < 1e-12);
        }
    }

    #[test]
    fn paths_use_independent_spaces_and_paint_flags() {
        let paths = parse(r#"<a:path w="200" h="100" fill="none"><a:moveTo><a:pt x="20" y="10"/></a:moveTo><a:quadBezTo><a:pt x="40" y="30"/><a:pt x="60" y="50"/></a:quadBezTo></a:path><a:path w="100" h="50" stroke="0"><a:moveTo><a:pt x="20" y="10"/></a:moveTo><a:lnTo><a:pt x="40" y="30"/></a:lnTo></a:path><a:path><a:moveTo><a:pt x="20" y="10"/></a:moveTo></a:path>"#).unwrap();
        assert_eq!(
            paths,
            vec![
                CustomGeometryPath {
                    commands: vec![
                        GeometryPathCommand::Move { x: 0.1, y: 0.1 },
                        GeometryPathCommand::Quad {
                            cpx: 0.2,
                            cpy: 0.3,
                            x: 0.3,
                            y: 0.5
                        }
                    ],
                    no_fill: true,
                    no_stroke: false
                },
                CustomGeometryPath {
                    commands: vec![
                        GeometryPathCommand::Move { x: 0.2, y: 0.2 },
                        GeometryPathCommand::Line { x: 0.4, y: 0.6 }
                    ],
                    no_fill: false,
                    no_stroke: true
                },
                CustomGeometryPath {
                    commands: vec![GeometryPathCommand::Move { x: 0.05, y: 0.05 }],
                    no_fill: false,
                    no_stroke: false
                },
            ]
        );
    }

    #[test]
    fn traced_outlines_with_thousands_of_points_keep_every_point() {
        let points: String = (0..6_000)
            .map(|index| {
                format!(
                    "<a:lnTo><a:pt x=\"{}\" y=\"{}\"/></a:lnTo>",
                    index % 200,
                    index % 100
                )
            })
            .collect();
        let paths = parse(&format!(
            "<a:path w=\"200\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo>{points}<a:close/></a:path>"
        ))
        .expect("a long freeform keeps its path");
        assert_eq!(paths[0].commands.len(), 6_002);
    }

    #[test]
    fn a_package_stops_expanding_geometry_once_its_budget_is_spent() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let path = |points: usize| {
            let lines = "<a:lnTo><a:pt x=\"10\" y=\"10\"/></a:lnTo>".repeat(points - 1);
            format!(
                "<a:path w=\"200\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo>{lines}</a:path>"
            )
        };
        let left = budget.geometry_commands_left();
        budget.spend_geometry_commands(left - 10);
        assert!(parse_in(&path(11), "", Some((400.0, 200.0)), &mut budget).is_none());
        assert_eq!(budget.geometry_commands_left(), 0);

        let mut budget = ParseBudget::new(&limits);
        budget.spend_geometry_commands(left - 10);
        assert_eq!(
            parse_in(&path(10), "", Some((400.0, 200.0)), &mut budget).unwrap()[0]
                .commands
                .len(),
            10
        );
        assert_eq!(budget.geometry_commands_left(), 0);
    }

    #[test]
    fn package_budget_keeps_two_path_shapes_whole() {
        let path = format!(
            "<a:path w=\"200\" h=\"100\">{}</a:path>",
            "<a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo>".repeat(6)
        );
        let limits = ParseLimits::default();
        for second in [
            path.as_str(),
            r#"<a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:arcTo wR="30" hR="30" stAng="0" swAng="27000000"/></a:path>"#,
        ] {
            let paths = format!("{path}{second}");
            for left in [10, 6, 12] {
                let mut budget = ParseBudget::new(&limits);
                budget.spend_geometry_commands(budget.geometry_commands_left() - left);
                let parsed = parse_in(&paths, "", Some((400.0, 200.0)), &mut budget);
                if left < 12 {
                    assert!(parsed.is_none(), "{left} commands cannot keep both paths");
                } else {
                    let parsed = parsed.expect("both paths fit exactly");
                    assert_eq!(parsed.len(), 2);
                    assert!(parsed.iter().all(|path| path.commands.len() == 6));
                }
                assert_eq!(budget.geometry_commands_left(), 0);
            }
        }
    }

    #[test]
    fn one_arc_cannot_expand_past_its_segment_bound() {
        let path = |turns: i64| {
            format!(
                r#"<a:path w="200" h="100"><a:moveTo><a:pt x="180" y="20"/></a:moveTo><a:arcTo wR="80" hR="30" stAng="0" swAng="{}"/></a:path>"#,
                turns * 21_600_000
            )
        };
        assert_eq!(parse(&path(16)).unwrap()[0].commands.len(), 65);
        assert!(parse(&path(17)).is_none());
    }

    #[test]
    fn invalid_paths_are_skipped_and_expansion_is_bounded() {
        let valid = r#"<a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="50"/></a:lnTo></a:path>"#;
        for invalid in [
            valid.replace("h=\"100\"", "h=\"inf\""),
            valid.replace("x=\"100\"", "x=\"NaN\""),
            valid.replace("x=\"100\"", "x=\"guide\""),
            valid.replace("w=\"200\"", "w=\"1e-320\""),
            valid.replace("lnTo", "unsupported"),
            valid.replace("moveTo", "lnTo"),
            valid.replace("<a:pt x=\"100\" y=\"50\"/>", ""),
            r#"<a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:arcTo wR="80" hR="30" stAng="0" swAng="1e300"/></a:path>"#.to_owned(),
        ] {
            let kept = parse(&format!("{valid}{invalid}")).expect("the valid path is kept");
            assert_eq!(kept.len(), 1, "{invalid}");
            assert!(parse(&invalid).is_none(), "{invalid}");
        }
        assert!(parse("").is_none());
        assert!(parse("<a:path w=\"200\" h=\"100\"/>").is_none());
        let moves = "<a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo>";
        let at_limit = format!(
            "<a:path w=\"200\" h=\"100\">{}</a:path>",
            moves.repeat(MAX_CUSTOM_PATH_COMMANDS)
        );
        assert_eq!(
            parse(&at_limit).unwrap()[0].commands.len(),
            MAX_CUSTOM_PATH_COMMANDS
        );
        let capped = parse(&format!("{at_limit}{valid}")).unwrap();
        assert_eq!(capped.len(), 1);
        assert_eq!(capped[0].commands.len(), MAX_CUSTOM_PATH_COMMANDS);
        let arcs = r#"<a:arcTo wR="80" hR="30" stAng="0" swAng="21600000"/>"#;
        assert!(
            parse(&format!(
                "<a:path w=\"200\" h=\"100\">{moves}{}</a:path>",
                arcs.repeat(MAX_CUSTOM_PATH_COMMANDS / 4)
            ))
            .is_none()
        );
    }

    #[test]
    fn guide_names_resolve_in_points() {
        let guides = r#"<a:avLst><a:gd name="adj" fmla="val 25"/></a:avLst><a:gdLst><a:gd name="x1" fmla="*/ w 1 2"/><a:gd name="x2" fmla="+- x1 adj 0"/></a:gdLst>"#;
        let paths = parse_with(
            r#"<a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="x1" y="adj"/></a:lnTo><a:lnTo><a:pt x="x2" y="vc"/></a:lnTo><a:lnTo><a:pt x="w" y="h"/></a:lnTo></a:path>"#,
            guides,
            Some((400.0, 200.0)),
        )
        .unwrap();
        assert_eq!(
            paths[0].commands,
            vec![
                GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                GeometryPathCommand::Line { x: 1.0, y: 0.25 },
                GeometryPathCommand::Line { x: 1.125, y: 1.0 },
                GeometryPathCommand::Line { x: 2.0, y: 2.0 },
            ]
        );
    }

    #[test]
    fn guides_are_shared_across_differently_sized_paths() {
        let guides = r#"<a:gdLst><a:gd name="x1" fmla="*/ w 1 2"/></a:gdLst>"#;
        let paths = parse_with(
            r#"<a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="x1" y="vc"/></a:lnTo></a:path><a:path w="100" h="50"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="x1" y="vc"/></a:lnTo></a:path>"#,
            guides,
            Some((400.0, 200.0)),
        )
        .unwrap();
        assert_eq!(
            paths[0].commands,
            vec![
                GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                GeometryPathCommand::Line { x: 1.0, y: 1.0 },
            ]
        );
        assert_eq!(
            paths[1].commands,
            vec![
                GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                GeometryPathCommand::Line { x: 2.0, y: 2.0 },
            ]
        );
    }

    #[test]
    fn guide_names_resolve_in_arc_to() {
        let guides =
            r#"<a:gdLst><a:gd name="rw" fmla="val 80"/><a:gd name="rh" fmla="val 30"/></a:gdLst>"#;
        let paths = parse_with(
            r#"<a:path w="200" h="100"><a:moveTo><a:pt x="180" y="20"/></a:moveTo><a:arcTo wR="rw" hR="rh" stAng="0" swAng="cd2"/></a:path>"#,
            guides,
            Some((400.0, 200.0)),
        )
        .unwrap();
        let commands = &paths[0].commands;
        assert_eq!(commands.len(), 3);
        let GeometryPathCommand::Cubic { x, y, .. } = commands[2] else {
            panic!("expected a cubic");
        };
        assert!((x - 0.1).abs() < 1e-12 && (y - 0.2).abs() < 1e-12);
    }

    #[test]
    fn one_bad_path_keeps_its_siblings() {
        let paths = parse(
            r#"<a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo></a:path><a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="missing" y="0"/></a:lnTo></a:path><a:path w="200" h="100"><a:moveTo><a:pt x="200" y="100"/></a:moveTo></a:path>"#,
        )
        .unwrap();
        assert_eq!(
            paths
                .iter()
                .map(|path| path.commands.clone())
                .collect::<Vec<_>>(),
            vec![
                vec![GeometryPathCommand::Move { x: 0.0, y: 0.0 }],
                vec![GeometryPathCommand::Move { x: 1.0, y: 1.0 }],
            ]
        );
    }

    #[test]
    fn zero_path_dimensions_use_the_shape_extent() {
        let paths = parse(
            r#"<a:path w="0" h="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="400" y="200"/></a:lnTo></a:path>"#,
        )
        .unwrap();
        assert_eq!(
            paths[0].commands[1],
            GeometryPathCommand::Line { x: 1.0, y: 1.0 }
        );

        let vertical = parse_with(
            r#"<a:path><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="200"/></a:lnTo></a:path>"#,
            "",
            Some((0.0, 200.0)),
        )
        .unwrap();
        assert_eq!(
            vertical[0].commands[1],
            GeometryPathCommand::Line { x: 0.0, y: 1.0 }
        );

        let horizontal = parse_with(
            r#"<a:path><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="400" y="0"/></a:lnTo></a:path>"#,
            "",
            Some((400.0, 0.0)),
        )
        .unwrap();
        assert_eq!(
            horizontal[0].commands[1],
            GeometryPathCommand::Line { x: 1.0, y: 0.0 }
        );
    }
}
