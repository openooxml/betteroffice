use std::f64::consts::{FRAC_PI_2, TAU};

use ooxml_drawingml::GeometryPathCommand;

use crate::CustomGeometryPath;
use crate::xml::XmlElement;

const MAX_CUSTOM_PATH_COMMANDS: usize = 2_048;

pub(crate) fn parse_custom_geometry(
    custom: &XmlElement,
    extent: Option<(f64, f64)>,
) -> Option<Vec<CustomGeometryPath>> {
    let mut paths = Vec::new();
    let mut remaining = MAX_CUSTOM_PATH_COMMANDS;
    for path in custom.child("pathLst")?.child_elements() {
        if path.local_name() != "path" {
            continue;
        }
        let width = dimension(path, "w", extent.map(|size| size.0))?;
        let height = dimension(path, "h", extent.map(|size| size.1))?;
        let mut commands = Vec::new();
        let mut current = None;
        let mut start = None;
        for command in path.child_elements() {
            remaining = remaining.checked_sub(1)?;
            let points: Vec<_> = command
                .child_elements()
                .take(4)
                .map(|point| {
                    (point.local_name() == "pt").then_some(())?;
                    normalized_point(number(point, "x")?, number(point, "y")?, width, height)
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
                    let arc = arc_to_cubics(command, current?, width, height, remaining + 1)?;
                    remaining = remaining.checked_sub(arc.len().saturating_sub(1))?;
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
        if !commands.is_empty() {
            paths.push(CustomGeometryPath {
                commands,
                no_fill: path.attribute("fill") == Some("none"),
                no_stroke: matches!(path.attribute("stroke"), Some("0" | "false" | "off")),
            });
        }
    }
    (!paths.is_empty()).then_some(paths)
}

fn number(element: &XmlElement, name: &str) -> Option<f64> {
    let value = element.attribute(name)?.parse::<f64>().ok()?;
    value.is_finite().then_some(value)
}

fn dimension(path: &XmlElement, name: &str, fallback: Option<f64>) -> Option<f64> {
    let value = if path.attribute(name).is_some() {
        number(path, name)?
    } else {
        fallback?
    };
    (value.is_finite() && value > 0.0).then_some(value)
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
) -> Option<Vec<GeometryPathCommand>> {
    let rx = number(arc, "wR")?;
    let ry = number(arc, "hR")?;
    let start = (number(arc, "stAng")?.rem_euclid(21_600_000.0) / 60_000.0).to_radians();
    let sweep = (number(arc, "swAng")? / 60_000.0).to_radians();
    if rx <= 0.0 || ry <= 0.0 || sweep.abs() > remaining as f64 * FRAC_PI_2 {
        return None;
    }
    if sweep == 0.0 {
        return Some(Vec::new());
    }
    let end = ellipse_angle(start + sweep, rx, ry);
    let start = ellipse_angle(start, rx, ry);
    let sweep = end - start;
    let count = (sweep.abs() / FRAC_PI_2).ceil() as usize;
    if count == 0 || count > remaining {
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
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let xml = parse_xml(
            format!("<a:custGeom><a:pathLst>{paths}</a:pathLst></a:custGeom>").as_bytes(),
            "custom.xml",
            &mut budget,
        )
        .unwrap();
        parse_custom_geometry(&xml, Some((400.0, 200.0)))
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
    fn invalid_paths_fall_back_as_a_whole_and_expansion_is_bounded() {
        let valid = r#"<a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="50"/></a:lnTo></a:path>"#;
        for invalid in [
            valid.replace("w=\"200\"", "w=\"0\""),
            valid.replace("h=\"100\"", "h=\"inf\""),
            valid.replace("x=\"100\"", "x=\"NaN\""),
            valid.replace("x=\"100\"", "x=\"guide\""),
            valid.replace("w=\"200\"", "w=\"1e-320\""),
            valid.replace("lnTo", "unsupported"),
            valid.replace("moveTo", "lnTo"),
            valid.replace("<a:pt x=\"100\" y=\"50\"/>", ""),
            r#"<a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:arcTo wR="80" hR="30" stAng="0" swAng="1e300"/></a:path>"#.to_owned(),
        ] {
            assert!(parse(&format!("{valid}{invalid}")).is_none(), "{invalid}");
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
        assert!(parse(&format!("{at_limit}{valid}")).is_none());
        let arcs = r#"<a:arcTo wR="80" hR="30" stAng="0" swAng="21600000"/>"#;
        assert!(
            parse(&format!(
                "<a:path w=\"200\" h=\"100\">{moves}{}</a:path>",
                arcs.repeat(MAX_CUSTOM_PATH_COMMANDS / 4)
            ))
            .is_none()
        );
    }
}
