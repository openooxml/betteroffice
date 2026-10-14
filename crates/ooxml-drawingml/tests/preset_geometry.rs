use std::collections::HashMap;

use ooxml_drawingml::{
    GeometryPathCommand, preset_geometry_default_adjustments, preset_geometry_layers,
    preset_geometry_to_path,
};
use proptest::prelude::*;
use proptest::sample::select;

/// Absolute slack on unit-frame coordinates; rounding here is around 1e-15.
const TOLERANCE: f64 = 1e-9;

/// Every preset name `preset_geometry_to_path` draws.
const PRESETS: &[&str] = &[
    "rect",
    "roundRect",
    "ellipse",
    "line",
    "straightConnector1",
    "triangle",
    "isosTriangle",
    "rtTriangle",
    "diamond",
    "flowChartDecision",
    "parallelogram",
    "trapezoid",
    "pentagon",
    "flowChartOffpageConnector",
    "hexagon",
    "heptagon",
    "octagon",
    "decagon",
    "dodecagon",
    "star4",
    "star5",
    "star6",
    "star7",
    "star8",
    "star10",
    "star12",
    "star16",
    "star24",
    "star32",
    "bentConnector2",
    "bentConnector3",
    "bentConnector4",
    "bentConnector5",
    "curvedConnector2",
    "curvedConnector3",
    "curvedConnector4",
    "curvedConnector5",
    "rightArrow",
    "leftArrow",
    "upArrow",
    "downArrow",
    "leftRightArrow",
    "upDownArrow",
    "chevron",
    "homePlate",
    "plus",
    "flowChartProcess",
    "flowChartAlternateProcess",
    "flowChartPredefinedProcess",
    "flowChartInternalStorage",
    "flowChartPreparation",
    "flowChartManualOperation",
    "flowChartMagneticTape",
    "flowChartMagneticDisk",
    "flowChartMagneticDrum",
    "flowChartDisplay",
    "textBox",
    "flowChartConnector",
    "flowChartInputOutput",
    "flowChartManualInput",
    "flowChartTerminator",
    "donut",
    "noSmoking",
    "corner",
    "foldedCorner",
    "mathMultiply",
    "bentArrow",
    "ribbon",
    "ellipseRibbon",
    "cloudCallout",
    "wedgeEllipseCallout",
    "wedgeRoundRectCallout",
    "wedgeRectCallout",
    "swooshArrow",
    "circularArrow",
];

/// Presets added here whose ECMA-376 definition curves rather than only turning corners.
const NEW_PRESETS: &[&str] = &[
    "donut",
    "noSmoking",
    "corner",
    "foldedCorner",
    "mathMultiply",
    "bentArrow",
    "ribbon",
    "ellipseRibbon",
    "cloudCallout",
    "wedgeEllipseCallout",
    "wedgeRoundRectCallout",
];

/// Callouts whose ECMA-376 tail target is unpinned, so it points outside the frame by design.
const CALLOUTS: &[&str] = &[
    "cloudCallout",
    "wedgeRectCallout",
    "wedgeEllipseCallout",
    "wedgeRoundRectCallout",
];

/// Presets whose ECMA-376 curve stays inside the frame while a control point of it does not.
const HULL_OUTSIDE_FRAME: &[&str] = &["ellipseRibbon", "noSmoking"];

/// Points sampled along each curve when a test needs the outline rather than its hull.
const CURVE_SAMPLES: usize = 24;

/// These presets can exceed their frames at valid adjustments or aspect ratios.
const SPEC_OVERRUNS_FRAME: &[&str] = &["mathMultiply", "swooshArrow"];

/// `cloudCallout`'s body overruns the 43200 frame its own path declares by about 1%.
const BODY_SLACK: f64 = 0.02;

/// Connectors whose ECMA-376 adjusts are unpinned, so they may route outside the frame.
const UNPINNED_CONNECTORS: &[&str] = &["bentConnector", "curvedConnector"];

/// Stars ECMA-376 scales by `hf`/`vf`; their outline leaves the frame near the top of `adj`.
const FRAME_SCALED_STARS: &[&str] = &["star5", "star6", "star7", "star10"];

/// Adjusts whose feature must move strictly with the value across the ECMA-376 range.
const KNOBS: &[(&str, &str)] = &[
    ("chevron", "adj"),
    ("homePlate", "adj"),
    ("rightArrow", "adj2"),
    ("leftArrow", "adj2"),
    ("upArrow", "adj2"),
    ("downArrow", "adj2"),
    ("rightArrow", "adj1"),
    ("leftArrow", "adj1"),
    ("upArrow", "adj1"),
    ("downArrow", "adj1"),
    ("roundRect", "adj"),
    ("triangle", "adj"),
    ("octagon", "adj"),
    ("star5", "adj"),
    ("star8", "adj"),
    ("star32", "adj"),
];

/// Width over the shortest side, `w / ss`, for a width-over-height aspect.
fn width_in_shortest_sides(aspect: f64) -> f64 {
    aspect.max(1.0)
}

/// Height over the shortest side, `h / ss`, for a width-over-height aspect.
fn height_in_shortest_sides(aspect: f64) -> f64 {
    (1.0 / aspect).max(1.0)
}

/// The ECMA-376 upper pin of one adjust, in guide units over 100000.
fn spec_max_adjust(shape: &str, adjust: &str, aspect: f64) -> f64 {
    match (shape, adjust) {
        ("chevron" | "homePlate" | "rightArrow" | "leftArrow", "adj" | "adj2") => {
            width_in_shortest_sides(aspect)
        }
        ("upArrow" | "downArrow", "adj2") => height_in_shortest_sides(aspect),
        ("roundRect" | "octagon", _) => 0.5,
        (star, _) if star.starts_with("star") => 0.5,
        _ => 1.0,
    }
}

/// Aspect ratios from 1:64 to 64:1, where shortest-side mistakes are largest.
fn aspect() -> impl Strategy<Value = f64> {
    (-6.0f64..6.0).prop_map(f64::exp2)
}

/// Any `f64` aspect, including the zero, infinite and NaN a degenerate extent yields.
fn any_aspect() -> BoxedStrategy<f64> {
    prop_oneof![
        4 => aspect(),
        2 => any::<f64>(),
        1 => Just(0.0),
        1 => Just(f64::INFINITY),
        1 => Just(f64::NAN),
        1 => Just(f64::MIN_POSITIVE),
    ]
    .boxed()
}

fn any_adjust() -> impl Strategy<Value = f64> {
    prop_oneof![3 => -1.0f64..8.0, 1 => any::<f64>()]
}

fn adjustments() -> impl Strategy<Value = HashMap<String, f64>> {
    (
        proptest::option::of(any_adjust()),
        proptest::option::of(any_adjust()),
        proptest::option::of(any_adjust()),
        proptest::option::of(any_adjust()),
        proptest::option::of(any_adjust()),
        proptest::option::of(any_adjust()),
    )
        .prop_map(|(adj, adj1, adj2, adj3, adj4, vf)| {
            [
                ("adj", adj),
                ("adj1", adj1),
                ("adj2", adj2),
                ("adj3", adj3),
                ("adj4", adj4),
                ("vf", vf),
            ]
            .into_iter()
            .filter_map(|(name, value)| Some((name.to_owned(), value?)))
            .collect()
        })
}

fn named(pairs: &[(&str, f64)]) -> HashMap<String, f64> {
    pairs
        .iter()
        .map(|&(name, value)| (name.to_owned(), value))
        .collect()
}

fn draw(shape: &str, adjustments: &HashMap<String, f64>, aspect: f64) -> Vec<GeometryPathCommand> {
    preset_geometry_to_path(shape, adjustments, aspect)
        .unwrap_or_else(|| panic!("{shape} is not a preset"))
}

fn coordinates(path: &[GeometryPathCommand]) -> Vec<(f64, f64)> {
    use GeometryPathCommand as C;
    path.iter()
        .flat_map(|command| match *command {
            C::Move { x, y } | C::Line { x, y } => vec![(x, y)],
            C::Quad { cpx, cpy, x, y } => vec![(cpx, cpy), (x, y)],
            C::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => vec![(cp1x, cp1y), (cp2x, cp2y), (x, y)],
            C::Close => Vec::new(),
        })
        .collect()
}

fn vertices(path: &[GeometryPathCommand]) -> Vec<(f64, f64)> {
    path.iter()
        .filter_map(|command| match *command {
            GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => Some((x, y)),
            _ => None,
        })
        .collect()
}

fn near(a: (f64, f64), b: (f64, f64)) -> bool {
    (a.0 - b.0).abs() < TOLERANCE && (a.1 - b.1).abs() < TOLERANCE
}

/// Whether two closed polygons trace the same outline, from any start in either winding.
fn same_polygon(a: &[(f64, f64)], b: &[(f64, f64)]) -> bool {
    let n = a.len();
    n == b.len()
        && (0..n).any(|shift| {
            (0..n).all(|i| near(a[i], b[(shift + i) % n]))
                || (0..n).all(|i| near(a[i], b[(shift + n - i) % n]))
        })
}

type Rotation = fn((f64, f64)) -> (f64, f64);

/// Maps a right-pointing arrow's frame onto a turned arrow's, and gives the turned aspect.
fn turn(shape: &str, aspect: f64) -> (Rotation, f64) {
    match shape {
        "upArrow" => (|(x, y)| (y, 1.0 - x), 1.0 / aspect),
        "downArrow" => (|(x, y)| (1.0 - y, x), 1.0 / aspect),
        "leftArrow" => (|(x, y)| (1.0 - x, 1.0 - y), aspect),
        _ => unreachable!("{shape} is not a turned arrow"),
    }
}

fn distance(a: (f64, f64), b: (f64, f64)) -> f64 {
    (a.0 - b.0).hypot(a.1 - b.1)
}

/// The size of the feature an adjust controls, growing with the adjust.
fn feature(shape: &str, adjust: &str, path: &[GeometryPathCommand]) -> f64 {
    let v = vertices(path);
    match (shape, adjust) {
        ("chevron" | "homePlate", _) => 1.0 - v[1].0,
        (_, "adj2") => 1.0 - distance(v[0], v[1]),
        (_, "adj1") => distance(v[0], v[6]),
        (star, _) if star.starts_with("star") => {
            let points: usize = star["star".len()..].parse().unwrap();
            distance(v[1], v[1 + points])
        }
        _ => v[0].0,
    }
}

/// Splits a path at each `Move`, so a ring and its hole can be compared.
fn subpaths(path: &[GeometryPathCommand]) -> Vec<Vec<GeometryPathCommand>> {
    let mut out: Vec<Vec<GeometryPathCommand>> = Vec::new();
    for command in path {
        if matches!(command, GeometryPathCommand::Move { .. }) || out.is_empty() {
            out.push(Vec::new());
        }
        out.last_mut()
            .expect("a subpath is open")
            .push(command.clone());
    }
    out
}

/// The outline itself, with every curve flattened.
fn sample(path: &[GeometryPathCommand]) -> Vec<(f64, f64)> {
    use GeometryPathCommand as C;
    let mut points = Vec::new();
    let (mut cursor, mut opened) = ((0.0, 0.0), (0.0, 0.0));
    for command in path {
        match *command {
            C::Move { x, y } => {
                cursor = (x, y);
                opened = cursor;
                points.push(cursor);
            }
            C::Line { x, y } => {
                cursor = (x, y);
                points.push(cursor);
            }
            C::Quad { cpx, cpy, x, y } => {
                for step in 1..=CURVE_SAMPLES {
                    let t = step as f64 / CURVE_SAMPLES as f64;
                    let u = 1.0 - t;
                    points.push((
                        u * u * cursor.0 + 2.0 * u * t * cpx + t * t * x,
                        u * u * cursor.1 + 2.0 * u * t * cpy + t * t * y,
                    ));
                }
                cursor = (x, y);
            }
            C::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                for step in 1..=CURVE_SAMPLES {
                    let t = step as f64 / CURVE_SAMPLES as f64;
                    let u = 1.0 - t;
                    points.push((
                        u * u * u * cursor.0
                            + 3.0 * u * u * t * cp1x
                            + 3.0 * u * t * t * cp2x
                            + t * t * t * x,
                        u * u * u * cursor.1
                            + 3.0 * u * u * t * cp1y
                            + 3.0 * u * t * t * cp2y
                            + t * t * t * y,
                    ));
                }
                cursor = (x, y);
            }
            C::Close => {
                cursor = opened;
                points.push(cursor);
            }
        }
    }
    points
}

fn signed_area(points: &[(f64, f64)]) -> f64 {
    points
        .iter()
        .zip(points.iter().cycle().skip(1))
        .map(|(a, b)| a.0 * b.1 - b.0 * a.1)
        .sum::<f64>()
        / 2.0
}

fn bounds(points: &[(f64, f64)]) -> (f64, f64, f64, f64) {
    points.iter().fold(
        (f64::MAX, f64::MAX, f64::MIN, f64::MIN),
        |(l, t, r, b), &(x, y)| (l.min(x), t.min(y), r.max(x), b.max(y)),
    )
}

proptest! {
    #[test]
    fn every_preset_emits_finite_coordinates(
        adjustments in adjustments(),
        aspect in any_aspect(),
    ) {
        for shape in PRESETS {
            for (x, y) in coordinates(&draw(shape, &adjustments, aspect)) {
                prop_assert!(x.is_finite() && y.is_finite(), "{shape} emitted ({x}, {y})");
            }
        }
        for shape in PRESETS.iter().chain(["arc", "cube", "leftBrace", "rightBrace", "ribbon2"].iter()) {
            for layer in preset_geometry_layers(shape, &adjustments, aspect).into_iter().flatten() {
                for (x, y) in coordinates(&layer.commands) {
                    prop_assert!(x.is_finite() && y.is_finite(), "{shape} layer emitted ({x}, {y})");
                }
            }
        }
    }

    #[test]
    fn a_turned_arrow_is_a_turned_right_arrow(
        shape in select(&["upArrow", "downArrow", "leftArrow"][..]),
        adj1 in -0.25f64..1.25,
        adj2 in -0.25f64..4.0,
        aspect in aspect(),
    ) {
        let adjustments = named(&[("adj1", adj1), ("adj2", adj2)]);
        let (rotate, right_aspect) = turn(shape, aspect);
        let expected = vertices(&draw("rightArrow", &adjustments, right_aspect))
            .into_iter()
            .map(rotate)
            .collect::<Vec<_>>();
        let actual = vertices(&draw(shape, &adjustments, aspect));
        prop_assert!(
            same_polygon(&actual, &expected),
            "{shape}\n  drawn:    {actual:?}\n  expected: {expected:?}"
        );
    }

    #[test]
    fn a_point_is_measured_off_the_shortest_side(
        shape in select(&["rightArrow", "chevron", "homePlate"][..]),
        fraction in -0.1f64..1.25,
        adj1 in 0.0f64..=1.0,
        aspect in aspect(),
    ) {
        let width = width_in_shortest_sides(aspect);
        let adj = fraction * width;
        let adjustments = named(&[("adj", adj), ("adj1", adj1), ("adj2", adj)]);
        let v = vertices(&draw(shape, &adjustments, aspect));
        let depth = (1.0 - v[1].0) * width;
        prop_assert!(
            (depth - adj.clamp(0.0, width)).abs() < TOLERANCE,
            "{shape} point is {depth} shortest sides deep for adj {adj}"
        );
        if shape == "rightArrow" {
            let shaft = v[6].1 - v[0].1;
            prop_assert!((shaft - adj1).abs() < TOLERANCE, "shaft {shaft} for adj1 {adj1}");
        }
    }

    #[test]
    fn a_pinned_preset_stays_inside_its_frame(
        adjustments in adjustments(),
        aspect in aspect(),
    ) {
        let pinned = PRESETS
            .iter()
            .filter(|shape| !UNPINNED_CONNECTORS.iter().any(|prefix| shape.starts_with(prefix)))
            .filter(|shape| !FRAME_SCALED_STARS.contains(shape))
            .filter(|shape| !CALLOUTS.contains(shape))
            .filter(|shape| !SPEC_OVERRUNS_FRAME.contains(shape))
            .filter(|shape| !HULL_OUTSIDE_FRAME.contains(shape));
        for shape in pinned {
            for (x, y) in coordinates(&draw(shape, &adjustments, aspect)) {
                prop_assert!(
                    (-TOLERANCE..=1.0 + TOLERANCE).contains(&x)
                        && (-TOLERANCE..=1.0 + TOLERANCE).contains(&y),
                    "{shape} left its frame at ({x}, {y})"
                );
            }
        }
    }

    #[test]
    fn raising_an_adjust_grows_its_feature(
        (shape, adjust) in select(KNOBS),
        a in 0.0f64..=1.0,
        b in 0.0f64..=1.0,
        aspect in aspect(),
    ) {
        prop_assume!((a - b).abs() > 1e-6);
        let max = spec_max_adjust(shape, adjust, aspect);
        let (low, high) = (a.min(b) * max, a.max(b) * max);
        let size = |value| feature(shape, adjust, &draw(shape, &named(&[(adjust, value)]), aspect));
        let (smaller, larger) = (size(low), size(high));
        prop_assert!(
            larger > smaller,
            "{shape} {adjust}: {low} gives {smaller}, {high} gives {larger} (spec max {max})"
        );
    }
}

proptest! {
    #[test]
    fn a_pinned_outline_stays_inside_its_frame(
        adjustments in adjustments(),
        aspect in aspect(),
    ) {
        let pinned = PRESETS
            .iter()
            .filter(|shape| !UNPINNED_CONNECTORS.iter().any(|prefix| shape.starts_with(prefix)))
            .filter(|shape| !FRAME_SCALED_STARS.contains(shape))
            .filter(|shape| !CALLOUTS.contains(shape))
            .filter(|shape| !SPEC_OVERRUNS_FRAME.contains(shape));
        for shape in pinned {
            for (x, y) in sample(&draw(shape, &adjustments, aspect)) {
                prop_assert!(
                    (-TOLERANCE..=1.0 + TOLERANCE).contains(&x)
                        && (-TOLERANCE..=1.0 + TOLERANCE).contains(&y),
                    "{shape} outline left its frame at ({x}, {y})"
                );
            }
        }
    }

    #[test]
    fn a_new_preset_closes_every_subpath(
        adjustments in adjustments(),
        aspect in any_aspect(),
    ) {
        for shape in NEW_PRESETS {
            for part in subpaths(&draw(shape, &adjustments, aspect)) {
                prop_assert!(
                    matches!(part.first(), Some(GeometryPathCommand::Move { .. })),
                    "{shape} opened a subpath without a move"
                );
                prop_assert!(
                    matches!(part.last(), Some(GeometryPathCommand::Close)),
                    "{shape} left a subpath open: {part:?}"
                );
            }
        }
    }

    #[test]
    fn a_ring_hole_is_wound_against_its_ring(
        shape in select(&["donut", "noSmoking"][..]),
        adj in 0.02f64..0.45,
        aspect in aspect(),
    ) {
        let parts = subpaths(&draw(shape, &named(&[("adj", adj)]), aspect));
        prop_assert!(parts.len() >= 2, "{shape} drew no hole");
        let ring = signed_area(&sample(&parts[0]));
        for hole in &parts[1..] {
            let area = signed_area(&sample(hole));
            prop_assert!(
                ring * area < 0.0,
                "{shape} hole winds with its ring: ring {ring}, hole {area}"
            );
        }
    }

    #[test]
    fn a_donut_hole_shrinks_as_its_adjust_grows(
        a in 0.02f64..0.45,
        b in 0.02f64..0.45,
        aspect in aspect(),
    ) {
        prop_assume!((a - b).abs() > 1e-3);
        let hole = |adj: f64| {
            let parts = subpaths(&draw("donut", &named(&[("adj", adj)]), aspect));
            signed_area(&sample(&parts[1])).abs()
        };
        let (low, high) = (a.min(b), a.max(b));
        prop_assert!(
            hole(high) < hole(low),
            "donut hole grew from adj {low} to {high}: {} then {}",
            hole(low),
            hole(high)
        );
    }

    #[test]
    fn a_callout_tail_leaves_its_body(
        shape in select(CALLOUTS),
        aspect in aspect(),
    ) {
        let defaults = preset_geometry_default_adjustments(shape);
        let reach = defaults["adj2"] + 0.5;
        let (_, _, _, bottom) = bounds(&sample(&draw(shape, &defaults, aspect)));
        prop_assert!(
            bottom >= reach - TOLERANCE && bottom < reach + BODY_SLACK * 2.0,
            "{shape} tail reached {bottom}, not the default adjust's {reach}"
        );
    }

    #[test]
    fn a_callout_body_stays_inside_its_frame(
        shape in select(CALLOUTS),
        aspect in aspect(),
    ) {
        let centred = draw(shape, &named(&[("adj1", 0.0), ("adj2", 0.0)]), aspect);
        let (left, top, right, bottom) = bounds(&sample(&centred));
        prop_assert!(
            left > -BODY_SLACK && top > -BODY_SLACK
                && right < 1.0 + BODY_SLACK && bottom < 1.0 + BODY_SLACK,
            "{shape} body left its frame: {left}..{right} by {top}..{bottom}"
        );
        prop_assert!(
            right - left > 0.95 && bottom - top > 0.95,
            "{shape} body does not fill its frame: {left}..{right} by {top}..{bottom}"
        );
    }

    #[test]
    fn a_corner_notch_is_measured_off_the_shortest_side(
        a1 in 0.0f64..=1.0,
        a2 in 0.0f64..=1.0,
        aspect in aspect(),
    ) {
        let (width, height) = (width_in_shortest_sides(aspect), height_in_shortest_sides(aspect));
        let (adj1, adj2) = (a1 * height, a2 * width);
        let v = vertices(&draw("corner", &named(&[("adj1", adj1), ("adj2", adj2)]), aspect));
        prop_assert!((v[1].0 * width - adj2).abs() < TOLERANCE, "arm {:?} for adj2 {adj2}", v[1]);
        prop_assert!(
            ((1.0 - v[2].1) * height - adj1).abs() < TOLERANCE,
            "leg {:?} for adj1 {adj1}",
            v[2]
        );
    }

    #[test]
    fn math_multiply_turns_onto_itself(adj in 0.0f64..=0.519_65, aspect in aspect()) {
        let v = vertices(&draw("mathMultiply", &named(&[("adj1", adj)]), aspect));
        let turned = v.iter().map(|&(x, y)| (1.0 - x, 1.0 - y)).collect::<Vec<_>>();
        prop_assert!(same_polygon(&v, &turned), "mathMultiply is lopsided: {v:?}");
    }

    #[test]
    fn a_folded_corner_winds_its_fold_with_its_body(
        adj in 0.0f64..=0.5,
        aspect in aspect(),
    ) {
        let parts = subpaths(&draw("foldedCorner", &named(&[("adj", adj)]), aspect));
        prop_assert_eq!(parts.len(), 2, "foldedCorner lost its fold");
        let (body, fold) = (signed_area(&sample(&parts[0])), signed_area(&sample(&parts[1])));
        prop_assert!(
            body * fold >= 0.0,
            "foldedCorner fold would punch a hole: body {body}, fold {fold}"
        );
    }

    #[test]
    fn a_bent_arrow_reaches_its_frame_edge(
        adjustments in adjustments(),
        aspect in aspect(),
    ) {
        let v = vertices(&draw("bentArrow", &adjustments, aspect));
        let (left, top, right, bottom) = bounds(&v);
        prop_assert!((left).abs() < TOLERANCE && (right - 1.0).abs() < TOLERANCE, "x span {left}..{right}");
        prop_assert!((top).abs() < TOLERANCE && (bottom - 1.0).abs() < TOLERANCE, "y span {top}..{bottom}");
    }

    #[test]
    fn a_curved_preset_is_not_a_polygon(
        adjustments in adjustments(),
        aspect in aspect(),
    ) {
        let curved = ["donut", "noSmoking", "ribbon", "ellipseRibbon", "cloudCallout",
                      "wedgeEllipseCallout", "wedgeRoundRectCallout", "bentArrow"];
        for shape in curved {
            let path = draw(shape, &adjustments, aspect);
            let curves = path
                .iter()
                .filter(|command| {
                    matches!(
                        command,
                        GeometryPathCommand::Quad { .. } | GeometryPathCommand::Cubic { .. }
                    )
                })
                .count();
            prop_assert!(curves > 0, "{shape} drew only straight edges");
        }
    }
}

#[test]
fn single_path_consumers_do_not_receive_partial_layered_presets() {
    for shape in ["arc", "cube", "leftBrace", "rightBrace", "ribbon2"] {
        assert!(preset_geometry_to_path(shape, &HashMap::new(), 1.0).is_none());
        assert!(
            preset_geometry_layers(shape, &HashMap::new(), 1.0)
                .unwrap()
                .len()
                > 1
        );
    }
}
