use std::collections::HashMap;

use ooxml_drawingml::{
    GeometryPathCommand, preset_geometry_default_adjustments, preset_geometry_layers,
    preset_geometry_to_path,
};
use proptest::prelude::*;
use proptest::sample::select;

/// Absolute slack on unit-frame coordinates; rounding here is around 1e-15.
const TOLERANCE: f64 = 1e-9;

/// Preset names covered by geometry properties.
const PRESETS: &[&str] = &[
    "accentCallout3",
    "accentBorderCallout1",
    "accentBorderCallout2",
    "accentBorderCallout3",
    "borderCallout2",
    "borderCallout3",
    "callout2",
    "callout3",
    "leftArrowCallout",
    "rightArrowCallout",
    "upDownArrowCallout",
    "leftRightArrowCallout",
    "quadArrowCallout",
    "actionButtonBlank",
    "actionButtonHome",
    "actionButtonHelp",
    "actionButtonInformation",
    "actionButtonForwardNext",
    "actionButtonBackPrevious",
    "actionButtonEnd",
    "actionButtonBeginning",
    "actionButtonReturn",
    "actionButtonSound",
    "actionButtonDocument",
    "actionButtonMovie",
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
    "round2SameRect",
    "round2DiagRect",
    "can",
    "blockArc",
    "pie",
    "flowChartExtract",
    "frame",
    "flowChartDelay",
    "mathEqual",
    "round1Rect",
    "snip2DiagRect",
    "stripedRightArrow",
    "snip1Rect",
    "bentUpArrow",
    "notchedRightArrow",
    "snip2SameRect",
    "mathPlus",
    "rightBracket",
    "teardrop",
    "leftBracket",
    "chord",
    "flowChartMerge",
    "accentCallout2",
    "cloud",
    "leftCircularArrow",
    "accentCallout1",
    "borderCallout1",
    "curvedLeftArrow",
    "flowChartOnlineStorage",
    "gear6",
    "curvedRightArrow",
    "uturnArrow",
    "moon",
    "gear9",
    "callout1",
    "upArrowCallout",
    "downArrowCallout",
    "quadArrow",
    "curvedUpArrow",
    "curvedDownArrow",
    "flowChartDocument",
    "verticalScroll",
    "horizontalScroll",
    "wave",
    "doubleWave",
    "heart",
    "irregularSeal1",
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
const HULL_OUTSIDE_FRAME: &[&str] = &[
    "ellipseRibbon",
    "noSmoking",
    "flowChartDocument",
    "wave",
    "doubleWave",
];

/// Points sampled along each curve when a test needs the outline rather than its hull.
const CURVE_SAMPLES: usize = 24;

/// These presets can exceed their frames at valid adjustments or aspect ratios.
const SPEC_OVERRUNS_FRAME: &[&str] = &[
    "accentCallout3",
    "accentBorderCallout1",
    "accentBorderCallout2",
    "accentBorderCallout3",
    "borderCallout2",
    "borderCallout3",
    "callout2",
    "callout3",
    "mathMultiply",
    "swooshArrow",
    "teardrop",
    "bentUpArrow",
    "accentCallout1",
    "accentCallout2",
    "borderCallout1",
    "callout1",
    "cloud",
    "heart",
    "curvedDownArrow",
];

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
        .or_else(|| {
            preset_geometry_layers(shape, adjustments, aspect).map(|layers| {
                layers
                    .into_iter()
                    .flat_map(|layer| layer.commands)
                    .collect()
            })
        })
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
        for shape in PRESETS.iter().chain(["arc", "cube", "leftBrace", "rightBrace", "ribbon2", "can", "leftBracket", "rightBracket"].iter()) {
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
        let curved = [
            "donut",
            "noSmoking",
            "ribbon",
            "ellipseRibbon",
            "cloudCallout",
            "wedgeEllipseCallout",
            "wedgeRoundRectCallout",
            "bentArrow",
            "cloud",
            "leftCircularArrow",
            "curvedLeftArrow",
            "flowChartOnlineStorage",
            "gear6",
            "curvedRightArrow",
            "uturnArrow",
            "moon",
            "gear9",
            "curvedUpArrow",
            "curvedDownArrow",
            "flowChartDocument",
            "verticalScroll",
            "horizontalScroll",
            "wave",
            "doubleWave",
            "heart",
        ];
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
fn rectangle_variants_follow_ecma_vertices() {
    assert_points(
        &endpoints(&draw("round2SameRect", &HashMap::new(), 2.0)),
        &[
            (0.083335, 0.0),
            (0.916665, 0.0),
            (1.0, 0.16667),
            (1.0, 1.0),
            (1.0, 1.0),
            (0.0, 1.0),
            (0.0, 1.0),
            (0.0, 0.16667),
            (0.083335, 0.0),
        ],
    );
    assert_points(
        &vertices(&draw("snip1Rect", &HashMap::new(), 2.0)),
        &[
            (0.0, 0.0),
            (0.916665, 0.0),
            (1.0, 0.16667),
            (1.0, 1.0),
            (0.0, 1.0),
        ],
    );
}

#[test]
fn frame_has_oppositely_wound_outer_and_inner_rectangles() {
    let parts = subpaths(&draw("frame", &HashMap::new(), 2.0));
    assert_eq!(parts.len(), 2);
    assert_points(
        &vertices(&parts[0]),
        &[(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)],
    );
    assert_points(
        &vertices(&parts[1]),
        &[
            (0.0625, 0.125),
            (0.0625, 0.875),
            (0.9375, 0.875),
            (0.9375, 0.125),
        ],
    );
    assert!(signed_area(&sample(&parts[0])) * signed_area(&sample(&parts[1])) < 0.0);
}

#[test]
fn math_equal_has_two_ecma_bars() {
    let parts = subpaths(&draw("mathEqual", &HashMap::new(), 2.0));
    assert_eq!(parts.len(), 2);
    assert_points(
        &vertices(&parts[0]),
        &[
            (0.13255, 0.206),
            (0.86745, 0.206),
            (0.86745, 0.4412),
            (0.13255, 0.4412),
        ],
    );
    assert_points(
        &vertices(&parts[1]),
        &[
            (0.13255, 0.5588),
            (0.86745, 0.5588),
            (0.86745, 0.794),
            (0.13255, 0.794),
        ],
    );
}

#[test]
fn flowchart_triangles_use_their_declared_coordinate_space() {
    for aspect in [0.5, 2.0] {
        assert_points(
            &vertices(&draw("flowChartExtract", &HashMap::new(), aspect)),
            &[(0.0, 1.0), (0.5, 0.0), (1.0, 1.0)],
        );
        assert_points(
            &vertices(&draw("flowChartMerge", &HashMap::new(), aspect)),
            &[(0.0, 0.0), (1.0, 0.0), (0.5, 1.0)],
        );
    }
}

#[test]
fn pie_defaults_trace_three_quadrants_in_a_wide_frame() {
    assert_eq!(
        preset_geometry_default_adjustments("pie"),
        named(&[("adj1", 0.0), ("adj2", 162.0)])
    );
    assert_points(
        &endpoints(&draw("pie", &HashMap::new(), 2.0)),
        &[(1.0, 0.5), (0.5, 1.0), (0.0, 0.5), (0.5, 0.0), (0.5, 0.5)],
    );
    let path = draw("pie", &named(&[("adj1", 27.0), ("adj2", 54.0)]), 2.0);
    assert_points(
        &endpoints(&path),
        &[
            (0.7236067977499789, 0.9472135954999579),
            (0.5, 1.0),
            (0.5, 0.5),
        ],
    );
}

#[test]
fn snips_and_block_arc_apply_non_default_adjustments() {
    let path = draw(
        "snip2DiagRect",
        &named(&[("adj1", 0.1), ("adj2", 0.3)]),
        2.0,
    );
    assert_points(
        &vertices(&path),
        &[
            (0.05, 0.0),
            (0.85, 0.0),
            (1.0, 0.3),
            (1.0, 0.9),
            (0.95, 1.0),
            (0.15, 1.0),
            (0.0, 0.7),
            (0.0, 0.1),
        ],
    );
    assert_points(
        &endpoints(&draw("blockArc", &named(&[("adj3", 0.1)]), 2.0)),
        &[
            (0.0, 0.5),
            (0.5, 0.0),
            (1.0, 0.5),
            (0.95, 0.5),
            (0.5, 0.1),
            (0.05, 0.5),
        ],
    );
    assert_points(
        &endpoints(&draw("blockArc", &HashMap::new(), 2.0)),
        &[
            (0.0, 0.5),
            (0.5, 0.0),
            (1.0, 0.5),
            (0.875, 0.5),
            (0.5, 0.25),
            (0.125, 0.5),
        ],
    );
}

#[test]
fn block_arc_with_a_collapsed_inner_radius_keeps_its_center() {
    assert_points(
        &endpoints(&draw(
            "blockArc",
            &named(&[("adj1", 0.0), ("adj2", 27.0), ("adj3", 0.5)]),
            2.0,
        )),
        &[
            (1.0, 0.5),
            (0.7236067977499789, 0.9472135954999579),
            (0.5, 0.5),
            (0.75, 0.5),
        ],
    );
}

#[test]
fn block_arc_with_subnormal_end_angle_stays_inside_its_frame() {
    for aspect in [
        1.0 / 64.0,
        0.5,
        1.0 - f64::EPSILON,
        1.0,
        1.0 + f64::EPSILON,
        2.0,
        64.0,
    ] {
        let expected = draw(
            "blockArc",
            &named(&[("adj2", 0.0), ("adj3", 4.394663328410801)]),
            aspect,
        );
        for adj2 in [
            f64::from_bits(1),
            4.86946011005727e-309,
            f64::MIN_POSITIVE,
            1e-20,
        ] {
            let path = draw(
                "blockArc",
                &named(&[("adj2", adj2), ("adj3", 4.394663328410801)]),
                aspect,
            );
            assert_points(&endpoints(&path), &endpoints(&expected));
            for (x, y) in coordinates(&path).into_iter().chain(sample(&path)) {
                assert!(
                    (-TOLERANCE..=1.0 + TOLERANCE).contains(&x)
                        && (-TOLERANCE..=1.0 + TOLERANCE).contains(&y),
                    "blockArc left its frame at ({x}, {y}) for adj2 {adj2} and aspect {aspect}"
                );
            }
        }
    }
}

#[test]
fn ribbon_at_zero_height_adjustment_keeps_its_curves() {
    for aspect in [1.0 / 64.0, 0.5, 1.0, 2.0, 64.0] {
        let path = draw("ribbon", &named(&[("adj1", 0.0)]), aspect);
        assert!(
            path.iter()
                .any(|command| matches!(command, GeometryPathCommand::Cubic { .. }))
        );
        for (x, y) in coordinates(&path).into_iter().chain(sample(&path)) {
            assert!(
                (-TOLERANCE..=1.0 + TOLERANCE).contains(&x)
                    && (-TOLERANCE..=1.0 + TOLERANCE).contains(&y),
                "ribbon left its frame at ({x}, {y}) for aspect {aspect}"
            );
        }
    }
}

#[test]
fn ecma_allows_teardrop_tips_and_bent_arrow_shafts_outside_the_frame() {
    assert_points(
        &endpoints(&draw("teardrop", &named(&[("adj", 2.0)]), 2.0)),
        &[
            (0.0, 0.5),
            (0.5, 0.0),
            (1.5, -0.5),
            (1.0, 0.5),
            (0.5, 1.0),
            (0.0, 0.5),
        ],
    );
    assert_points(
        &vertices(&draw(
            "bentUpArrow",
            &named(&[("adj1", 0.5), ("adj2", 0.0), ("adj3", 0.25)]),
            2.0,
        )),
        &[
            (0.0, 0.5),
            (0.875, 0.5),
            (0.875, 0.25),
            (1.0, 0.25),
            (1.0, 0.0),
            (1.0, 0.25),
            (1.125, 0.25),
            (1.125, 1.0),
            (0.0, 1.0),
        ],
    );
}

#[test]
fn can_at_zero_adjustment_keeps_every_layer_inside_its_frame() {
    for aspect in [1.0 / 64.0, 0.5, 1.0, 2.0, 64.0] {
        let layers = preset_geometry_layers("can", &named(&[("adj", 0.0)]), aspect).unwrap();
        assert_eq!(layers.len(), 3);
        for layer in &layers {
            for (x, y) in coordinates(&layer.commands)
                .into_iter()
                .chain(sample(&layer.commands))
            {
                assert!(
                    (-TOLERANCE..=1.0 + TOLERANCE).contains(&x)
                        && (-TOLERANCE..=1.0 + TOLERANCE).contains(&y),
                    "can left its frame at ({x}, {y}) for aspect {aspect}"
                );
            }
            assert!(layer.commands.iter().all(|command| match *command {
                GeometryPathCommand::Cubic { cp1y, cp2y, y, .. } => near((cp1y, cp2y), (y, y)),
                GeometryPathCommand::Move { .. }
                | GeometryPathCommand::Line { .. }
                | GeometryPathCommand::Close => true,
                _ => false,
            }));
        }
        assert_points(
            &endpoints(&layers[0].commands),
            &[(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)],
        );
        assert_points(
            &endpoints(&layers[1].commands),
            &[(0.0, 0.0), (1.0, 0.0), (0.0, 0.0)],
        );
        assert_points(
            &endpoints(&layers[2].commands),
            &[
                (1.0, 0.0),
                (0.0, 0.0),
                (1.0, 0.0),
                (1.0, 1.0),
                (0.0, 1.0),
                (0.0, 0.0),
            ],
        );
    }
}

#[test]
fn can_and_brackets_preserve_fill_and_open_stroke_layers() {
    use ooxml_drawingml::PresetPathFill;
    let layers = preset_geometry_layers("can", &HashMap::new(), 2.0).unwrap();
    assert_eq!(layers.len(), 3);
    assert_eq!(layers[0].fill, PresetPathFill::Normal);
    assert_eq!(layers[1].fill, PresetPathFill::Lighten);
    assert_eq!(layers[2].fill, PresetPathFill::None);
    assert!(!layers[0].stroke && !layers[1].stroke && layers[2].stroke);
    assert_points(
        &endpoints(&layers[0].commands),
        &[
            (0.0, 0.125),
            (0.5, 0.25),
            (1.0, 0.125),
            (1.0, 0.875),
            (0.5, 1.0),
            (0.0, 0.875),
        ],
    );
    for shape in ["can", "leftBracket", "rightBracket"] {
        let layers = preset_geometry_layers(shape, &HashMap::new(), 2.0).unwrap();
        let outline = layers.last().unwrap();
        assert_eq!(outline.fill, PresetPathFill::None);
        assert!(outline.stroke);
        assert!(
            !outline
                .commands
                .iter()
                .any(|command| matches!(command, GeometryPathCommand::Close))
        );
    }
}

#[test]
fn callouts_follow_ecma_vertices_and_keep_leader_lines_stroke_only() {
    use ooxml_drawingml::PresetPathFill;
    let one_segment_defaults = named(&[
        ("adj1", 0.1875),
        ("adj2", -0.08333),
        ("adj3", 1.125),
        ("adj4", -0.38333),
    ]);
    for (shape, body_stroked, leader_points) in [
        (
            "callout1",
            false,
            &[(-0.08333, 0.1875), (-0.38333, 1.125)][..],
        ),
        (
            "borderCallout1",
            true,
            &[(-0.08333, 0.1875), (-0.38333, 1.125)][..],
        ),
        (
            "accentCallout1",
            false,
            &[
                (-0.08333, 0.0),
                (-0.08333, 1.0),
                (-0.08333, 0.1875),
                (-0.38333, 1.125),
            ][..],
        ),
    ] {
        assert_eq!(
            preset_geometry_default_adjustments(shape),
            one_segment_defaults
        );
        let layers = preset_geometry_layers(shape, &HashMap::new(), 2.0).unwrap();
        assert_eq!(layers.len(), 2, "{shape}");
        assert_eq!(layers[0].fill, PresetPathFill::Normal);
        assert_eq!(layers[0].stroke, body_stroked);
        assert_points(
            &vertices(&layers[0].commands),
            &[(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)],
        );
        assert_eq!(layers[1].fill, PresetPathFill::None);
        assert!(layers[1].stroke);
        assert!(
            !layers[1]
                .commands
                .iter()
                .any(|command| matches!(command, GeometryPathCommand::Close))
        );
        assert_points(&vertices(&layers[1].commands), leader_points);
    }

    assert_eq!(
        preset_geometry_default_adjustments("accentCallout2"),
        named(&[
            ("adj1", 0.1875),
            ("adj2", -0.08333),
            ("adj3", 0.1875),
            ("adj4", -0.16667),
            ("adj5", 1.125),
            ("adj6", -0.46667),
        ])
    );
    let accent2 = preset_geometry_layers("accentCallout2", &HashMap::new(), 2.0).unwrap();
    assert_eq!(accent2.len(), 2);
    assert_eq!(accent2[0].fill, PresetPathFill::Normal);
    assert!(!accent2[0].stroke);
    assert_points(
        &vertices(&accent2[0].commands),
        &[(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)],
    );
    assert_eq!(accent2[1].fill, PresetPathFill::None);
    assert!(accent2[1].stroke);
    assert!(
        !accent2[1]
            .commands
            .iter()
            .any(|command| matches!(command, GeometryPathCommand::Close))
    );
    assert_points(
        &vertices(&accent2[1].commands),
        &[
            (-0.08333, 0.0),
            (-0.08333, 1.0),
            (-0.08333, 0.1875),
            (-0.16667, 0.1875),
            (-0.46667, 1.125),
        ],
    );

    assert_eq!(
        preset_geometry_default_adjustments("upArrowCallout"),
        named(&[
            ("adj1", 0.25),
            ("adj2", 0.25),
            ("adj3", 0.25),
            ("adj4", 0.64977),
        ])
    );
    assert_points(
        &vertices(&draw("upArrowCallout", &HashMap::new(), 2.0)),
        &[
            (0.0, 0.35023),
            (0.4375, 0.35023),
            (0.4375, 0.25),
            (0.375, 0.25),
            (0.5, 0.0),
            (0.625, 0.25),
            (0.5625, 0.25),
            (0.5625, 0.35023),
            (1.0, 0.35023),
            (1.0, 1.0),
            (0.0, 1.0),
        ],
    );

    assert_eq!(
        preset_geometry_default_adjustments("downArrowCallout"),
        named(&[
            ("adj1", 0.25),
            ("adj2", 0.25),
            ("adj3", 0.25),
            ("adj4", 0.64977),
        ])
    );
    assert_points(
        &vertices(&draw("downArrowCallout", &HashMap::new(), 2.0)),
        &[
            (0.0, 0.0),
            (1.0, 0.0),
            (1.0, 0.64977),
            (0.5625, 0.64977),
            (0.5625, 0.75),
            (0.625, 0.75),
            (0.5, 1.0),
            (0.375, 0.75),
            (0.4375, 0.75),
            (0.4375, 0.64977),
            (0.0, 0.64977),
        ],
    );
}

#[test]
fn uturn_arrow_follows_ecma_endpoints() {
    assert_eq!(
        preset_geometry_default_adjustments("uturnArrow"),
        named(&[
            ("adj1", 0.25),
            ("adj2", 0.25),
            ("adj3", 0.25),
            ("adj4", 0.4375),
            ("adj5", 0.75),
        ])
    );
    assert_points(
        &endpoints(&draw("uturnArrow", &HashMap::new(), 2.0)),
        &[
            (0.0, 1.0),
            (0.0, 0.4375),
            (0.21875, 0.0),
            (0.71875, 0.0),
            (0.9375, 0.4375),
            (0.9375, 0.5),
            (1.0, 0.5),
            (0.875, 0.75),
            (0.75, 0.5),
            (0.8125, 0.5),
            (0.8125, 0.4375),
            (0.71875, 0.25),
            (0.21875, 0.25),
            (0.125, 0.4375),
            (0.125, 1.0),
        ],
    );

    assert_eq!(
        preset_geometry_default_adjustments("quadArrow"),
        named(&[("adj1", 0.225), ("adj2", 0.225), ("adj3", 0.225)])
    );
    assert_points(
        &vertices(&draw("quadArrow", &HashMap::new(), 2.0)),
        &[
            (0.0, 0.5),
            (0.1125, 0.275),
            (0.1125, 0.3875),
            (0.44375, 0.3875),
            (0.44375, 0.225),
            (0.3875, 0.225),
            (0.5, 0.0),
            (0.6125, 0.225),
            (0.55625, 0.225),
            (0.55625, 0.3875),
            (0.8875, 0.3875),
            (0.8875, 0.275),
            (1.0, 0.5),
            (0.8875, 0.725),
            (0.8875, 0.6125),
            (0.55625, 0.6125),
            (0.55625, 0.775),
            (0.6125, 0.775),
            (0.5, 1.0),
            (0.3875, 0.775),
            (0.44375, 0.775),
            (0.44375, 0.6125),
            (0.1125, 0.6125),
            (0.1125, 0.725),
        ],
    );
}

#[test]
fn flowchart_online_storage_and_moon_follow_ecma_endpoints() {
    assert!(preset_geometry_default_adjustments("flowChartDocument").is_empty());
    for aspect in [0.5, 2.0] {
        assert_ecma_endpoints(
            &endpoints(&draw("flowChartOnlineStorage", &HashMap::new(), aspect)),
            &[
                (1.0 / 6.0, 0.0),
                (1.0, 0.0),
                (5.0 / 6.0, 0.5),
                (1.0, 1.0),
                (1.0 / 6.0, 1.0),
                (0.0, 0.5),
                (1.0 / 6.0, 0.0),
            ],
        );
        assert_points(
            &coordinates(&draw("flowChartDocument", &HashMap::new(), aspect)),
            &[
                (0.0, 0.0),
                (1.0, 0.0),
                (1.0, 17322.0 / 21600.0),
                (0.5, 17322.0 / 21600.0),
                (0.5, 23922.0 / 21600.0),
                (0.0, 20172.0 / 21600.0),
            ],
        );
    }
    assert_eq!(
        preset_geometry_default_adjustments("moon"),
        named(&[("adj", 0.5)])
    );
    for aspect in [0.5, 1.0, 2.0] {
        assert_ecma_endpoints(
            &endpoints(&draw("moon", &HashMap::new(), aspect)),
            &[(1.0, 1.0), (0.0, 0.5), (1.0, 0.0), (0.5, 0.5), (1.0, 1.0)],
        );
    }
}

#[test]
fn curved_arrows_follow_ecma_endpoints() {
    use ooxml_drawingml::PresetPathFill;
    let defaults = named(&[("adj1", 0.25), ("adj2", 0.5), ("adj3", 0.25)]);
    for shape in [
        "curvedLeftArrow",
        "curvedRightArrow",
        "curvedUpArrow",
        "curvedDownArrow",
    ] {
        assert_eq!(preset_geometry_default_adjustments(shape), defaults);
        let layers = preset_geometry_layers(shape, &HashMap::new(), 1.0).unwrap();
        assert_eq!(layers.len(), 3);
        assert_eq!(layers[0].fill, PresetPathFill::Normal);
        assert_eq!(layers[1].fill, PresetPathFill::DarkenLess);
        assert_eq!(layers[2].fill, PresetPathFill::None);
        assert!(!layers[0].stroke && !layers[1].stroke && layers[2].stroke);
    }

    let left_default = preset_geometry_layers("curvedLeftArrow", &HashMap::new(), 1.0).unwrap();
    assert_ecma_endpoints(
        &endpoints(&left_default[0].commands),
        &[
            (0.0, 0.75),
            (0.25, 0.4900768239224544),
            (0.25, 0.6150768239224544),
            (0.916515138991168, 0.4375),
            (1.0, 0.5625),
            (0.25, 0.8650768239224544),
            (0.25, 0.9900768239224544),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&left_default[1].commands),
        &[(1.0, 0.5625), (0.0, 0.25), (0.0, 0.0), (1.0, 0.3125)],
    );
    assert_ecma_endpoints(
        &endpoints(&left_default[2].commands),
        &[
            (1.0, 0.5625),
            (0.0, 0.25),
            (0.0, 0.0),
            (1.0, 0.3125),
            (1.0, 0.5625),
            (0.25, 0.8650768239224544),
            (0.25, 0.9900768239224544),
            (0.0, 0.75),
            (0.25, 0.4900768239224544),
            (0.25, 0.6150768239224544),
            (0.916515138991168, 0.4375),
        ],
    );

    let custom = named(&[("adj1", 0.36), ("adj2", 0.44), ("adj3", 0.6)]);
    let left_custom = preset_geometry_layers("curvedLeftArrow", &custom, 1.0).unwrap();
    assert_ecma_endpoints(
        &endpoints(&left_custom[0].commands),
        &[
            (0.0, 0.78),
            (0.6, 0.5),
            (0.6, 0.54),
            (0.8, 0.48),
            (1.0, 0.66),
            (0.6, 0.9),
            (0.6, 0.94),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&left_custom[1].commands),
        &[(1.0, 0.66), (0.0, 0.36), (0.0, 0.0), (1.0, 0.3)],
    );
    assert_ecma_endpoints(
        &endpoints(&left_custom[2].commands),
        &[
            (1.0, 0.66),
            (0.0, 0.36),
            (0.0, 0.0),
            (1.0, 0.3),
            (1.0, 0.66),
            (0.6, 0.9),
            (0.6, 0.94),
            (0.0, 0.78),
            (0.6, 0.5),
            (0.6, 0.54),
            (0.8, 0.48),
        ],
    );

    let right_default = preset_geometry_layers("curvedRightArrow", &HashMap::new(), 1.0).unwrap();
    assert_ecma_endpoints(
        &endpoints(&right_default[0].commands),
        &[
            (0.0, 0.3125),
            (0.75, 0.6150768239224544),
            (0.75, 0.4900768239224544),
            (1.0, 0.75),
            (0.75, 0.9900768239224544),
            (0.75, 0.8650768239224544),
            (0.0, 0.5625),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&right_default[1].commands),
        &[
            (1.0, 0.25),
            (0.08348486100883202, 0.4375),
            (0.0, 0.3125),
            (1.0, 0.0),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&right_default[2].commands),
        &[
            (0.0, 0.3125),
            (0.75, 0.6150768239224544),
            (0.75, 0.4900768239224544),
            (1.0, 0.75),
            (0.75, 0.9900768239224544),
            (0.75, 0.8650768239224544),
            (0.0, 0.5625),
            (0.0, 0.3125),
            (1.0, 0.0),
            (1.0, 0.25),
            (0.08348486100883202, 0.4375),
        ],
    );

    let right_custom = preset_geometry_layers("curvedRightArrow", &custom, 1.0).unwrap();
    assert_ecma_endpoints(
        &endpoints(&right_custom[0].commands),
        &[
            (0.0, 0.3),
            (0.4, 0.54),
            (0.4, 0.5),
            (1.0, 0.78),
            (0.4, 0.94),
            (0.4, 0.9),
            (0.0, 0.66),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&right_custom[1].commands),
        &[(1.0, 0.36), (0.2, 0.48), (0.0, 0.3), (1.0, 0.0)],
    );
    assert_ecma_endpoints(
        &endpoints(&right_custom[2].commands),
        &[
            (0.0, 0.3),
            (0.4, 0.54),
            (0.4, 0.5),
            (1.0, 0.78),
            (0.4, 0.94),
            (0.4, 0.9),
            (0.0, 0.66),
            (0.0, 0.3),
            (1.0, 0.0),
            (1.0, 0.36),
            (0.2, 0.48),
        ],
    );

    let up_default = preset_geometry_layers("curvedUpArrow", &HashMap::new(), 1.0).unwrap();
    assert_ecma_endpoints(
        &endpoints(&up_default[0].commands),
        &[
            (0.75, 0.0),
            (0.9900768239224544, 0.25),
            (0.8650768239224544, 0.25),
            (0.5625, 1.0),
            (0.4375, 0.916515138991168),
            (0.6150768239224544, 0.25),
            (0.4900768239224544, 0.25),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&up_default[1].commands),
        &[(0.3125, 1.0), (0.0, 0.0), (0.25, 0.0), (0.5625, 1.0)],
    );
    assert_ecma_endpoints(
        &endpoints(&up_default[2].commands),
        &[
            (0.4375, 0.916515138991168),
            (0.6150768239224544, 0.25),
            (0.4900768239224544, 0.25),
            (0.75, 0.0),
            (0.9900768239224544, 0.25),
            (0.8650768239224544, 0.25),
            (0.5625, 1.0),
            (0.3125, 1.0),
            (0.0, 0.0),
            (0.25, 0.0),
            (0.5625, 1.0),
        ],
    );

    // a2=0.32, a1=th=aw=0.32, q1=0.16, wR=0.34, q7=0.68, q11=0.6, idy=maxAdj3=0.8823529411764706,
    // a3=ah=y1=0.8, x3=0.66, q5=0.6, dx=0.204, x4=x5=0.544, x7=x8=0.864, dh=0.0, x6=0.84, ix=0.5.
    let up_adj = named(&[("adj1", 0.48), ("adj2", 0.32), ("adj3", 0.8)]);
    let up_custom = preset_geometry_layers("curvedUpArrow", &up_adj, 1.0).unwrap();
    assert_ecma_endpoints(
        &endpoints(&up_custom[0].commands),
        &[
            (0.84, 0.0),
            (0.864, 0.8),
            (0.864, 0.8),
            (0.66, 1.0),
            (0.5, 0.8823529411764706),
            (0.544, 0.8),
            (0.544, 0.8),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&up_custom[1].commands),
        &[(0.34, 1.0), (0.0, 0.0), (0.32, 0.0), (0.66, 1.0)],
    );
    assert_ecma_endpoints(
        &endpoints(&up_custom[2].commands),
        &[
            (0.5, 0.8823529411764706),
            (0.544, 0.8),
            (0.544, 0.8),
            (0.84, 0.0),
            (0.864, 0.8),
            (0.864, 0.8),
            (0.66, 1.0),
            (0.34, 1.0),
            (0.0, 0.0),
            (0.32, 0.0),
            (0.66, 1.0),
        ],
    );

    let up_over_adj = named(&[("adj3", 1.5)]);
    let up_pinned = preset_geometry_layers("curvedUpArrow", &up_over_adj, 1.0).unwrap();
    let up_max_adj = named(&[("adj3", 0.916515138991168)]);
    let up_at_max = preset_geometry_layers("curvedUpArrow", &up_max_adj, 1.0).unwrap();
    for (layer, max_layer) in up_pinned.iter().zip(&up_at_max) {
        assert_points(&endpoints(&layer.commands), &endpoints(&max_layer.commands));
        for (x, y) in coordinates(&layer.commands)
            .into_iter()
            .chain(sample(&layer.commands))
        {
            assert!(
                (-TOLERANCE..=1.0 + TOLERANCE).contains(&x)
                    && (-TOLERANCE..=1.0 + TOLERANCE).contains(&y),
                "curvedUpArrow left its frame at ({x}, {y})"
            );
        }
    }
    assert_ecma_endpoints(
        &endpoints(&up_pinned[0].commands),
        &[
            (0.75, 0.0),
            (0.8125, 0.916515138991168),
            (0.6875, 0.916515138991168),
            (0.5625, 1.0),
            (0.4375, 0.916515138991168),
            (0.4375, 0.916515138991168),
            (0.3125, 0.916515138991168),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&up_pinned[1].commands),
        &[(0.3125, 1.0), (0.0, 0.0), (0.25, 0.0), (0.5625, 1.0)],
    );
    assert_ecma_endpoints(
        &endpoints(&up_pinned[2].commands),
        &[
            (0.4375, 0.916515138991168),
            (0.4375, 0.916515138991168),
            (0.3125, 0.916515138991168),
            (0.75, 0.0),
            (0.8125, 0.916515138991168),
            (0.6875, 0.916515138991168),
            (0.5625, 1.0),
            (0.3125, 1.0),
            (0.0, 0.0),
            (0.25, 0.0),
            (0.5625, 1.0),
        ],
    );

    let down_default = preset_geometry_layers("curvedDownArrow", &HashMap::new(), 1.0).unwrap();
    assert_ecma_endpoints(
        &endpoints(&down_default[0].commands),
        &[
            (0.75, 1.0),
            (0.4900768239224544, 0.75),
            (0.6150768239224544, 0.75),
            (0.3125, 0.0),
            (0.5625, 0.0),
            (0.8650768239224544, 0.75),
            (0.9900768239224544, 0.75),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&down_default[1].commands),
        &[
            (0.4375, 0.08348486100883202),
            (0.25, 1.0),
            (0.0, 1.0),
            (0.3125, 0.0),
            (0.4375, 0.08348486100883202),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&down_default[2].commands),
        &[
            (0.4375, 0.08348486100883202),
            (0.25, 1.0),
            (0.0, 1.0),
            (0.3125, 0.0),
            (0.5625, 0.0),
            (0.8650768239224544, 0.75),
            (0.9900768239224544, 0.75),
            (0.75, 1.0),
            (0.4900768239224544, 0.75),
            (0.6150768239224544, 0.75),
            (0.3125, 0.0),
        ],
    );

    let down_custom = preset_geometry_layers("curvedDownArrow", &custom, 1.0).unwrap();
    assert_ecma_endpoints(
        &endpoints(&down_custom[0].commands),
        &[
            (0.78, 1.0),
            (0.5, 0.4),
            (0.54, 0.4),
            (0.3, 0.0),
            (0.66, 0.0),
            (0.9, 0.4),
            (0.94, 0.4),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&down_custom[1].commands),
        &[
            (0.48, 0.2),
            (0.36, 1.0),
            (0.0, 1.0),
            (0.3, 0.0),
            (0.48, 0.2),
        ],
    );
    assert_ecma_endpoints(
        &endpoints(&down_custom[2].commands),
        &[
            (0.48, 0.2),
            (0.36, 1.0),
            (0.0, 1.0),
            (0.3, 0.0),
            (0.66, 0.0),
            (0.9, 0.4),
            (0.94, 0.4),
            (0.78, 1.0),
            (0.5, 0.4),
            (0.54, 0.4),
            (0.3, 0.0),
        ],
    );
}

#[test]
fn left_circular_arrow_gear6_and_heart_follow_ecma_endpoints() {
    assert_eq!(
        preset_geometry_default_adjustments("leftCircularArrow"),
        named(&[
            ("adj1", 0.125),
            ("adj2", -11.42319),
            ("adj3", 11.42319),
            ("adj4", 108.0),
            ("adj5", 0.125),
        ])
    );
    assert_ecma_endpoints(
        &endpoints(&draw("leftCircularArrow", &HashMap::new(), 1.0)),
        &[
            (0.0625, 0.5),
            (0.1875, 0.5),
            (0.5, 0.8125),
            (0.7875626944045718, 0.6223272119635808),
            (0.7294870282721458, 0.6223272119635808),
            (0.875, 0.5),
            (0.9794870282721458, 0.6223272119635808),
            (0.9200503579491605, 0.6223272119635808),
            (0.5, 0.9375),
            (0.0625, 0.5),
        ],
    );
    let circular_adj = named(&[
        ("adj1", 0.2),
        ("adj2", 0.0),
        ("adj3", 27.0),
        ("adj4", 108.0),
        ("adj5", 0.2),
    ]);
    assert_ecma_endpoints(
        &endpoints(&draw("leftCircularArrow", &circular_adj, 1.0)),
        &[
            (0.1, 0.5),
            (0.3, 0.5),
            (0.5, 0.7),
            (0.6414213562373095, 0.6414213562373095),
            (0.5707106781186548, 0.5707106781186548),
            (0.7121320343559643, 0.7121320343559643),
            (0.8535533905932737, 0.8535533905932737),
            (0.782842712474619, 0.782842712474619),
            (0.5, 0.9),
            (0.1, 0.5),
        ],
    );

    assert_eq!(
        preset_geometry_default_adjustments("gear6"),
        named(&[("adj1", 0.15), ("adj2", 0.03526)])
    );
    assert_ecma_endpoints(
        &endpoints(&draw("gear6", &HashMap::new(), 1.0)),
        &[
            (0.748246990037985, 0.2532745818990671),
            (0.8957808006056508, 0.2088106376365064),
            (0.9500677854214032, 0.3028384535271097),
            (0.8377939748537374, 0.408374509264549),
            (0.85, 0.5),
            (0.8377939748537374, 0.591625490735451),
            (0.9500677854214032, 0.6971615464728903),
            (0.8957808006056508, 0.7911893623634936),
            (0.748246990037985, 0.7467254181009329),
            (0.5895469848157524, 0.8383509088363839),
            (0.5542869848157524, 0.9883509088363839),
            (0.4457130151842476, 0.9883509088363839),
            (0.4104530151842476, 0.8383509088363839),
            (0.251753009962015, 0.7467254181009329),
            (0.1042191993943492, 0.7911893623634936),
            (0.0499322145785968, 0.6971615464728903),
            (0.1622060251462626, 0.591625490735451),
            (0.15, 0.5),
            (0.1622060251462626, 0.408374509264549),
            (0.0499322145785968, 0.3028384535271097),
            (0.1042191993943492, 0.2088106376365064),
            (0.251753009962015, 0.2532745818990671),
            (0.4104530151842476, 0.1616490911636161),
            (0.4457130151842476, 0.0116490911636161),
            (0.5542869848157524, 0.0116490911636161),
            (0.5895469848157524, 0.1616490911636161),
            (0.748246990037985, 0.2532745818990671),
        ],
    );
    let gear_custom = draw("gear6", &named(&[("adj1", 0.2), ("adj2", 0.0)]), 1.0);
    assert_ecma_endpoints(
        &endpoints(&gear_custom),
        &[
            (0.6990409859747991, 0.275539121666648),
            (0.8722460667316868, 0.175539121666648),
            (0.9671143965367382, 0.3398558889181979),
            (0.7939093157798505, 0.4398558889181979),
            (0.8, 0.5),
            (0.7939093157798505, 0.5601441110818021),
            (0.9671143965367382, 0.6601441110818021),
            (0.8722460667316868, 0.824460878333352),
            (0.6990409859747991, 0.724460878333352),
            (0.5948683298050514, 0.7846049894151541),
            (0.5948683298050514, 0.9846049894151541),
            (0.4051316701949486, 0.9846049894151541),
            (0.4051316701949486, 0.7846049894151541),
            (0.3009590140252009, 0.724460878333352),
            (0.1277539332683132, 0.824460878333352),
            (0.0328856034632618, 0.6601441110818021),
            (0.2060906842201495, 0.5601441110818021),
            (0.2, 0.5),
            (0.2060906842201495, 0.4398558889181979),
            (0.0328856034632618, 0.3398558889181979),
            (0.1277539332683132, 0.175539121666648),
            (0.3009590140252009, 0.275539121666648),
            (0.4051316701949486, 0.2153950105848459),
            (0.4051316701949486, 0.01539501058484589),
            (0.5948683298050514, 0.01539501058484589),
            (0.5948683298050514, 0.2153950105848459),
            (0.6990409859747991, 0.275539121666648),
        ],
    );

    assert!(preset_geometry_default_adjustments("heart").is_empty());
    let heart_default = draw("heart", &HashMap::new(), 2.0);
    assert_points(
        &endpoints(&heart_default),
        &[(0.5, 0.25), (0.5, 1.0), (0.5, 0.25)],
    );
    assert_points(
        &coordinates(&heart_default),
        &[
            (0.5, 0.25),
            (17.0 / 24.0, -1.0 / 3.0),
            (73.0 / 48.0, 0.25),
            (0.5, 1.0),
            (-25.0 / 48.0, 0.25),
            (7.0 / 24.0, -1.0 / 3.0),
            (0.5, 0.25),
        ],
    );
    let heart_custom = draw("heart", &named(&[("adj", 0.75)]), 0.5);
    assert_points(
        &endpoints(&heart_custom),
        &[(0.5, 0.25), (0.5, 1.0), (0.5, 0.25)],
    );
    assert_points(
        &coordinates(&heart_custom),
        &[
            (0.5, 0.25),
            (17.0 / 24.0, -1.0 / 3.0),
            (73.0 / 48.0, 0.25),
            (0.5, 1.0),
            (-25.0 / 48.0, 0.25),
            (7.0 / 24.0, -1.0 / 3.0),
            (0.5, 0.25),
        ],
    );
}

fn endpoints(path: &[GeometryPathCommand]) -> Vec<(f64, f64)> {
    path.iter()
        .filter_map(|command| match *command {
            GeometryPathCommand::Move { x, y }
            | GeometryPathCommand::Line { x, y }
            | GeometryPathCommand::Quad { x, y, .. }
            | GeometryPathCommand::Cubic { x, y, .. } => Some((x, y)),
            GeometryPathCommand::Close => None,
        })
        .collect()
}

fn assert_points(actual: &[(f64, f64)], expected: &[(f64, f64)]) {
    assert_eq!(actual.len(), expected.len(), "{actual:?}");
    for (&actual, &expected) in actual.iter().zip(expected) {
        assert!(near(actual, expected), "{actual:?} != {expected:?}");
    }
}

fn assert_ecma_endpoints(actual: &[(f64, f64)], expected: &[(f64, f64)]) {
    let mut rest = actual;
    for &expected_point in expected {
        let Some(index) = rest
            .iter()
            .position(|&actual_point| near(actual_point, expected_point))
        else {
            panic!("missing {expected_point:?} in {actual:?}");
        };
        rest = &rest[index + 1..];
    }
}

#[test]
fn single_path_consumers_do_not_receive_partial_layered_presets() {
    for shape in [
        "accentCallout3",
        "accentBorderCallout1",
        "accentBorderCallout2",
        "accentBorderCallout3",
        "borderCallout2",
        "borderCallout3",
        "callout2",
        "callout3",
        "actionButtonBackPrevious",
        "actionButtonBeginning",
        "actionButtonDocument",
        "actionButtonEnd",
        "actionButtonForwardNext",
        "actionButtonHelp",
        "actionButtonHome",
        "actionButtonInformation",
        "actionButtonMovie",
        "actionButtonReturn",
        "actionButtonSound",
        "arc",
        "cube",
        "leftBrace",
        "rightBrace",
        "ribbon2",
        "can",
        "leftBracket",
        "rightBracket",
        "accentCallout1",
        "accentCallout2",
        "borderCallout1",
        "callout1",
        "cloud",
        "curvedLeftArrow",
        "curvedRightArrow",
        "curvedUpArrow",
        "curvedDownArrow",
        "verticalScroll",
        "horizontalScroll",
    ] {
        assert!(preset_geometry_to_path(shape, &HashMap::new(), 1.0).is_none());
        assert!(
            preset_geometry_layers(shape, &HashMap::new(), 1.0)
                .unwrap()
                .len()
                > 1
        );
    }
}

fn assert_inside_frame(shape: &str, path: &[GeometryPathCommand], label: &str) {
    let hull = if HULL_OUTSIDE_FRAME.contains(&shape) {
        Vec::new()
    } else {
        coordinates(path)
    };
    for (x, y) in sample(path).into_iter().chain(hull) {
        assert!(
            (-TOLERANCE..=1.0 + TOLERANCE).contains(&x)
                && (-TOLERANCE..=1.0 + TOLERANCE).contains(&y),
            "{shape} {label} left its frame at ({x}, {y})"
        );
    }
}

#[test]
fn a_can_with_a_flat_top_stays_inside_its_frame() {
    for adj in [
        9.343293760265696e-15,
        5.484785289152219e-13,
        1.4054251999661466e-12,
    ] {
        let path = draw("can", &named(&[("adj", adj)]), 1.0);
        assert_inside_frame("can", &path, &format!("adj {adj}"));
    }
}

#[test]
fn a_near_zero_adjust_stays_inside_its_frame() {
    let pinned = PRESETS
        .iter()
        .filter(|shape| {
            !UNPINNED_CONNECTORS
                .iter()
                .any(|prefix| shape.starts_with(prefix))
        })
        .filter(|shape| !FRAME_SCALED_STARS.contains(shape))
        .filter(|shape| !CALLOUTS.contains(shape))
        .filter(|shape| !SPEC_OVERRUNS_FRAME.contains(shape));
    for shape in pinned {
        for adjust in ["adj", "adj1", "adj2", "adj3", "adj4", "vf"] {
            for value in [1e-16, 1e-15, 1e-14, 1e-12, 1e-9, 1e-6] {
                for aspect in [1.0 / 64.0, 1.0, 64.0] {
                    let path = draw(shape, &named(&[(adjust, value)]), aspect);
                    assert_inside_frame(shape, &path, &format!("{adjust} {value} aspect {aspect}"));
                }
            }
        }
    }
}

#[test]
fn scrolls_follow_ecma_endpoints() {
    use ooxml_drawingml::PresetPathFill;
    for shape in ["verticalScroll", "horizontalScroll"] {
        assert_eq!(
            preset_geometry_default_adjustments(shape),
            named(&[("adj", 0.125)])
        );
        let layers = preset_geometry_layers(shape, &HashMap::new(), 1.0).unwrap();
        assert_eq!(layers.len(), 3);
        assert_eq!(layers[0].fill, PresetPathFill::Normal);
        assert_eq!(layers[1].fill, PresetPathFill::DarkenLess);
        assert_eq!(layers[2].fill, PresetPathFill::None);
        assert!(!layers[0].stroke && !layers[1].stroke && layers[2].stroke);
        for layer in &layers[..2] {
            let parts = subpaths(&layer.commands);
            assert_eq!(parts.len(), 2);
            assert!(
                parts
                    .iter()
                    .all(|part| matches!(part.last(), Some(GeometryPathCommand::Close)))
            );
        }
        assert_eq!(subpaths(&layers[2].commands).len(), 5);
    }
    let vertical = preset_geometry_layers("verticalScroll", &HashMap::new(), 1.0).unwrap();
    assert_points(
        &endpoints(&vertical[0].commands),
        &[
            (0.0625, 1.0),
            (0.125, 0.9375),
            (0.0625, 0.9375),
            (0.09375, 0.90625),
            (0.0625, 0.875),
            (0.125, 0.875),
            (0.125, 0.0625),
            (0.1875, 0.0),
            (0.9375, 0.0),
            (1.0, 0.0625),
            (0.9375, 0.125),
            (0.875, 0.125),
            (0.875, 0.9375),
            (0.8125, 1.0),
            (0.25, 0.0625),
            (0.1875, 0.125),
            (0.15625, 0.09375),
            (0.1875, 0.0625),
        ],
    );
    assert_points(
        &endpoints(&vertical[1].commands),
        &[
            (0.25, 0.0625),
            (0.1875, 0.125),
            (0.15625, 0.09375),
            (0.1875, 0.0625),
            (0.125, 0.9375),
            (0.0625, 1.0),
            (0.0, 0.9375),
            (0.0625, 0.875),
            (0.09375, 0.90625),
            (0.0625, 0.9375),
        ],
    );
    assert_points(
        &endpoints(&vertical[2].commands),
        &[
            (0.125, 0.875),
            (0.125, 0.0625),
            (0.1875, 0.0),
            (0.9375, 0.0),
            (1.0, 0.0625),
            (0.9375, 0.125),
            (0.875, 0.125),
            (0.875, 0.9375),
            (0.8125, 1.0),
            (0.0625, 1.0),
            (0.0, 0.9375),
            (0.0625, 0.875),
            (0.1875, 0.0),
            (0.25, 0.0625),
            (0.1875, 0.125),
            (0.15625, 0.09375),
            (0.1875, 0.0625),
            (0.25, 0.0625),
            (0.875, 0.125),
            (0.1875, 0.125),
            (0.0625, 0.875),
            (0.09375, 0.90625),
            (0.0625, 0.9375),
            (0.125, 0.9375),
            (0.0625, 1.0),
            (0.125, 0.9375),
            (0.125, 0.875),
        ],
    );

    let horizontal = preset_geometry_layers("horizontalScroll", &HashMap::new(), 1.0).unwrap();
    assert_points(
        &endpoints(&horizontal[0].commands),
        &[
            (1.0, 0.0625),
            (0.9375, 0.125),
            (0.9375, 0.0625),
            (0.90625, 0.09375),
            (0.875, 0.0625),
            (0.875, 0.125),
            (0.0625, 0.125),
            (0.0, 0.1875),
            (0.0, 0.9375),
            (0.0625, 1.0),
            (0.125, 0.9375),
            (0.125, 0.875),
            (0.9375, 0.875),
            (1.0, 0.8125),
            (0.0625, 0.25),
            (0.125, 0.1875),
            (0.09375, 0.15625),
            (0.0625, 0.1875),
        ],
    );
    assert_points(
        &endpoints(&horizontal[1].commands),
        &[
            (0.0625, 0.25),
            (0.125, 0.1875),
            (0.09375, 0.15625),
            (0.0625, 0.1875),
            (0.9375, 0.125),
            (1.0, 0.0625),
            (0.9375, 0.0),
            (0.875, 0.0625),
            (0.90625, 0.09375),
            (0.9375, 0.0625),
        ],
    );
    assert_points(
        &endpoints(&horizontal[2].commands),
        &[
            (0.0, 0.1875),
            (0.0625, 0.125),
            (0.875, 0.125),
            (0.875, 0.0625),
            (0.9375, 0.0),
            (1.0, 0.0625),
            (1.0, 0.8125),
            (0.9375, 0.875),
            (0.125, 0.875),
            (0.125, 0.9375),
            (0.0625, 1.0),
            (0.0, 0.9375),
            (0.875, 0.125),
            (0.9375, 0.125),
            (1.0, 0.0625),
            (0.9375, 0.125),
            (0.9375, 0.0625),
            (0.90625, 0.09375),
            (0.875, 0.0625),
            (0.0625, 0.25),
            (0.0625, 0.1875),
            (0.09375, 0.15625),
            (0.125, 0.1875),
            (0.0625, 0.25),
            (0.0, 0.1875),
            (0.125, 0.1875),
            (0.125, 0.875),
        ],
    );
}

#[test]
fn waves_follow_ecma_coordinates() {
    assert_eq!(
        preset_geometry_default_adjustments("wave"),
        named(&[("adj1", 0.125), ("adj2", 0.0)])
    );
    assert_points(
        &coordinates(&draw("wave", &HashMap::new(), 2.0)),
        &[
            (0.0, 0.125),
            (1.0 / 3.0, -7.0 / 24.0),
            (2.0 / 3.0, 13.0 / 24.0),
            (1.0, 0.125),
            (1.0, 0.875),
            (2.0 / 3.0, 31.0 / 24.0),
            (1.0 / 3.0, 11.0 / 24.0),
            (0.0, 0.875),
        ],
    );
    assert_points(
        &coordinates(&draw("wave", &named(&[("adj1", 0.2), ("adj2", 0.1)]), 0.5)),
        &[
            (0.0, 0.2),
            (4.0 / 15.0, -7.0 / 15.0),
            (8.0 / 15.0, 13.0 / 15.0),
            (0.8, 0.2),
            (1.0, 0.8),
            (11.0 / 15.0, 22.0 / 15.0),
            (7.0 / 15.0, 2.0 / 15.0),
            (0.2, 0.8),
        ],
    );
    assert_points(
        &coordinates(&draw("wave", &named(&[("adj1", 0.2), ("adj2", -0.1)]), 2.0)),
        &[
            (0.2, 0.2),
            (7.0 / 15.0, -7.0 / 15.0),
            (11.0 / 15.0, 13.0 / 15.0),
            (1.0, 0.2),
            (0.8, 0.8),
            (8.0 / 15.0, 22.0 / 15.0),
            (4.0 / 15.0, 2.0 / 15.0),
            (0.0, 0.8),
        ],
    );
}

#[test]
fn double_wave_follows_ecma_coordinates() {
    assert_eq!(
        preset_geometry_default_adjustments("doubleWave"),
        named(&[("adj1", 0.0625), ("adj2", 0.0)])
    );
    assert_points(
        &coordinates(&draw("doubleWave", &HashMap::new(), 0.5)),
        &[
            (0.0, 0.0625),
            (1.0 / 6.0, -7.0 / 48.0),
            (1.0 / 3.0, 13.0 / 48.0),
            (0.5, 0.0625),
            (2.0 / 3.0, -7.0 / 48.0),
            (5.0 / 6.0, 13.0 / 48.0),
            (1.0, 0.0625),
            (1.0, 0.9375),
            (5.0 / 6.0, 55.0 / 48.0),
            (2.0 / 3.0, 35.0 / 48.0),
            (0.5, 0.9375),
            (1.0 / 3.0, 55.0 / 48.0),
            (1.0 / 6.0, 35.0 / 48.0),
            (0.0, 0.9375),
        ],
    );
    for (adj2, top, bottom) in [(0.1, 0.0, 0.2), (-0.1, 0.2, 0.0)] {
        assert_points(
            &coordinates(&draw(
                "doubleWave",
                &named(&[("adj1", 0.125), ("adj2", adj2)]),
                2.0,
            )),
            &[
                (top, 0.125),
                (top + 2.0 / 15.0, -7.0 / 24.0),
                (top + 4.0 / 15.0, 13.0 / 24.0),
                (top + 0.4, 0.125),
                (top + 8.0 / 15.0, -7.0 / 24.0),
                (top + 2.0 / 3.0, 13.0 / 24.0),
                (top + 0.8, 0.125),
                (bottom + 0.8, 0.875),
                (bottom + 2.0 / 3.0, 31.0 / 24.0),
                (bottom + 8.0 / 15.0, 11.0 / 24.0),
                (bottom + 0.4, 0.875),
                (bottom + 4.0 / 15.0, 31.0 / 24.0),
                (bottom + 2.0 / 15.0, 11.0 / 24.0),
                (bottom, 0.875),
            ],
        );
    }
}

#[test]
fn wave_adjustments_use_ecma_limits() {
    for (shape, max_adj1) in [("wave", 0.2), ("doubleWave", 0.125)] {
        for aspect in [0.5, 1.0, 2.0] {
            for (adj1, pinned_adj1) in [(-1.0, 0.0), (1.0, max_adj1)] {
                for (adj2, pinned_adj2) in [(-1.0, -0.1), (1.0, 0.1)] {
                    assert_points(
                        &coordinates(&draw(
                            shape,
                            &named(&[("adj1", adj1), ("adj2", adj2)]),
                            aspect,
                        )),
                        &coordinates(&draw(
                            shape,
                            &named(&[("adj1", pinned_adj1), ("adj2", pinned_adj2)]),
                            aspect,
                        )),
                    );
                }
            }
        }
    }
}

#[test]
fn curved_down_arrow_uses_reference_thickness_and_head_guides() {
    let layers = preset_geometry_layers(
        "curvedDownArrow",
        &named(&[("adj1", 0.36), ("adj2", 0.04), ("adj3", 0.96)]),
        1.0,
    )
    .unwrap();
    assert_points(
        &vertices(&layers[0].commands),
        &[
            (0.98, 1.0),
            (0.672, 0.04),
            (0.512, 0.04),
            (0.76, 0.0),
            (0.712, 0.04),
        ],
    );
}

#[test]
fn irregular_seal1_follows_ecma_vertices() {
    use ooxml_drawingml::PresetPathFill;
    assert!(preset_geometry_default_adjustments("irregularSeal1").is_empty());
    let expected = [
        (10800.0 / 21600.0, 5800.0 / 21600.0),
        (14522.0 / 21600.0, 0.0),
        (14155.0 / 21600.0, 5325.0 / 21600.0),
        (18380.0 / 21600.0, 4457.0 / 21600.0),
        (16702.0 / 21600.0, 7315.0 / 21600.0),
        (21097.0 / 21600.0, 8137.0 / 21600.0),
        (17607.0 / 21600.0, 10475.0 / 21600.0),
        (1.0, 13290.0 / 21600.0),
        (16837.0 / 21600.0, 12942.0 / 21600.0),
        (18145.0 / 21600.0, 18095.0 / 21600.0),
        (14020.0 / 21600.0, 14457.0 / 21600.0),
        (13247.0 / 21600.0, 19737.0 / 21600.0),
        (10532.0 / 21600.0, 14935.0 / 21600.0),
        (8485.0 / 21600.0, 1.0),
        (7715.0 / 21600.0, 15627.0 / 21600.0),
        (4762.0 / 21600.0, 17617.0 / 21600.0),
        (5667.0 / 21600.0, 13937.0 / 21600.0),
        (135.0 / 21600.0, 14587.0 / 21600.0),
        (3722.0 / 21600.0, 11775.0 / 21600.0),
        (0.0, 8615.0 / 21600.0),
        (4627.0 / 21600.0, 7617.0 / 21600.0),
        (370.0 / 21600.0, 2295.0 / 21600.0),
        (7312.0 / 21600.0, 6320.0 / 21600.0),
        (8352.0 / 21600.0, 2295.0 / 21600.0),
    ];
    for aspect in [0.5, 1.0, 2.0] {
        let layers = preset_geometry_layers("irregularSeal1", &HashMap::new(), aspect).unwrap();
        assert_eq!(layers.len(), 1);
        assert_eq!(layers[0].fill, PresetPathFill::Normal);
        assert!(layers[0].stroke);
        let path = draw("irregularSeal1", &HashMap::new(), aspect);
        assert_points(&coordinates(&path), &expected);
        assert!(matches!(path.last(), Some(GeometryPathCommand::Close)));
    }
}

#[test]
fn remaining_line_callouts_follow_ecma_paths() {
    use ooxml_drawingml::PresetPathFill;
    for (shape, segments, border, accent) in [
        ("accentCallout3", 3, false, true),
        ("accentBorderCallout1", 1, true, true),
        ("accentBorderCallout2", 2, true, true),
        ("accentBorderCallout3", 3, true, true),
        ("borderCallout2", 2, true, false),
        ("borderCallout3", 3, true, false),
        ("callout2", 2, false, false),
        ("callout3", 3, false, false),
    ] {
        let expected = match segments {
            1 => vec![(-0.08333, 0.1875), (-0.38333, 1.125)],
            2 => vec![(-0.08333, 0.1875), (-0.16667, 0.1875), (-0.46667, 1.125)],
            _ => vec![
                (-0.08333, 0.1875),
                (-0.16667, 0.1875),
                (-0.16667, 1.0),
                (-0.08333, 1.12963),
            ],
        };
        let defaults: HashMap<_, _> = expected
            .iter()
            .enumerate()
            .flat_map(|(i, &(x, y))| {
                [
                    (format!("adj{}", 2 * i + 1), y),
                    (format!("adj{}", 2 * i + 2), x),
                ]
            })
            .collect();
        assert_eq!(
            preset_geometry_default_adjustments(shape),
            defaults,
            "{shape}"
        );
        for aspect in [2.0, 0.5] {
            let layers = preset_geometry_layers(shape, &HashMap::new(), aspect).unwrap();
            assert_eq!(layers.len(), 2, "{shape}");
            assert_eq!(layers[0].fill, PresetPathFill::Normal);
            assert_eq!(layers[0].stroke, border);
            assert_points(
                &vertices(&layers[0].commands),
                &[(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)],
            );
            assert_eq!(layers[1].fill, PresetPathFill::None);
            assert!(layers[1].stroke);
            let mut leader = Vec::new();
            if accent {
                leader.extend([(-0.08333, 0.0), (-0.08333, 1.0)]);
                assert_eq!(layers[1].commands[1], GeometryPathCommand::Close);
            }
            leader.extend(&expected);
            assert_points(&vertices(&layers[1].commands), &leader);
            let custom = named(&[
                ("adj1", 0.3),
                ("adj2", -0.2),
                ("adj3", 0.4),
                ("adj4", -0.4),
                ("adj5", 0.7),
                ("adj6", -0.6),
                ("adj7", 1.2),
                ("adj8", -0.8),
            ]);
            let layers = preset_geometry_layers(shape, &custom, aspect).unwrap();
            let mut leader = Vec::new();
            if accent {
                leader.extend([(-0.2, 0.0), (-0.2, 1.0)]);
            }
            leader.extend(&[(-0.2, 0.3), (-0.4, 0.4), (-0.6, 0.7), (-0.8, 1.2)][..segments + 1]);
            assert_points(&vertices(&layers[1].commands), &leader);
        }
    }
}

#[test]
fn remaining_arrow_callouts_follow_ecma_vertices() {
    for (shape, expected) in [
        (
            "leftArrowCallout",
            &[
                (0.0, 0.5),
                (0.125, 0.25),
                (0.125, 0.375),
                (0.35023, 0.375),
                (0.35023, 0.0),
                (1.0, 0.0),
                (1.0, 1.0),
                (0.35023, 1.0),
                (0.35023, 0.625),
                (0.125, 0.625),
                (0.125, 0.75),
            ][..],
        ),
        (
            "rightArrowCallout",
            &[
                (0.0, 0.0),
                (0.64977, 0.0),
                (0.64977, 0.375),
                (0.875, 0.375),
                (0.875, 0.25),
                (1.0, 0.5),
                (0.875, 0.75),
                (0.875, 0.625),
                (0.64977, 0.625),
                (0.64977, 1.0),
                (0.0, 1.0),
            ][..],
        ),
        (
            "upDownArrowCallout",
            &[
                (0.0, 0.259385),
                (0.4375, 0.259385),
                (0.4375, 0.25),
                (0.375, 0.25),
                (0.5, 0.0),
                (0.625, 0.25),
                (0.5625, 0.25),
                (0.5625, 0.259385),
                (1.0, 0.259385),
                (1.0, 0.740615),
                (0.5625, 0.740615),
                (0.5625, 0.75),
                (0.625, 0.75),
                (0.5, 1.0),
                (0.375, 0.75),
                (0.4375, 0.75),
                (0.4375, 0.740615),
                (0.0, 0.740615),
            ][..],
        ),
        (
            "leftRightArrowCallout",
            &[
                (0.0, 0.5),
                (0.125, 0.25),
                (0.125, 0.375),
                (0.259385, 0.375),
                (0.259385, 0.0),
                (0.740615, 0.0),
                (0.740615, 0.375),
                (0.875, 0.375),
                (0.875, 0.25),
                (1.0, 0.5),
                (0.875, 0.75),
                (0.875, 0.625),
                (0.740615, 0.625),
                (0.740615, 1.0),
                (0.259385, 1.0),
                (0.259385, 0.625),
                (0.125, 0.625),
                (0.125, 0.75),
            ][..],
        ),
        (
            "quadArrowCallout",
            &[
                (0.0, 0.5),
                (0.092575, 0.31485),
                (0.092575, 0.407425),
                (0.259385, 0.407425),
                (0.259385, 0.259385),
                (0.4537125, 0.259385),
                (0.4537125, 0.18515),
                (0.407425, 0.18515),
                (0.5, 0.0),
                (0.592575, 0.18515),
                (0.5462875, 0.18515),
                (0.5462875, 0.259385),
                (0.740615, 0.259385),
                (0.740615, 0.407425),
                (0.907425, 0.407425),
                (0.907425, 0.31485),
                (1.0, 0.5),
                (0.907425, 0.68515),
                (0.907425, 0.592575),
                (0.740615, 0.592575),
                (0.740615, 0.740615),
                (0.5462875, 0.740615),
                (0.5462875, 0.81485),
                (0.592575, 0.81485),
                (0.5, 1.0),
                (0.407425, 0.81485),
                (0.4537125, 0.81485),
                (0.4537125, 0.740615),
                (0.259385, 0.740615),
                (0.259385, 0.592575),
                (0.092575, 0.592575),
                (0.092575, 0.68515),
            ][..],
        ),
    ] {
        let defaults = if shape == "quadArrowCallout" {
            [0.18515, 0.18515, 0.18515, 0.48123]
        } else if matches!(shape, "leftArrowCallout" | "rightArrowCallout") {
            [0.25, 0.25, 0.25, 0.64977]
        } else {
            [0.25, 0.25, 0.25, 0.48123]
        };
        assert_eq!(
            preset_geometry_default_adjustments(shape),
            named(&[
                ("adj1", defaults[0]),
                ("adj2", defaults[1]),
                ("adj3", defaults[2]),
                ("adj4", defaults[3])
            ])
        );
        assert_points(&vertices(&draw(shape, &HashMap::new(), 2.0)), expected);
    }
}

#[test]
fn arrow_callouts_apply_ecma_adjustment_pins() {
    let over = named(&[("adj1", 9.0), ("adj2", 9.0), ("adj3", 9.0), ("adj4", 9.0)]);
    assert_points(
        &vertices(&draw("leftArrowCallout", &over, 2.0)),
        &[
            (0.0, 0.5),
            (1.0, 0.0),
            (1.0, 0.0),
            (1.0, 0.0),
            (1.0, 0.0),
            (1.0, 0.0),
            (1.0, 1.0),
            (1.0, 1.0),
            (1.0, 1.0),
            (1.0, 1.0),
            (1.0, 1.0),
        ],
    );
    assert_points(
        &vertices(&draw("rightArrowCallout", &over, 2.0)),
        &[
            (0.0, 0.0),
            (0.0, 0.0),
            (0.0, 0.0),
            (0.0, 0.0),
            (0.0, 0.0),
            (1.0, 0.5),
            (0.0, 1.0),
            (0.0, 1.0),
            (0.0, 1.0),
            (0.0, 1.0),
            (0.0, 1.0),
        ],
    );
    let quad = draw(
        "quadArrowCallout",
        &named(&[("adj1", 0.2), ("adj2", 0.25), ("adj3", 0.2), ("adj4", -0.4)]),
        2.0,
    );
    let v = vertices(&quad);
    assert!(near(v[3], (0.4, 0.4)));
    assert!(near(v[4], (0.4, 0.4)));
    assert!(near(v[20], (0.6, 0.6)));
    let quad = draw(
        "quadArrowCallout",
        &named(&[("adj1", 0.2), ("adj2", 0.25), ("adj3", 9.0), ("adj4", 9.0)]),
        2.0,
    );
    let v = vertices(&quad);
    assert!(near(v[1], (0.125, 0.25)));
    assert!(near(v[4], (0.25, 0.25)));
}

#[test]
fn action_buttons_follow_ecma_icon_coordinates_and_layers() {
    use ooxml_drawingml::PresetPathFill;
    let rectangle = [(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)];
    assert_points(
        &vertices(&draw("actionButtonBlank", &HashMap::new(), 2.0)),
        &rectangle,
    );
    for (shape, first, fills) in [
        (
            "actionButtonHome",
            (0.6171875, 0.359375),
            &[
                PresetPathFill::Normal,
                PresetPathFill::DarkenLess,
                PresetPathFill::Darken,
                PresetPathFill::None,
            ][..],
        ),
        (
            "actionButtonHelp",
            (11.0 / 28.0, 19.0 / 56.0),
            &[
                PresetPathFill::Normal,
                PresetPathFill::Darken,
                PresetPathFill::None,
            ][..],
        ),
        (
            "actionButtonInformation",
            (0.5, 0.125),
            &[
                PresetPathFill::Normal,
                PresetPathFill::Darken,
                PresetPathFill::Lighten,
                PresetPathFill::None,
            ][..],
        ),
        (
            "actionButtonForwardNext",
            (0.6875, 0.5),
            &[
                PresetPathFill::Normal,
                PresetPathFill::Darken,
                PresetPathFill::None,
            ][..],
        ),
        (
            "actionButtonBackPrevious",
            (0.3125, 0.5),
            &[
                PresetPathFill::Normal,
                PresetPathFill::Darken,
                PresetPathFill::None,
            ][..],
        ),
        (
            "actionButtonEnd",
            (0.59375, 0.5),
            &[
                PresetPathFill::Normal,
                PresetPathFill::Darken,
                PresetPathFill::None,
            ][..],
        ),
        (
            "actionButtonBeginning",
            (0.40625, 0.5),
            &[
                PresetPathFill::Normal,
                PresetPathFill::Darken,
                PresetPathFill::None,
            ][..],
        ),
        (
            "actionButtonReturn",
            (0.6875, 0.3125),
            &[
                PresetPathFill::Normal,
                PresetPathFill::Darken,
                PresetPathFill::None,
            ][..],
        ),
        (
            "actionButtonSound",
            (0.3125, 0.359375),
            &[
                PresetPathFill::Normal,
                PresetPathFill::Darken,
                PresetPathFill::None,
            ][..],
        ),
        (
            "actionButtonDocument",
            (0.359375, 0.125),
            &[
                PresetPathFill::Normal,
                PresetPathFill::DarkenLess,
                PresetPathFill::Darken,
                PresetPathFill::None,
            ][..],
        ),
        (
            "actionButtonMovie",
            (0.3125, 37.0 / 120.0),
            &[
                PresetPathFill::Normal,
                PresetPathFill::Darken,
                PresetPathFill::None,
            ][..],
        ),
    ] {
        assert!(preset_geometry_default_adjustments(shape).is_empty());
        let layers = preset_geometry_layers(shape, &HashMap::new(), 2.0).unwrap();
        assert_eq!(layers.len(), fills.len(), "{shape}");
        for (layer, fill) in layers.iter().zip(fills) {
            assert_eq!(layer.fill, *fill, "{shape}");
            assert_eq!(layer.stroke, layer.fill == PresetPathFill::None, "{shape}");
        }
        assert_points(&vertices(&layers[0].commands)[..4], &rectangle);
        assert!(near(endpoints(&layers[1].commands)[0], first), "{shape}");
        let outline = subpaths(&layers.last().unwrap().commands);
        assert_points(&vertices(outline.last().unwrap()), &rectangle);
    }
}

#[test]
fn action_button_back_previous_follows_ecma_endpoints() {
    let path0 = [(0.3125, 0.5), (0.6875, 0.125), (0.6875, 0.875)];
    assert_action_button_endpoints("actionButtonBackPrevious", 2.0, &[&path0, &path0, &path0]);
}

#[test]
fn action_button_beginning_follows_ecma_endpoints() {
    let path0 = [
        (0.40625, 0.5),
        (0.6875, 0.125),
        (0.6875, 0.875),
        (0.359375, 0.125),
        (0.3125, 0.125),
        (0.3125, 0.875),
        (0.359375, 0.875),
    ];
    let path1 = [
        (0.40625, 0.5),
        (0.6875, 0.125),
        (0.6875, 0.875),
        (0.359375, 0.125),
        (0.359375, 0.875),
        (0.3125, 0.875),
        (0.3125, 0.125),
    ];
    assert_action_button_endpoints("actionButtonBeginning", 2.0, &[&path0, &path0, &path1]);
}

#[test]
fn action_button_document_follows_ecma_endpoints() {
    let path0 = [
        (0.359375, 0.125),
        (0.546875, 0.125),
        (0.640625, 0.3125),
        (0.640625, 0.875),
        (0.359375, 0.875),
    ];
    let path1 = [
        (0.359375, 0.125),
        (0.546875, 0.125),
        (0.546875, 0.3125),
        (0.640625, 0.3125),
        (0.640625, 0.875),
        (0.359375, 0.875),
    ];
    let path2 = [(0.546875, 0.125), (0.546875, 0.3125), (0.640625, 0.3125)];
    let path3 = [
        (0.359375, 0.125),
        (0.546875, 0.125),
        (0.640625, 0.3125),
        (0.640625, 0.875),
        (0.359375, 0.875),
        (0.640625, 0.3125),
        (0.546875, 0.3125),
        (0.546875, 0.125),
    ];
    assert_action_button_endpoints(
        "actionButtonDocument",
        2.0,
        &[&path0, &path1, &path2, &path3],
    );
}

#[test]
fn action_button_end_follows_ecma_endpoints() {
    let path0 = [
        (0.59375, 0.5),
        (0.3125, 0.125),
        (0.3125, 0.875),
        (0.640625, 0.125),
        (0.6875, 0.125),
        (0.6875, 0.875),
        (0.640625, 0.875),
    ];
    let path1 = [
        (0.59375, 0.5),
        (0.3125, 0.875),
        (0.3125, 0.125),
        (0.640625, 0.125),
        (0.6875, 0.125),
        (0.6875, 0.875),
        (0.640625, 0.875),
    ];
    assert_action_button_endpoints("actionButtonEnd", 2.0, &[&path0, &path0, &path1]);
}

#[test]
fn action_button_forward_next_follows_ecma_endpoints() {
    let path0 = [(0.6875, 0.5), (0.3125, 0.125), (0.3125, 0.875)];
    let path1 = [(0.6875, 0.5), (0.3125, 0.875), (0.3125, 0.125)];
    assert_action_button_endpoints("actionButtonForwardNext", 2.0, &[&path0, &path0, &path1]);
}

#[test]
fn action_button_help_follows_ecma_endpoints() {
    let path0 = [
        (11.0 / 28.0, 19.0 / 56.0),
        (17.0 / 28.0, 19.0 / 56.0),
        (31.0 / 56.0, 0.5),
        (59.0 / 112.0, 65.0 / 112.0),
        (59.0 / 112.0, 0.6875),
        (53.0 / 112.0, 0.6875),
        (53.0 / 112.0, 65.0 / 112.0),
        (59.0 / 112.0, 47.0 / 112.0),
        (31.0 / 56.0, 19.0 / 56.0),
        (25.0 / 56.0, 19.0 / 56.0),
        (0.5, 5.0 / 7.0),
        (0.5, 5.0 / 7.0),
    ];
    assert_action_button_endpoints("actionButtonHelp", 2.0, &[&path0, &path0, &path0]);
}

#[test]
fn action_button_home_follows_ecma_endpoints() {
    let path0 = [
        (0.5, 0.125),
        (0.3125, 0.5),
        (0.359375, 0.5),
        (0.359375, 0.875),
        (0.640625, 0.875),
        (0.640625, 0.5),
        (0.6875, 0.5),
        (0.6171875, 0.359375),
        (0.6171875, 0.171875),
        (0.5703125, 0.171875),
        (0.5703125, 0.265625),
    ];
    let path1 = [
        (0.6171875, 0.359375),
        (0.6171875, 0.171875),
        (0.5703125, 0.171875),
        (0.5703125, 0.265625),
        (0.359375, 0.5),
        (0.359375, 0.875),
        (0.4765625, 0.875),
        (0.4765625, 0.6875),
        (0.5234375, 0.6875),
        (0.5234375, 0.875),
        (0.640625, 0.875),
        (0.640625, 0.5),
    ];
    let path2 = [
        (0.5, 0.125),
        (0.3125, 0.5),
        (0.6875, 0.5),
        (0.4765625, 0.6875),
        (0.5234375, 0.6875),
        (0.5234375, 0.875),
        (0.4765625, 0.875),
    ];
    let path3 = [
        (0.5, 0.125),
        (0.5703125, 0.265625),
        (0.5703125, 0.171875),
        (0.6171875, 0.171875),
        (0.6171875, 0.359375),
        (0.6875, 0.5),
        (0.640625, 0.5),
        (0.640625, 0.875),
        (0.359375, 0.875),
        (0.359375, 0.5),
        (0.3125, 0.5),
        (0.5703125, 0.265625),
        (0.6171875, 0.359375),
        (0.640625, 0.5),
        (0.359375, 0.5),
        (0.4765625, 0.875),
        (0.4765625, 0.6875),
        (0.5234375, 0.6875),
        (0.5234375, 0.875),
    ];
    assert_action_button_endpoints("actionButtonHome", 2.0, &[&path0, &path1, &path2, &path3]);
}

#[test]
fn action_button_information_follows_ecma_endpoints() {
    let path0 = [(0.5, 0.125), (0.5, 0.125)];
    let path1 = [
        (0.5, 0.125),
        (0.5, 0.125),
        (0.5, 0.1484375),
        (0.5, 0.1484375),
        (0.4296875, 0.359375),
        (0.4296875, 0.40625),
        (0.46484375, 0.40625),
        (0.46484375, 0.734375),
        (0.4296875, 0.734375),
        (0.4296875, 0.78125),
        (0.5703125, 0.78125),
        (0.5703125, 0.734375),
        (0.53515625, 0.734375),
        (0.53515625, 0.359375),
    ];
    let path2 = [
        (0.5, 0.1484375),
        (0.5, 0.1484375),
        (0.4296875, 0.359375),
        (0.53515625, 0.359375),
        (0.53515625, 0.734375),
        (0.5703125, 0.734375),
        (0.5703125, 0.78125),
        (0.4296875, 0.78125),
        (0.4296875, 0.734375),
        (0.46484375, 0.734375),
        (0.46484375, 0.40625),
        (0.4296875, 0.40625),
    ];
    let path3 = [
        (0.5, 0.125),
        (0.5, 0.125),
        (0.5, 0.1484375),
        (0.5, 0.1484375),
        (0.4296875, 0.359375),
        (0.53515625, 0.359375),
        (0.53515625, 0.734375),
        (0.5703125, 0.734375),
        (0.5703125, 0.78125),
        (0.4296875, 0.78125),
        (0.4296875, 0.734375),
        (0.46484375, 0.734375),
        (0.46484375, 0.40625),
        (0.4296875, 0.40625),
    ];
    assert_action_button_endpoints(
        "actionButtonInformation",
        2.0,
        &[&path0, &path1, &path2, &path3],
    );
}

#[test]
fn action_button_movie_follows_ecma_endpoints() {
    let path0 = [
        (0.3125, 37.0 / 120.0),
        (0.3125, 877.0 / 1920.0),
        (1297.0 / 3840.0, 877.0 / 1920.0),
        (1327.0 / 3840.0, 12667.0 / 28800.0),
        (271.0 / 768.0, 12667.0 / 28800.0),
        (271.0 / 768.0, 2399.0 / 3600.0),
        (0.6078125, 2399.0 / 3600.0),
        (0.6078125, 8471.0 / 14400.0),
        (2489.0 / 3840.0, 8471.0 / 14400.0),
        (2573.0 / 3840.0, 0.63125),
        (0.6875, 0.63125),
        (0.6875, 341.0 / 960.0),
        (2573.0 / 3840.0, 341.0 / 960.0),
        (503.0 / 768.0, 2773.0 / 7200.0),
        (0.6078125, 2773.0 / 7200.0),
        (0.6078125, 341.0 / 960.0),
        (0.59296875, 311.0 / 960.0),
        (1327.0 / 3840.0, 311.0 / 960.0),
        (1297.0 / 3840.0, 37.0 / 120.0),
    ];
    let path1 = [
        (0.3125, 37.0 / 120.0),
        (1297.0 / 3840.0, 37.0 / 120.0),
        (1327.0 / 3840.0, 311.0 / 960.0),
        (0.59296875, 311.0 / 960.0),
        (0.6078125, 341.0 / 960.0),
        (0.6078125, 2773.0 / 7200.0),
        (503.0 / 768.0, 2773.0 / 7200.0),
        (2573.0 / 3840.0, 341.0 / 960.0),
        (0.6875, 341.0 / 960.0),
        (0.6875, 0.63125),
        (2573.0 / 3840.0, 0.63125),
        (2489.0 / 3840.0, 8471.0 / 14400.0),
        (0.6078125, 8471.0 / 14400.0),
        (0.6078125, 2399.0 / 3600.0),
        (271.0 / 768.0, 2399.0 / 3600.0),
        (271.0 / 768.0, 12667.0 / 28800.0),
        (1327.0 / 3840.0, 12667.0 / 28800.0),
        (1297.0 / 3840.0, 877.0 / 1920.0),
        (0.3125, 877.0 / 1920.0),
    ];
    assert_action_button_endpoints("actionButtonMovie", 2.0, &[&path0, &path0, &path1]);
}

#[test]
fn action_button_return_follows_ecma_endpoints() {
    let path0 = [
        (0.6875, 0.3125),
        (0.59375, 0.125),
        (0.5, 0.3125),
        (0.546875, 0.3125),
        (0.546875, 0.59375),
        (0.5, 0.6875),
        (0.453125, 0.6875),
        (0.40625, 0.59375),
        (0.40625, 0.3125),
        (0.3125, 0.3125),
        (0.3125, 0.59375),
        (0.453125, 0.875),
        (0.5, 0.875),
        (0.640625, 0.59375),
        (0.640625, 0.3125),
    ];
    let path1 = [
        (0.6875, 0.3125),
        (0.640625, 0.3125),
        (0.640625, 0.59375),
        (0.5, 0.875),
        (0.453125, 0.875),
        (0.3125, 0.59375),
        (0.3125, 0.3125),
        (0.40625, 0.3125),
        (0.40625, 0.59375),
        (0.453125, 0.6875),
        (0.5, 0.6875),
        (0.546875, 0.59375),
        (0.546875, 0.3125),
        (0.5, 0.3125),
        (0.59375, 0.125),
    ];
    assert_action_button_endpoints("actionButtonReturn", 2.0, &[&path0, &path0, &path1]);
}

#[test]
fn action_button_sound_follows_ecma_endpoints() {
    let path0 = [
        (0.3125, 0.359375),
        (0.3125, 0.640625),
        (0.4296875, 0.640625),
        (0.546875, 0.875),
        (0.546875, 0.125),
        (0.4296875, 0.359375),
    ];
    let path1 = [
        (0.3125, 0.359375),
        (0.4296875, 0.359375),
        (0.546875, 0.125),
        (0.546875, 0.875),
        (0.4296875, 0.640625),
        (0.3125, 0.640625),
        (0.59375, 0.359375),
        (0.6875, 0.21875),
        (0.59375, 0.5),
        (0.6875, 0.5),
        (0.59375, 0.640625),
        (0.6875, 0.78125),
    ];
    assert_action_button_endpoints("actionButtonSound", 2.0, &[&path0, &path0, &path1]);
}

fn assert_action_button_endpoints(shape: &str, aspect: f64, expected: &[&[(f64, f64)]]) {
    let layers = preset_geometry_layers(shape, &HashMap::new(), aspect).unwrap();
    let rectangle = [(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)];
    assert_eq!(layers.len(), expected.len(), "{shape}");
    for (index, (layer, expected)) in layers.iter().zip(expected).enumerate() {
        let mut points = Vec::new();
        if index == 0 {
            points.extend(rectangle);
        }
        points.extend_from_slice(expected);
        if index == layers.len() - 1 {
            points.extend(rectangle);
        }
        assert_ecma_endpoints(&endpoints(&layer.commands), &points);
    }
}

#[test]
fn action_button_home_follows_ecma_endpoints_at_tall_aspect() {
    let layers = preset_geometry_layers("actionButtonHome", &HashMap::new(), 0.5).unwrap();
    assert_points(
        &endpoints(&layers[1].commands),
        &[
            (0.734375, 0.4296875),
            (0.734375, 0.3359375),
            (0.640625, 0.3359375),
            (0.640625, 0.3828125),
            (0.21875, 0.5),
            (0.21875, 0.6875),
            (0.453125, 0.6875),
            (0.453125, 0.59375),
            (0.546875, 0.59375),
            (0.546875, 0.6875),
            (0.78125, 0.6875),
            (0.78125, 0.5),
        ],
    );
}

#[test]
fn up_down_arrow_callout_applies_ecma_adjustment_pins() {
    let over = named(&[("adj1", 9.0), ("adj2", 9.0), ("adj3", 9.0), ("adj4", 9.0)]);
    assert_points(
        &vertices(&draw("upDownArrowCallout", &over, 2.0)),
        &[
            (0.0, 0.5),
            (0.0, 0.5),
            (0.0, 0.5),
            (0.0, 0.5),
            (0.5, 0.0),
            (1.0, 0.5),
            (1.0, 0.5),
            (1.0, 0.5),
            (1.0, 0.5),
            (1.0, 0.5),
            (1.0, 0.5),
            (1.0, 0.5),
            (1.0, 0.5),
            (0.5, 1.0),
            (0.0, 0.5),
            (0.0, 0.5),
            (0.0, 0.5),
            (0.0, 0.5),
        ],
    );
}

#[test]
fn left_right_arrow_callout_applies_ecma_adjustment_pins() {
    let over = named(&[("adj1", 9.0), ("adj2", 9.0), ("adj3", 9.0), ("adj4", 9.0)]);
    assert_points(
        &vertices(&draw("leftRightArrowCallout", &over, 2.0)),
        &[
            (0.0, 0.5),
            (0.5, 0.0),
            (0.5, 0.0),
            (0.5, 0.0),
            (0.5, 0.0),
            (0.5, 0.0),
            (0.5, 0.0),
            (0.5, 0.0),
            (0.5, 0.0),
            (1.0, 0.5),
            (0.5, 1.0),
            (0.5, 1.0),
            (0.5, 1.0),
            (0.5, 1.0),
            (0.5, 1.0),
            (0.5, 1.0),
            (0.5, 1.0),
            (0.5, 1.0),
        ],
    );
}

proptest! {
    #[test]
    fn line_callouts_keep_all_eight_adjustments_finite(
        offsets in prop::array::uniform8(any_adjust()),
        aspect in any_aspect(),
    ) {
        let adjustments = offsets.into_iter().enumerate()
            .map(|(i, value)| (format!("adj{}", i + 1), value))
            .collect();
        for shape in [
            "accentCallout3", "accentBorderCallout1", "accentBorderCallout2",
            "accentBorderCallout3", "borderCallout2", "borderCallout3", "callout2", "callout3",
        ] {
            let layers = preset_geometry_layers(shape, &adjustments, aspect).unwrap();
            for layer in &layers {
                for (x, y) in coordinates(&layer.commands) {
                    prop_assert!(x.is_finite() && y.is_finite(), "{shape} emitted ({x}, {y})");
                }
            }
            for (x, y) in coordinates(&layers[0].commands) {
                prop_assert!((-TOLERANCE..=1.0 + TOLERANCE).contains(&x)
                    && (-TOLERANCE..=1.0 + TOLERANCE).contains(&y), "{shape} body left its frame");
            }
        }
    }
}
