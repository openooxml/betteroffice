use std::collections::HashMap;

use ooxml_drawingml::{GeometryPathCommand, preset_geometry_to_path};
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
];

/// Connectors whose ECMA-376 adjusts are unpinned, so they may route outside the frame.
const UNPINNED_CONNECTORS: &[&str] = &["bentConnector", "curvedConnector"];

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
    )
        .prop_map(|(adj, adj1, adj2)| {
            [("adj", adj), ("adj1", adj1), ("adj2", adj2)]
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
        (star, _) if star.starts_with("star") => distance(v[1], (0.5, 0.5)),
        _ => v[0].0,
    }
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
            .filter(|shape| !UNPINNED_CONNECTORS.iter().any(|prefix| shape.starts_with(prefix)));
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
