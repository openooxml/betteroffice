//! What a shape becomes before anything paints it: `kurbo` subdivides every
//! arc by its radius, and `tiny-skia`'s stroker splits every curve until its
//! outline meets a tolerance. Neither grows with a document's bytes, so both
//! are read here from the numbers the document writes, with the parsers `usvg`
//! reads them with.

use std::str::FromStr;

use resvg::usvg::roxmltree::Node;
use svgtypes::{Length, LengthListParser, LengthUnit, PathParser, PathSegment};

use super::SvgRefusal;
use super::audit::SVG_NS;

/// Cubics one arc may become. `svgtypes` buffers an arc's cubics and drains
/// them from the front, so an arc costs the square of its cubics in copies.
const MAX_ARC_CUBICS: f64 = 64.0;
/// The tolerance `svgtypes` and `usvg` hand `kurbo` to subdivide an arc.
const ARC_TOLERANCE: f64 = 0.1;
/// Bytes per piece the stroker emits: its points and verb, in the outer and
/// inner outlines it builds and the path it joins them into, at most 43
/// measured over a million pieces.
pub(super) const OUTLINE_VERB_BYTES: f64 = 48.0;
/// The tolerance `usvg` strokes a shape with to measure it: a quarter of a
/// user unit, at a resolution scale of one.
pub(super) const MEASURE_TOLERANCE: f64 = 0.25;
/// Outline pieces the stroker may emit for any path besides its segments, and
/// its setup, about half a microsecond a path.
const PATH_VERBS: f64 = 16.0;
/// Pieces a curve may add besides one per tolerance of its length: the spiral
/// of lines the stroker leaves around a cusp as it splits down to the float
/// precision of `t`, and the circle it adds there.
const CURVE_VERBS: f64 = 64.0;

/// A shape's outline as the arc converter and the stroker see it.
#[derive(Clone, Copy, Default)]
pub(super) struct Outline {
    pub(super) segments: f64,
    pub(super) contours: f64,
    pub(super) curves: f64,
    /// Control-polygon length of the curves.
    pub(super) length: f64,
    /// How far from the origin any curve's points lie.
    pub(super) reach: f64,
    /// Cubics its arcs become.
    pub(super) arc_cubics: f64,
}

impl Outline {
    /// Adds a curve through its control points.
    pub(super) fn curve(&mut self, points: &[(f64, f64)]) {
        self.segments += 1.0;
        self.curves += 1.0;
        for pair in points.windows(2) {
            self.length += distance(pair[0], pair[1]);
        }
        for point in points {
            self.reach = self.reach.max(magnitude(*point));
        }
    }

    /// Adds an arc of up to `radius` about a centre within `radius` of
    /// `middle`, cut into `cubics`.
    fn arc(&mut self, cubics: f64, radius: f64, middle: f64) {
        self.segments += cubics;
        self.curves += cubics;
        self.arc_cubics += cubics;
        self.length += 8.0 * radius;
        self.reach = self.reach.max(middle + 2.0 * radius);
    }
}

/// Pieces `tiny-skia`'s stroker may emit for `outline` stroked at `radius`
/// with `tolerance`, all in the units it strokes in, while its curves stay
/// within [`super::MAX_SVG_STROKE_SPAN`] tolerances of the origin. Round
/// joins, caps and cusp circles are conics of up to 32 quads, and a curve
/// splits until its pieces meet the tolerance: about a piece per tolerance of
/// control polygon at worst, six times what a search over adversarial cubics
/// reached.
pub(super) fn stroke_verbs(outline: &Outline, radius: f64, tolerance: f64) -> f64 {
    let quads = (1.0 + (radius / tolerance).sqrt()).min(32.0);
    PATH_VERBS
        + outline.segments * (4.0 + 2.0 * quads)
        + outline.contours * 4.0 * quads
        + outline.curves * (CURVE_VERBS + 4.0 * quads)
        + outline.length / tolerance
}

/// Path data as `svgtypes` simplifies it for `usvg`, which stops at the first
/// error. Positions are tracked as it tracks them; after an arc it resumes
/// from `kurbo`'s last point, so later positions carry that drift.
pub(super) fn path(data: &str) -> Result<Outline, SvgRefusal> {
    let mut outline = Outline::default();
    let (mut at, mut start) = ((0.0, 0.0), (0.0, 0.0));
    let mut cubic: Option<(f64, f64)> = None;
    let mut quad: Option<(f64, f64)> = None;
    let mut drift = 0.0f64;
    let mut closed = false;
    for segment in PathParser::from(data) {
        let Ok(segment) = segment else {
            break;
        };
        let origin = if segment.is_abs() { (0.0, 0.0) } else { at };
        let point = |x: f64, y: f64| (x + origin.0, y + origin.1);
        if closed && !matches!(segment, PathSegment::MoveTo { .. }) {
            outline.contours += 1.0;
        }
        closed = false;
        let (mut next_cubic, mut next_quad) = (None, None);
        match segment {
            PathSegment::MoveTo { x, y, .. } => {
                at = point(x, y);
                start = at;
                outline.contours += 1.0;
            }
            PathSegment::LineTo { x, y, .. } => {
                at = point(x, y);
                outline.segments += 1.0;
            }
            PathSegment::HorizontalLineTo { x, .. } => {
                at.0 = point(x, 0.0).0;
                outline.segments += 1.0;
            }
            PathSegment::VerticalLineTo { y, .. } => {
                at.1 = point(0.0, y).1;
                outline.segments += 1.0;
            }
            PathSegment::CurveTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
                ..
            } => {
                let (first, second, end) = (point(x1, y1), point(x2, y2), point(x, y));
                outline.curve(&[at, first, second, end]);
                (at, next_cubic) = (end, Some(second));
            }
            PathSegment::SmoothCurveTo { x2, y2, x, y, .. } => {
                let first = cubic.map_or(at, |control| reflect(at, control));
                let (second, end) = (point(x2, y2), point(x, y));
                outline.curve(&[at, first, second, end]);
                (at, next_cubic) = (end, Some(second));
            }
            PathSegment::Quadratic { x1, y1, x, y, .. } => {
                let (control, end) = (point(x1, y1), point(x, y));
                outline.curve(&[at, control, end]);
                (at, next_quad) = (end, Some(control));
            }
            PathSegment::SmoothQuadratic { x, y, .. } => {
                let control = quad.map_or(at, |control| reflect(at, control));
                let end = point(x, y);
                outline.curve(&[at, control, end]);
                (at, next_quad) = (end, Some(control));
            }
            PathSegment::EllipticalArc {
                rx,
                ry,
                large_arc,
                x,
                y,
                ..
            } => {
                let end = point(x, y);
                let (rx, ry) = (rx.abs(), ry.abs());
                if rx <= 1e-5 || ry <= 1e-5 {
                    outline.segments += 1.0;
                } else {
                    let half = distance(at, end) / 2.0 + drift;
                    let radius = rx.max(ry) * (half / rx.min(ry)).max(1.0);
                    let cubics = arc_cubics(radius, large_arc)?;
                    let middle = magnitude(((at.0 + end.0) / 2.0, (at.1 + end.1) / 2.0));
                    outline.arc(cubics, radius, middle + drift);
                    drift += 1e-6 * (radius + magnitude(at) + magnitude(end));
                }
                at = end;
            }
            PathSegment::ClosePath { .. } => {
                outline.segments += 1.0;
                at = start;
                closed = true;
            }
        }
        (cubic, quad) = (next_cubic, next_quad);
    }
    outline.reach += drift;
    if !outline.reach.is_finite() || !outline.length.is_finite() {
        return Err(SvgRefusal::ExpansionTooLarge);
    }
    Ok(outline)
}

/// The outline `usvg` builds for a circle, ellipse or rectangle, from its
/// geometry attributes. A shape with curves must give them in absolute
/// units: an `em` or `%` resolves against context the audit does not track.
pub(super) fn shape(node: Node<'_, '_>) -> Result<Outline, SvgRefusal> {
    let mut outline = Outline::default();
    match node.tag_name().name() {
        "circle" => {
            let radius = absolute(node, "r")?.unwrap_or(0.0);
            if radius > 0.0 {
                let centre = absolute(node, "cx")?.unwrap_or(0.0).abs()
                    + absolute(node, "cy")?.unwrap_or(0.0).abs();
                ellipse(&mut outline, radius, radius, centre)?;
            }
        }
        "ellipse" => {
            let (rx, ry) = radii(node)?;
            if rx > 0.0 && ry > 0.0 {
                let centre = absolute(node, "cx")?.unwrap_or(0.0).abs()
                    + absolute(node, "cy")?.unwrap_or(0.0).abs();
                ellipse(&mut outline, rx, ry, centre)?;
            }
        }
        "rect" => {
            let (rx, ry) = radii(node)?;
            outline.segments = 5.0;
            outline.contours = 1.0;
            if rx > 0.0 || ry > 0.0 {
                let corner = ["x", "y", "width", "height"]
                    .iter()
                    .map(|name| absolute(node, name).map(|value| value.unwrap_or(0.0).abs()))
                    .sum::<Result<f64, _>>()?;
                ellipse(&mut outline, rx.max(ry), rx.max(ry), corner)?;
            }
        }
        "line" => {
            outline.segments = 1.0;
            outline.contours = 1.0;
        }
        "polyline" | "polygon" => {
            let points = attributes(node, "points").map(str::len).sum::<usize>();
            outline.segments = points as f64 / 2.0 + 1.0;
            outline.contours = 1.0;
        }
        _ => {}
    }
    Ok(outline)
}

/// Four quarter arcs of an ellipse `usvg` joins end to end. Their ends are
/// rounded to `f32` and each starts where `kurbo` left the last, so `kurbo`
/// may widen the radii to reach them: never past three times the larger.
fn ellipse(outline: &mut Outline, rx: f64, ry: f64, centre: f64) -> Result<(), SvgRefusal> {
    let radius = rx.max(ry);
    let cubics = 4.0 * arc_cubics(3.0 * radius, false)?;
    outline.arc(cubics, radius, centre);
    outline.segments += 1.0;
    outline.contours = outline.contours.max(1.0);
    Ok(())
}

/// Cubics `kurbo` cuts an arc into, at most: it picks a count per full turn
/// from the corrected radius, and an arc that is not large sweeps at most half
/// a turn. Refuses an arc past [`MAX_ARC_CUBICS`].
fn arc_cubics(radius: f64, large: bool) -> Result<f64, SvgRefusal> {
    let per_turn = (1.1163 * radius / ARC_TOLERANCE).powf(1.0 / 6.0);
    let share = if large { 1.0 } else { 0.5 };
    let cubics = (per_turn.max(3.999_999) * share * (1.0 + 1e-7)).ceil();
    if !radius.is_finite() || cubics.is_nan() || cubics > MAX_ARC_CUBICS {
        return Err(SvgRefusal::ExpansionTooLarge);
    }
    Ok(cubics)
}

/// `rx` and `ry` as `usvg` resolves them for a rectangle or an ellipse: a
/// negative one is dropped, and a missing one takes the other's value.
fn radii(node: Node<'_, '_>) -> Result<(f64, f64), SvgRefusal> {
    let positive =
        |name| absolute(node, name).map(|value| value.filter(|value| !value.is_sign_negative()));
    Ok(match (positive("rx")?, positive("ry")?) {
        (None, None) => (0.0, 0.0),
        (Some(rx), None) => (rx, rx),
        (None, Some(ry)) => (ry, ry),
        (Some(rx), Some(ry)) => (rx, ry),
    })
}

/// The largest length `name` gives in user units, as `svgtypes::Length`
/// reads it for `usvg`; one it cannot parse is one `usvg` ignores. `em`, `ex`
/// and `%` are refused.
fn absolute(node: Node<'_, '_>, name: &str) -> Result<Option<f64>, SvgRefusal> {
    let mut largest: Option<f64> = None;
    for value in attributes(node, name) {
        let Ok(length) = Length::from_str(value) else {
            continue;
        };
        let value = length.number * unit(length.unit)?;
        largest = Some(largest.map_or(value, |largest| largest.max(value)));
    }
    Ok(largest)
}

/// A `stroke-width` in user units, as `svgtypes::Length` reads it for `usvg`;
/// zero for one `usvg` ignores. `em`, `ex` and `%` are refused.
pub(super) fn stroke_width(value: &str) -> Result<f64, SvgRefusal> {
    let Ok(length) = Length::from_str(value) else {
        return Ok(0.0);
    };
    Ok(length.number * unit(length.unit)?)
}

/// Bytes of a dash list, which `usvg` parses anew for every shape that
/// strokes with it. `em` and `ex` are refused: each entry in them walks up
/// the tree for a font size.
pub(super) fn dash_list(value: &str) -> Result<u64, SvgRefusal> {
    for length in LengthListParser::from(value).flatten() {
        if matches!(length.unit, LengthUnit::Em | LengthUnit::Ex) {
            return Err(SvgRefusal::ExpansionTooLarge);
        }
    }
    Ok(value.len() as u64)
}

/// User units per `unit` at `usvg`'s 96 dpi; `em`, `ex` and `%` are refused.
fn unit(unit: LengthUnit) -> Result<f64, SvgRefusal> {
    Ok(match unit {
        LengthUnit::None | LengthUnit::Px => 1.0,
        LengthUnit::In => 96.0,
        LengthUnit::Cm => 96.0 / 2.54,
        LengthUnit::Mm => 96.0 / 25.4,
        LengthUnit::Pt => 96.0 / 72.0,
        LengthUnit::Pc => 16.0,
        LengthUnit::Em | LengthUnit::Ex | LengthUnit::Percent => {
            return Err(SvgRefusal::ExpansionTooLarge);
        }
    })
}

/// Every value of `name` in no namespace or the SVG one, the two `usvg` reads.
pub(super) fn attributes<'a>(node: Node<'a, '_>, name: &'a str) -> impl Iterator<Item = &'a str> {
    node.attributes()
        .filter(move |attribute| {
            attribute.name() == name && matches!(attribute.namespace(), None | Some(SVG_NS))
        })
        .map(|attribute| attribute.value())
}

fn reflect(at: (f64, f64), control: (f64, f64)) -> (f64, f64) {
    (at.0 * 2.0 - control.0, at.1 * 2.0 - control.1)
}

fn distance(a: (f64, f64), b: (f64, f64)) -> f64 {
    (b.0 - a.0).hypot(b.1 - a.1)
}

fn magnitude(point: (f64, f64)) -> f64 {
    point.0.abs() + point.1.abs()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_small_arc_is_a_few_cubics_and_a_huge_one_is_refused() {
        let arcs = path("M0 0a1 1 0 0 0 1 1A5 5 0 1 1 0 0").unwrap();
        assert_eq!(arcs.arc_cubics, 2.0 + 4.0);
        let radius = format!("M0 0a{}1 1 0 0 0 1 1", "6".repeat(58));
        assert_eq!(path(&radius).err(), Some(SvgRefusal::ExpansionTooLarge));
    }

    #[test]
    fn radii_too_small_for_the_chord_are_widened_as_kurbo_widens_them() {
        let widened = format!("M0 0A1 1 0 0 0 {} 0", 1e12);
        assert_eq!(path(&widened).err(), Some(SvgRefusal::ExpansionTooLarge));
        let eccentric = "M0 0A0.00002 1e6 0 0 0 1000 0";
        assert_eq!(path(eccentric).err(), Some(SvgRefusal::ExpansionTooLarge));
        assert!(path("M0 0A0.000001 1e30 0 0 0 1 1").is_ok());
    }

    #[test]
    fn positions_follow_relative_commands_and_reflections() {
        let outline = path("M10 10l10 0c0 10 10 10 10 0s10-10 10 0zm5 5q5 5 10 0t10 0").unwrap();
        assert_eq!(outline.contours, 2.0);
        assert_eq!(outline.curves, 4.0);
        assert_eq!(outline.segments, 6.0);
        assert_eq!(outline.reach, 50.0);
    }

    #[test]
    fn a_shape_in_relative_units_is_refused_only_where_it_curves() {
        let document = |body: &str| format!("<svg xmlns='{SVG_NS}'>{body}</svg>");
        let outline = |body: &str| {
            let source = document(body);
            let document = resvg::usvg::roxmltree::Document::parse(&source).unwrap();
            shape(document.root_element().first_element_child().unwrap())
        };
        assert!(outline("<rect width='100%' height='50%'/>").is_ok());
        assert_eq!(
            outline("<rect width='100%' height='50%' rx='2'/>").err(),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert_eq!(
            outline("<circle r='1em'/>").err(),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert_eq!(
            outline("<circle r='1e30'/>").err(),
            Some(SvgRefusal::ExpansionTooLarge)
        );
        assert_eq!(
            outline("<ellipse rx='3' ry='-1'/>").unwrap().arc_cubics,
            8.0
        );
        assert_eq!(outline("<circle r='1in'/>").unwrap().arc_cubics, 8.0);
    }
}
