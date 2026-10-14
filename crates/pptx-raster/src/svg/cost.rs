//! What rendering a converted tree costs at one raster size, measured on the
//! tree `usvg` resolved (CSS applied, references expanded) before `resvg` runs.

use resvg::usvg::{self, Node, Paint};
use tiny_skia::{IntSize, PathSegment, PathStroker, Point, Rect, Transform};

use super::geometry::{self, OUTLINE_VERB_BYTES, Outline};
use super::{
    MAX_SVG_EXPANDED_NODES, MAX_SVG_LAYER_DEPTH, MAX_SVG_OVERDRAW, MAX_SVG_RASTER_DIM,
    MAX_SVG_RENDER_WORK, MAX_SVG_STROKE_SPAN, SVG_TRANSIENT_BYTES, SvgRefusal,
};

/// Work units, each about one painted pixel, per drawn path before any pixel.
const PATH_WORK: u64 = 128;
/// Per device pixel of outline length the rasteriser walks for a fill.
const FILL_EDGE_WORK: f64 = 4.0;
/// The same for a hairline, which is drawn a line at a time.
const HAIRLINE_WORK: f64 = 16.0;
/// Per dash `tiny-skia` cuts, each stroked as its own contour.
const DASH_WORK: f64 = 128.0;
/// Per piece the stroker emits, as [`super::VERB_NS`] prices it.
const VERB_WORK: f64 = 8.0;
/// Per step of the rasteriser's insertion sort of its active edges, about a
/// nanosecond where a painted pixel is four.
const SORT_WORK: f64 = 0.25;
/// Edges a line, quad and cubic become. Inside the surface `tiny-skia` fills
/// unclipped, splitting a curve into pieces monotonic in y: one, two, three.
/// Clipped, it also splits them at their x extrema, one, three, five, turning
/// a piece beyond a side into a vertical edge; and a piece straddling a side
/// gains a vertical edge on either end, up to three a line, nine a quad and
/// the clipper's own cap of eighteen a cubic.
const EDGES: [[f64; 3]; 3] = [[1.0, 1.0, 3.0], [2.0, 3.0, 9.0], [3.0, 5.0, 18.0]];
/// Bytes per edge: an 84-byte edge in a vector that doubles as it grows, and
/// the scratch its stable sort takes.
const EDGE_BYTES: f64 = 256.0;
/// Bytes of dashed path per dash, before any outline.
const DASH_PATH_BYTES: f64 = 48.0;
/// Below this many edge pairs a path is charged every pair without sweeping
/// its rows.
const SWEEP_PAIRS: f64 = 1_048_576.0;
/// How far, in device pixels, a stroke's outline may reach past its path:
/// eight times the largest raster. `tiny-skia` rasterises in 16.16 fixed
/// point, and an outline far past that overflows its coverage runs.
const MAX_REACH: f64 = 65_536.0;

pub(super) struct Cost {
    /// The output raster plus the deepest stack of layers and clip masks.
    pub(super) pixels: u64,
    /// Painting work in units of one painted pixel.
    pub(super) work: u64,
}

/// A device-space rectangle in output pixels.
#[derive(Clone, Copy)]
struct Area {
    left: f64,
    top: f64,
    right: f64,
    bottom: f64,
}

impl Area {
    fn pixels(self) -> f64 {
        (self.right - self.left).max(0.0) * (self.bottom - self.top).max(0.0)
    }

    fn meet(self, other: Area) -> Area {
        Area {
            left: self.left.max(other.left),
            top: self.top.max(other.top),
            right: self.right.min(other.right),
            bottom: self.bottom.min(other.bottom),
        }
    }

    fn of(rect: Rect, transform: Transform) -> Option<Area> {
        let rect = rect.transform(transform)?;
        Some(Area {
            left: f64::from(rect.left()),
            top: f64::from(rect.top()),
            right: f64::from(rect.right()),
            bottom: f64::from(rect.bottom()),
        })
    }

    /// Whether `tiny-skia` fills a path with these bounds unclipped: they fit
    /// the surface once rounded out, with a pixel to spare.
    fn holds(self, bounds: Area) -> bool {
        bounds.left - 1.0 >= self.left
            && bounds.top - 1.0 >= self.top
            && bounds.right + 1.0 <= self.right
            && bounds.bottom + 1.0 <= self.bottom
    }
}

struct Frame<'a> {
    group: &'a usvg::Group,
    surface: Area,
    live: u64,
    depth: usize,
}

#[derive(Default)]
struct Tally {
    nodes: u64,
    painted: f64,
    work: f64,
    peak: u64,
    /// The most any one path's edges, outline and dashes take while drawn.
    transient: f64,
}

impl Tally {
    fn node(&mut self) -> Result<(), SvgRefusal> {
        self.nodes += 1;
        if self.nodes > MAX_SVG_EXPANDED_NODES {
            return Err(SvgRefusal::ExpansionTooLarge);
        }
        Ok(())
    }

    fn paint(&mut self, pixels: f64, paint: &Paint) -> Result<(), SvgRefusal> {
        let stops = match paint {
            Paint::Color(_) => 0,
            Paint::LinearGradient(gradient) => gradient.stops().len(),
            Paint::RadialGradient(gradient) => gradient.stops().len(),
            Paint::Pattern(_) => return Err(SvgRefusal::UnsupportedElement),
        };
        let painted = pixels * (1.0 + stops as f64 / 8.0);
        self.painted += painted;
        self.work += painted;
        Ok(())
    }

    /// Prices filling `data`, placed by `place`, into `surface`.
    fn fill(&mut self, data: &tiny_skia::Path, place: Transform, surface: Area, extra: f64) {
        let (steps, edges) = sorting(data, place, surface);
        self.work += FILL_EDGE_WORK * length(closed(data), place) + SORT_WORK * steps;
        self.transient = self.transient.max(extra + edges * EDGE_BYTES);
    }

    /// Marks this size as past the budget: the raster steps down, or at its
    /// intrinsic size is refused.
    fn exceed(&mut self) {
        self.work = f64::INFINITY;
    }
}

/// Refuses a tree whose render is unbounded in kind (masks, filters, nested
/// clips, layers past [`MAX_SVG_LAYER_DEPTH`], overdraw past
/// [`MAX_SVG_OVERDRAW`]) and prices the rest for a `size` raster. Layers are
/// sized the way `resvg` allocates them: the group's device bounds grown by two
/// pixels a side, clamped to five canvases a side around the parent layer.
///
/// Each stroke wider than a hairline is priced for the pieces the stroker may
/// emit. With `strokes` it is instead stroked as `resvg` will stroke it, when
/// that many are still left, and charged the pieces it took, its outline
/// priced as the fill it becomes.
pub(super) fn measure(
    tree: &usvg::Tree,
    size: IntSize,
    mut strokes: Option<&mut f64>,
) -> Result<Cost, SvgRefusal> {
    let (width, height) = (f64::from(size.width()), f64::from(size.height()));
    let scale = Transform::from_scale(
        size.width() as f32 / tree.size().width(),
        size.height() as f32 / tree.size().height(),
    );
    let canvas = Area {
        left: 0.0,
        top: 0.0,
        right: width,
        bottom: height,
    };
    let output = u64::from(size.width()) * u64::from(size.height());
    let mut tally = Tally::default();
    let mut stack = vec![Frame {
        group: tree.root(),
        surface: canvas,
        live: 0,
        depth: 0,
    }];
    while let Some(frame) = stack.pop() {
        for node in frame.group.children() {
            tally.node()?;
            match node {
                Node::Group(group) => {
                    if group.mask().is_some() || !group.filters().is_empty() {
                        return Err(SvgRefusal::UnsupportedElement);
                    }
                    let mut inner = Frame { group, ..frame };
                    if group.should_isolate() {
                        inner.depth += 1;
                        if inner.depth > MAX_SVG_LAYER_DEPTH {
                            return Err(SvgRefusal::TooManyLayers);
                        }
                        let bounds = Area::of(group.abs_layer_bounding_box().to_rect(), scale);
                        let Some(layer) = layer(bounds, frame.surface, width, height) else {
                            continue;
                        };
                        let side = f64::from(MAX_SVG_RASTER_DIM);
                        if layer.right - layer.left > side || layer.bottom - layer.top > side {
                            tally.exceed();
                        }
                        let pixels = layer.pixels();
                        let masks = u64::from(group.clip_path().is_some());
                        tally.work += pixels * (1 + 2 * masks) as f64;
                        tally.painted += layer.meet(frame.surface).pixels();
                        inner.live = frame
                            .live
                            .saturating_add((pixels as u64).saturating_mul(1 + masks));
                        tally.peak = tally.peak.max(inner.live);
                        if let Some(clip) = group.clip_path() {
                            let place = scale
                                .pre_concat(group.abs_transform())
                                .pre_concat(clip.transform());
                            measure_clip(clip, place, layer, &mut tally)?;
                        }
                        inner.surface = layer;
                    }
                    stack.push(inner);
                }
                Node::Path(path) => {
                    if !path.is_visible() {
                        continue;
                    }
                    let place = scale.pre_concat(path.abs_transform());
                    tally.work += PATH_WORK as f64;
                    if let Some(fill) = path.fill() {
                        let area = Area::of(path.abs_bounding_box(), scale)
                            .map_or(frame.surface, |area| area.meet(frame.surface));
                        tally.paint(area.pixels(), fill.paint())?;
                        if area.pixels() > 0.0 {
                            tally.fill(path.data(), place, frame.surface, 0.0);
                        }
                    }
                    if let Some(stroke) = path.stroke() {
                        let area = Area::of(path.abs_stroke_bounding_box(), scale)
                            .map_or(frame.surface, |area| area.meet(frame.surface));
                        tally.paint(area.pixels(), stroke.paint())?;
                        let budget = strokes.as_deref_mut();
                        measure_stroke(path, stroke, place, frame.surface, &mut tally, budget)?;
                    }
                }
                Node::Image(_) | Node::Text(_) => return Err(SvgRefusal::UnsupportedElement),
            }
            if tally.work > MAX_SVG_RENDER_WORK as f64 {
                return Ok(Cost {
                    pixels: output,
                    work: u64::MAX,
                });
            }
        }
    }
    if tally.painted > (MAX_SVG_OVERDRAW as f64) * width * height {
        return Err(SvgRefusal::RenderTooCostly);
    }
    if tally.transient > SVG_TRANSIENT_BYTES as f64 {
        tally.exceed();
    }
    Ok(Cost {
        pixels: output.saturating_add(tally.peak),
        work: if tally.work.is_finite() {
            tally.work as u64
        } else {
            u64::MAX
        },
    })
}

/// Prices a stroke, on the surface or off it: `resvg` dashes and outlines a
/// stroke before anything culls it. A hairline is drawn a line at a time; anything wider is
/// dashed, outlined by the stroker and filled, so it is charged the pieces
/// [`geometry::stroke_verbs`] allows it, and with `strokes` left it is
/// stroked here once, as `resvg` will stroke it, for the outline to be priced.
/// Its curves must sit where the stroker's `f32` arithmetic stays well inside
/// its tolerance: within [`MAX_SVG_STROKE_SPAN`] tolerances of the origin. So
/// must a dashed hairline's, which the dasher measures by splitting each curve
/// until it meets a tolerance too.
fn measure_stroke(
    path: &usvg::Path,
    stroke: &usvg::Stroke,
    place: Transform,
    surface: Area,
    tally: &mut Tally,
    strokes: Option<&mut f64>,
) -> Result<(), SvgRefusal> {
    let dashes = dashes(path.data(), stroke);
    tally.work += DASH_WORK * dashes;
    let dashed = dashes * DASH_PATH_BYTES;
    let resolution = PathStroker::compute_resolution_scale(&place);
    let tolerance = 0.25 / f64::from(resolution);
    let radius = f64::from(stroke.width().get()) / 2.0;
    let mut outline = outline(path.data());
    let span = (outline.reach + radius) / tolerance;
    let far = span.is_nan() || span > MAX_SVG_STROKE_SPAN;
    if hairline(path, stroke, place) {
        tally.work += HAIRLINE_WORK * length(path.data().segments(), place);
        tally.transient = tally.transient.max(dashed);
        if far && dashes > 0.0 {
            tally.exceed();
        }
        return Ok(());
    }
    let reach = reach(stroke, place);
    if reach.is_nan() || reach > MAX_REACH {
        return Err(SvgRefusal::RenderTooCostly);
    }
    outline.segments += dashes;
    outline.contours += dashes;
    if outline.curves > 0.0 {
        outline.curves += dashes;
    }
    let verbs = geometry::stroke_verbs(&outline, radius, tolerance);
    let pieces = dashed + verbs * OUTLINE_VERB_BYTES;
    if far || pieces > SVG_TRANSIENT_BYTES as f64 {
        tally.exceed();
        return Ok(());
    }
    let Some(strokes) = strokes else {
        tally.work += VERB_WORK * verbs;
        tally.transient = tally.transient.max(pieces);
        return Ok(());
    };
    if verbs > *strokes || tally.work > MAX_SVG_RENDER_WORK as f64 {
        tally.exceed();
        return Ok(());
    }
    let outline_stroke = stroke.to_tiny_skia();
    let dashed_path;
    let source = match &outline_stroke.dash {
        Some(dash) => match path.data().dash(dash, resolution) {
            Some(data) => {
                dashed_path = data;
                &dashed_path
            }
            None => return Ok(()),
        },
        None => path.data(),
    };
    if let Some(drawn) = source.stroke(&outline_stroke, resolution) {
        let emitted = drawn.len() as f64;
        *strokes -= emitted;
        tally.work += VERB_WORK * emitted;
        tally.fill(
            &drawn,
            place,
            surface,
            dashed + emitted * OUTLINE_VERB_BYTES,
        );
    }
    Ok(())
}

/// Where `resvg` allocates a group's layer: its device bounds grown two pixels
/// a side (one more for rounding), fitted to five canvases a side around the
/// surface it composites onto. `None` is a layer `resvg` skips.
fn layer(bounds: Option<Area>, parent: Area, width: f64, height: f64) -> Option<Area> {
    let reach = Area {
        left: parent.left - 2.0 * width,
        top: parent.top - 2.0 * height,
        right: parent.left + 3.0 * width,
        bottom: parent.top + 3.0 * height,
    };
    let grown = bounds
        .filter(|area| {
            [area.left, area.top, area.right, area.bottom]
                .iter()
                .all(|v| v.is_finite())
        })
        .map_or(reach, |area| Area {
            left: area.left.floor() - 2.0,
            top: area.top.floor() - 2.0,
            right: area.left.floor() + (area.right - area.left).ceil() + 3.0,
            bottom: area.top.floor() + (area.bottom - area.top).ceil() + 3.0,
        });
    let layer = grown.meet(reach);
    (layer.pixels() > 0.0).then_some(layer)
}

/// A clip path fills its children into a raster the size of the layer it
/// clips. One level only: a clip on a clip, or on a group inside one, stacks
/// another raster per level.
fn measure_clip(
    clip: &usvg::ClipPath,
    place: Transform,
    layer: Area,
    tally: &mut Tally,
) -> Result<(), SvgRefusal> {
    if clip.clip_path().is_some() {
        return Err(SvgRefusal::TooManyLayers);
    }
    let mut stack = vec![clip.root()];
    while let Some(group) = stack.pop() {
        for node in group.children() {
            tally.node()?;
            match node {
                Node::Group(group) => {
                    if group.clip_path().is_some() {
                        return Err(SvgRefusal::TooManyLayers);
                    }
                    stack.push(group);
                }
                Node::Path(path) => {
                    let area = Area::of(path.abs_bounding_box(), place)
                        .map_or(layer, |area| area.meet(layer));
                    let drawn = place.pre_concat(path.abs_transform());
                    tally.painted += area.pixels();
                    tally.work += PATH_WORK as f64 + area.pixels();
                    tally.fill(path.data(), drawn, layer, 0.0);
                }
                Node::Image(_) | Node::Text(_) => return Err(SvgRefusal::UnsupportedElement),
            }
        }
    }
    Ok(())
}

/// A converted path's outline, in its own units, as the stroker reads it.
fn outline(data: &tiny_skia::Path) -> Outline {
    let point = |point: Point| (f64::from(point.x), f64::from(point.y));
    let mut outline = Outline::default();
    let (mut start, mut current) = ((0.0, 0.0), (0.0, 0.0));
    for segment in data.segments() {
        match segment {
            PathSegment::MoveTo(to) => {
                (start, current) = (point(to), point(to));
                outline.contours += 1.0;
            }
            PathSegment::LineTo(to) => {
                current = point(to);
                outline.segments += 1.0;
            }
            PathSegment::QuadTo(control, to) => {
                outline.curve(&[current, point(control), point(to)]);
                current = point(to);
            }
            PathSegment::CubicTo(first, second, to) => {
                outline.curve(&[current, point(first), point(second), point(to)]);
                current = point(to);
            }
            PathSegment::Close => {
                current = start;
                outline.segments += 1.0;
            }
        }
    }
    outline
}

/// A path's segments as `tiny-skia` fills them: every contour left open is
/// closed by a line back to its start, before the next contour and at the end.
fn closed(data: &tiny_skia::Path) -> impl Iterator<Item = PathSegment> + '_ {
    let mut open = false;
    data.segments()
        .map(Some)
        .chain(std::iter::once(None))
        .flat_map(move |segment| {
            let close = match segment {
                Some(PathSegment::MoveTo(_)) | None => std::mem::take(&mut open),
                Some(PathSegment::Close) => {
                    open = false;
                    false
                }
                Some(_) => {
                    open = true;
                    false
                }
            };
            close
                .then_some(PathSegment::Close)
                .into_iter()
                .chain(segment)
        })
}

/// The control-polygon length of `segments` under `transform`, which a curve
/// never exceeds.
fn length(segments: impl Iterator<Item = PathSegment>, transform: Transform) -> f64 {
    let map = |point: Point| {
        (
            f64::from(transform.sx * point.x + transform.kx * point.y + transform.tx),
            f64::from(transform.ky * point.x + transform.sy * point.y + transform.ty),
        )
    };
    let distance =
        |(ax, ay): (f64, f64), (bx, by): (f64, f64)| ((bx - ax).powi(2) + (by - ay).powi(2)).sqrt();
    let (mut start, mut current) = ((0.0, 0.0), (0.0, 0.0));
    let mut total = 0.0;
    for segment in segments {
        let points: &[Point] = match &segment {
            PathSegment::MoveTo(point) => {
                start = map(*point);
                current = start;
                continue;
            }
            PathSegment::LineTo(point) => std::slice::from_ref(point),
            PathSegment::QuadTo(control, point) => &[*control, *point],
            PathSegment::CubicTo(first, second, point) => &[*first, *second, *point],
            PathSegment::Close => {
                total += distance(current, start);
                current = start;
                continue;
            }
        };
        for point in points {
            let next = map(*point);
            total += distance(current, next);
            current = next;
        }
    }
    total
}

/// The edges `tiny-skia` builds to fill `data` into `surface`, the lines
/// closing each open contour included, and the steps
/// it may take insertion-sorting its active edges. A segment wholly above or
/// below a clipped surface is dropped. Sorting takes at most one step per pair
/// of edges, since two monotone edges swap at most once, and at most every
/// pair active together on each of the four anti-aliasing sub-rows of every
/// row a segment's control points span.
fn sorting(data: &tiny_skia::Path, place: Transform, surface: Area) -> (f64, f64) {
    let inside = Area::of(data.bounds(), place).is_some_and(|bounds| surface.holds(bounds));
    let map = |point: &Point| {
        (
            f64::from(place.sx * point.x + place.kx * point.y + place.tx),
            f64::from(place.ky * point.x + place.sy * point.y + place.ty),
        )
    };
    let mut spans = Vec::new();
    let mut edges = 0.0;
    let (mut start, mut current) = (Point::zero(), Point::zero());
    for segment in closed(data) {
        let points: &[Point] = match &segment {
            PathSegment::MoveTo(point) => {
                start = *point;
                current = start;
                continue;
            }
            PathSegment::LineTo(point) => std::slice::from_ref(point),
            PathSegment::QuadTo(control, point) => &[*control, *point],
            PathSegment::CubicTo(first, second, point) => &[*first, *second, *point],
            PathSegment::Close => std::slice::from_ref(&start),
        };
        let (mut columns, mut rows) = (
            (f64::INFINITY, f64::NEG_INFINITY),
            (f64::INFINITY, f64::NEG_INFINITY),
        );
        for point in std::iter::once(&current).chain(points) {
            let (x, y) = map(point);
            columns = (columns.0.min(x), columns.1.max(x));
            rows = (rows.0.min(y), rows.1.max(y));
        }
        current = *points.last().unwrap_or(&current);
        let kind = &EDGES[points.len().min(3) - 1];
        let count = if inside {
            kind[0]
        } else if rows.1 <= surface.top || rows.0 >= surface.bottom {
            continue;
        } else if columns.1 <= surface.left
            || columns.0 >= surface.right
            || (columns.0 >= surface.left && columns.1 <= surface.right)
        {
            kind[1]
        } else {
            kind[2]
        };
        edges += count;
        let top = rows.0.floor().max(surface.top.floor());
        let bottom = rows.1.ceil().min(surface.bottom.ceil());
        if top <= bottom {
            spans.push((top, count));
            spans.push((bottom + 1.0, -count));
        }
    }
    let pairs = 2.0 * edges * edges;
    if !pairs.is_finite() || pairs <= SWEEP_PAIRS {
        return (pairs, edges);
    }
    spans.sort_by(|a, b| a.0.total_cmp(&b.0));
    let (mut active, mut row, mut crowded) = (0.0f64, f64::NEG_INFINITY, 0.0);
    for (at, change) in spans {
        if active > 0.0 {
            crowded += 4.0 * (at - row) * active * active;
        }
        active += change;
        row = at;
    }
    (pairs.min(crowded), edges)
}

/// Whether `tiny-skia` draws a stroke as a hairline, which it never outlines:
/// its width maps within a pixel and the path is anti-aliased.
fn hairline(path: &usvg::Path, stroke: &usvg::Stroke, place: Transform) -> bool {
    let width = stroke.width().get();
    let spread = |x: f32, y: f32| {
        let (x, y) = (
            (place.sx * x + place.kx * y).abs(),
            (place.ky * x + place.sy * y).abs(),
        );
        x.max(y) + x.min(y) / 2.0
    };
    path.rendering_mode().use_shape_antialiasing()
        && spread(width, 0.0) <= 1.0
        && spread(0.0, width) <= 1.0
}

/// How far a stroke's outline reaches past its path, in device pixels: half
/// its width, stretched by a miter or a square cap, and a pixel of rounding.
fn reach(stroke: &usvg::Stroke, place: Transform) -> f64 {
    let scale = f64::from((place.sx.abs() + place.kx.abs()).max(place.ky.abs() + place.sy.abs()));
    let miter = match stroke.linejoin() {
        usvg::LineJoin::Miter | usvg::LineJoin::MiterClip => f64::from(stroke.miterlimit().get()),
        _ => 1.0,
    };
    f64::from(stroke.width().get()) / 2.0 * scale * miter.max(std::f64::consts::SQRT_2) + 1.0
}

/// Dashes `tiny-skia` cuts from a stroke, at most, in the path's own units: a
/// pattern's dashes per period of length, and since it restarts the pattern at
/// every contour and walks each interval, a pattern and a part more per
/// contour, however short the contour and however long the pattern's gaps.
fn dashes(data: &tiny_skia::Path, stroke: &usvg::Stroke) -> f64 {
    let Some(array) = stroke.dasharray() else {
        return 0.0;
    };
    let period: f64 = array.iter().map(|value| f64::from(*value)).sum();
    if period.is_nan() || period <= 0.0 {
        return 0.0;
    }
    let contours = data
        .segments()
        .filter(|segment| matches!(segment, PathSegment::MoveTo(_)))
        .count() as f64;
    let length = length(data.segments(), Transform::identity());
    (array.len() / 2) as f64 * (length / period + 2.0 * contours)
}
