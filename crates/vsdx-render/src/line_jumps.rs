use std::collections::BTreeMap;

use ooxml_drawingml::GeometryPathCommand;
use vsdx_eval::{Evaluation, Value, evaluate_cell_with_shape_package_theme};
use vsdx_parse::{ParseLimits, Sheet, VsdxPackage};
use vsdx_resolve::{Lookup, ResolvedShape, Resolver};

use crate::display_list::{Affine, Primitive};

const END_EPSILON: f64 = 1e-9;
const MAX_STYLE_DEPTH: usize = 8;
const SEGMENT_INDEX_LEAF_SIZE: usize = 8;
const MAX_CROSSING_CANDIDATES: usize = 20_000_000;
const MAX_PLACED_JUMPS: usize = 100_000;

pub struct PageJumpSettings {
    pub code: i32,
    pub style: i32,
    pub factor_x: f64,
    pub factor_y: f64,
    pub line_to_x: f64,
    pub line_to_y: f64,
    pub dir_x: i32,
    pub dir_y: i32,
}

pub struct ConnectorJumpOverride {
    pub id: String,
    pub code: i32,
    pub style: i32,
    pub dir_x: i32,
    pub dir_y: i32,
}

pub fn page_jump_settings(
    package: &VsdxPackage,
    resolver: &Resolver<'_>,
    page_part: &str,
) -> PageJumpSettings {
    let resolved = package
        .page_part_ids
        .get(page_part)
        .and_then(|id| package.page_sheets.get(id))
        .and_then(|sheet| resolver.resolve_sheet(sheet).ok());
    let number = |name: &str| {
        resolved
            .as_ref()
            .and_then(|shape| resolved_number(package, shape, name))
            .or_else(|| no_style_number(package, name))
    };
    PageJumpSettings {
        code: number("LineJumpCode").unwrap_or(0.0) as i32,
        style: number("LineJumpStyle").unwrap_or(0.0) as i32,
        factor_x: number("LineJumpFactorX").unwrap_or(0.5),
        factor_y: number("LineJumpFactorY").unwrap_or(0.5),
        line_to_x: number("LineToLineX").unwrap_or(0.1),
        line_to_y: number("LineToLineY").unwrap_or(0.1),
        dir_x: number("PageLineJumpDirX").unwrap_or(0.0) as i32,
        dir_y: number("PageLineJumpDirY").unwrap_or(0.0) as i32,
    }
}

pub fn connector_override(package: &VsdxPackage, resolved: &ResolvedShape, name: &str) -> i32 {
    crate::paint::number(resolved, name)
        .or_else(|| no_style_number(package, name))
        .unwrap_or(0.0) as i32
}

pub(crate) fn no_style_number(package: &VsdxPackage, name: &str) -> Option<f64> {
    let sheet = package
        .style_sheets
        .iter()
        .find(|sheet| sheet.id == Some(0))
        .or_else(|| {
            package.style_sheets.iter().find(|sheet| {
                sheet
                    .other_attrs
                    .iter()
                    .any(|(key, value)| key == "NameU" && value == "No Style")
            })
        })?;
    no_style_chain(package, sheet, name, 0)
}

pub fn apply_line_jumps(
    primitives: &mut [Primitive],
    overrides: &[ConnectorJumpOverride],
    settings: &PageJumpSettings,
) {
    let lookup: BTreeMap<&str, &ConnectorJumpOverride> = overrides
        .iter()
        .map(|item| (item.id.as_str(), item))
        .collect();
    let mut routes = Vec::new();
    collect_routes(primitives, &lookup, Affine::identity(), &mut routes);
    if routes.len() < 2 {
        return;
    }
    let mut jumps: BTreeMap<usize, Vec<PlacedJump>> = BTreeMap::new();
    let segments = route_segments(&routes);
    let index = SegmentIndex::new(&segments);
    let mut candidates = MAX_CROSSING_CANDIDATES;
    let mut budget = MAX_PLACED_JUMPS;
    for (segment_index, segment) in segments.iter().enumerate() {
        let within_budget = index.overlapping(&segments, segment.bounds, &mut |other_index| {
            if other_index <= segment_index || segments[other_index].route == segment.route {
                return true;
            }
            if candidates == 0 {
                return false;
            }
            candidates -= 1;
            let other = &segments[other_index];
            let Some(crossing) = crossing(&routes, segment, other) else {
                return true;
            };
            let winner = carrier(
                lookup[routes[crossing.a].id.as_str()].code,
                lookup[routes[crossing.b].id.as_str()].code,
                page_pick(
                    settings.code,
                    crossing.dir_a,
                    crossing.dir_b,
                    routes[crossing.a].z,
                    routes[crossing.b].z,
                ),
            );
            if let Some(winner) = winner {
                let route = if winner == Side::A {
                    crossing.a
                } else {
                    crossing.b
                };
                let item = &lookup[routes[route].id.as_str()];
                if let Some(jump) = placed_jump(item, settings, &crossing, winner == Side::A) {
                    if budget == 0 {
                        return false;
                    }
                    budget -= 1;
                    jumps.entry(route).or_default().push(jump);
                }
            }
            true
        });
        if !within_budget {
            return;
        }
    }
    for (index, mut placed) in jumps {
        let route = &routes[index];
        if let Some(path) = shape_mut(primitives, &route.id) {
            *path = splice(&route.points, &mut placed);
            if let Some(inverse) = route.local {
                for command in path.iter_mut() {
                    crate::transform_affine(command, inverse);
                }
            }
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum Side {
    A,
    B,
}

struct Route {
    id: String,
    z: u32,
    /// Page space, so routes from different groups share one frame.
    points: Vec<(f64, f64)>,
    /// Maps page space back onto the path's own frame; `None` when they are the same.
    local: Option<Affine>,
}

struct Segment {
    route: usize,
    index: usize,
    bounds: (f64, f64, f64, f64),
}

struct SegmentIndex {
    root: Option<Box<SegmentIndexNode>>,
}

struct SegmentIndexNode {
    bounds: (f64, f64, f64, f64),
    entries: Vec<usize>,
    left: Option<Box<SegmentIndexNode>>,
    right: Option<Box<SegmentIndexNode>>,
}

struct Crossing {
    a: usize,
    b: usize,
    x: f64,
    y: f64,
    seg_a: usize,
    seg_b: usize,
    t_a: f64,
    t_b: f64,
    dir_a: (f64, f64),
    dir_b: (f64, f64),
}

struct PlacedJump {
    seg: usize,
    t: f64,
    x: f64,
    y: f64,
    half: f64,
    nx: f64,
    ny: f64,
    gap: bool,
}

fn resolved_number(package: &VsdxPackage, resolved: &ResolvedShape, name: &str) -> Option<f64> {
    let Lookup::Found(cell) = resolved.cell(name)? else {
        return None;
    };
    if cell
        .cell
        .formula
        .as_deref()
        .is_some_and(|formula| formula.eq_ignore_ascii_case("Inh"))
    {
        return None;
    }
    if let Some(formula) = cell.cell.formula.as_deref()
        && let Evaluation::Evaluated(result) = evaluate_cell_with_shape_package_theme(
            name,
            formula,
            resolved,
            &ParseLimits::default(),
            resolved,
            package,
        )
        && let Value::Number(number) = result.value
        && number.number.is_finite()
    {
        return Some(number.number);
    }
    cell.cell
        .value
        .as_deref()?
        .parse()
        .ok()
        .filter(|value: &f64| value.is_finite())
}

fn no_style_chain(package: &VsdxPackage, sheet: &Sheet, name: &str, depth: usize) -> Option<f64> {
    if depth > MAX_STYLE_DEPTH {
        return None;
    }
    let mut inherit = false;
    for cell in sheet.cells() {
        if cell.name == name {
            if cell
                .formula
                .as_deref()
                .is_some_and(|formula| formula.eq_ignore_ascii_case("Inh"))
            {
                inherit = true;
                break;
            }
            return cell
                .value
                .as_deref()?
                .parse()
                .ok()
                .filter(|value: &f64| value.is_finite());
        }
    }
    if !inherit && depth > 0 {
        return None;
    }
    let based: u32 = sheet
        .other_attrs
        .iter()
        .find(|(key, _)| key == "BasedOn")?
        .1
        .parse()
        .ok()?;
    let parent = package
        .style_sheets
        .iter()
        .find(|sheet| sheet.id == Some(based))?;
    no_style_chain(package, parent, name, depth + 1)
}

fn collect_routes(
    primitives: &[Primitive],
    lookup: &BTreeMap<&str, &ConnectorJumpOverride>,
    ancestors: Affine,
    routes: &mut Vec<Route>,
) {
    for primitive in primitives {
        match primitive {
            Primitive::Shape {
                id,
                z_order,
                path,
                stroke: Some(_),
                transform,
                ..
            } if lookup.contains_key(id.as_str()) => {
                let page = ancestors.compose(*transform);
                let local = if page.is_identity() {
                    None
                } else {
                    let Some(inverse) = page.invert() else {
                        continue;
                    };
                    Some(inverse)
                };
                let Some(mut points) = route_points(path) else {
                    continue;
                };
                if local.is_some() {
                    for point in &mut points {
                        let (x, y) = page.apply_point(point.0 as f32, point.1 as f32);
                        *point = (x as f64, y as f64);
                    }
                }
                routes.push(Route {
                    id: id.clone(),
                    z: *z_order,
                    points,
                    local,
                });
            }
            Primitive::Group {
                primitives,
                transform,
                ..
            } => collect_routes(primitives, lookup, ancestors.compose(*transform), routes),
            _ => {}
        }
    }
}

fn route_points(path: &[GeometryPathCommand]) -> Option<Vec<(f64, f64)>> {
    let mut points = Vec::new();
    for command in path {
        match *command {
            GeometryPathCommand::Move { x, y } => {
                if !points.is_empty() {
                    return None;
                }
                points.push((x, y));
            }
            GeometryPathCommand::Line { x, y } => points.push((x, y)),
            _ => return None,
        }
    }
    if points.len() >= 2 && points.iter().all(|(x, y)| x.is_finite() && y.is_finite()) {
        Some(points)
    } else {
        None
    }
}

fn route_bounds(points: &[(f64, f64)]) -> (f64, f64, f64, f64) {
    points.iter().fold(
        (
            f64::INFINITY,
            f64::INFINITY,
            f64::NEG_INFINITY,
            f64::NEG_INFINITY,
        ),
        |(min_x, min_y, max_x, max_y), (x, y)| {
            (min_x.min(*x), min_y.min(*y), max_x.max(*x), max_y.max(*y))
        },
    )
}

fn bounds_overlap(a: (f64, f64, f64, f64), b: (f64, f64, f64, f64)) -> bool {
    a.0 <= b.2 && b.0 <= a.2 && a.1 <= b.3 && b.1 <= a.3
}

fn route_segments(routes: &[Route]) -> Vec<Segment> {
    routes
        .iter()
        .enumerate()
        .flat_map(|(route, item)| {
            item.points
                .windows(2)
                .enumerate()
                .map(move |(index, points)| Segment {
                    route,
                    index,
                    bounds: route_bounds(points),
                })
        })
        .collect()
}

impl SegmentIndex {
    fn new(segments: &[Segment]) -> Self {
        let entries = (0..segments.len()).collect();
        Self {
            root: SegmentIndexNode::new(segments, entries),
        }
    }

    fn overlapping(
        &self,
        segments: &[Segment],
        bounds: (f64, f64, f64, f64),
        visitor: &mut impl FnMut(usize) -> bool,
    ) -> bool {
        if let Some(root) = &self.root {
            return root.overlapping(segments, bounds, visitor);
        }
        true
    }
}

impl SegmentIndexNode {
    fn new(segments: &[Segment], mut entries: Vec<usize>) -> Option<Box<Self>> {
        if entries.is_empty() {
            return None;
        }
        let bounds = entries.iter().fold(
            (
                f64::INFINITY,
                f64::INFINITY,
                f64::NEG_INFINITY,
                f64::NEG_INFINITY,
            ),
            |bounds, &entry| merge_bounds(bounds, segments[entry].bounds),
        );
        if entries.len() <= SEGMENT_INDEX_LEAF_SIZE {
            return Some(Box::new(Self {
                bounds,
                entries,
                left: None,
                right: None,
            }));
        }
        let split_x = bounds.2 - bounds.0 >= bounds.3 - bounds.1;
        entries.sort_by(|&a, &b| {
            let a = segments[a].bounds;
            let b = segments[b].bounds;
            let a = if split_x { a.0 + a.2 } else { a.1 + a.3 };
            let b = if split_x { b.0 + b.2 } else { b.1 + b.3 };
            a.total_cmp(&b)
        });
        let right_entries = entries.split_off(entries.len() / 2);
        Some(Box::new(Self {
            bounds,
            entries: Vec::new(),
            left: Self::new(segments, entries),
            right: Self::new(segments, right_entries),
        }))
    }

    fn overlapping(
        &self,
        segments: &[Segment],
        bounds: (f64, f64, f64, f64),
        visitor: &mut impl FnMut(usize) -> bool,
    ) -> bool {
        if !bounds_overlap(self.bounds, bounds) {
            return true;
        }
        if self.left.is_none() {
            for &entry in &self.entries {
                if bounds_overlap(segments[entry].bounds, bounds) && !visitor(entry) {
                    return false;
                }
            }
            return true;
        }
        if let Some(left) = &self.left
            && !left.overlapping(segments, bounds, visitor)
        {
            return false;
        }
        if let Some(right) = &self.right {
            return right.overlapping(segments, bounds, visitor);
        }
        true
    }
}

fn merge_bounds(a: (f64, f64, f64, f64), b: (f64, f64, f64, f64)) -> (f64, f64, f64, f64) {
    (a.0.min(b.0), a.1.min(b.1), a.2.max(b.2), a.3.max(b.3))
}

fn crossing(routes: &[Route], a: &Segment, b: &Segment) -> Option<Crossing> {
    let route_a = &routes[a.route];
    let route_b = &routes[b.route];
    let p = route_a.points[a.index];
    let q = route_a.points[a.index + 1];
    let r = route_b.points[b.index];
    let s = route_b.points[b.index + 1];
    let (t_a, t_b, x, y) = segment_intersection(p, q, r, s)?;
    if route_endpoint(a.index, route_a.points.len() - 2, t_a)
        || route_endpoint(b.index, route_b.points.len() - 2, t_b)
    {
        return None;
    }
    Some(Crossing {
        a: a.route,
        b: b.route,
        x,
        y,
        seg_a: a.index,
        seg_b: b.index,
        t_a,
        t_b,
        dir_a: (q.0 - p.0, q.1 - p.1),
        dir_b: (s.0 - r.0, s.1 - r.1),
    })
}

fn route_endpoint(segment: usize, last_segment: usize, t: f64) -> bool {
    (segment == 0 && t < END_EPSILON) || (segment == last_segment && t > 1.0 - END_EPSILON)
}

#[cfg(test)]
fn segments_cross(
    p: (f64, f64),
    q: (f64, f64),
    r: (f64, f64),
    s: (f64, f64),
) -> Option<(f64, f64, f64, f64)> {
    let hit = segment_intersection(p, q, r, s)?;
    ((END_EPSILON..=1.0 - END_EPSILON).contains(&hit.0)
        && (END_EPSILON..=1.0 - END_EPSILON).contains(&hit.1))
    .then_some(hit)
}

fn segment_intersection(
    p: (f64, f64),
    q: (f64, f64),
    r: (f64, f64),
    s: (f64, f64),
) -> Option<(f64, f64, f64, f64)> {
    let dx_p = q.0 - p.0;
    let dy_p = q.1 - p.1;
    let dx_r = s.0 - r.0;
    let dy_r = s.1 - r.1;
    let denom = dx_p * dy_r - dy_p * dx_r;
    if denom == 0.0 || !denom.is_finite() {
        return None;
    }
    let t = ((r.0 - p.0) * dy_r - (r.1 - p.1) * dx_r) / denom;
    let u = ((r.0 - p.0) * dy_p - (r.1 - p.1) * dx_p) / denom;
    if !(0.0..=1.0).contains(&t) || !(0.0..=1.0).contains(&u) {
        return None;
    }
    let (x, y) = (p.0 + t * dx_p, p.1 + t * dy_p);
    (x.is_finite() && y.is_finite()).then_some((t, u, x, y))
}

fn page_pick(code: i32, dir_a: (f64, f64), dir_b: (f64, f64), z_a: u32, z_b: u32) -> Option<Side> {
    match code {
        1 => slope_order(dir_a, dir_b),
        2 => slope_order(dir_b, dir_a),
        4 => Some(if z_a > z_b { Side::A } else { Side::B }),
        5 => Some(if z_a < z_b { Side::A } else { Side::B }),
        _ => None,
    }
}

fn slope_order(first: (f64, f64), second: (f64, f64)) -> Option<Side> {
    let left = first.1.abs() * second.0.abs();
    let right = second.1.abs() * first.0.abs();
    if !(left.is_finite() && right.is_finite()) {
        return None;
    }
    if left < right {
        Some(Side::A)
    } else if left > right {
        Some(Side::B)
    } else {
        None
    }
}

fn carrier(a: i32, b: i32, pick: Option<Side>) -> Option<Side> {
    if a == 4 || b == 4 {
        return None;
    }
    let claim_a = a == 2 || b == 3;
    let claim_b = b == 2 || a == 3;
    let pick = match (claim_a, claim_b) {
        (true, false) => Some(Side::A),
        (false, true) => Some(Side::B),
        _ => pick,
    };
    match pick {
        Some(Side::A) if a == 1 || a == 3 => {
            if b == 1 || b == 3 || b == 4 {
                None
            } else {
                Some(Side::B)
            }
        }
        Some(Side::B) if b == 1 || b == 3 => {
            if a == 1 || a == 3 || a == 4 {
                None
            } else {
                Some(Side::A)
            }
        }
        _ => pick,
    }
}

fn placed_jump(
    item: &ConnectorJumpOverride,
    settings: &PageJumpSettings,
    crossing: &Crossing,
    is_a: bool,
) -> Option<PlacedJump> {
    let (seg, t) = if is_a {
        (crossing.seg_a, crossing.t_a)
    } else {
        (crossing.seg_b, crossing.t_b)
    };
    let (dx, dy) = if is_a { crossing.dir_a } else { crossing.dir_b };
    let length = dx.hypot(dy);
    if length <= 0.0 || !length.is_finite() {
        return None;
    }
    let horizontal = dx.abs() >= dy.abs();
    let half = if horizontal {
        settings.factor_x * settings.line_to_x
    } else {
        settings.factor_y * settings.line_to_y
    } / 2.0;
    if half <= 0.0 || !half.is_finite() {
        return None;
    }
    if t * length <= half || (1.0 - t) * length <= half {
        return None;
    }
    let (ux, uy) = (dx / length, dy / length);
    let dir = if horizontal {
        connector_dir(item.dir_x, settings.dir_x)
    } else {
        connector_dir(item.dir_y, settings.dir_y)
    };
    let (nx, ny) = jump_normal(ux, uy, horizontal, dir);
    let style = if item.style == 0 {
        settings.style
    } else {
        item.style
    };
    match style {
        0 | 1 => Some(PlacedJump {
            seg,
            t,
            x: crossing.x,
            y: crossing.y,
            half,
            nx,
            ny,
            gap: false,
        }),
        2 => Some(PlacedJump {
            seg,
            t,
            x: crossing.x,
            y: crossing.y,
            half,
            nx,
            ny,
            gap: true,
        }),
        _ => None,
    }
}

fn connector_dir(local: i32, page: i32) -> i32 {
    if local == 1 || local == 2 {
        local
    } else {
        page
    }
}

fn jump_normal(ux: f64, uy: f64, horizontal: bool, dir: i32) -> (f64, f64) {
    match (horizontal, dir) {
        (true, 1) => (0.0, 1.0),
        (true, 2) => (0.0, -1.0),
        (false, 1) => (-1.0, 0.0),
        (false, 2) => (1.0, 0.0),
        _ => (-uy, ux),
    }
}

fn splice(points: &[(f64, f64)], jumps: &mut [PlacedJump]) -> Vec<GeometryPathCommand> {
    jumps.sort_by(|a, b| a.seg.cmp(&b.seg).then_with(|| a.t.total_cmp(&b.t)));
    let mut by_seg: BTreeMap<usize, Vec<&PlacedJump>> = BTreeMap::new();
    for jump in jumps.iter() {
        by_seg.entry(jump.seg).or_default().push(jump);
    }
    let mut out = vec![GeometryPathCommand::Move {
        x: points[0].0,
        y: points[0].1,
    }];
    for seg in 0..points.len() - 1 {
        let (x0, y0) = points[seg];
        let (x1, y1) = points[seg + 1];
        let length = (x1 - x0).hypot(y1 - y0);
        let mut kept: Vec<&&PlacedJump> = Vec::new();
        if let Some(candidates) = by_seg.get(&seg) {
            for jump in candidates {
                if jump.t * length <= jump.half || (1.0 - jump.t) * length <= jump.half {
                    continue;
                }
                if let Some(last) = kept.last()
                    && (jump.t - last.t) * length < last.half + jump.half
                {
                    continue;
                }
                kept.push(jump);
            }
        }
        let (ux, uy) = if length > 0.0 {
            ((x1 - x0) / length, (y1 - y0) / length)
        } else {
            (0.0, 0.0)
        };
        for jump in kept {
            let (ex, ey) = (jump.x - jump.half * ux, jump.y - jump.half * uy);
            let (xx, xy) = (jump.x + jump.half * ux, jump.y + jump.half * uy);
            out.push(GeometryPathCommand::Line { x: ex, y: ey });
            if jump.gap {
                out.push(GeometryPathCommand::Move { x: xx, y: xy });
            } else {
                let (ax, ay) = (jump.x + jump.half * jump.nx, jump.y + jump.half * jump.ny);
                out.push(GeometryPathCommand::Quad {
                    cpx: ex + jump.half * jump.nx,
                    cpy: ey + jump.half * jump.ny,
                    x: ax,
                    y: ay,
                });
                out.push(GeometryPathCommand::Quad {
                    cpx: xx + jump.half * jump.nx,
                    cpy: xy + jump.half * jump.ny,
                    x: xx,
                    y: xy,
                });
            }
        }
        out.push(GeometryPathCommand::Line { x: x1, y: y1 });
    }
    out
}

fn shape_mut<'a>(
    primitives: &'a mut [Primitive],
    id: &str,
) -> Option<&'a mut Vec<GeometryPathCommand>> {
    primitives.iter_mut().find_map(|primitive| match primitive {
        Primitive::Shape {
            id: actual, path, ..
        } if actual == id => Some(path),
        Primitive::Group { primitives, .. } => shape_mut(primitives, id),
        _ => None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn settings(code: i32) -> PageJumpSettings {
        PageJumpSettings {
            code,
            style: 1,
            factor_x: 0.5,
            factor_y: 0.5,
            line_to_x: 0.1,
            line_to_y: 0.1,
            dir_x: 0,
            dir_y: 0,
        }
    }

    fn stroke() -> crate::display_list::Stroke {
        crate::display_list::Stroke {
            color: String::new(),
            width: 0.01,
            dashed: false,
        }
    }

    #[test]
    fn crossing_rejects_parallel_and_endpoint_touch() {
        assert!(segments_cross((0.0, 0.0), (2.0, 0.0), (0.0, 1.0), (2.0, 1.0)).is_none());
        assert!(segments_cross((0.0, 0.0), (2.0, 0.0), (1.0, 0.0), (3.0, 0.0)).is_none());
        assert!(segments_cross((0.0, 0.0), (2.0, 0.0), (2.0, 0.0), (2.0, 2.0)).is_none());
        let hit = segments_cross((0.0, 0.0), (2.0, 0.0), (1.0, -1.0), (1.0, 1.0)).unwrap();
        assert!((hit.0 - 0.5).abs() < 1e-12);
        assert!((hit.1 - 0.5).abs() < 1e-12);
        assert!((hit.2 - 1.0).abs() < 1e-12);
        assert!((hit.3 - 0.0).abs() < 1e-12);
    }

    #[test]
    fn page_pick_prefers_strictly_more_horizontal_or_vertical() {
        assert_eq!(page_pick(1, (4.0, 1.0), (1.0, 4.0), 0, 1), Some(Side::A));
        assert_eq!(page_pick(1, (1.0, 4.0), (4.0, 1.0), 0, 1), Some(Side::B));
        assert_eq!(page_pick(1, (1.0, 1.0), (2.0, 2.0), 0, 1), None);
        assert_eq!(page_pick(2, (4.0, 1.0), (1.0, 4.0), 0, 1), Some(Side::B));
        assert_eq!(page_pick(0, (4.0, 1.0), (1.0, 4.0), 0, 1), None);
        assert_eq!(page_pick(3, (4.0, 1.0), (1.0, 4.0), 0, 1), None);
        assert_eq!(page_pick(4, (4.0, 1.0), (1.0, 4.0), 0, 1), Some(Side::B));
        assert_eq!(page_pick(5, (4.0, 1.0), (1.0, 4.0), 0, 1), Some(Side::A));
    }

    #[test]
    fn carrier_honours_never_neither_always_and_other() {
        assert_eq!(carrier(0, 0, Some(Side::A)), Some(Side::A));
        assert_eq!(carrier(0, 0, None), None);
        assert_eq!(carrier(1, 0, Some(Side::A)), Some(Side::B));
        assert_eq!(carrier(1, 1, Some(Side::A)), None);
        assert_eq!(carrier(0, 4, Some(Side::A)), None);
        assert_eq!(carrier(2, 0, Some(Side::B)), Some(Side::A));
        assert_eq!(carrier(0, 3, Some(Side::B)), Some(Side::A));
        assert_eq!(carrier(3, 3, Some(Side::A)), None);
        assert_eq!(carrier(2, 2, Some(Side::B)), Some(Side::B));
        assert_eq!(carrier(3, 1, Some(Side::B)), None);
    }

    #[test]
    fn splice_orders_jumps_and_skips_overlaps() {
        let points = [(0.0, 0.0), (4.0, 0.0)];
        let mut jumps = vec![
            PlacedJump {
                seg: 0,
                t: 0.7,
                x: 2.8,
                y: 0.0,
                half: 0.2,
                nx: 0.0,
                ny: 1.0,
                gap: false,
            },
            PlacedJump {
                seg: 0,
                t: 0.3,
                x: 1.2,
                y: 0.0,
                half: 0.2,
                nx: 0.0,
                ny: 1.0,
                gap: false,
            },
            PlacedJump {
                seg: 0,
                t: 0.38,
                x: 1.52,
                y: 0.0,
                half: 0.2,
                nx: 0.0,
                ny: 1.0,
                gap: false,
            },
        ];
        let path = splice(&points, &mut jumps);
        assert!(matches!(path[0], GeometryPathCommand::Move { .. }));
        assert_eq!(
            path.iter()
                .filter(|command| matches!(command, GeometryPathCommand::Quad { .. }))
                .count(),
            4
        );
        let apex = path.iter().find_map(|command| match *command {
            GeometryPathCommand::Quad { x, y, .. } if (y - 0.2).abs() < 1e-12 => Some((x, y)),
            _ => None,
        });
        assert_eq!(apex, Some((1.2, 0.2)));
    }

    #[test]
    fn apply_skips_shared_endpoints_and_honours_code_zero() {
        let mut primitives = vec![
            Primitive::Shape {
                id: "a".into(),
                z_order: 0,
                path: vec![
                    GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                    GeometryPathCommand::Line { x: 2.0, y: 0.0 },
                ],
                fill: None,
                stroke: Some(stroke()),
                transform: crate::Affine::identity(),
                diagnostics: Vec::new(),
            },
            Primitive::Shape {
                id: "b".into(),
                z_order: 1,
                path: vec![
                    GeometryPathCommand::Move { x: 2.0, y: 0.0 },
                    GeometryPathCommand::Line { x: 2.0, y: 2.0 },
                ],
                fill: None,
                stroke: Some(stroke()),
                transform: crate::Affine::identity(),
                diagnostics: Vec::new(),
            },
        ];
        let overrides = vec![
            ConnectorJumpOverride {
                id: "a".into(),
                code: 0,
                style: 0,
                dir_x: 0,
                dir_y: 0,
            },
            ConnectorJumpOverride {
                id: "b".into(),
                code: 0,
                style: 0,
                dir_x: 0,
                dir_y: 0,
            },
        ];
        apply_line_jumps(&mut primitives, &overrides, &settings(1));
        for primitive in &primitives {
            let Primitive::Shape { path, .. } = primitive else {
                unreachable!()
            };
            assert_eq!(path.len(), 2);
        }
    }

    #[test]
    fn apply_ignores_crossings_with_invisible_connectors() {
        let mut primitives = vec![
            Primitive::Shape {
                id: "visible".into(),
                z_order: 0,
                path: vec![
                    GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                    GeometryPathCommand::Line { x: 2.0, y: 0.0 },
                ],
                fill: None,
                stroke: Some(stroke()),
                transform: crate::Affine::identity(),
                diagnostics: Vec::new(),
            },
            Primitive::Shape {
                id: "invisible".into(),
                z_order: 1,
                path: vec![
                    GeometryPathCommand::Move { x: 1.0, y: -1.0 },
                    GeometryPathCommand::Line { x: 1.0, y: 1.0 },
                ],
                fill: None,
                stroke: None,
                transform: crate::Affine::identity(),
                diagnostics: Vec::new(),
            },
        ];
        let overrides = ["visible", "invisible"]
            .into_iter()
            .map(|id| ConnectorJumpOverride {
                id: id.into(),
                code: 0,
                style: 0,
                dir_x: 0,
                dir_y: 0,
            })
            .collect::<Vec<_>>();
        apply_line_jumps(&mut primitives, &overrides, &settings(1));
        let Primitive::Shape { path, .. } = &primitives[0] else {
            unreachable!()
        };
        assert_eq!(path.len(), 2);
    }

    #[test]
    fn apply_draws_jump_when_crossing_an_internal_bend() {
        let mut primitives = vec![
            Primitive::Shape {
                id: "bent".into(),
                z_order: 0,
                path: vec![
                    GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                    GeometryPathCommand::Line { x: 1.0, y: 0.0 },
                    GeometryPathCommand::Line { x: 1.0, y: 1.0 },
                ],
                fill: None,
                stroke: Some(stroke()),
                transform: crate::Affine::identity(),
                diagnostics: Vec::new(),
            },
            Primitive::Shape {
                id: "crossing".into(),
                z_order: 1,
                path: vec![
                    GeometryPathCommand::Move { x: 0.5, y: -0.5 },
                    GeometryPathCommand::Line { x: 1.5, y: 0.5 },
                ],
                fill: None,
                stroke: Some(stroke()),
                transform: crate::Affine::identity(),
                diagnostics: Vec::new(),
            },
        ];
        let overrides = ["bent", "crossing"]
            .into_iter()
            .map(|id| ConnectorJumpOverride {
                id: id.into(),
                code: 0,
                style: 0,
                dir_x: 0,
                dir_y: 0,
            })
            .collect::<Vec<_>>();
        apply_line_jumps(&mut primitives, &overrides, &settings(1));
        let Primitive::Shape { path, .. } = &primitives[1] else {
            unreachable!()
        };
        assert!(
            path.iter()
                .any(|command| matches!(command, GeometryPathCommand::Quad { .. }))
        );
    }

    #[test]
    fn apply_preserves_late_crossings_after_many_collinear_candidates() {
        let count = 1_416;
        let mut primitives = (0..count)
            .map(|index| Primitive::Shape {
                id: format!("horizontal-{index}"),
                z_order: index as u32,
                path: vec![
                    GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                    GeometryPathCommand::Line { x: 2.0, y: 0.0 },
                ],
                fill: None,
                stroke: Some(stroke()),
                transform: crate::Affine::identity(),
                diagnostics: Vec::new(),
            })
            .collect::<Vec<_>>();
        primitives.push(Primitive::Shape {
            id: "vertical".into(),
            z_order: count as u32,
            path: vec![
                GeometryPathCommand::Move { x: 1.0, y: -1.0 },
                GeometryPathCommand::Line { x: 1.0, y: 1.0 },
            ],
            fill: None,
            stroke: Some(stroke()),
            transform: crate::Affine::identity(),
            diagnostics: Vec::new(),
        });
        let overrides = primitives
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Shape { id, .. } => Some(ConnectorJumpOverride {
                    id: id.clone(),
                    code: 0,
                    style: 0,
                    dir_x: 0,
                    dir_y: 0,
                }),
                _ => None,
            })
            .collect::<Vec<_>>();

        apply_line_jumps(&mut primitives, &overrides, &settings(1));

        let Primitive::Shape { path, .. } = &primitives[count - 1] else {
            unreachable!()
        };
        assert!(
            path.iter()
                .any(|command| matches!(command, GeometryPathCommand::Quad { .. }))
        );
    }

    fn route(id: &str, z: u32, points: &[(f64, f64)]) -> Primitive {
        let mut path = vec![GeometryPathCommand::Move {
            x: points[0].0,
            y: points[0].1,
        }];
        for (x, y) in &points[1..] {
            path.push(GeometryPathCommand::Line { x: *x, y: *y });
        }
        Primitive::Shape {
            id: id.into(),
            z_order: z,
            path,
            fill: None,
            stroke: Some(stroke()),
            transform: crate::Affine::identity(),
            diagnostics: Vec::new(),
        }
    }

    fn plain(ids: &[&str]) -> Vec<ConnectorJumpOverride> {
        ids.iter()
            .map(|id| ConnectorJumpOverride {
                id: (*id).into(),
                code: 0,
                style: 0,
                dir_x: 0,
                dir_y: 0,
            })
            .collect()
    }

    fn arcs(primitives: &[Primitive]) -> usize {
        primitives
            .iter()
            .map(|p| match p {
                Primitive::Shape { path, .. } => path
                    .iter()
                    .filter(|c| matches!(c, GeometryPathCommand::Quad { .. }))
                    .count(),
                _ => 0,
            })
            .sum()
    }

    #[test]
    fn degenerate_route_geometry_never_bridges() {
        for (name, primitives, ids) in [
            (
                "collinear overlap",
                vec![
                    route("a", 0, &[(0.0, 0.0), (4.0, 0.0)]),
                    route("b", 1, &[(1.0, 0.0), (5.0, 0.0)]),
                ],
                &["a", "b"][..],
            ),
            (
                "zero length",
                vec![
                    route("a", 0, &[(0.0, 0.0), (4.0, 0.0)]),
                    route("b", 1, &[(2.0, 0.0), (2.0, 0.0)]),
                ],
                &["a", "b"][..],
            ),
            (
                "endpoint touch",
                vec![
                    route("a", 0, &[(0.0, 0.0), (2.0, 0.0)]),
                    route("b", 1, &[(2.0, 0.0), (2.0, 2.0)]),
                ],
                &["a", "b"][..],
            ),
            (
                "bend touch",
                vec![
                    route("a", 0, &[(0.0, 0.0), (2.0, 0.0), (2.0, 2.0)]),
                    route("b", 1, &[(0.0, 2.0), (2.0, 0.0)]),
                ],
                &["a", "b"][..],
            ),
            (
                "self crossing",
                vec![route(
                    "a",
                    0,
                    &[(0.0, 0.0), (2.0, 2.0), (2.0, 0.0), (0.0, 2.0)],
                )],
                &["a"][..],
            ),
        ] {
            for code in [0, 1, 2, 4, 5] {
                let mut primitives = primitives.clone();
                apply_line_jumps(&mut primitives, &plain(ids), &settings(code));
                assert_eq!(arcs(&primitives), 0, "{name} code {code}");
            }
        }
    }

    #[test]
    fn unusable_jump_factors_leave_the_route_intact() {
        for (factor, line_to) in [
            (0.0, 0.1),
            (-0.5, 0.1),
            (0.5, 0.0),
            (0.5, -0.1),
            (1e300, 0.1),
        ] {
            let mut primitives = vec![
                route("a", 0, &[(0.0, 1.0), (4.0, 1.0)]),
                route("b", 1, &[(2.0, 0.0), (2.0, 2.0)]),
            ];
            let mut jump = settings(1);
            jump.factor_x = factor;
            jump.factor_y = factor;
            jump.line_to_x = line_to;
            jump.line_to_y = line_to;
            apply_line_jumps(&mut primitives, &plain(&["a", "b"]), &jump);
            assert_eq!(arcs(&primitives), 0, "factor {factor} line_to {line_to}");
        }
    }

    #[test]
    fn exhausting_the_jump_budget_leaves_every_route_unbridged() {
        let count = MAX_PLACED_JUMPS / 2 + 8;
        let mut primitives: Vec<Primitive> = (0..count)
            .map(|index| {
                let y = index as f64 * 1e-4;
                route(&format!("h{index}"), index as u32, &[(0.0, y), (4.0, y)])
            })
            .collect();
        primitives.extend((0..2).map(|index| {
            let x = 1.0 + index as f64;
            route(
                &format!("v{index}"),
                (count + index) as u32,
                &[(x, -1.0), (x, count as f64)],
            )
        }));
        let ids: Vec<String> = (0..count)
            .map(|index| format!("h{index}"))
            .chain((0..2).map(|index| format!("v{index}")))
            .collect();
        let overrides: Vec<ConnectorJumpOverride> = ids
            .iter()
            .map(|id| ConnectorJumpOverride {
                id: id.clone(),
                code: 0,
                style: 0,
                dir_x: 0,
                dir_y: 0,
            })
            .collect();
        apply_line_jumps(&mut primitives, &overrides, &settings(1));
        assert_eq!(arcs(&primitives), 0);
    }
}
