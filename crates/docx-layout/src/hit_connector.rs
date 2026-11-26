use serde_json::{Number, Value};

use crate::display_list::{ShapePathCommand, ShapePrimitive};

type Point = (f64, f64);

pub(super) fn contains(shape: &ShapePrimitive, x: f64, y: f64, min_radius: f64) -> bool {
    let attrs = &shape.attrs;
    let num = |value: &Number| value.as_f64().unwrap_or_default();
    if attrs
        .primitive_opacity
        .as_ref()
        .is_some_and(|opacity| num(opacity) <= 0.0)
    {
        return false;
    }
    if let Some(group) = &attrs.clip_group
        && let Some(clip) = &group.clip
    {
        let px = |value: &Option<Number>| value.as_ref().map_or(0.0, num);
        let (left, top, width, height) = (px(&clip.x), px(&clip.y), px(&clip.w), px(&clip.h));
        if width <= 0.0
            || height <= 0.0
            || x < left
            || x > left + width
            || y < top
            || y > top + height
            || group
                .opacity
                .as_ref()
                .is_some_and(|opacity| num(opacity) <= 0.0)
        {
            return false;
        }
    }
    let stroke = shape.stroke.as_ref();
    let paint = |key: &str| attrs.stroke_paint.as_ref().and_then(|paint| paint.get(key));
    let width = paint("width")
        .and_then(Value::as_f64)
        .or_else(|| stroke.map(|stroke| num(&stroke.width)))
        .unwrap_or(if attrs.stroke_paint.is_some() {
            1.0
        } else {
            0.0
        });
    let color = paint("color")
        .and_then(Value::as_str)
        .or_else(|| stroke.map(|stroke| stroke.color.as_str()));
    if width <= 0.0
        || color.is_some_and(|color| {
            matches!(color, "transparent" | "none")
                || matches!(
                    color.as_bytes(),
                    [b'#', _, _, _, b'0'] | [b'#', _, _, _, _, _, _, b'0', b'0']
                )
        })
    {
        return false;
    }
    let point = if let Some(transform) = &shape.transform {
        let cx = num(&shape.x) + num(&shape.w) / 2.0;
        let cy = num(&shape.y) + num(&shape.h) / 2.0;
        let degrees = transform.rotation.as_ref().map_or(0.0, num);
        let (sin, cos) = crate::fmath::sin_cos(degrees.to_radians());
        let (dx, dy) = (x - cx, y - cy);
        let (u, v) = (dx * cos + dy * sin, dy * cos - dx * sin);
        (
            cx + if transform.flip_h { -u } else { u },
            cy + if transform.flip_v { -v } else { v },
        )
    } else {
        (x, y)
    };
    let cap = paint("cap").and_then(Value::as_str).unwrap_or("flat");
    let join = paint("join").and_then(Value::as_str).unwrap_or("miter");
    let miter = paint("miterLimit")
        .and_then(Value::as_f64)
        .unwrap_or(4.0)
        .max(1.0);
    let dash = paint("dash")
        .and_then(Value::as_str)
        .or_else(|| stroke.and_then(|stroke| stroke.dash.as_deref()));
    let stroke = Stroke {
        point,
        radius: (width / 2.0).max(min_radius),
        cap,
        join,
        miter,
        dash: dash_pattern(paint("customDash"), dash, width),
    };
    let mut paths: Vec<Vec<Point>> = Vec::new();
    let mut tangents = Vec::new();
    for command in &shape.geometry_path {
        use ShapePathCommand::*;
        let (end, controls) = match command {
            Move { x, y } => {
                let point = (num(x), num(y));
                paths.push(vec![point]);
                tangents.push(point);
                continue;
            }
            Line { x, y } => ((num(x), num(y)), None),
            Quad { cpx, cpy, x, y } => ((num(x), num(y)), Some(((num(cpx), num(cpy)), None))),
            Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => (
                (num(x), num(y)),
                Some(((num(cp1x), num(cp1y)), Some((num(cp2x), num(cp2y))))),
            ),
            Close => {
                if let Some(path) = paths.last_mut() {
                    path.push(path[0]);
                }
                continue;
            }
        };
        let Some(path) = paths.last_mut() else {
            continue;
        };
        let start = *path.last().unwrap();
        if let Some((a, b)) = controls {
            tangents.push(a);
            let (a, b) = if let Some(b) = b {
                tangents.push(b);
                (a, b)
            } else {
                (lerp(start, a, 2.0 / 3.0), lerp(end, a, 2.0 / 3.0))
            };
            flatten([start, a, b, end], path, 0);
        } else {
            path.push(end);
        }
        tangents.push(end);
    }
    if paths.iter_mut().any(|path| {
        path.dedup();
        stroke.contains(path, false)
            || (min_radius > 0.0
                && cap == "flat"
                && path.first() != path.last()
                && path
                    .first()
                    .into_iter()
                    .chain(path.last())
                    .any(|end| (point.0 - end.0).hypot(point.1 - end.1) <= stroke.radius))
    }) {
        return true;
    }
    tangents.dedup();
    let len = tangents.len();
    len >= 2
        && [
            (paint("headEnd"), tangents[0], tangents[1]),
            (paint("tailEnd"), tangents[len - 1], tangents[len - 2]),
        ]
        .into_iter()
        .any(|(end, tip, toward)| end.is_some_and(|end| line_end(end, tip, toward, width, &stroke)))
}

fn dash_pattern(custom: Option<&Value>, name: Option<&str>, width: f64) -> Vec<f64> {
    let gap = (width * 2.0).max(2.0);
    let dash = (width * 3.0).max(2.0);
    let custom = custom
        .and_then(Value::as_array)
        .filter(|values| !values.is_empty());
    let mut pattern = if let Some(values) = custom {
        values
            .iter()
            .map(|value| value.as_f64().unwrap_or(0.0).max(0.0) * width)
            .collect()
    } else {
        match name.unwrap_or_default() {
            "dot" | "dotted" | "sysDot" => vec![width.max(1.0), gap],
            "dash" | "dashed" | "dashSmallGap" | "sysDash" => vec![dash, gap],
            "lgDash" | "dashLong" | "dashLongHeavy" => vec![(width * 6.0).max(4.0), gap],
            "dashDot" | "lgDashDot" | "sysDashDot" | "dashDotHeavy" => vec![dash, gap, width, gap],
            "dashDotDot" | "lgDashDotDot" | "sysDashDotDot" | "dashDotDotHeavy" => {
                vec![dash, gap, width, gap, width, gap]
            }
            _ => Vec::new(),
        }
    };
    if !pattern.iter().sum::<f64>().is_finite() || pattern.iter().all(|length| *length == 0.0) {
        pattern.clear();
    } else if pattern.len() % 2 != 0 {
        pattern.extend_from_within(..);
    }
    pattern
}

struct Stroke<'a> {
    point: Point,
    radius: f64,
    cap: &'a str,
    join: &'a str,
    miter: f64,
    dash: Vec<f64>,
}

impl Stroke<'_> {
    fn contains(&self, path: &[Point], solid: bool) -> bool {
        let closed = path.len() > 2 && path.first() == path.last();
        let period = if solid {
            0.0
        } else {
            self.dash.iter().sum::<f64>()
        };
        let mut phase = 0.0;
        for (index, pair) in path.windows(2).enumerate() {
            let (a, b) = (pair[0], pair[1]);
            let (dx, dy) = (b.0 - a.0, b.1 - a.1);
            let length = dx.hypot(dy);
            if length == 0.0 {
                continue;
            }
            let (u, v) = (dx / length, dy / length);
            let (x, y) = (self.point.0 - a.0, self.point.1 - a.1);
            let along = x * u + y * v;
            let across = (y * u - x * v).abs();
            if index > 0
                && (period == 0.0 || self.dash_crosses_vertex(phase))
                && self.join_contains(path[index - 1], a, b)
            {
                return true;
            }
            if across <= self.radius {
                let first = index == 0 && !closed;
                let last = index + 2 == path.len() && !closed;
                if period == 0.0 {
                    if self.interval(along, across, 0.0, length, first, last) {
                        return true;
                    }
                } else if self.dash_contains(along, across, length, phase, period, (first, last)) {
                    return true;
                }
            }
            if period > 0.0 {
                phase = (phase + length.rem_euclid(period)).rem_euclid(period);
            }
        }
        closed
            && period == 0.0
            && path.len() > 3
            && self.join_contains(path[path.len() - 2], path[0], path[1])
    }

    fn dash_crosses_vertex(&self, phase: f64) -> bool {
        let mut start = 0.0;
        self.dash.chunks_exact(2).any(|pair| {
            let hit = phase > start && phase < start + pair[0];
            start += pair[0] + pair[1];
            hit
        })
    }

    fn dash_contains(
        &self,
        along: f64,
        across: f64,
        length: f64,
        phase: f64,
        period: f64,
        (starts_path, ends_path): (bool, bool),
    ) -> bool {
        let nearest = along.clamp(0.0, length);
        let offset = (phase + nearest.rem_euclid(period)).rem_euclid(period);
        let mut start = 0.0;
        for pair in self.dash.chunks_exact(2) {
            let relative = start - offset;
            for shift in [-period, 0.0, period] {
                let first = nearest + (relative + shift);
                let last = nearest + (relative + shift + pair[0]);
                if last >= 0.0
                    && first <= length
                    && self.interval(
                        along,
                        across,
                        first.max(0.0),
                        last.min(length),
                        starts_path || first >= 0.0,
                        ends_path || last <= length,
                    )
                {
                    return true;
                }
            }
            start += pair[0] + pair[1];
        }
        false
    }

    fn interval(
        &self,
        along: f64,
        across: f64,
        start: f64,
        end: f64,
        first: bool,
        last: bool,
    ) -> bool {
        if along >= start && along <= end && end > start {
            return true;
        }
        let distance = if along <= start && first {
            start - along
        } else if along >= end && last {
            along - end
        } else {
            return false;
        };
        match self.cap {
            "square" => distance <= self.radius,
            "round" => distance.hypot(across) <= self.radius,
            _ => false,
        }
    }

    fn join_contains(&self, a: Point, b: Point, c: Point) -> bool {
        let (u, v) = (unit(a, b), unit(b, c));
        let turn = u.0 * v.1 - u.1 * v.0;
        if turn.abs() < 1e-9 {
            return false;
        }
        let r = self.radius * turn.signum();
        let n = (u.1 * r, -u.0 * r);
        let m = (v.1 * r, -v.0 * r);
        let delta = (self.point.0 - b.0, self.point.1 - b.1);
        if self.join == "round" {
            return delta.0.hypot(delta.1) <= self.radius;
        }
        if polygon_contains(&[(0.0, 0.0), n, m], delta) {
            return true;
        }
        let scale = 1.0 + u.0 * v.0 + u.1 * v.1;
        if self.join != "miter" || scale <= 0.0 {
            return false;
        }
        let tip = ((n.0 + m.0) / scale, (n.1 + m.1) / scale);
        tip.0.hypot(tip.1) <= self.miter * self.radius && polygon_contains(&[n, tip, m], delta)
    }
}

fn line_end(end: &Value, tip: Point, toward: Point, width: f64, stroke: &Stroke<'_>) -> bool {
    let Some(kind) = end
        .get("type")
        .and_then(Value::as_str)
        .filter(|kind| *kind != "none")
    else {
        return false;
    };
    let scale = |key| match end.get(key).and_then(Value::as_str) {
        Some("sm") => 0.75,
        Some("lg") => 1.5,
        _ => 1.0,
    };
    let (u, v) = unit(tip, toward);
    let half_width = 2.0 * width * scale("width");
    if kind == "oval" {
        return (tip.0 - stroke.point.0).hypot(tip.1 - stroke.point.1) <= half_width;
    }
    let length = 4.0 * width * scale("length");
    let base = (tip.0 + u * length, tip.1 + v * length);
    let normal = (-v * half_width, u * half_width);
    let mut points = vec![tip, (base.0 + normal.0, base.1 + normal.1)];
    if kind == "diamond" {
        points.push((base.0 + u * length, base.1 + v * length));
    }
    points.push((base.0 - normal.0, base.1 - normal.1));
    if kind == "arrow" {
        points.push(tip);
        stroke.contains(&points, true)
    } else {
        polygon_contains(&points, stroke.point)
    }
}

fn polygon_contains(points: &[Point], point: Point) -> bool {
    let (mut positive, mut negative) = (false, false);
    for (&a, &b) in points.iter().zip(points.iter().cycle().skip(1)) {
        let cross = (b.0 - a.0) * (point.1 - a.1) - (b.1 - a.1) * (point.0 - a.0);
        positive |= cross > 0.0;
        negative |= cross < 0.0;
    }
    !(positive && negative)
}

fn unit(a: Point, b: Point) -> Point {
    let (dx, dy) = (b.0 - a.0, b.1 - a.1);
    let length = dx.hypot(dy);
    (dx / length, dy / length)
}

fn lerp(a: Point, b: Point, t: f64) -> Point {
    (a.0 + (b.0 - a.0) * t, a.1 + (b.1 - a.1) * t)
}

fn flatten([a, b, c, d]: [Point; 4], points: &mut Vec<Point>, depth: u8) {
    let distance = |point: Point| {
        let (dx, dy) = (d.0 - a.0, d.1 - a.1);
        let length_squared = dx * dx + dy * dy;
        let t = if length_squared > 0.0 {
            (((point.0 - a.0) * dx + (point.1 - a.1) * dy) / length_squared).clamp(0.0, 1.0)
        } else {
            0.0
        };
        let nearest = lerp(a, d, t);
        (point.0 - nearest.0).hypot(point.1 - nearest.1)
    };
    if depth == 12 || distance(b).max(distance(c)) <= 0.125 {
        points.push(d);
        return;
    }
    let (ab, bc, cd) = (lerp(a, b, 0.5), lerp(b, c, 0.5), lerp(c, d, 0.5));
    let (abc, bcd) = (lerp(ab, bc, 0.5), lerp(bc, cd, 0.5));
    let middle = lerp(abc, bcd, 0.5);
    flatten([a, ab, abc, middle], points, depth + 1);
    flatten([middle, bcd, cd, d], points, depth + 1);
}
