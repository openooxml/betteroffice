use tiny_skia::{Path, PathBuilder, PathSegment, Point, Stroke, Transform};

/// Clips an open polyline and keeps the dash phase measured from its original start.
pub fn clip_dashed_polyline(
    path: &Path,
    stroke: &Stroke,
    transform: Transform,
    size: (u32, u32),
    period: f64,
) -> Option<Vec<(Path, f32)>> {
    let project = |point: Point| {
        let (x, y) = (f64::from(point.x), f64::from(point.y));
        [
            x * f64::from(transform.sx) + y * f64::from(transform.kx) + f64::from(transform.tx),
            x * f64::from(transform.ky) + y * f64::from(transform.sy) + f64::from(transform.ty),
        ]
    };
    let padding = |a: f32, b: f32| {
        2.0 + f64::from(stroke.width)
            * f64::from(stroke.miter_limit.max(1.0))
            * f64::from(a).hypot(f64::from(b))
    };
    let pad_x = padding(transform.sx, transform.kx);
    let pad_y = padding(transform.ky, transform.sy);
    let low = [-pad_x, -pad_y];
    let high = [f64::from(size.0) + pad_x, f64::from(size.1) + pad_y];
    if !period.is_finite() || period <= 0.0 {
        return None;
    }
    let inside = |point: Point| {
        let [x, y] = project(point);
        (low[0]..=high[0]).contains(&x) && (low[1]..=high[1]).contains(&y)
    };
    let mut segments = path.segments();
    let Some(PathSegment::MoveTo(mut previous)) = segments.next() else {
        return None;
    };
    let mut all_inside = inside(previous);
    for segment in segments {
        let PathSegment::LineTo(point) = segment else {
            return None;
        };
        all_inside &= inside(point);
    }
    if all_inside {
        return None;
    }
    let mut fragments = Vec::new();
    let mut builder = PathBuilder::new();
    let mut active = false;
    let mut phase = 0.0;
    let mut distance = 0.0;
    for segment in path.segments().skip(1) {
        let PathSegment::LineTo(point) = segment else {
            unreachable!();
        };
        let dx = f64::from(point.x) - f64::from(previous.x);
        let dy = f64::from(point.y) - f64::from(previous.y);
        let length = dx.hypot(dy);
        if length == 0.0 {
            continue;
        }
        if let Some((enter, exit)) = clip_interval(project(previous), project(point), low, high) {
            if !active {
                phase = (distance + length * enter).rem_euclid(period) as f32;
                builder.move_to(
                    (f64::from(previous.x) + dx * enter) as f32,
                    (f64::from(previous.y) + dy * enter) as f32,
                );
            }
            builder.line_to(
                (f64::from(previous.x) + dx * exit) as f32,
                (f64::from(previous.y) + dy * exit) as f32,
            );
            active = exit == 1.0;
        } else {
            active = false;
        }
        if !active && let Some(path) = std::mem::replace(&mut builder, PathBuilder::new()).finish()
        {
            fragments.push((path, phase));
        }
        distance = (distance + length).rem_euclid(period);
        previous = point;
    }
    if let Some(path) = builder.finish() {
        fragments.push((path, phase));
    }
    Some(fragments)
}

fn clip_interval(a: [f64; 2], b: [f64; 2], low: [f64; 2], high: [f64; 2]) -> Option<(f64, f64)> {
    let (mut enter, mut exit) = (0.0_f64, 1.0_f64);
    for axis in 0..2 {
        let delta = b[axis] - a[axis];
        if delta == 0.0 {
            if !(low[axis]..=high[axis]).contains(&a[axis]) {
                return None;
            }
            continue;
        }
        let t0 = (low[axis] - a[axis]) / delta;
        let t1 = (high[axis] - a[axis]) / delta;
        enter = enter.max(t0.min(t1));
        exit = exit.min(t0.max(t1));
    }
    (enter < exit).then_some((enter, exit))
}

/// An upper bound on the length a stroke walks, since a curve is no longer
/// than the control polygon it is drawn from.
pub fn stroke_path_length(path: &Path) -> f32 {
    let mut total = 0.0;
    let mut cursor = Point::zero();
    let mut start = Point::zero();
    let mut step = |from: Point, to: Point| {
        total += (to.x - from.x).hypot(to.y - from.y);
    };
    for segment in path.segments() {
        match segment {
            PathSegment::MoveTo(point) => {
                cursor = point;
                start = point;
            }
            PathSegment::LineTo(point) => {
                step(cursor, point);
                cursor = point;
            }
            PathSegment::QuadTo(control, point) => {
                step(cursor, control);
                step(control, point);
                cursor = point;
            }
            PathSegment::CubicTo(first, second, point) => {
                step(cursor, first);
                step(first, second);
                step(second, point);
                cursor = point;
            }
            PathSegment::Close => {
                step(cursor, start);
                cursor = start;
            }
        }
    }
    total
}
