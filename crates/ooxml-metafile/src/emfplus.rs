//! EMF+ (`[MS-EMFPLUS]`) records carried in EMF comments.
//!
//! A dual metafile carries a GDI rendition of the same picture as well; the
//! replay prefers the EMF+ records, as Office does, and falls back to the GDI
//! ones when the EMF+ records hold something it cannot draw. An EMF+-only
//! metafile plays its EMF+ records here, and its GDI records only inside the
//! spans an `EmfPlusGetDC` record opens.

use std::rc::Rc;
use std::sync::Arc;

use crate::drawing::{
    Bitmap, Clip, ClipRegion, Drawing, Font, Image, LineCap, LineJoin, LinearGradient, Op, Paint,
    PathCommand, Pixels, Rgba, Shape, Spread, Stroke, Text, TextAnchor,
};
use crate::player::{ARC_SEGMENTS, IDENTITY, Player, Xform, apply, chain, concat, rect_path};
use crate::read::{finite_at, i16_at, i32_at, u8_at, u16_at, u32_at};

const HEADER: u16 = 0x4001;
const MAX_OBJECT_BYTES: usize = 16 * 1024 * 1024;
const MAX_STACK: usize = 1024;

#[derive(Clone)]
enum Object {
    Brush(Paint),
    Pen(Pen),
    Path(Arc<Path>),
    Region(Arc<Region>),
    Image(Arc<PlusImage>),
    Font(Arc<PlusFont>),
    Format(Format),
    Other,
}

#[derive(Clone)]
struct Pen {
    width: f64,
    world_width: bool,
    dash: Option<Vec<f64>>,
    cap: LineCap,
    join: LineJoin,
    miter_limit: f64,
    paint: Paint,
}

/// A GDI+ path in world units: figures of straight and cubic segments.
struct Path {
    figures: Vec<Figure>,
}

struct Figure {
    start: (f64, f64),
    segments: Vec<Segment>,
    closed: bool,
}

#[derive(Clone, Copy)]
enum Segment {
    Line((f64, f64)),
    Cubic([(f64, f64); 3]),
}

enum Region {
    Infinite,
    Empty,
    Path(Arc<Path>),
    /// A boolean combination, with its extent when its Infinite and Empty
    /// operands decide it (see [`extent`]).
    Combine(u32, Box<Region>, Box<Region>, Option<bool>),
}

enum PlusImage {
    Bitmap(Arc<Bitmap>),
    Metafile(Nested),
}

/// A metafile image, replayed once when its object is defined.
enum Nested {
    TooDeep,
    Failed,
    /// The drawing, and the path commands each placement of it copies.
    Drawn(Drawing, usize),
}

struct PlusFont {
    size: f64,
    unit: u32,
    style: i32,
    family: String,
}

#[derive(Clone, Copy, Default)]
struct Format {
    flags: u32,
    align: u32,
    line_align: u32,
}

#[derive(Clone)]
struct Graphics {
    world: Xform,
    page_unit: u32,
    page_scale: f64,
    clip: Clip,
}

pub(crate) struct State {
    dual: bool,
    dpi: (f64, f64),
    objects: Vec<Option<Object>>,
    graphics: Graphics,
    stack: Vec<(u32, Graphics)>,
    continued: Option<(u16, usize, Vec<u8>)>,
    records: usize,
}

impl State {
    fn new(dual: bool, dpi: (f64, f64)) -> Self {
        Self {
            dual,
            dpi,
            objects: vec![None; 64],
            graphics: Graphics {
                world: IDENTITY,
                page_unit: 2,
                page_scale: 1.0,
                clip: None,
            },
            stack: Vec::new(),
            continued: None,
            records: 0,
        }
    }

    /// Device pixels per `unit` (`UnitType`) along x and y, at the reference
    /// device's DPI.
    fn unit_scale(&self, unit: u32) -> (f64, f64) {
        let per_inch = match unit {
            3 => 72.0,
            4 => 1.0,
            5 => 300.0,
            6 => 25.4,
            _ => return (1.0, 1.0),
        };
        (self.dpi.0 / per_inch, self.dpi.1 / per_inch)
    }

    /// World units to device pixels.
    fn to_device(&self) -> Xform {
        let (x, y) = self.unit_scale(self.graphics.page_unit);
        let scale = self.graphics.page_scale;
        concat(
            self.graphics.world,
            [x * scale, 0.0, 0.0, y * scale, 0.0, 0.0],
        )
    }
}

/// An EMF comment record; an EMF+ one plays or opens EMF+ playback.
pub(crate) fn comment<const FULL: bool>(player: &mut Player<FULL>, bytes: &[u8]) -> Option<()> {
    if u32_at(bytes, 12) != Some(crate::emf::EMF_PLUS) {
        return Some(());
    }
    let size = u32_at(bytes, 8)? as usize;
    let Some(end) = 12usize.checked_add(size).filter(|end| *end <= bytes.len()) else {
        return player.refuse("an EMF+ comment runs past its record");
    };
    let mut at = 16;
    while at + 12 <= end {
        let kind = u16_at(bytes, at)?;
        let flags = u16_at(bytes, at + 2)?;
        let size = u32_at(bytes, at + 4)? as usize;
        let data_size = u32_at(bytes, at + 8)? as usize;
        if size < 12 || size > end - at || data_size > size - 12 {
            return player.refuse("an EMF+ record has an invalid size");
        }
        let data = &bytes[at + 12..at + 12 + data_size];
        player.spend(1, 0)?;
        match player.plus.as_deref() {
            None if kind == HEADER => {
                let dpi = (
                    f64::from(u32_at(data, 8).unwrap_or(96).max(1)),
                    f64::from(u32_at(data, 12).unwrap_or(96).max(1)),
                );
                let dual = flags & 1 != 0 && player.gdi_records && !player.prefer_plus;
                player.plus = Some(Box::new(State::new(dual, dpi)));
                player.plus_only = !dual;
                if dual {
                    return Some(());
                }
            }
            None => return Some(()),
            Some(state) if state.dual => return Some(()),
            Some(_) => {
                let records = player.plus.as_deref_mut().map_or(0, |state| {
                    state.records += 1;
                    state.records
                });
                if records > player.limits.records {
                    return player.refuse("the EMF+ records exceed the replay limit");
                }
                if player.plus_gdi {
                    player.flush_pending();
                    player.plus_gdi = false;
                }
                if !record(player, kind, flags, data)? {
                    return player.refuse(format!("EMF+ record type {kind:#06x} is not supported"));
                }
                if player.overflowed {
                    return player.refuse("the metafile draws more than the replay limits");
                }
            }
        }
        at += size;
    }
    Some(())
}

/// Plays one EMF+ record; `Some(false)` when the record type is unknown.
fn record<const FULL: bool>(
    player: &mut Player<FULL>,
    kind: u16,
    flags: u16,
    data: &[u8],
) -> Option<bool> {
    let mut state = player.plus.take()?;
    let known = play(player, &mut state, kind, flags, data);
    player.plus = Some(state);
    known
}

fn play<const FULL: bool>(
    player: &mut Player<FULL>,
    state: &mut State,
    kind: u16,
    flags: u16,
    data: &[u8],
) -> Option<bool> {
    let compressed = flags & 0x4000 != 0;
    let relative = flags & 0x0800 != 0;
    let solid = flags & 0x8000 != 0;
    let id = usize::from(flags & 0xff);
    match kind {
        0x4001 | 0x4002 | 0x4003 | 0x4005 | 0x4006 | 0x4007 | 0x401D | 0x401E | 0x401F | 0x4020
        | 0x4021 | 0x4022 | 0x4023 | 0x4024 | 0x4038 | 0x4039 | 0x403A => {}
        0x4004 => player.plus_gdi = true,
        0x4008 => object(player, state, flags, data)?,
        0x4009 => {
            let color = Rgba::from_argb(u32_at(data, 0)?);
            let (w, h) = player.unit;
            let clip = state.graphics.clip.clone();
            push_shape(
                player,
                rect_path([0.0, 0.0, w, h]),
                Some(Paint::Solid(color)),
                None,
                false,
                clip,
            );
        }
        0x400A | 0x400B => {
            let (paint, at) = if kind == 0x400A {
                (Some(brush(player, state, solid, u32_at(data, 0)?)?), 4)
            } else {
                (None, 0)
            };
            let count = u32_at(data, at)? as usize;
            let rects = rects(
                data,
                at + 4,
                count,
                compressed,
                player.limits.points_per_record,
            )?;
            let mut path = Vec::with_capacity(rects.len() * 5);
            for rect in rects {
                figure_path(state, player, &rect_figure(rect), &mut path);
            }
            let stroke = if kind == 0x400B {
                pen(state, id)?.map(|pen| stroke(state, player, &pen))
            } else {
                None
            };
            let clip = state.graphics.clip.clone();
            push_shape(player, path, paint, stroke, false, clip);
        }
        0x400C | 0x400D => {
            let (paint, at) = if kind == 0x400C {
                (Some(brush(player, state, solid, u32_at(data, 0)?)?), 4)
            } else {
                (None, 0)
            };
            let count = u32_at(data, at)? as usize;
            let points = points(
                data,
                at + 4,
                count,
                compressed,
                relative,
                player.limits.points_per_record,
            )?;
            let closed = kind == 0x400C || flags & 0x2000 != 0;
            let figure = polyline(&points, closed)?;
            let mut path = Vec::new();
            figure_path(state, player, &figure, &mut path);
            let stroke = if kind == 0x400D {
                pen(state, id)?.map(|pen| stroke(state, player, &pen))
            } else {
                None
            };
            let clip = state.graphics.clip.clone();
            push_shape(player, path, paint, stroke, kind == 0x400C, clip);
        }
        0x400E..=0x4012 => {
            let (paint, mut at) = match kind {
                0x400E | 0x4010 => (Some(brush(player, state, solid, u32_at(data, 0)?)?), 4),
                _ => (None, 0),
            };
            let (start, sweep) = if kind >= 0x4010 {
                let angles = (finite_at(data, at)?, finite_at(data, at + 4)?);
                at += 8;
                Some(angles)
            } else {
                None
            }
            .unwrap_or((0.0, 360.0));
            let rect = rects(data, at, 1, compressed, 1)?[0];
            let pie = matches!(kind, 0x4010 | 0x4011);
            let figure = arc_figure(rect, start, sweep, pie, kind < 0x4010 || pie);
            let mut path = Vec::new();
            figure_path(state, player, &figure, &mut path);
            let stroke = if paint.is_none() {
                pen(state, id)?.map(|pen| stroke(state, player, &pen))
            } else {
                None
            };
            let clip = state.graphics.clip.clone();
            push_shape(player, path, paint, stroke, false, clip);
        }
        0x4013 => {
            let paint = brush(player, state, solid, u32_at(data, 0)?)?;
            let Some(Some(Object::Region(region))) = state.objects.get(id).cloned() else {
                return Some(true);
            };
            let mut path = Vec::new();
            if !region_path(state, player, &region, &mut path)? {
                player.omit("EMF+ regions approximated")?;
            }
            let clip = state.graphics.clip.clone();
            push_shape(player, path, Some(paint), None, true, clip);
        }
        0x4014 | 0x4015 => {
            let Some(Some(Object::Path(shape))) = state.objects.get(id).cloned() else {
                return Some(true);
            };
            let (paint, stroke) = if kind == 0x4014 {
                (Some(brush(player, state, solid, u32_at(data, 0)?)?), None)
            } else {
                let pen = pen(state, u32_at(data, 0)? as usize)?;
                (None, pen.map(|pen| stroke(state, player, &pen)))
            };
            if paint.is_none() && stroke.is_none() {
                return Some(true);
            }
            let mut path = Vec::new();
            for figure in &shape.figures {
                figure_path(state, player, figure, &mut path);
            }
            let clip = state.graphics.clip.clone();
            push_shape(player, path, paint, stroke, kind == 0x4014, clip);
        }
        0x4016..=0x4018 => {
            let (paint, at) = if kind == 0x4016 {
                (Some(brush(player, state, solid, u32_at(data, 0)?)?), 4)
            } else {
                (None, 0)
            };
            let tension = finite_at(data, at)?;
            let (at, segments) = if kind == 0x4018 {
                (
                    at + 12,
                    Some((
                        u32_at(data, at + 4)? as usize,
                        u32_at(data, at + 8)? as usize,
                    )),
                )
            } else {
                (at + 4, None)
            };
            let count = u32_at(data, at)? as usize;
            let points = points(
                data,
                at + 4,
                count,
                compressed,
                relative && kind != 0x4018,
                player.limits.points_per_record,
            )?;
            let figure = cardinal(&points, tension, kind != 0x4018, segments)?;
            let mut path = Vec::new();
            figure_path(state, player, &figure, &mut path);
            let stroke = if paint.is_none() {
                pen(state, id)?.map(|pen| stroke(state, player, &pen))
            } else {
                None
            };
            let clip = state.graphics.clip.clone();
            push_shape(
                player,
                path,
                paint,
                stroke,
                kind == 0x4016 && flags & 0x2000 == 0,
                clip,
            );
        }
        0x4019 => {
            let count = u32_at(data, 0)? as usize;
            let points = points(
                data,
                4,
                count,
                compressed,
                relative,
                player.limits.points_per_record,
            )?;
            let (first, rest) = points.split_first()?;
            let figure = Figure {
                start: *first,
                segments: rest
                    .as_chunks::<3>()
                    .0
                    .iter()
                    .map(|chunk| Segment::Cubic(*chunk))
                    .collect(),
                closed: false,
            };
            let mut path = Vec::new();
            figure_path(state, player, &figure, &mut path);
            let stroke = pen(state, id)?.map(|pen| stroke(state, player, &pen));
            let clip = state.graphics.clip.clone();
            push_shape(player, path, None, stroke, false, clip);
        }
        0x401A | 0x401B => {
            let source = (
                finite_at(data, 8)?,
                finite_at(data, 12)?,
                finite_at(data, 16)?,
                finite_at(data, 20)?,
            );
            let corners = if kind == 0x401A {
                let rect = rects(data, 24, 1, compressed, 1)?[0];
                [
                    (rect.0, rect.1),
                    (rect.0 + rect.2, rect.1),
                    (rect.0, rect.1 + rect.3),
                ]
            } else {
                if u32_at(data, 24)? != 3 {
                    return player.refuse("an EMF+ image destination is not a parallelogram");
                }
                let corners = points(data, 28, 3, compressed, relative, 3)?;
                [corners[0], corners[1], corners[2]]
            };
            if u32_at(data, 4)? != 2 {
                player.omit("EMF+ image source rectangles in physical units")?;
            }
            draw_image(player, state, id, source, corners)?;
        }
        0x401C => draw_string(player, state, id, solid, data)?,
        0x4036 => draw_driver_string(player, state, id, solid, data)?,
        0x4025 | 0x4027 | 0x4028 => {
            let index = match kind {
                0x4027 => u32_at(data, 32)?,
                _ => u32_at(data, 0)?,
            };
            if state.stack.len() >= MAX_STACK {
                return player.refuse("EMF+ graphics states nest past the save limit");
            }
            state.stack.push((index, state.graphics.clone()));
            if kind == 0x4027 {
                let dest = (
                    finite_at(data, 0)?,
                    finite_at(data, 4)?,
                    finite_at(data, 8)?,
                    finite_at(data, 12)?,
                );
                let source = (
                    finite_at(data, 16)?,
                    finite_at(data, 20)?,
                    finite_at(data, 24)?,
                    finite_at(data, 28)?,
                );
                let (from, page) = (
                    state.unit_scale(u32::from(flags & 0xff)),
                    state.unit_scale(state.graphics.page_unit),
                );
                let unit = (
                    from.0 / page.0 / state.graphics.page_scale,
                    from.1 / page.1 / state.graphics.page_scale,
                );
                if source.2 != 0.0 && source.3 != 0.0 {
                    let (sx, sy) = (dest.2 / source.2 * unit.0, dest.3 / source.3 * unit.1);
                    let container = [
                        sx,
                        0.0,
                        0.0,
                        sy,
                        dest.0 * unit.0 - source.0 * sx,
                        dest.1 * unit.1 - source.1 * sy,
                    ];
                    state.graphics.world = concat(container, state.graphics.world);
                }
            }
        }
        0x4026 | 0x4029 => {
            let index = u32_at(data, 0)?;
            player.charge(state.stack.len(), 0)?;
            if let Some(position) = state.stack.iter().rposition(|(saved, _)| *saved == index) {
                state.graphics = state.stack[position].1.clone();
                state.stack.truncate(position);
            }
        }
        0x402A => state.graphics.world = matrix(data, 0)?,
        0x402B => state.graphics.world = IDENTITY,
        0x402C..=0x402F => {
            let m = match kind {
                0x402C => matrix(data, 0)?,
                0x402D => [1.0, 0.0, 0.0, 1.0, finite_at(data, 0)?, finite_at(data, 4)?],
                0x402E => [finite_at(data, 0)?, 0.0, 0.0, finite_at(data, 4)?, 0.0, 0.0],
                _ => {
                    let (sin, cos) = finite_at(data, 0)?.to_radians().sin_cos();
                    [cos, sin, -sin, cos, 0.0, 0.0]
                }
            };
            state.graphics.world = if flags & 0x2000 != 0 {
                concat(state.graphics.world, m)
            } else {
                concat(m, state.graphics.world)
            };
        }
        0x4030 => {
            state.graphics.page_unit = u32::from(flags & 0xff);
            let scale = finite_at(data, 0)?;
            state.graphics.page_scale = if scale > 0.0 { scale } else { 1.0 };
        }
        0x4031 => state.graphics.clip = None,
        0x4032..=0x4034 => {
            let mode = u32::from((flags >> 8) & 0x0f);
            let region = match kind {
                0x4032 => {
                    let rect = rects(data, 0, 1, false, 1)?[0];
                    Region::Path(Arc::new(Path {
                        figures: vec![rect_figure(rect)],
                    }))
                }
                0x4033 => match state.objects.get(id).cloned().flatten() {
                    Some(Object::Path(path)) => Region::Path(path),
                    _ => return Some(true),
                },
                _ => match state.objects.get(id).cloned().flatten() {
                    Some(Object::Region(region)) => {
                        return set_clip(player, state, &region, mode).map(|()| true);
                    }
                    _ => return Some(true),
                },
            };
            set_clip(player, state, &region, mode)?;
        }
        0x4035 => {
            let (dx, dy) = (finite_at(data, 0)?, finite_at(data, 4)?);
            let m = concat(state.to_device(), player.device_to_output());
            let (ox, oy) = (dx * m[0] + dy * m[2], dx * m[1] + dy * m[3]);
            let mut at = state.graphics.clip.as_deref();
            while let Some(link) = at {
                player.charge(link.region.path.len().saturating_add(1), 0)?;
                at = link.parent.as_deref();
            }
            state.graphics.clip = offset_chain(&state.graphics.clip, ox, oy);
        }
        0x4037 => player.omit("EMF+ stroke-and-fill paths")?,
        _ => return Some(false),
    }
    Some(true)
}

fn matrix(data: &[u8], at: usize) -> Option<Xform> {
    let mut m = IDENTITY;
    for (index, slot) in m.iter_mut().enumerate() {
        *slot = finite_at(data, at + index * 4)?;
    }
    Some(m)
}

fn push_shape<const FULL: bool>(
    player: &mut Player<FULL>,
    path: Vec<PathCommand>,
    fill: Option<Paint>,
    stroke: Option<Stroke>,
    even_odd: bool,
    clip: Clip,
) {
    if player.spend(path.len() as u64, 0).is_none() {
        return;
    }
    player.commands = player.commands.saturating_add(path.len());
    if player.commands > player.limits.commands {
        player.overflowed = true;
        return;
    }
    if path.is_empty() || (fill.is_none() && stroke.is_none()) {
        return;
    }
    player.push_op(Op::Shape(Shape {
        path,
        fill,
        stroke,
        even_odd,
        clip,
    }));
}

/// A brush object, or an ARGB colour when `solid`.
fn brush<const FULL: bool>(
    player: &Player<FULL>,
    state: &State,
    solid: bool,
    value: u32,
) -> Option<Paint> {
    if solid {
        return Some(Paint::Solid(Rgba::from_argb(value)));
    }
    Some(match state.objects.get(value as usize).cloned().flatten() {
        Some(Object::Brush(paint)) => placed(player, state, paint),
        _ => Paint::Solid(Rgba {
            a: 0,
            ..Rgba::BLACK
        }),
    })
}

/// A brush's gradient, kept in world units, at the current graphics transform.
fn placed<const FULL: bool>(player: &Player<FULL>, state: &State, paint: Paint) -> Paint {
    let Paint::Linear(gradient) = &paint else {
        return paint;
    };
    let m = concat(state.to_device(), player.device_to_output());
    Paint::Linear(Arc::new(crate::transform::linear(gradient, m)))
}

/// The pen object `id`, or `Some(None)` when the slot holds none.
fn pen(state: &State, id: usize) -> Option<Option<Pen>> {
    Some(match state.objects.get(id).cloned().flatten() {
        Some(Object::Pen(pen)) => Some(pen),
        _ => None,
    })
}

fn stroke<const FULL: bool>(state: &State, player: &Player<FULL>, pen: &Pen) -> Stroke {
    let device = state.to_device();
    let world_scale = (device[0] * device[3] - device[1] * device[2]).abs().sqrt();
    let pixels = if pen.world_width {
        pen.width * world_scale
    } else {
        pen.width
    };
    let pixel = player.device_pixel();
    let width = (pixels * pixel).max(pixel);
    Stroke {
        paint: placed(player, state, pen.paint.clone()),
        width,
        dash: pen
            .dash
            .as_ref()
            .map(|dash| dash.iter().map(|length| length * width).collect()),
        cap: pen.cap,
        join: pen.join,
        miter_limit: pen.miter_limit,
    }
}

fn read_point(data: &[u8], at: &mut usize, compressed: bool, relative: bool) -> Option<(f64, f64)> {
    if relative {
        let mut value = || -> Option<f64> {
            let first = u8_at(data, *at)?;
            if first & 0x80 == 0 {
                *at += 1;
                Some(f64::from(((first << 1) as i8) >> 1))
            } else {
                let raw = u16::from_be_bytes([first, u8_at(data, *at + 1)?]) & 0x7fff;
                *at += 2;
                Some(f64::from(((raw << 1) as i16) >> 1))
            }
        };
        return Some((value()?, value()?));
    }
    if compressed {
        let point = (
            f64::from(i16_at(data, *at)?),
            f64::from(i16_at(data, *at + 2)?),
        );
        *at += 4;
        Some(point)
    } else {
        let point = (finite_at(data, *at)?, finite_at(data, *at + 4)?);
        *at += 8;
        Some(point)
    }
}

fn points(
    data: &[u8],
    at: usize,
    count: usize,
    compressed: bool,
    relative: bool,
    limit: usize,
) -> Option<Vec<(f64, f64)>> {
    if count > limit || count > data.len() {
        return None;
    }
    let mut at = at;
    let mut points = Vec::with_capacity(count);
    let mut last = (0.0, 0.0);
    for _ in 0..count {
        let point = read_point(data, &mut at, compressed, relative)?;
        last = if relative {
            (last.0 + point.0, last.1 + point.1)
        } else {
            point
        };
        points.push(last);
    }
    Some(points)
}

fn rects(
    data: &[u8],
    at: usize,
    count: usize,
    compressed: bool,
    limit: usize,
) -> Option<Vec<(f64, f64, f64, f64)>> {
    if count > limit || count > data.len() {
        return None;
    }
    let mut rects = Vec::with_capacity(count);
    for index in 0..count {
        rects.push(if compressed {
            let at = at + index * 8;
            let side = |offset| i16_at(data, at + offset).map(f64::from);
            (side(0)?, side(2)?, side(4)?, side(6)?)
        } else {
            let at = at + index * 16;
            (
                finite_at(data, at)?,
                finite_at(data, at + 4)?,
                finite_at(data, at + 8)?,
                finite_at(data, at + 12)?,
            )
        });
    }
    Some(rects)
}

fn rect_figure((x, y, w, h): (f64, f64, f64, f64)) -> Figure {
    Figure {
        start: (x, y),
        segments: vec![
            Segment::Line((x + w, y)),
            Segment::Line((x + w, y + h)),
            Segment::Line((x, y + h)),
        ],
        closed: true,
    }
}

fn polyline(points: &[(f64, f64)], closed: bool) -> Option<Figure> {
    let (first, rest) = points.split_first()?;
    Some(Figure {
        start: *first,
        segments: rest.iter().map(|point| Segment::Line(*point)).collect(),
        closed,
    })
}

/// An arc of the ellipse inscribed in `rect`, angles in degrees clockwise from +x.
fn arc_figure(
    rect: (f64, f64, f64, f64),
    start: f64,
    sweep: f64,
    pie: bool,
    closed: bool,
) -> Figure {
    let (x, y, w, h) = rect;
    let (cx, cy, rx, ry) = (x + w / 2.0, y + h / 2.0, w / 2.0, h / 2.0);
    let at = |degrees: f64| {
        let (sin, cos) = degrees.to_radians().sin_cos();
        (cx + rx * cos, cy + ry * sin)
    };
    let sweep = sweep.clamp(-360.0, 360.0);
    let steps = ((sweep.abs() / 360.0) * ARC_SEGMENTS as f64)
        .ceil()
        .max(2.0) as usize;
    let arc: Vec<Segment> = (1..=steps)
        .map(|step| Segment::Line(at(start + sweep * step as f64 / steps as f64)))
        .collect();
    if pie {
        let mut segments = vec![Segment::Line(at(start))];
        segments.extend(arc);
        Figure {
            start: (cx, cy),
            segments,
            closed: true,
        }
    } else {
        Figure {
            start: at(start),
            segments: arc,
            closed,
        }
    }
}

/// A cardinal spline through `points` as cubic segments.
fn cardinal(
    points: &[(f64, f64)],
    tension: f64,
    closed: bool,
    segments: Option<(usize, usize)>,
) -> Option<Figure> {
    let count = points.len();
    if count < 2 {
        return None;
    }
    let k = tension / 3.0;
    let point = |index: isize| -> (f64, f64) {
        if closed {
            points[index.rem_euclid(count as isize) as usize]
        } else {
            points[index.clamp(0, count as isize - 1) as usize]
        }
    };
    let spans = if closed { count } else { count - 1 };
    let (first, taken) = segments.unwrap_or((0, spans));
    let mut figure = Figure {
        start: point(first as isize),
        segments: Vec::new(),
        closed,
    };
    for index in first..first.saturating_add(taken).min(spans) {
        let i = index as isize;
        let (p0, p1, p2, p3) = (point(i - 1), point(i), point(i + 1), point(i + 2));
        figure.segments.push(Segment::Cubic([
            (p1.0 + k * (p2.0 - p0.0), p1.1 + k * (p2.1 - p0.1)),
            (p2.0 - k * (p3.0 - p1.0), p2.1 - k * (p3.1 - p1.1)),
            p2,
        ]));
    }
    Some(figure)
}

/// Appends `figure`, in world units, to `path` in drawing units.
fn figure_path<const FULL: bool>(
    state: &State,
    player: &Player<FULL>,
    figure: &Figure,
    path: &mut Vec<PathCommand>,
) {
    let m = concat(state.to_device(), player.device_to_output());
    let (x, y) = apply(m, figure.start);
    path.push(PathCommand::Move { x, y });
    for segment in &figure.segments {
        match segment {
            Segment::Line(point) => {
                let (x, y) = apply(m, *point);
                path.push(PathCommand::Line { x, y });
            }
            Segment::Cubic([a, b, c]) => {
                let (a, b, c) = (apply(m, *a), apply(m, *b), apply(m, *c));
                path.push(PathCommand::Cubic {
                    cp1x: a.0,
                    cp1y: a.1,
                    cp2x: b.0,
                    cp2y: b.1,
                    x: c.0,
                    y: c.1,
                });
            }
        }
    }
    if figure.closed {
        path.push(PathCommand::Close);
    }
}

fn object<const FULL: bool>(
    player: &mut Player<FULL>,
    state: &mut State,
    flags: u16,
    data: &[u8],
) -> Option<()> {
    let id = usize::from(flags & 0xff);
    let kind = (flags >> 8) & 0x7f;
    let owned;
    let data = if flags & 0x8000 != 0 {
        let total = u32_at(data, 0)? as usize;
        if total > MAX_OBJECT_BYTES {
            return player.refuse("an EMF+ object exceeds the size limit");
        }
        let chunk = data.get(4..)?;
        let mut buffer = match state.continued.take() {
            Some((pending, expected, buffer)) if pending == flags && expected == total => buffer,
            _ => Vec::new(),
        };
        buffer.extend_from_slice(&chunk[..chunk.len().min(total - buffer.len())]);
        if buffer.len() < total {
            state.continued = Some((flags, total, buffer));
            return Some(());
        }
        owned = buffer;
        Some(&owned[..])
    } else {
        match state.continued.take() {
            Some((pending, total, mut buffer)) if pending & 0x7fff == flags => {
                buffer.extend_from_slice(&data[..data.len().min(total - buffer.len())]);
                owned = buffer;
                (owned.len() == total).then_some(&owned[..])
            }
            _ => Some(data),
        }
    };
    let parsed = data.and_then(|data| match kind {
        1 => parse_brush(player, data).map(Object::Brush),
        2 => parse_pen(player, state, data).map(Object::Pen),
        3 => parse_path(data, player.limits.points_per_record)
            .map(|(path, _)| Object::Path(Arc::new(path))),
        4 => {
            let mut nodes = (u32_at(data, 4).unwrap_or(0) as usize).saturating_add(1);
            parse_region(data, 8, &mut nodes, 0, player.limits.points_per_record)
                .map(|(region, _)| Object::Region(Arc::new(region)))
        }
        5 => parse_image(player, data).map(|image| Object::Image(Arc::new(image))),
        6 => parse_font(data).map(|font| Object::Font(Arc::new(font))),
        7 => Some(Object::Format(Format {
            flags: u32_at(data, 4).unwrap_or(0),
            align: u32_at(data, 12).unwrap_or(0),
            line_align: u32_at(data, 16).unwrap_or(0),
        })),
        _ => Some(Object::Other),
    });
    let object = match parsed {
        Some(object) => object,
        None if player.refusal.is_some() => return None,
        None => {
            player.omit("EMF+ objects that could not be decoded")?;
            Object::Other
        }
    };
    if let Some(slot) = state.objects.get_mut(id) {
        *slot = Some(object);
    }
    Some(())
}

fn parse_brush<const FULL: bool>(player: &mut Player<FULL>, data: &[u8]) -> Option<Paint> {
    let kind = u32_at(data, 4)?;
    let body = 8;
    Some(match kind {
        0 => Paint::Solid(Rgba::from_argb(u32_at(data, body)?)),
        1 => {
            let style = u32_at(data, body)?;
            let fore = Rgba::from_argb(u32_at(data, body + 4)?);
            let back = Rgba::from_argb(u32_at(data, body + 8)?);
            if style > 5 {
                return Some(Paint::Solid(mix(fore, back, 0.5)));
            }
            Paint::Hatch {
                style: crate::player::hatch_style(style),
                color: fore,
                background: (back.a > 0).then_some(back),
                cell: 8.0 * player.device_pixel(),
            }
        }
        2 => {
            let flags = u32_at(data, body)?;
            let mut at = body + 8;
            if flags & 0x02 != 0 {
                at += 24;
            }
            let image = parse_image(player, data.get(at..)?)?;
            let PlusImage::Bitmap(tile) = image else {
                return None;
            };
            let pixel = player.device_pixel();
            Paint::Pattern {
                width: f64::from(tile.width) * pixel,
                height: f64::from(tile.height) * pixel,
                tile,
            }
        }
        3 => {
            player.omit("EMF+ path gradients drawn flat")?;
            let center = Rgba::from_argb(u32_at(data, body + 8)?);
            let count = u32_at(data, body + 20)?;
            let surround = if count > 0 {
                Rgba::from_argb(u32_at(data, body + 24)?)
            } else {
                center
            };
            Paint::Solid(mix(center, surround, 0.5))
        }
        4 => {
            let flags = u32_at(data, body)?;
            let wrap = i32_at(data, body + 4)?;
            let rect = (
                finite_at(data, body + 8)?,
                finite_at(data, body + 12)?,
                finite_at(data, body + 16)?,
                finite_at(data, body + 20)?,
            );
            let (start, end) = (
                Rgba::from_argb(u32_at(data, body + 24)?),
                Rgba::from_argb(u32_at(data, body + 28)?),
            );
            let mut at = body + 40;
            let transform = if flags & 0x02 != 0 {
                let m = matrix(data, at)?;
                at += 24;
                m
            } else {
                IDENTITY
            };
            let mut stops = vec![(0.0, start), (1.0, end)];
            if flags & 0x04 != 0 {
                let count = (u32_at(data, at)? as usize).min(256);
                let mut preset = Vec::with_capacity(count);
                for index in 0..count {
                    let offset = finite_at(data, at + 4 + index * 4)?;
                    let color = Rgba::from_argb(u32_at(data, at + 4 + count * 4 + index * 4)?);
                    preset.push((offset, color));
                }
                if !preset.is_empty() {
                    stops = preset;
                }
            } else if flags & 0x18 != 0 {
                let count = (u32_at(data, at)? as usize).min(256);
                let mut blended = Vec::with_capacity(count);
                for index in 0..count {
                    let offset = finite_at(data, at + 4 + index * 4)?;
                    let factor = finite_at(data, at + 4 + count * 4 + index * 4)?;
                    blended.push((offset, mix(start, end, factor)));
                }
                if !blended.is_empty() {
                    stops = blended;
                }
            }
            stops.sort_by(|a, b| a.0.total_cmp(&b.0));
            let gradient = LinearGradient {
                start: (rect.0, rect.1),
                end: (rect.0 + rect.2, rect.1),
                stops: stops.into(),
                spread: match wrap {
                    4 => Spread::Pad,
                    1 | 3 => Spread::Reflect,
                    _ => Spread::Repeat,
                },
            };
            Paint::Linear(Arc::new(crate::transform::linear(&gradient, transform)))
        }
        _ => return None,
    })
}

fn mix(a: Rgba, b: Rgba, at: f64) -> Rgba {
    let at = at.clamp(0.0, 1.0);
    let channel = |a: u8, b: u8| (f64::from(a) + (f64::from(b) - f64::from(a)) * at).round() as u8;
    Rgba {
        r: channel(a.r, b.r),
        g: channel(a.g, b.g),
        b: channel(a.b, b.b),
        a: channel(a.a, b.a),
    }
}

fn parse_pen<const FULL: bool>(
    player: &mut Player<FULL>,
    state: &State,
    data: &[u8],
) -> Option<Pen> {
    let flags = u32_at(data, 8)?;
    let unit = u32_at(data, 12)?;
    let width = finite_at(data, 16)?;
    let mut at = 20;
    let mut cap = LineCap::Butt;
    let mut join = LineJoin::Miter;
    let mut miter_limit = 10.0;
    let mut style = 0;
    let mut custom = None;
    let mut decorated = false;
    if flags & 0x0001 != 0 {
        at += 24;
    }
    if flags & 0x0002 != 0 {
        let start = i32_at(data, at)?;
        decorated |= start >= 0x10;
        at += 4;
    }
    if flags & 0x0004 != 0 {
        let end = i32_at(data, at)?;
        cap = match end {
            1 => LineCap::Square,
            2 => LineCap::Round,
            _ => LineCap::Butt,
        };
        decorated |= end >= 0x10;
        at += 4;
    }
    if flags & 0x0008 != 0 {
        join = match u32_at(data, at)? {
            1 => LineJoin::Bevel,
            2 => LineJoin::Round,
            _ => LineJoin::Miter,
        };
        at += 4;
    }
    if flags & 0x0010 != 0 {
        miter_limit = finite_at(data, at)?.max(1.0);
        at += 4;
    }
    if flags & 0x0020 != 0 {
        style = i32_at(data, at)?;
        at += 4;
    }
    if flags & 0x0040 != 0 {
        at += 4;
    }
    if flags & 0x0080 != 0 {
        at += 4;
    }
    if flags & 0x0100 != 0 {
        let count = (u32_at(data, at)? as usize).min(16);
        let mut dash = Vec::with_capacity(count);
        for index in 0..count {
            dash.push(finite_at(data, at + 4 + index * 4)?.max(0.0));
        }
        custom = Some(dash);
        at = skip_counted(data, at, 4)?;
    }
    if flags & 0x0200 != 0 {
        at += 4;
    }
    if flags & 0x0400 != 0 {
        at = skip_counted(data, at, 4)?;
    }
    for bit in [0x0800, 0x1000] {
        if flags & bit != 0 {
            decorated = true;
            at = skip_counted(data, at, 1)?;
        }
    }
    let paint = parse_brush(player, data.get(at..)?)?;
    let dash = custom
        .filter(|dash| dash.iter().any(|length| *length > 0.0))
        .or(match style {
            1 => Some(vec![3.0, 1.0]),
            2 => Some(vec![1.0, 1.0]),
            3 => Some(vec![3.0, 1.0, 1.0, 1.0]),
            4 => Some(vec![3.0, 1.0, 1.0, 1.0, 1.0, 1.0]),
            _ => None,
        });
    if decorated {
        player.omit("EMF+ line end decorations")?;
    }
    Some(Pen {
        width: width.abs(),
        world_width: unit == 0,
        dash,
        cap,
        join,
        miter_limit,
        paint,
    })
    .map(|mut pen| {
        if !pen.world_width {
            let (x, y) = state.unit_scale(unit);
            pen.width *= (x * y).sqrt();
        }
        pen
    })
}

/// The offset past a `u32` count at `at` and that many `unit`-byte items.
fn skip_counted(data: &[u8], at: usize, unit: usize) -> Option<usize> {
    let count = u32_at(data, at)? as usize;
    at.checked_add(4)?
        .checked_add(count.checked_mul(unit)?)
        .filter(|end| *end <= data.len())
}

/// An `EmfPlusPath` object, and the bytes it spans.
fn parse_path(data: &[u8], limit: usize) -> Option<(Path, usize)> {
    let count = u32_at(data, 4)? as usize;
    let flags = u32_at(data, 8)?;
    if count > limit || count > data.len() {
        return None;
    }
    let relative = flags & 0x0800 != 0;
    let (compressed, relative, rle) = (!relative && flags & 0x4000 != 0, relative, relative);
    let mut at = 12;
    let mut points = Vec::with_capacity(count);
    let mut last = (0.0, 0.0);
    for _ in 0..count {
        let point = read_point(data, &mut at, compressed, relative)?;
        last = if relative {
            (last.0 + point.0, last.1 + point.1)
        } else {
            point
        };
        points.push(last);
    }
    let mut types = Vec::with_capacity(count);
    if rle {
        while types.len() < count {
            let run = usize::from(u8_at(data, at)? & 0x3f).max(1);
            let kind = u8_at(data, at + 1)?;
            at += 2;
            types.extend(std::iter::repeat_n(kind, run.min(count - types.len())));
        }
    } else {
        types.extend_from_slice(crate::read::span(data, at, count, 1)?);
        at += count;
    }
    let mut figures: Vec<Figure> = Vec::new();
    let mut index = 0;
    while index < count {
        let (point, kind) = (points[index], types[index]);
        match kind & 0x07 {
            0 => figures.push(Figure {
                start: point,
                segments: Vec::new(),
                closed: false,
            }),
            1 => match figures.last_mut() {
                Some(figure) => figure.segments.push(Segment::Line(point)),
                None => figures.push(Figure {
                    start: point,
                    segments: Vec::new(),
                    closed: false,
                }),
            },
            3 if index + 2 < count => {
                let figure = figures.last_mut()?;
                figure.segments.push(Segment::Cubic([
                    points[index],
                    points[index + 1],
                    points[index + 2],
                ]));
                index += 2;
            }
            _ => return None,
        }
        if types[index] & 0x80 != 0
            && let Some(figure) = figures.last_mut()
        {
            figure.closed = true;
        }
        index += 1;
    }
    Some((Path { figures }, at.div_ceil(4) * 4))
}

/// Deepest region combination tree replayed.
const MAX_REGION_DEPTH: usize = 64;

/// A region node and the offset after it, spending one of `nodes` per node.
fn parse_region(
    data: &[u8],
    at: usize,
    nodes: &mut usize,
    depth: usize,
    limit: usize,
) -> Option<(Region, usize)> {
    if *nodes == 0 || depth > MAX_REGION_DEPTH {
        return None;
    }
    *nodes -= 1;
    let kind = u32_at(data, at)?;
    Some(match kind {
        1..=5 => {
            let (left, after) = parse_region(data, at.checked_add(4)?, nodes, depth + 1, limit)?;
            let (right, after) = parse_region(data, after, nodes, depth + 1, limit)?;
            (
                Region::Combine(kind, Box::new(left), Box::new(right), None).settled(),
                after,
            )
        }
        0x1000_0000 => {
            let rect = rects(data, at + 4, 1, false, 1)?[0];
            (
                Region::Path(Arc::new(Path {
                    figures: vec![rect_figure(rect)],
                })),
                at + 20,
            )
        }
        0x1000_0001 => {
            let length = u32_at(data, at.checked_add(4)?)? as usize;
            let start = at.checked_add(8)?;
            let end = start.checked_add(length)?;
            let (path, _) = parse_path(data.get(start..end)?, limit)?;
            let region = if path.figures.is_empty() {
                Region::Empty
            } else {
                Region::Path(Arc::new(path))
            };
            (region, end)
        }
        0x1000_0002 => (Region::Empty, at + 4),
        0x1000_0003 => (Region::Infinite, at + 4),
        _ => return None,
    })
}

impl Region {
    /// The combination with its extent filled in from its operands'.
    fn settled(self) -> Self {
        let Region::Combine(mode, left, right, _) = self else {
            return self;
        };
        let known = match (mode, extent(&left), extent(&right)) {
            (1, Some(false), _) | (1, _, Some(false)) => Some(false),
            (1, Some(true), Some(true)) => Some(true),
            (2, Some(true), _) | (2, _, Some(true)) => Some(true),
            (2, Some(false), Some(false)) => Some(false),
            (3, Some(a), Some(b)) => Some(a != b),
            (4, Some(false), _) | (4, _, Some(true)) => Some(false),
            (4, Some(true), Some(false)) => Some(true),
            (5, _, Some(false)) | (5, Some(true), _) => Some(false),
            (5, Some(false), Some(true)) => Some(true),
            _ => None,
        };
        Region::Combine(mode, left, right, known)
    }
}

/// Whether `region` is everything (`Some(true)`) or nothing (`Some(false)`)
/// whatever its paths.
fn extent(region: &Region) -> Option<bool> {
    match region {
        Region::Infinite => Some(true),
        Region::Empty => Some(false),
        Region::Path(_) => None,
        Region::Combine(.., known) => *known,
    }
}

/// Clip intersections; `Some(false)` when approximated.
/// `spent` counts every command built, kept or not.
fn region_clips<const FULL: bool>(
    state: &State,
    player: &mut Player<FULL>,
    region: &Region,
    out: &mut Vec<ClipRegion>,
    spent: &mut usize,
) -> Option<bool> {
    player.charge(1, 0)?;
    let empty = |spent: &mut usize| {
        let path = rect_path([0.0; 4]);
        *spent = spent.saturating_add(path.len());
        ClipRegion {
            path,
            even_odd: false,
            exclude: false,
        }
    };
    if let Region::Combine(..) = region {
        match extent(region) {
            Some(true) => return Some(true),
            Some(false) => {
                out.push(empty(spent));
                return Some(true);
            }
            None => {}
        }
    }
    Some(match region {
        Region::Infinite => true,
        Region::Empty => {
            out.push(empty(spent));
            true
        }
        Region::Path(path) => {
            let mut commands = Vec::new();
            for figure in &path.figures {
                figure_path(state, player, figure, &mut commands);
            }
            *spent = spent.saturating_add(commands.len());
            out.push(ClipRegion {
                path: commands,
                even_odd: false,
                exclude: false,
            });
            true
        }
        Region::Combine(1, left, right, _) => {
            let a = region_clips(state, player, left, out, spent)?;
            region_clips(state, player, right, out, spent)? && a
        }
        Region::Combine(2 | 3, left, right, _) if extent(left) == Some(false) => {
            region_clips(state, player, right, out, spent)?
        }
        Region::Combine(2 | 3, left, right, _) if extent(right) == Some(false) => {
            region_clips(state, player, left, out, spent)?
        }
        Region::Combine(mode @ 3..=5, left, right, _)
            if *mode != 3 || extent(left) == Some(true) || extent(right) == Some(true) =>
        {
            let (keep, cut) = match mode {
                4 => (left, right),
                5 => (right, left),
                _ if extent(left) == Some(true) => (left, right),
                _ => (right, left),
            };
            let exact = region_clips(state, player, keep, out, spent)?;
            let mut cuts = Vec::new();
            let simple = region_clips(state, player, cut, &mut cuts, spent)? && cuts.len() <= 1;
            for mut cut in cuts {
                cut.exclude = !cut.exclude;
                out.push(cut);
            }
            exact && simple
        }
        Region::Combine(mode, left, right, _) => {
            let (mut a, mut b) = (Vec::new(), Vec::new());
            region_clips(state, player, left, &mut a, spent)?;
            region_clips(state, player, right, &mut b, spent)?;
            if a.is_empty() || b.is_empty() {
                return Some(false);
            }
            let mut path = Vec::new();
            for region in a.into_iter().chain(b) {
                if !region.exclude {
                    path.extend(region.path);
                }
            }
            out.push(ClipRegion {
                path,
                even_odd: *mode == 3,
                exclude: false,
            });
            false
        }
    })
}

/// The region's area as one path, for filling; `false` when approximated.
/// Geometry built but not kept is charged here, the kept path when drawn.
fn region_path<const FULL: bool>(
    state: &State,
    player: &mut Player<FULL>,
    region: &Region,
    out: &mut Vec<PathCommand>,
) -> Option<bool> {
    let (mut clips, mut spent) = (Vec::new(), 0);
    let exact = region_clips(state, player, region, &mut clips, &mut spent)?;
    let simple = clips.len() == 1 && !clips[0].exclude;
    if clips.is_empty() {
        player.charge(spent, 0)?;
        let (w, h) = player.unit;
        out.extend(rect_path([0.0, 0.0, w, h]));
        return Some(exact);
    }
    let kept = out.len();
    for clip in clips.into_iter().filter(|clip| !clip.exclude) {
        out.extend(clip.path);
    }
    player.charge(spent.saturating_sub(out.len() - kept), 0)?;
    Some(exact && simple)
}

fn set_clip<const FULL: bool>(
    player: &mut Player<FULL>,
    state: &mut State,
    region: &Region,
    mode: u32,
) -> Option<()> {
    let (mut clips, mut spent) = (Vec::new(), 0);
    let mut exact = region_clips(state, player, region, &mut clips, &mut spent)?;
    player.charge(spent, 0)?;
    let kept = if mode == 1 || mode == 4 {
        player.clip_depth(&state.graphics.clip)?
    } else {
        0
    };
    if kept.saturating_add(clips.len()) > player.limits.clip_depth {
        return player.refuse("clip regions nest past the depth limit");
    }
    let base = match mode {
        0 => None,
        1 => state.graphics.clip.clone(),
        4 => {
            let mut excluded = match clips.len() {
                0 => {
                    player.charge(5, 0)?;
                    ClipRegion {
                        path: rect_path([0.0; 4]),
                        even_odd: false,
                        exclude: false,
                    }
                }
                1 => {
                    let mut clip = clips.pop()?;
                    clip.exclude = !clip.exclude;
                    clip
                }
                _ => {
                    exact = false;
                    let mut path = Vec::new();
                    for clip in clips.drain(..).filter(|clip| !clip.exclude) {
                        path.extend(clip.path);
                    }
                    ClipRegion {
                        path,
                        even_odd: false,
                        exclude: true,
                    }
                }
            };
            if !exact {
                player.omit("EMF+ clip regions approximated")?;
            }
            if excluded.path.is_empty() {
                player.charge(5, 0)?;
                excluded.path = rect_path([0.0; 4]);
                excluded.exclude = false;
            }
            state.graphics.clip = chain(state.graphics.clip.clone(), excluded);
            return check_depth(player, state);
        }
        _ => {
            exact = false;
            if mode == 2 && state.graphics.clip.is_none() {
                return Some(());
            }
            None
        }
    };
    if !exact {
        player.omit("EMF+ clip regions approximated")?;
    }
    let mut chain_clip = base;
    for clip in clips {
        chain_clip = chain(chain_clip, clip);
    }
    state.graphics.clip = chain_clip;
    check_depth(player, state)
}

fn check_depth<const FULL: bool>(player: &mut Player<FULL>, state: &State) -> Option<()> {
    if player.clip_depth(&state.graphics.clip)? > player.limits.clip_depth {
        return player.refuse("clip regions nest past the depth limit");
    }
    Some(())
}

fn offset_chain(clip: &Clip, dx: f64, dy: f64) -> Clip {
    let link = clip.as_deref()?;
    let mut region = link.region.clone();
    for command in &mut region.path {
        crate::player::translate(command, dx, dy);
    }
    chain(offset_chain(&link.parent, dx, dy), region)
}

fn parse_font(data: &[u8]) -> Option<PlusFont> {
    let length = (u32_at(data, 20)? as usize).min(64);
    let units: Vec<u16> = (0..length)
        .map_while(|index| u16_at(data, 24 + index * 2))
        .collect();
    Some(PlusFont {
        size: finite_at(data, 4)?.abs(),
        unit: u32_at(data, 8)?,
        style: i32_at(data, 12)?,
        family: String::from_utf16_lossy(&units)
            .trim_end_matches('\0')
            .to_owned(),
    })
}

fn parse_image<const FULL: bool>(player: &mut Player<FULL>, data: &[u8]) -> Option<PlusImage> {
    match u32_at(data, 4)? {
        1 => {
            let (width, height) = (i32_at(data, 8)?, i32_at(data, 12)?);
            let stride = i32_at(data, 16)?;
            let format = u32_at(data, 20)?;
            let compressed = u32_at(data, 24)? == 1;
            let bits = data.get(28..)?;
            let (mime, (width, height)) = if compressed {
                let mime = if bits.starts_with(b"\x89PNG") {
                    "image/png"
                } else if bits.starts_with(&[0xFF, 0xD8]) {
                    "image/jpeg"
                } else {
                    return None;
                };
                (Some(mime), crate::dib::encoded_size(bits, mime)?)
            } else if width > 0 && height > 0 {
                (None, (width as u32, height as u32))
            } else {
                return None;
            };
            let pixels = u64::from(width) * u64::from(height);
            let budget = player
                .limits
                .bitmap_pixels
                .saturating_sub(player.bitmap_pixels);
            if pixels > budget.min(crate::dib::MAX_BITMAP_PIXELS) {
                player.refuse::<()>("the metafile's bitmaps exceed the pixel budget");
                return None;
            }
            player.charge(0, pixels)?;
            let pixels = match mime {
                Some(mime) => Pixels::Encoded {
                    mime,
                    bytes: bits.to_vec(),
                },
                None => Pixels::Rgba(pixel_data(bits, width, height, stride, format)?),
            };
            let bitmap = Bitmap {
                width,
                height,
                pixels,
            };
            Some(PlusImage::Bitmap(Arc::new(bitmap)))
        }
        2 => {
            let size = u32_at(data, 12)? as usize;
            let bytes = data.get(16..16usize.checked_add(size)?)?;
            Some(PlusImage::Metafile(nested(player, bytes)?))
        }
        _ => None,
    }
}

/// Replays an embedded image with the enclosing picture's allowances.
fn nested<const FULL: bool>(player: &mut Player<FULL>, bytes: &[u8]) -> Option<Nested> {
    if player.depth + 1 >= crate::MAX_NESTING {
        return Some(Nested::TooDeep);
    }
    let budget = Rc::clone(player.budget.as_ref()?);
    let drawing = crate::play_nested(bytes, player.depth + 1, Rc::clone(&budget));
    if budget.exceeded.get() {
        return player.refuse("the metafile draws more than the replay limits");
    }
    let Ok(drawing) = drawing else {
        return Some(Nested::Failed);
    };
    let mut commands = 0;
    let mut clips = std::collections::HashSet::new();
    for op in &drawing.ops {
        let clip = match op {
            Op::Shape(shape) => {
                commands += shape.path.len();
                &shape.clip
            }
            Op::Text(text) => {
                commands += text.text.len();
                &text.clip
            }
            Op::Image(image) => &image.clip,
        };
        let mut at = clip.as_deref();
        while let Some(link) = at
            && clips.insert(std::ptr::from_ref(link))
        {
            commands += link.region.path.len().saturating_add(1);
            at = link.parent.as_deref();
        }
    }
    Some(Nested::Drawn(drawing, commands))
}

/// GDI+ pixel formats to straight RGBA.
fn pixel_data(bits: &[u8], width: u32, height: u32, stride: i32, format: u32) -> Option<Vec<u8>> {
    let (width, height) = (width as usize, height as usize);
    let indexed = format & 0x0001_0000 != 0;
    let (palette, bits) = if indexed {
        let count = (u32_at(bits, 4)? as usize).min(256);
        let mut palette = Vec::with_capacity(count);
        for index in 0..count {
            palette.push(Rgba::from_argb(u32_at(bits, 8 + index * 4)?));
        }
        (palette, bits.get(8 + count * 4..)?)
    } else {
        (Vec::new(), bits)
    };
    let bpp = ((format >> 8) & 0xff) as usize;
    let stride = if stride > 0 {
        stride as usize
    } else {
        (width * bpp).div_ceil(32) * 4
    };
    if stride < (width * bpp).div_ceil(8) || bits.len() < stride.checked_mul(height)? {
        return None;
    }
    let mut out = Vec::with_capacity(width * height * 4);
    for row in 0..height {
        let line = &bits[row * stride..row * stride + stride];
        for x in 0..width {
            let color = match format {
                0x0026_200A | 0x000E_200B | 0x0002_2009 => {
                    let p = &line[x * 4..x * 4 + 4];
                    let mut color = Rgba {
                        r: p[2],
                        g: p[1],
                        b: p[0],
                        a: p[3],
                    };
                    if format == 0x0002_2009 {
                        color.a = 255;
                    } else if format == 0x000E_200B && color.a > 0 && color.a < 255 {
                        let un = |c: u8| {
                            ((u16::from(c) * 255 + u16::from(color.a) / 2) / u16::from(color.a))
                                .min(255) as u8
                        };
                        color = Rgba {
                            r: un(color.r),
                            g: un(color.g),
                            b: un(color.b),
                            a: color.a,
                        };
                    }
                    color
                }
                0x0002_1808 => {
                    let p = &line[x * 3..x * 3 + 3];
                    Rgba::opaque(p[2], p[1], p[0])
                }
                0x0002_1005 | 0x0002_1006 | 0x0006_1007 => {
                    let value = u16::from_le_bytes([line[x * 2], line[x * 2 + 1]]);
                    let five = |v: u16| ((v & 0x1f) * 255 / 31) as u8;
                    if format == 0x0002_1006 {
                        Rgba::opaque(
                            five(value >> 11),
                            ((value >> 5 & 0x3f) * 255 / 63) as u8,
                            five(value),
                        )
                    } else {
                        let alpha = if format == 0x0006_1007 && value & 0x8000 == 0 {
                            0
                        } else {
                            255
                        };
                        Rgba {
                            r: five(value >> 10),
                            g: five(value >> 5),
                            b: five(value),
                            a: alpha,
                        }
                    }
                }
                0x0010_1004 => {
                    let value = line[x * 2 + 1];
                    Rgba::opaque(value, value, value)
                }
                0x0003_0101 | 0x0003_0402 | 0x0003_0803 => {
                    let bit = x * bpp;
                    let byte = line[bit / 8];
                    let index = (byte >> (8 - bpp - bit % 8)) & ((1u16 << bpp) - 1) as u8;
                    palette
                        .get(usize::from(index))
                        .copied()
                        .unwrap_or(Rgba::BLACK)
                }
                _ => return None,
            };
            out.extend_from_slice(&[color.r, color.g, color.b, color.a]);
        }
    }
    Some(out)
}

fn draw_image<const FULL: bool>(
    player: &mut Player<FULL>,
    state: &State,
    id: usize,
    source: (f64, f64, f64, f64),
    corners: [(f64, f64); 3],
) -> Option<()> {
    let Some(Some(Object::Image(image))) = state.objects.get(id) else {
        return Some(());
    };
    let world = concat(state.to_device(), player.device_to_output());
    let [a, b, c] = corners.map(|corner| apply(world, corner));
    let (sw, sh) = (source.2, source.3);
    if sw == 0.0 || sh == 0.0 {
        return Some(());
    }
    match image.as_ref() {
        PlusImage::Bitmap(bitmap) => {
            let (x0, y0) = (source.0.round().max(0.0), source.1.round().max(0.0));
            let (x1, y1) = (
                (source.0 + sw).round().min(f64::from(bitmap.width)),
                (source.1 + sh).round().min(f64::from(bitmap.height)),
            );
            if x1 <= x0 || y1 <= y0 {
                return Some(());
            }
            let whole = source == (0.0, 0.0, f64::from(bitmap.width), f64::from(bitmap.height));
            let mut clip = state.graphics.clip.clone();
            let (placed, (ox, oy)) = match &bitmap.pixels {
                Pixels::Encoded { .. } => {
                    if !whole {
                        player.charge(5, 0)?;
                        let d = (b.0 + c.0 - a.0, b.1 + c.1 - a.1);
                        clip = chain(clip, crate::blit::parallelogram([a, b, d, c]));
                    }
                    (Arc::clone(bitmap), (-source.0, -source.1))
                }
                Pixels::Rgba(_) => {
                    if !whole {
                        player.charge(0, ((x1 - x0) * (y1 - y0)) as u64)?;
                    }
                    let cropped = crop(
                        bitmap,
                        x0 as u32,
                        y0 as u32,
                        (x1 - x0) as u32,
                        (y1 - y0) as u32,
                    )?;
                    (cropped, (x0 - source.0, y0 - source.1))
                }
            };
            let (u, v) = ((b.0 - a.0) / sw, (b.1 - a.1) / sw);
            let (p, q) = ((c.0 - a.0) / sh, (c.1 - a.1) / sh);
            let transform = [u, v, p, q, a.0 + u * ox + p * oy, a.1 + v * ox + q * oy];
            player.push_op(Op::Image(Image {
                transform,
                bitmap: placed,
                opacity: 1.0,
                clip,
            }));
        }
        PlusImage::Metafile(Nested::TooDeep) => player.omit("metafiles nested too deeply")?,
        PlusImage::Metafile(Nested::Failed) => {
            player.omit("embedded metafiles that could not be replayed")?
        }
        PlusImage::Metafile(Nested::Drawn(nested, commands)) => {
            player.charge(*commands, 0)?;
            let (u, v) = ((b.0 - a.0) / sw, (b.1 - a.1) / sw);
            let (p, q) = ((c.0 - a.0) / sh, (c.1 - a.1) / sh);
            let m = [
                u,
                v,
                p,
                q,
                a.0 - u * source.0 - p * source.1,
                a.1 - v * source.0 - q * source.1,
            ];
            player.charge(5, 0)?;
            let d = (b.0 + c.0 - a.0, b.1 + c.1 - a.1);
            let clip = chain(
                state.graphics.clip.clone(),
                crate::blit::parallelogram([a, b, d, c]),
            );
            let mut cache = std::collections::HashMap::new();
            for op in &nested.ops {
                let op = crate::transform::op(op.clone(), m, &clip, &mut cache);
                player.push_op(op);
            }
            for omission in &nested.omissions {
                let count = player.omissions.entry(omission.what).or_default();
                *count = count.saturating_add(omission.count);
            }
        }
    }
    Some(())
}

fn crop(bitmap: &Arc<Bitmap>, x: u32, y: u32, width: u32, height: u32) -> Option<Arc<Bitmap>> {
    if x == 0 && y == 0 && width == bitmap.width && height == bitmap.height {
        return Some(Arc::clone(bitmap));
    }
    let Pixels::Rgba(rgba) = &bitmap.pixels else {
        return Some(Arc::clone(bitmap));
    };
    let stride = bitmap.width as usize * 4;
    let mut out = Vec::with_capacity(width as usize * height as usize * 4);
    for row in y as usize..(y + height) as usize {
        let start = row * stride + x as usize * 4;
        out.extend_from_slice(rgba.get(start..start + width as usize * 4)?);
    }
    Some(Arc::new(Bitmap {
        width,
        height,
        pixels: Pixels::Rgba(out),
    }))
}

/// The font's em size in world units, and its face.
fn world_font(state: &State, font: &PlusFont) -> Font {
    let size = if font.unit == 0 {
        font.size
    } else {
        font.size * state.unit_scale(font.unit).1
            / (state.unit_scale(state.graphics.page_unit).1 * state.graphics.page_scale)
    };
    Font {
        family: font.family.clone(),
        size,
        weight: if font.style & 1 != 0 { 700 } else { 400 },
        italic: font.style & 2 != 0,
        underline: font.style & 4 != 0,
        strike: font.style & 8 != 0,
    }
}

fn draw_string<const FULL: bool>(
    player: &mut Player<FULL>,
    state: &State,
    id: usize,
    solid: bool,
    data: &[u8],
) -> Option<()> {
    let Some(Some(Object::Font(font))) = state.objects.get(id).cloned() else {
        return Some(());
    };
    let fill = brush(player, state, solid, u32_at(data, 0)?)?;
    let format = match state
        .objects
        .get(u32_at(data, 4)? as usize)
        .cloned()
        .flatten()
    {
        Some(Object::Format(format)) => format,
        _ => Format::default(),
    };
    let length = u32_at(data, 8)? as usize;
    if length > player.limits.text_chars {
        return player.refuse("a text record holds more characters than the limit");
    }
    let rect = (
        finite_at(data, 12)?,
        finite_at(data, 16)?,
        finite_at(data, 20)?,
        finite_at(data, 24)?,
    );
    player.spend(length as u64, 0)?;
    let units: Vec<u16> = (0..length)
        .map(|index| u16_at(data, 28 + index * 2))
        .collect::<Option<_>>()?;
    let (chars, _) = crate::text::decode_wide(&units, None);
    player.text_chars = player.text_chars.saturating_add(chars.len());
    if player.text_chars > player.limits.text_chars {
        return player.refuse("the metafile holds more text than the limit");
    }
    if format.flags & 0x0002 != 0 {
        player.omit("vertical EMF+ text")?;
    }
    let font = world_font(state, &font);
    let text: String = chars.into_iter().collect();
    let wrap = (format.flags & NO_WRAP == 0 && rect.2 > 0.0).then_some(rect.2);
    let (lines, estimated) = layout_lines(&text, &font, wrap);
    if estimated {
        player.omit("EMF+ text wrapped by estimated widths")?;
    }
    let (ascent, descent) = crate::text::font_metrics(&font.family);
    let line_height = font.size * (ascent + descent);
    let block = line_height * lines.len() as f64;
    let top = match format.line_align {
        1 if rect.3 > 0.0 => rect.1 + (rect.3 - block) / 2.0,
        2 if rect.3 > 0.0 => rect.1 + rect.3 - block,
        _ => rect.1,
    };
    let (x, anchor) = match format.align {
        1 if rect.2 > 0.0 => (rect.0 + rect.2 / 2.0, TextAnchor::Middle),
        2 if rect.2 > 0.0 => (rect.0 + rect.2, TextAnchor::End),
        _ => (rect.0, TextAnchor::Start),
    };
    let world = concat(state.to_device(), player.device_to_output());
    let mut clip = state.graphics.clip.clone();
    if format.flags & NO_CLIP == 0 && rect.2 > 0.0 && rect.3 > 0.0 {
        let (l, t, r, b) = (rect.0, rect.1, rect.0 + rect.2, rect.1 + rect.3);
        let corners = [(l, t), (r, t), (r, b), (l, b)].map(|corner| apply(world, corner));
        player.charge(5, 0)?;
        clip = chain(clip, crate::blit::parallelogram(corners));
    }
    for (index, line) in lines.iter().enumerate() {
        if line.is_empty() {
            continue;
        }
        let baseline = top + line_height * index as f64 + ascent * font.size;
        let transform = concat([1.0, 0.0, 0.0, 1.0, x, baseline], world);
        player.push_op(Op::Text(Text {
            transform,
            text: line.clone(),
            positions: None,
            anchor,
            font: font.clone(),
            fill: fill.clone(),
            clip: clip.clone(),
        }));
    }
    Some(())
}

/// `StringFormat` flags: no wrapping at the layout rectangle, and no
/// clipping to it.
const NO_WRAP: u32 = 0x0000_1000;
const NO_CLIP: u32 = 0x0000_4000;

/// `text` split at its line breaks, keeping blank lines, and wrapped at word
/// boundaries to `width` by estimated character widths; `true` when any
/// line needed wrapping.
fn layout_lines(text: &str, font: &Font, width: Option<f64>) -> (Vec<String>, bool) {
    let face = font.family.to_ascii_lowercase();
    let per_char = if face.contains("courier") || face.contains("mono") {
        0.6
    } else {
        0.5
    } * font.size;
    let fits = |chars: usize| width.is_none_or(|width| chars as f64 * per_char <= width);
    let mut lines = Vec::new();
    let mut estimated = false;
    for paragraph in text.replace("\r\n", "\n").split(['\n', '\r']) {
        if fits(paragraph.chars().count()) {
            lines.push(paragraph.to_owned());
            continue;
        }
        estimated = true;
        let (mut line, mut chars) = (String::new(), 0);
        for word in paragraph.split(' ') {
            let length = word.chars().count();
            if chars > 0 && !fits(chars + 1 + length) {
                lines.push(std::mem::take(&mut line));
                chars = 0;
            } else if chars > 0 {
                line.push(' ');
                chars += 1;
            }
            line.push_str(word);
            chars += length;
        }
        lines.push(line);
    }
    (lines, estimated)
}

fn draw_driver_string<const FULL: bool>(
    player: &mut Player<FULL>,
    state: &State,
    id: usize,
    solid: bool,
    data: &[u8],
) -> Option<()> {
    let Some(Some(Object::Font(font))) = state.objects.get(id).cloned() else {
        return Some(());
    };
    let fill = brush(player, state, solid, u32_at(data, 0)?)?;
    let options = u32_at(data, 4)?;
    let has_matrix = u32_at(data, 8)? != 0;
    let count = u32_at(data, 12)? as usize;
    if count > player.limits.text_chars {
        return player.refuse("a text record holds more characters than the limit");
    }
    if options & 0x1 == 0 {
        return player.omit("text given as glyph indexes");
    }
    if options & 0x2 != 0 {
        player.omit("vertical EMF+ text")?;
    }
    player.spend(count as u64 * 2, 0)?;
    let units: Vec<u16> = (0..count)
        .map(|index| u16_at(data, 16 + index * 2))
        .collect::<Option<_>>()?;
    let positions_at = 16 + count * 2;
    let mut positions = Vec::with_capacity(count);
    for index in 0..count {
        positions.push((
            finite_at(data, positions_at + index * 8)?,
            finite_at(data, positions_at + index * 8 + 4)?,
        ));
    }
    let glyph = if has_matrix {
        matrix(data, positions_at + count * 8)?
    } else {
        IDENTITY
    };
    let origin = *positions.first()?;
    let (chars, _) = crate::text::decode_wide(&units, None);
    player.text_chars = player.text_chars.saturating_add(chars.len());
    if player.text_chars > player.limits.text_chars {
        return player.refuse("the metafile holds more text than the limit");
    }
    let positions = (options & 0x4 == 0).then(|| {
        units
            .iter()
            .zip(&positions)
            .filter(|(unit, _)| !(0xDC00..0xE000).contains(*unit))
            .map(|(_, p)| (p.0 - origin.0, p.1 - origin.1))
            .collect::<Vec<_>>()
    });
    let font = world_font(state, &font);
    let world = concat(state.to_device(), player.device_to_output());
    let placed = concat(
        concat(glyph, [1.0, 0.0, 0.0, 1.0, origin.0, origin.1]),
        world,
    );
    let clip = state.graphics.clip.clone();
    player.push_op(Op::Text(Text {
        transform: placed,
        text: chars.into_iter().collect(),
        positions,
        anchor: TextAnchor::Start,
        font,
        fill,
        clip,
    }));
    Some(())
}

#[cfg(test)]
mod tests {
    use std::cell::Cell;

    use super::*;
    use crate::player::SharedBudget;
    use crate::test_records::*;

    #[test]
    fn empty_clip_offsets_charge_every_link_before_copying() {
        let bytes = |offsets| {
            let mut records = vec![plus_header(false), plus_path(1, &[])];
            for _ in 0..128 {
                records.push((0x4033, 0x0101, Vec::new()));
            }
            for _ in 0..offsets {
                records.push((0x4035, 0, f32s(&[1.0, 1.0])));
            }
            records.push(plus_eof());
            let (kind, body) = plus(&records);
            Emf::new(100, 100).rec(kind, &body).bytes()
        };
        let replay = |offsets, work| {
            let budget = Rc::new(SharedBudget {
                remaining: Cell::new(crate::ReplayBudget { work, pixels: 0 }),
                exceeded: Cell::new(false),
            });
            let result =
                crate::play_emf::<true>(&bytes(offsets), 0, true, Some(Rc::clone(&budget)));
            (result, budget)
        };
        let (base, budget) = replay(0, 100_000);
        let base = base.unwrap();
        assert_eq!(
            crate::drawing::ClipChain::depth(&base.plus.as_ref().unwrap().graphics.clip),
            128
        );
        let spent = 100_000 - budget.remaining.get().work;
        let (offset, budget) = replay(2, 100_000);
        assert_eq!(offset.unwrap().commands - base.commands, 2 * 128);
        assert_eq!(100_000 - budget.remaining.get().work - spent, 2 * 129);
        let (result, budget) = replay(8, spent + 2 * 129);
        assert!(result.is_err());
        assert!(budget.exceeded.get());
        assert_eq!(budget.remaining.get().work, 0);
    }

    #[test]
    fn restore_charges_scans_and_uses_the_latest_matching_state() {
        let budget = Rc::new(SharedBudget {
            remaining: Cell::new(crate::ReplayBudget {
                work: 20,
                pixels: 0,
            }),
            exceeded: Cell::new(false),
        });
        let mut player = Player::<true>::new((0.0, 0.0, 10.0, 10.0), 0, (10.0, 10.0));
        player.budget = Some(Rc::clone(&budget));
        let mut state = State::new(false, (96.0, 96.0));
        for (id, x) in [(1u32, 10.0), (2, 20.0), (1, 30.0)] {
            state.graphics.world[4] = x;
            assert_eq!(
                play(&mut player, &mut state, 0x4025, 0, &id.to_le_bytes()),
                Some(true)
            );
        }
        state.graphics.world[4] = 40.0;
        assert_eq!(
            play(&mut player, &mut state, 0x4026, 0, &99u32.to_le_bytes()),
            Some(true)
        );
        assert_eq!(state.graphics.world[4], 40.0);
        assert_eq!(state.stack.len(), 3);
        for (kind, id, x, depth) in [
            (0x4026, 1u32, 30.0, 2),
            (0x4029, 2, 20.0, 1),
            (0x4026, 1, 10.0, 0),
        ] {
            assert_eq!(
                play(&mut player, &mut state, kind, 0, &id.to_le_bytes()),
                Some(true)
            );
            assert_eq!(state.graphics.world[4], x);
            assert_eq!(state.stack.len(), depth);
        }
        assert_eq!(player.commands, 9);
        assert_eq!(budget.remaining.get().work, 11);
    }

    #[test]
    fn restore_refuses_before_scanning_past_the_budget() {
        let budget = Rc::new(SharedBudget {
            remaining: Cell::new(crate::ReplayBudget { work: 2, pixels: 0 }),
            exceeded: Cell::new(false),
        });
        let mut player = Player::<true>::new((0.0, 0.0, 10.0, 10.0), 0, (10.0, 10.0));
        player.budget = Some(Rc::clone(&budget));
        let mut state = State::new(false, (96.0, 96.0));
        for id in 0u32..3 {
            play(&mut player, &mut state, 0x4025, 0, &id.to_le_bytes()).unwrap();
        }
        assert_eq!(
            play(&mut player, &mut state, 0x4026, 0, &0u32.to_le_bytes()),
            None
        );
        assert_eq!(state.stack.len(), 3);
        assert!(budget.exceeded.get());
        assert_eq!(budget.remaining.get().work, 0);
    }
}
