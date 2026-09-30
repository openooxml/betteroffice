//! The GDI device context both profiles replay records into.
//!
//! `FULL = false` is the strict vector replay PowerPoint pictures use: any
//! record or parameter it cannot model refuses the whole metafile. `FULL = true`
//! adds text, bitmaps, general clipping and styled pens, and records ink it
//! cannot reproduce as an omission instead of refusing.

use std::cell::Cell;
use std::collections::BTreeMap;
use std::rc::Rc;
use std::sync::Arc;

use crate::drawing::{
    Bitmap, Clip, ClipChain, ClipRegion, Hatch, LineCap, LineJoin, Op, Paint, PathCommand, Rgba,
    Shape, Stroke,
};

pub(crate) const ARC_SEGMENTS: usize = 64;
pub(crate) const MAX_HANDLES: usize = 4096;
pub(crate) const MAX_SAVED: usize = 1024;

#[derive(Clone, Copy)]
pub(crate) struct Limits {
    pub records: usize,
    pub ops: usize,
    pub commands: usize,
    pub points_per_record: usize,
    pub clip_depth: usize,
    pub clip_rects: usize,
    pub text_chars: usize,
    pub bitmap_pixels: u64,
}

pub(crate) const SHAPES_LIMITS: Limits = Limits {
    records: 200_000,
    ops: 4_096,
    commands: 400_000,
    points_per_record: 65_536,
    clip_depth: 1,
    clip_rects: 0,
    text_chars: 0,
    bitmap_pixels: 0,
};

pub(crate) const FULL_LIMITS: Limits = Limits {
    records: 1_000_000,
    ops: 250_000,
    commands: 4_000_000,
    points_per_record: 1 << 20,
    clip_depth: 128,
    clip_rects: 16_384,
    text_chars: 1_000_000,
    bitmap_pixels: 16_777_216,
};

pub(crate) struct SharedBudget {
    pub remaining: Cell<crate::ReplayBudget>,
    pub exceeded: Cell<bool>,
}

impl SharedBudget {
    pub(crate) fn spend(&self, work: u64, pixels: u64) -> bool {
        if self.exceeded.get() {
            return false;
        }
        let remaining = self.remaining.get();
        let exceeded = work > remaining.work || pixels > remaining.pixels;
        self.remaining.set(crate::ReplayBudget {
            work: remaining.work.saturating_sub(work),
            pixels: remaining.pixels.saturating_sub(pixels),
        });
        self.exceeded.set(exceeded);
        !exceeded
    }
}

pub(crate) type Xform = [f64; 6];

pub(crate) const IDENTITY: Xform = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];

/// `first` then `then`, in GDI's row-vector convention.
pub(crate) fn concat(first: Xform, then: Xform) -> Xform {
    [
        first[0] * then[0] + first[1] * then[2],
        first[0] * then[1] + first[1] * then[3],
        first[2] * then[0] + first[3] * then[2],
        first[2] * then[1] + first[3] * then[3],
        first[4] * then[0] + first[5] * then[2] + then[4],
        first[4] * then[1] + first[5] * then[3] + then[5],
    ]
}

pub(crate) fn apply(m: Xform, (x, y): (f64, f64)) -> (f64, f64) {
    (x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5])
}

#[derive(Clone, PartialEq)]
pub(crate) struct Pen {
    pub color: u32,
    pub width: f64,
    pub visible: bool,
    /// The raw `PS_*` style: dash kind, end cap, join and pen type bits.
    pub style: u32,
    pub cosmetic: bool,
    pub user_dash: Option<Arc<[f64]>>,
}

impl Pen {
    pub(crate) fn solid(color: u32, width: f64, visible: bool) -> Self {
        Self {
            color,
            width,
            visible,
            style: 0,
            cosmetic: width <= 1.0,
            user_dash: None,
        }
    }
}

#[derive(Clone, PartialEq)]
pub(crate) struct Brush {
    pub color: u32,
    pub visible: bool,
    pub hatch: Option<u32>,
    pub pattern: Option<Arc<Bitmap>>,
}

impl Brush {
    pub(crate) const fn solid(color: u32, visible: bool) -> Self {
        Self {
            color,
            visible,
            hatch: None,
            pattern: None,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub(crate) struct LogFont {
    pub height: i32,
    pub escapement: i32,
    pub weight: i32,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub charset: u8,
    pub face: String,
}

#[derive(Clone)]
pub(crate) enum GdiObject {
    Pen(Pen),
    Brush(Brush),
    Font(Arc<LogFont>),
    /// A WMF region's rectangles, in logical units.
    Region(Arc<[[f64; 4]]>),
    /// An object this replay does not model. It holds its handle slot so the
    /// records that select by index keep addressing the objects they meant.
    Opaque,
}

#[derive(Clone)]
pub(crate) struct Dc {
    pub window_org: (f64, f64),
    pub window_ext: (f64, f64),
    pub viewport_org: (f64, f64),
    pub viewport_ext: (f64, f64),
    pub scaled: bool,
    pub window_ext_set: bool,
    pub viewport_ext_set: bool,
    pub map_mode: i32,
    pub even_odd: bool,
    pub clockwise: bool,
    pub xform: Xform,
    pub pen: Option<Pen>,
    pub brush: Option<Brush>,
    pub clip: Clip,
    pub clip_rect: Option<[f64; 4]>,
    /// The clip `SETMETARGN` fixed; the application clip builds on it.
    pub meta: Clip,
    pub meta_rect: Option<[f64; 4]>,
    pub font: Option<Arc<LogFont>>,
    pub text_color: u32,
    pub bk_color: u32,
    pub bk_opaque: bool,
    pub text_align: u32,
    pub miter_limit: f64,
    pub rop2: u32,
}

impl Default for Dc {
    fn default() -> Self {
        Self {
            window_org: (0.0, 0.0),
            window_ext: (1.0, 1.0),
            viewport_org: (0.0, 0.0),
            viewport_ext: (1.0, 1.0),
            scaled: false,
            window_ext_set: false,
            viewport_ext_set: false,
            map_mode: 1,
            even_odd: true,
            clockwise: false,
            xform: IDENTITY,
            pen: Some(Pen::solid(0, 0.0, true)),
            brush: Some(Brush::solid(0x00ff_ffff, true)),
            clip: None,
            clip_rect: None,
            meta: None,
            meta_rect: None,
            font: None,
            text_color: 0,
            bk_color: 0x00ff_ffff,
            bk_opaque: true,
            text_align: 0,
            miter_limit: 10.0,
            rop2: 13,
        }
    }
}

/// How a new region combines with the current clip (`RGN_*`).
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum Combine {
    And,
    Or,
    Xor,
    Diff,
    Copy,
}

impl Combine {
    pub(crate) fn from_gdi(mode: u32) -> Option<Self> {
        Some(match mode {
            1 => Self::And,
            2 => Self::Or,
            3 => Self::Xor,
            4 => Self::Diff,
            5 => Self::Copy,
            _ => return None,
        })
    }
}

pub(crate) struct Player<const FULL: bool> {
    pub dc: Dc,
    pub saved: Vec<Dc>,
    pub objects: Vec<Option<GdiObject>>,
    pub free_objects: crate::wmf::FreeObjects,
    /// The frame in device units, as `(x, y, width, height)`.
    pub frame: (f64, f64, f64, f64),
    /// Output units across the frame: `(1, 1)` for fractions of it.
    pub unit: (f64, f64),
    /// Device units per millimetre, for the fixed mapping modes.
    pub device_per_mm: (f64, f64),
    pub pixel: Option<f64>,
    pub path: Vec<PathCommand>,
    pub selected_path: Vec<PathCommand>,
    pub figure_start: (f64, f64),
    pub current: (f64, f64),
    pub bracketed: bool,
    pub pending_stroke: bool,
    /// `WIDENPATH` ran: the next `FILLPATH` paints the path's outline.
    pub widened: bool,
    pub commands: usize,
    pub ops: Vec<Op>,
    pub overflowed: bool,
    pub limits: Limits,
    pub omissions: BTreeMap<&'static str, usize>,
    pub text_chars: usize,
    pub bitmap_pixels: u64,
    pub budget: Option<Rc<SharedBudget>>,
    pub refusal: Option<String>,
    /// How many metafiles enclose this one.
    pub depth: usize,
    /// An EMF+-only file: GDI records play only inside `EmfPlusGetDC` spans.
    pub plus_only: bool,
    pub plus_gdi: bool,
    pub plus: Option<Box<crate::emfplus::State>>,
    /// Whether the EMF carries GDI records beyond its header and comments.
    pub gdi_records: bool,
    /// Plays a dual EMF+ metafile's EMF+ records instead of its GDI ones.
    pub prefer_plus: bool,
}

impl<const FULL: bool> Player<FULL> {
    pub(crate) fn new(frame: (f64, f64, f64, f64), handles: usize, unit: (f64, f64)) -> Self {
        Self {
            dc: Dc::default(),
            saved: Vec::new(),
            objects: vec![None; handles.min(MAX_HANDLES)],
            free_objects: crate::wmf::FreeObjects::new(),
            frame,
            unit,
            device_per_mm: (96.0 / 25.4, 96.0 / 25.4),
            pixel: None,
            path: Vec::new(),
            selected_path: Vec::new(),
            figure_start: (0.0, 0.0),
            current: (0.0, 0.0),
            bracketed: false,
            pending_stroke: false,
            widened: false,
            commands: 0,
            ops: Vec::new(),
            overflowed: false,
            limits: if FULL { FULL_LIMITS } else { SHAPES_LIMITS },
            omissions: BTreeMap::new(),
            text_chars: 0,
            bitmap_pixels: 0,
            budget: None,
            refusal: None,
            depth: 0,
            plus_only: false,
            plus_gdi: false,
            plus: None,
            gdi_records: true,
            prefer_plus: false,
        }
    }

    /// Records why the replay stops and yields `None` for `?`.
    pub(crate) fn refuse<T>(&mut self, why: impl Into<String>) -> Option<T> {
        if self.refusal.is_none() {
            self.refusal = Some(why.into());
        }
        None
    }

    /// Spends cumulative work and pixels in the full profile.
    pub(crate) fn spend(&mut self, work: u64, pixels: u64) -> Option<()> {
        if FULL
            && self
                .budget
                .as_ref()
                .is_some_and(|budget| !budget.spend(work, pixels))
        {
            self.overflowed = true;
            return self.refuse("the metafile draws more than the replay limits");
        }
        Some(())
    }

    /// Spends commands and pixels built outside `push`.
    pub(crate) fn charge(&mut self, commands: usize, pixels: u64) -> Option<()> {
        self.spend(commands as u64, pixels)?;
        self.commands = self.commands.saturating_add(commands);
        self.bitmap_pixels = self.bitmap_pixels.saturating_add(pixels);
        if self.commands > self.limits.commands || self.bitmap_pixels > self.limits.bitmap_pixels {
            return self.refuse("the metafile draws more than the replay limits");
        }
        Some(())
    }

    /// Notes ink drawn without; only the full profile tolerates omissions.
    pub(crate) fn omit(&mut self, what: &'static str) -> Option<()> {
        if !FULL {
            return None;
        }
        let count = self.omissions.entry(what).or_default();
        *count = count.saturating_add(1);
        Some(())
    }

    /// Page-space scale from logical to device units.
    pub(crate) fn page_scale(&self) -> (f64, f64) {
        if !FULL {
            return if self.dc.scaled {
                (
                    self.dc.viewport_ext.0 / self.dc.window_ext.0,
                    self.dc.viewport_ext.1 / self.dc.window_ext.1,
                )
            } else {
                (1.0, 1.0)
            };
        }
        let (mx, my) = self.device_per_mm;
        let ratio = |viewport: f64, window: f64| {
            if window == 0.0 {
                1.0
            } else {
                viewport / window
            }
        };
        let (vx, vy) = self.dc.viewport_ext;
        let (wx, wy) = self.dc.window_ext;
        match self.dc.map_mode {
            2 => (mx * 0.1, -my * 0.1),
            3 => (mx * 0.01, -my * 0.01),
            4 => (mx * 0.254, -my * 0.254),
            5 => (mx * 0.0254, -my * 0.0254),
            6 => (mx * 25.4 / 1440.0, -my * 25.4 / 1440.0),
            7 => {
                let (sx, sy) = (ratio(vx, wx), ratio(vy, wy));
                let scale = sx.abs().min(sy.abs());
                (scale.copysign(sx), scale.copysign(sy))
            }
            8 => (ratio(vx, wx), ratio(vy, wy)),
            _ => (1.0, 1.0),
        }
    }

    /// World then page transform: logical units to device units.
    pub(crate) fn logical_to_device(&self) -> Xform {
        let (sx, sy) = self.page_scale();
        let page = [
            sx,
            0.0,
            0.0,
            sy,
            self.dc.viewport_org.0 - self.dc.window_org.0 * sx,
            self.dc.viewport_org.1 - self.dc.window_org.1 * sy,
        ];
        concat(self.dc.xform, page)
    }

    pub(crate) fn device_to_output(&self) -> Xform {
        let sx = self.unit.0 / self.frame.2;
        let sy = self.unit.1 / self.frame.3;
        [sx, 0.0, 0.0, sy, -self.frame.0 * sx, -self.frame.1 * sy]
    }

    pub(crate) fn logical_to_output(&self) -> Xform {
        concat(self.logical_to_device(), self.device_to_output())
    }

    pub(crate) fn device_point(&self, x: f64, y: f64) -> (f64, f64) {
        (
            (x - self.frame.0) / self.frame.2 * self.unit.0,
            (y - self.frame.1) / self.frame.3 * self.unit.1,
        )
    }

    pub(crate) fn point(&self, x: f64, y: f64) -> (f64, f64) {
        let m = self.dc.xform;
        let wx = x * m[0] + y * m[2] + m[4];
        let wy = x * m[1] + y * m[3] + m[5];
        let (sx, sy) = self.page_scale();
        let dx = (wx - self.dc.window_org.0) * sx + self.dc.viewport_org.0;
        let dy = (wy - self.dc.window_org.1) * sy + self.dc.viewport_org.1;
        self.device_point(dx, dy)
    }

    /// A logical length along x, in output units.
    pub(crate) fn logical_length(&self, width: f64) -> f64 {
        let scale = if FULL {
            self.page_scale().0.abs()
        } else if self.dc.scaled {
            (self.dc.viewport_ext.0 / self.dc.window_ext.0).abs()
        } else {
            1.0
        };
        (width * scale * self.dc.xform[0].hypot(self.dc.xform[1]) / self.frame.2 * self.unit.0)
            .abs()
    }

    /// One device pixel along x, in output units: a reference-device pixel
    /// for an EMF, a CSS pixel for a WMF, which names no device.
    pub(crate) fn device_pixel(&self) -> f64 {
        self.pixel
            .unwrap_or_else(|| (self.unit.0 / self.frame.2).abs())
    }

    pub(crate) fn push(&mut self, command: PathCommand) {
        if self.spend(1, 0).is_none() {
            return;
        }
        self.commands += 1;
        if self.commands > self.limits.commands {
            self.overflowed = true;
            return;
        }
        self.path.push(command);
    }

    pub(crate) fn move_to(&mut self, x: f64, y: f64) {
        self.current = (x, y);
        self.figure_start = (x, y);
        let (px, py) = self.point(x, y);
        self.push(PathCommand::Move { x: px, y: py });
    }

    pub(crate) fn line_to(&mut self, x: f64, y: f64) {
        self.resume_figure();
        self.current = (x, y);
        let (px, py) = self.point(x, y);
        self.push(PathCommand::Line { x: px, y: py });
    }

    pub(crate) fn cubic_to(&mut self, points: [(f64, f64); 3]) {
        self.resume_figure();
        self.current = points[2];
        let a = self.point(points[0].0, points[0].1);
        let b = self.point(points[1].0, points[1].1);
        let c = self.point(points[2].0, points[2].1);
        self.push(PathCommand::Cubic {
            cp1x: a.0,
            cp1y: a.1,
            cp2x: b.0,
            cp2y: b.1,
            x: c.0,
            y: c.1,
        });
    }

    pub(crate) fn resume_figure(&mut self) {
        if self.path.is_empty() || matches!(self.path.last(), Some(PathCommand::Close)) {
            self.move_to(self.current.0, self.current.1);
        }
    }

    pub(crate) fn close_figure(&mut self) {
        self.push(PathCommand::Close);
        self.current = self.figure_start;
    }

    pub(crate) fn save(&mut self) -> Option<()> {
        if self.saved.len() >= MAX_SAVED {
            return self.refuse("device contexts nest past the save limit");
        }
        self.saved.push(self.dc.clone());
        Some(())
    }

    pub(crate) fn restore(&mut self, depth: i32) -> Option<()> {
        self.flush_pending();
        let index = if depth < 0 {
            self.saved
                .len()
                .checked_sub(depth.unsigned_abs() as usize)?
        } else {
            (depth as usize).checked_sub(1)?
        };
        self.dc = self.saved.get(index)?.clone();
        self.saved.truncate(index);
        Some(())
    }

    pub(crate) fn brush_fill(&self) -> Option<Paint> {
        let brush = self.dc.brush.as_ref().filter(|brush| brush.visible)?;
        if FULL {
            if let Some(forced) = self.rop2_color() {
                return forced;
            }
            if let Some(style) = brush.hatch {
                return Some(Paint::Hatch {
                    style: hatch_style(style),
                    color: Rgba::from_colorref(brush.color),
                    background: self
                        .dc
                        .bk_opaque
                        .then(|| Rgba::from_colorref(self.dc.bk_color)),
                    cell: 8.0 * self.device_pixel(),
                });
            }
            if let Some(tile) = &brush.pattern {
                let pixel = self.device_pixel();
                let tall = (self.unit.1 / self.frame.3).abs();
                return Some(Paint::Pattern {
                    width: f64::from(tile.width) * pixel,
                    height: f64::from(tile.height) * tall,
                    tile: Arc::clone(tile),
                });
            }
        }
        Some(Paint::Solid(Rgba::from_colorref(brush.color)))
    }

    /// The colour `SETROP2` forces on every pen and brush: `Some(None)` draws
    /// nothing, `None` leaves the object's own colour.
    fn rop2_color(&self) -> Option<Option<Paint>> {
        match self.dc.rop2 {
            1 => Some(Some(Paint::Solid(Rgba::BLACK))),
            16 => Some(Some(Paint::Solid(Rgba::WHITE))),
            11 => Some(None),
            _ => None,
        }
    }

    pub(crate) fn pen_stroke(&self) -> Option<Stroke> {
        let pen = self.dc.pen.as_ref().filter(|pen| pen.visible)?;
        let mut color = Paint::Solid(Rgba::from_colorref(pen.color));
        if !FULL {
            return Some(Stroke {
                paint: color,
                width: self.logical_length(pen.width),
                dash: None,
                cap: LineCap::Round,
                join: LineJoin::Round,
                miter_limit: 10.0,
            });
        }
        if let Some(forced) = self.rop2_color() {
            color = forced?;
        }
        let pixel = self.device_pixel();
        let width = if pen.cosmetic {
            pixel.max(1.0)
        } else {
            self.logical_length(pen.width).max(pixel)
        };
        let unit = if pen.cosmetic { pixel } else { width };
        let pattern: &[f64] = match pen.style & 0x0f {
            1 if pen.cosmetic => &[18.0, 6.0],
            1 => &[3.0, 1.0],
            2 if pen.cosmetic => &[3.0, 3.0],
            2 => &[1.0, 1.0],
            3 if pen.cosmetic => &[9.0, 6.0, 3.0, 6.0],
            3 => &[3.0, 1.0, 1.0, 1.0],
            4 if pen.cosmetic => &[9.0, 3.0, 3.0, 3.0, 3.0, 3.0],
            4 => &[3.0, 1.0, 1.0, 1.0, 1.0, 1.0],
            8 => &[1.0, 1.0],
            _ => &[],
        };
        let dash = match (&pen.user_dash, pen.style & 0x0f) {
            (Some(entries), 7) => {
                let scale = if pen.cosmetic {
                    pixel
                } else {
                    self.logical_length(1.0)
                };
                Some(entries.iter().map(|length| length * scale).collect())
            }
            _ if !pattern.is_empty() => Some(pattern.iter().map(|length| length * unit).collect()),
            _ => None,
        };
        let (cap, join) = if pen.cosmetic {
            (LineCap::Butt, LineJoin::Miter)
        } else {
            let cap = match pen.style & 0x0f00 {
                0x0100 => LineCap::Square,
                0x0200 => LineCap::Butt,
                _ => LineCap::Round,
            };
            let join = match pen.style & 0xf000 {
                0x1000 => LineJoin::Bevel,
                0x2000 => LineJoin::Miter,
                _ => LineJoin::Round,
            };
            (cap, join)
        };
        Some(Stroke {
            paint: color,
            width,
            dash,
            cap,
            join,
            miter_limit: self.dc.miter_limit.max(1.0),
        })
    }

    /// The pen's outline painted with the brush, as `FILLPATH` paints a widened path.
    pub(crate) fn widened_stroke(&self) -> Option<Stroke> {
        let paint = self.brush_fill()?;
        let mut stroke = self.pen_stroke().unwrap_or(Stroke {
            paint: paint.clone(),
            width: self.device_pixel(),
            dash: None,
            cap: LineCap::Round,
            join: LineJoin::Round,
            miter_limit: 10.0,
        });
        stroke.paint = paint;
        Some(stroke)
    }

    pub(crate) fn push_op(&mut self, op: Op) {
        if self.spend(1, 0).is_none() {
            return;
        }
        if self.ops.len() >= self.limits.ops {
            self.overflowed = true;
            return;
        }
        self.ops.push(op);
    }

    pub(crate) fn emit(&mut self, fill: Option<Paint>, stroke: Option<Stroke>) {
        self.pending_stroke = false;
        let path = std::mem::take(&mut self.path);
        if path.is_empty() || (fill.is_none() && stroke.is_none()) {
            return;
        }
        let op = Op::Shape(Shape {
            path,
            fill,
            stroke,
            even_odd: self.dc.even_odd,
            clip: self.dc.clip.clone(),
        });
        self.push_op(op);
    }

    pub(crate) fn flush_pending(&mut self) {
        if self.bracketed {
            return;
        }
        if !self.pending_stroke {
            self.path.clear();
            return;
        }
        let stroke = self.pen_stroke();
        self.emit(None, stroke);
    }

    /// Fills and strokes the path built so far, unless a path bracket collects it.
    pub(crate) fn paint_figure(&mut self, filled: bool) {
        if self.bracketed {
            return;
        }
        let fill = if filled { self.brush_fill() } else { None };
        let stroke = self.pen_stroke();
        self.emit(fill, stroke);
    }

    pub(crate) fn append_arc(
        &mut self,
        box_rect: (f64, f64, f64, f64),
        start: (f64, f64),
        end: (f64, f64),
    ) {
        let (x0, y0, x1, y1) = box_rect;
        let (cx, cy) = ((x0 + x1) / 2.0, (y0 + y1) / 2.0);
        let (rx, ry) = ((x1 - x0) / 2.0, (y1 - y0) / 2.0);
        if rx == 0.0 || ry == 0.0 {
            return;
        }
        let angle = |p: (f64, f64)| ((p.1 - cy) / ry).atan2((p.0 - cx) / rx);
        let from = angle(start);
        let mut sweep = angle(end) - from;
        if self.dc.clockwise {
            while sweep <= 0.0 {
                sweep += std::f64::consts::TAU;
            }
        } else {
            while sweep >= 0.0 {
                sweep -= std::f64::consts::TAU;
            }
        }
        self.append_sweep((cx, cy), (rx, ry), from, sweep);
    }

    /// Straight segments around an ellipse from `from` through `sweep` radians.
    pub(crate) fn append_sweep(
        &mut self,
        centre: (f64, f64),
        radii: (f64, f64),
        from: f64,
        sweep: f64,
    ) {
        let (cx, cy) = centre;
        let (rx, ry) = radii;
        let steps = ((sweep.abs() / std::f64::consts::TAU) * ARC_SEGMENTS as f64).ceil() as usize;
        let steps = steps.clamp(2, ARC_SEGMENTS);
        for step in 0..=steps {
            let theta = from + sweep * step as f64 / steps as f64;
            let (x, y) = (cx + rx * theta.cos(), cy + ry * theta.sin());
            if step == 0 && self.path.is_empty() {
                self.move_to(x, y);
            } else {
                self.line_to(x, y);
            }
        }
    }

    pub(crate) fn append_ellipse(&mut self, box_rect: (f64, f64, f64, f64)) {
        let (x0, y0, x1, y1) = box_rect;
        let (cx, cy) = ((x0 + x1) / 2.0, (y0 + y1) / 2.0);
        let (rx, ry) = ((x1 - x0) / 2.0, (y1 - y0) / 2.0);
        for step in 0..ARC_SEGMENTS {
            let theta = std::f64::consts::TAU * step as f64 / ARC_SEGMENTS as f64;
            let (x, y) = (cx + rx * theta.cos(), cy + ry * theta.sin());
            if step == 0 {
                self.move_to(x, y);
            } else {
                self.line_to(x, y);
            }
        }
        self.push(PathCommand::Close);
    }

    pub(crate) fn append_rect(&mut self, box_rect: (f64, f64, f64, f64)) {
        let (x0, y0, x1, y1) = box_rect;
        self.move_to(x0, y0);
        self.line_to(x1, y0);
        self.line_to(x1, y1);
        self.line_to(x0, y1);
        self.push(PathCommand::Close);
    }

    /// A rectangle whose corners are quarter ellipses `corner` wide and tall.
    pub(crate) fn append_round_rect(&mut self, box_rect: (f64, f64, f64, f64), corner: (f64, f64)) {
        let (x0, y0, x1, y1) = (
            box_rect.0.min(box_rect.2),
            box_rect.1.min(box_rect.3),
            box_rect.0.max(box_rect.2),
            box_rect.1.max(box_rect.3),
        );
        let rx = (corner.0.abs() / 2.0).min((x1 - x0) / 2.0);
        let ry = (corner.1.abs() / 2.0).min((y1 - y0) / 2.0);
        if rx <= 0.0 || ry <= 0.0 {
            self.append_rect((x0, y0, x1, y1));
            return;
        }
        use std::f64::consts::{FRAC_PI_2, PI};
        let quarter = FRAC_PI_2;
        let corners = [
            ((x1 - rx, y0 + ry), -quarter),
            ((x1 - rx, y1 - ry), 0.0),
            ((x0 + rx, y1 - ry), quarter),
            ((x0 + rx, y0 + ry), PI),
        ];
        for (centre, from) in corners {
            let steps = ARC_SEGMENTS / 4;
            for step in 0..=steps {
                let theta = from + quarter * step as f64 / steps as f64;
                let (x, y) = (centre.0 + rx * theta.cos(), centre.1 + ry * theta.sin());
                if self.path.is_empty() {
                    self.move_to(x, y);
                } else {
                    self.line_to(x, y);
                }
            }
        }
        self.push(PathCommand::Close);
    }

    pub(crate) fn select(&mut self, object: GdiObject) {
        match object {
            GdiObject::Pen(pen) => self.dc.pen = Some(pen),
            GdiObject::Brush(brush) => self.dc.brush = Some(brush),
            GdiObject::Font(font) => self.dc.font = Some(font),
            GdiObject::Region(_) | GdiObject::Opaque => {}
        }
    }

    pub(crate) fn store(&mut self, index: usize, object: GdiObject) {
        if index >= self.objects.len() {
            if index >= MAX_HANDLES {
                return;
            }
            self.objects.resize(index + 1, None);
        }
        self.objects[index] = Some(object);
    }

    /// The region `path` covers, in output units, combined into the clip.
    pub(crate) fn combine_clip(
        &mut self,
        path: Vec<PathCommand>,
        even_odd: bool,
        mode: Combine,
    ) -> Option<()> {
        if FULL {
            let copied = if matches!(mode, Combine::Or | Combine::Xor)
                && !same_clip(&self.dc.clip, &self.dc.meta)
            {
                self.dc
                    .clip
                    .as_ref()
                    .map_or(0, |clip| clip.region.path.len())
            } else {
                0
            };
            self.charge(path.len().saturating_add(copied), 0)?;
        }
        let region = |exclude| ClipRegion {
            path: path.clone(),
            even_odd,
            exclude,
        };
        let clip = match mode {
            Combine::Copy => chain(self.dc.meta.clone(), region(false)),
            Combine::And => chain(self.dc.clip.clone(), region(false)),
            Combine::Diff => chain(self.dc.clip.clone(), region(true)),
            Combine::Or | Combine::Xor => {
                if same_clip(&self.dc.clip, &self.dc.meta) {
                    if mode == Combine::Or {
                        return Some(());
                    }
                    chain(self.dc.meta.clone(), region(true))
                } else {
                    let own = self.dc.clip.clone()?;
                    if !same_clip(&own.parent, &self.dc.meta) || own.region.exclude {
                        self.omit("clip combinations approximated")?;
                    }
                    let mut joined = own.region.path.clone();
                    joined.extend(path);
                    chain(
                        own.parent.clone(),
                        ClipRegion {
                            path: joined,
                            even_odd: mode == Combine::Xor || own.region.even_odd || even_odd,
                            exclude: false,
                        },
                    )
                }
            }
        };
        if self.clip_depth(&clip)? > self.limits.clip_depth {
            return self.refuse("clip regions nest past the depth limit");
        }
        self.dc.clip_rect = clip
            .as_ref()
            .and_then(|clip| axis_aligned_rect(&clip.region.path));
        self.dc.clip = clip;
        Some(())
    }

    /// Intersects the clip with a device-aligned rectangle given in output
    /// units, merging it into the innermost rectangle when there is one.
    pub(crate) fn intersect_rect(&mut self, rect: [f64; 4]) -> Option<()> {
        if let Some(own) = self.dc.clip.as_deref()
            && !same_clip(&self.dc.clip, &self.dc.meta)
            && !own.region.exclude
            && let Some(inner) = self.dc.clip_rect
        {
            let merged = [
                inner[0].max(rect[0]),
                inner[1].max(rect[1]),
                inner[2].min(rect[2]),
                inner[3].min(rect[3]),
            ];
            let merged = [
                merged[0],
                merged[1],
                merged[2].max(merged[0]),
                merged[3].max(merged[1]),
            ];
            let parent = own.parent.clone();
            if FULL {
                self.charge(5, 0)?;
            }
            self.dc.clip = chain(
                parent,
                ClipRegion {
                    path: rect_path(merged),
                    even_odd: false,
                    exclude: false,
                },
            );
            self.dc.clip_rect = Some(merged);
            return Some(());
        }
        self.combine_clip(rect_path(rect), false, Combine::And)
    }

    pub(crate) fn reset_clip(&mut self) {
        self.dc.clip = self.dc.meta.clone();
        self.dc.clip_rect = self.dc.meta_rect;
    }

    pub(crate) fn set_meta_region(&mut self) {
        self.dc.meta = self.dc.clip.clone();
        self.dc.meta_rect = self.dc.clip_rect;
    }

    pub(crate) fn clip_depth(&mut self, clip: &Clip) -> Option<usize> {
        let mut depth = 0;
        let mut at = clip.as_deref();
        while let Some(link) = at {
            if FULL {
                self.charge(1, 0)?;
            }
            depth += 1;
            at = link.parent.as_deref();
        }
        Some(depth)
    }

    /// Moves the application clip by a device offset in output units.
    pub(crate) fn offset_clip(&mut self, dx: f64, dy: f64) -> Option<()> {
        let mut links = Vec::new();
        let mut at = self.dc.clip.clone();
        while let Some(link) = at.clone() {
            if same_clip(&at, &self.dc.meta) {
                break;
            }
            self.charge(link.region.path.len().saturating_add(1), 0)?;
            at = link.parent.clone();
            links.push(link);
        }
        if links.is_empty() {
            return Some(());
        }
        let mut clip = self.dc.meta.clone();
        for link in links.into_iter().rev() {
            let mut region = link.region.clone();
            for command in &mut region.path {
                translate(command, dx, dy);
            }
            clip = chain(clip, region);
        }
        self.dc.clip_rect = clip
            .as_ref()
            .and_then(|clip| axis_aligned_rect(&clip.region.path));
        self.dc.clip = clip;
        Some(())
    }

    /// The logical rectangle `(l, t, r, b)` as a path in output units.
    pub(crate) fn logical_rect_path(&self, rect: (f64, f64, f64, f64)) -> Vec<PathCommand> {
        let (l, t, r, b) = rect;
        let corners = [(l, t), (r, t), (r, b), (l, b)];
        let mut path = Vec::with_capacity(5);
        for (index, (x, y)) in corners.into_iter().enumerate() {
            let (x, y) = self.point(x, y);
            path.push(if index == 0 {
                PathCommand::Move { x, y }
            } else {
                PathCommand::Line { x, y }
            });
        }
        path.push(PathCommand::Close);
        path
    }
}

pub(crate) fn chain(parent: Clip, region: ClipRegion) -> Clip {
    Some(Arc::new(ClipChain { parent, region }))
}

/// Whether two clips are the same chain, not merely equal regions.
pub(crate) fn same_clip(a: &Clip, b: &Clip) -> bool {
    match (a, b) {
        (None, None) => true,
        (Some(a), Some(b)) => Arc::ptr_eq(a, b),
        _ => false,
    }
}

pub(crate) fn hatch_style(style: u32) -> Hatch {
    match style {
        1 => Hatch::Vertical,
        2 => Hatch::ForwardDiagonal,
        3 => Hatch::BackwardDiagonal,
        4 => Hatch::Cross,
        5 => Hatch::DiagonalCross,
        _ => Hatch::Horizontal,
    }
}

pub(crate) fn rect_path(rect: [f64; 4]) -> Vec<PathCommand> {
    let [l, t, r, b] = rect;
    vec![
        PathCommand::Move { x: l, y: t },
        PathCommand::Line { x: r, y: t },
        PathCommand::Line { x: r, y: b },
        PathCommand::Line { x: l, y: b },
        PathCommand::Close,
    ]
}

pub(crate) fn translate(command: &mut PathCommand, dx: f64, dy: f64) {
    match command {
        PathCommand::Move { x, y } | PathCommand::Line { x, y } => {
            *x += dx;
            *y += dy;
        }
        PathCommand::Quad { cpx, cpy, x, y } => {
            *cpx += dx;
            *cpy += dy;
            *x += dx;
            *y += dy;
        }
        PathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => {
            *cp1x += dx;
            *cp1y += dy;
            *cp2x += dx;
            *cp2y += dy;
            *x += dx;
            *y += dy;
        }
        PathCommand::Close => {}
    }
}

/// Reads a path back as an axis-aligned rectangle, or `None` if it is not one.
pub(crate) fn axis_aligned_rect(path: &[PathCommand]) -> Option<[f64; 4]> {
    let mut corners: Vec<(f64, f64)> = Vec::with_capacity(5);
    for command in path {
        match command {
            PathCommand::Move { x, y } | PathCommand::Line { x, y } => {
                if corners.last() != Some(&(*x, *y)) {
                    corners.push((*x, *y));
                }
            }
            PathCommand::Close => {}
            _ => return None,
        }
        if corners.len() > 5 {
            return None;
        }
    }
    if corners.len() == 5 && corners[4] == corners[0] {
        corners.pop();
    }
    let [a, b, c, d] = corners[..] else {
        return None;
    };
    let across = a.1 == b.1 && b.0 == c.0 && c.1 == d.1 && d.0 == a.0;
    let down = a.0 == b.0 && b.1 == c.1 && c.0 == d.0 && d.1 == a.1;
    if !(across || down) {
        return None;
    }
    Some([a.0.min(c.0), a.1.min(c.1), a.0.max(c.0), a.1.max(c.1)])
}

pub(crate) fn stock_object(index: u32) -> Option<GdiObject> {
    let brush = |color| Some(GdiObject::Brush(Brush::solid(color, true)));
    let pen = |color| Some(GdiObject::Pen(Pen::solid(color, 0.0, true)));
    match index {
        0 => brush(0x00ff_ffff),
        1 => brush(0x00c0_c0c0),
        2 => brush(0x0080_8080),
        3 => brush(0x0040_4040),
        4 => brush(0x0000_0000),
        5 => Some(GdiObject::Brush(Brush::solid(0, false))),
        6 => pen(0x00ff_ffff),
        7 => pen(0x0000_0000),
        8 => Some(GdiObject::Pen(Pen::solid(0, 0.0, false))),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_records::*;

    #[test]
    fn rectangle_recognition_accepts_repeated_points() {
        let rect = [0.0, 0.0, 10.0, 10.0];
        let mut path = rect_path(rect);
        path.splice(1..1, vec![PathCommand::Line { x: 0.0, y: 0.0 }; 64]);
        assert_eq!(axis_aligned_rect(&path), Some(rect));
        assert_eq!(
            axis_aligned_rect(&vec![PathCommand::Line { x: 0.0, y: 0.0 }; 64]),
            None
        );
    }

    #[test]
    fn empty_clip_offsets_charge_every_link_before_copying() {
        let bytes = |offsets: usize| {
            let mut records = Vec::new();
            for _ in 0..FULL_LIMITS.clip_depth {
                records.extend([bare(59), bare(60), value(67, 1)]);
            }
            for _ in 0..offsets {
                records.push((26, i32s(&[1, 1])));
            }
            Emf::new(100, 100).recs(records).bytes()
        };
        let replay = |offsets: usize, work| {
            let budget = Rc::new(SharedBudget {
                remaining: Cell::new(crate::ReplayBudget { work, pixels: 0 }),
                exceeded: Cell::new(false),
            });
            let result =
                crate::play_emf::<true>(&bytes(offsets), 0, false, Some(Rc::clone(&budget)));
            (result, budget)
        };
        let (base, budget) = replay(0, 100_000);
        let base = base.unwrap();
        assert_eq!(ClipChain::depth(&base.dc.clip), 128);
        let spent = 100_000 - budget.remaining.get().work;
        let (offset, budget) = replay(2, 100_000);
        let offset = offset.unwrap();
        assert_eq!(offset.commands - base.commands, 2 * 128);
        assert_eq!(100_000 - budget.remaining.get().work - spent, 2 * 129);
        let (result, budget) = replay(8, spent + 2 * 129);
        assert!(result.is_err());
        assert!(budget.exceeded.get());
        assert_eq!(budget.remaining.get().work, 0);
    }

    #[test]
    fn repeated_rectangle_clips_decode_and_cache_text_intersections() {
        let clip = |repeats| {
            let mut points = vec![(0, 0); repeats];
            points.extend([(10, 0), (10, 10), (0, 10)]);
            vec![
                bare(59),
                poly16(87, &points),
                bare(61),
                bare(60),
                value(67, 5),
            ]
        };
        let strict = |repeats| {
            let mut records = clip(repeats);
            records.push(rect(0, 0, 20, 20));
            crate::decode(&Emf::new(100, 100).recs(records).bytes()).unwrap()
        };
        assert_eq!(strict(64), strict(1));
        let replay = |repeats, count| {
            let mut records = clip(repeats);
            for _ in 0..count {
                records.extend([
                    bare(33),
                    intersect_clip(0, 0, 10, 10),
                    value(34, u32::MAX),
                    text_out(0, 0, "x", None, 4, [0, 0, 10, 10]),
                ]);
            }
            let budget = Rc::new(SharedBudget {
                remaining: Cell::new(crate::ReplayBudget {
                    work: 10_000,
                    pixels: 0,
                }),
                exceeded: Cell::new(false),
            });
            let player = crate::play_emf::<true>(
                &Emf::new(100, 100).recs(records).bytes(),
                0,
                false,
                Some(Rc::clone(&budget)),
            )
            .unwrap();
            assert_eq!(player.ops.len(), count);
            assert_eq!(
                player.dc.clip.as_ref().unwrap().region.path.len(),
                repeats + 4
            );
            assert!(player.dc.clip_rect.is_some());
            for op in &player.ops {
                let Op::Text(text) = op else {
                    panic!("a clipped text run");
                };
                assert_eq!(ClipChain::depth(&text.clip), 1);
                assert_eq!(
                    axis_aligned_rect(&text.clip.as_ref().unwrap().region.path),
                    player.dc.clip_rect
                );
            }
            10_000 - budget.remaining.get().work
        };
        assert_eq!(replay(64, 8) - replay(64, 1), replay(1, 8) - replay(1, 1));
        assert_eq!(replay(64, 8) - replay(1, 8), 2 * 63);
    }

    #[test]
    fn ignored_clip_unions_charge_the_input_path() {
        let budget = Rc::new(SharedBudget {
            remaining: Cell::new(crate::ReplayBudget {
                work: 100,
                pixels: 0,
            }),
            exceeded: Cell::new(false),
        });
        let mut player = Player::<true>::new((0.0, 0.0, 10.0, 10.0), 0, (10.0, 10.0));
        player.budget = Some(Rc::clone(&budget));
        player
            .combine_clip(
                vec![PathCommand::Line { x: 0.0, y: 0.0 }; 64],
                false,
                Combine::Or,
            )
            .unwrap();
        assert!(player.dc.clip.is_none());
        assert_eq!(player.commands, 64);
        assert_eq!(budget.remaining.get().work, 36);
    }

    #[test]
    fn repeated_intersections_and_clipped_text_keep_work_linear() {
        let replay = |count: usize| {
            let mut records = vec![bare(59), poly16(87, &[(0, 0); 64]), bare(60), value(67, 5)];
            let mut small = i32s(&[0, 0, 1, 0x0204, 1]);
            small.extend(f32s(&[1.0, 1.0]));
            small.extend(i32s(&[0, 0, 10, 10]));
            small.extend(b"x\0\0\0");
            for _ in 0..count {
                records.extend([
                    bare(33),
                    intersect_clip(0, 0, 10, 10),
                    value(34, u32::MAX),
                    text_out(0, 0, "x", None, 4, [0, 0, 10, 10]),
                    (108, small.clone()),
                ]);
            }
            let budget = Rc::new(SharedBudget {
                remaining: Cell::new(crate::ReplayBudget {
                    work: 10_000,
                    pixels: 0,
                }),
                exceeded: Cell::new(false),
            });
            let player = crate::play_emf::<true>(
                &Emf::new(100, 100).recs(records).bytes(),
                0,
                false,
                Some(Rc::clone(&budget)),
            )
            .unwrap();
            assert_eq!(player.commands, 129 + count * 21);
            assert_eq!(player.ops.len(), count * 2);
            assert_eq!(player.dc.clip.as_ref().unwrap().region.path.len(), 64);
            for op in &player.ops {
                let Op::Text(text) = op else {
                    panic!("a clipped text run");
                };
                assert_eq!(ClipChain::depth(&text.clip), 2);
            }
            10_000 - budget.remaining.get().work
        };
        assert_eq!(replay(5) - replay(1), 4 * 30);
    }
}
