//! Replays EMF and WMF picture parts into display-list paths.
//!
//! Neither `image` nor a browser `CanvasImageSource` decodes a metafile, so a
//! picture whose media is EMF or WMF paints nothing at all. The records are
//! vector drawing operations, so the layout pass replays them into the
//! `Primitive::Shape` paths both backends already draw rather than handing
//! bytes to a raster decoder that will reject them.
//!
//! Coordinates come out as fractions of the metafile's frame rectangle, which
//! is what a `Primitive::Shape` path is measured in.

use ooxml_drawingml::GeometryPathCommand;

/// Records read from one part. A metafile that keeps drawing past this is
/// treated as undecodable rather than allowed to spend the render budget.
const MAX_RECORDS: usize = 200_000;
/// Fill/stroke operations one metafile may contribute to a slide.
const MAX_OPS: usize = 4_096;
/// Path commands across the whole drawing.
const MAX_COMMANDS: usize = 400_000;
/// Points one record may carry, so a corrupt count cannot allocate.
const MAX_POINTS_PER_RECORD: usize = 65_536;
/// Line segments an arc is flattened to over a full turn.
const ARC_SEGMENTS: usize = 64;

/// One fill or stroke, with its path in fractions of the frame rectangle.
#[derive(Debug, Clone, PartialEq)]
pub struct MetafileOp {
    pub path: Vec<GeometryPathCommand>,
    pub fill: Option<String>,
    pub stroke: Option<MetafileStroke>,
}

/// A pen, with its width as a fraction of the frame rectangle's width.
#[derive(Debug, Clone, PartialEq)]
pub struct MetafileStroke {
    pub color: String,
    pub width: f64,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct MetafileDrawing {
    pub ops: Vec<MetafileOp>,
}

/// Decodes EMF or WMF bytes, or `None` when the bytes are neither, are
/// malformed, or draw nothing this module understands.
pub fn decode(bytes: &[u8]) -> Option<MetafileDrawing> {
    let drawing = decode_emf(bytes).or_else(|| decode_wmf(bytes))?;
    (!drawing.ops.is_empty()).then_some(drawing)
}

fn u16_at(bytes: &[u8], offset: usize) -> Option<u16> {
    let slice = bytes.get(offset..offset + 2)?;
    Some(u16::from_le_bytes([slice[0], slice[1]]))
}

fn u32_at(bytes: &[u8], offset: usize) -> Option<u32> {
    let slice = bytes.get(offset..offset + 4)?;
    Some(u32::from_le_bytes([slice[0], slice[1], slice[2], slice[3]]))
}

fn i32_at(bytes: &[u8], offset: usize) -> Option<i32> {
    u32_at(bytes, offset).map(|value| value as i32)
}

fn i16_at(bytes: &[u8], offset: usize) -> Option<i16> {
    u16_at(bytes, offset).map(|value| value as i16)
}

fn f32_at(bytes: &[u8], offset: usize) -> Option<f32> {
    u32_at(bytes, offset).map(f32::from_bits)
}

/// A GDI `COLORREF` is `0x00bbggrr`.
fn colorref_hex(value: u32) -> String {
    format!(
        "#{:02x}{:02x}{:02x}",
        value & 0xff,
        (value >> 8) & 0xff,
        (value >> 16) & 0xff
    )
}

#[derive(Clone, Copy, PartialEq)]
struct Pen {
    color: u32,
    /// Logical units; zero means one device unit.
    width: f64,
    visible: bool,
}

#[derive(Clone, Copy, PartialEq)]
struct Brush {
    color: u32,
    visible: bool,
}

#[derive(Clone, Copy)]
enum GdiObject {
    Pen(Pen),
    Brush(Brush),
}

/// Row-major 2x3 affine: `[m11, m12, m21, m22, dx, dy]`.
type Xform = [f64; 6];

const IDENTITY: Xform = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];

fn concat(first: Xform, then: Xform) -> Xform {
    [
        first[0] * then[0] + first[1] * then[2],
        first[0] * then[1] + first[1] * then[3],
        first[2] * then[0] + first[3] * then[2],
        first[2] * then[1] + first[3] * then[3],
        first[4] * then[0] + first[5] * then[2] + then[4],
        first[4] * then[1] + first[5] * then[3] + then[5],
    ]
}

#[derive(Clone, Copy)]
struct Dc {
    window_org: (f64, f64),
    window_ext: (f64, f64),
    viewport_org: (f64, f64),
    viewport_ext: (f64, f64),
    /// Both extents seen. A file that sets only a window extent is still
    /// mapping logical units 1:1 — `project17/image18.emf` sets one after it has
    /// finished drawing, and honouring it would shrink the pie to nothing.
    scaled: bool,
    xform: Xform,
    pen: Option<Pen>,
    brush: Option<Brush>,
}

impl Default for Dc {
    fn default() -> Self {
        Self {
            window_org: (0.0, 0.0),
            window_ext: (1.0, 1.0),
            viewport_org: (0.0, 0.0),
            viewport_ext: (1.0, 1.0),
            scaled: false,
            xform: IDENTITY,
            pen: Some(Pen {
                color: 0,
                width: 0.0,
                visible: true,
            }),
            brush: Some(Brush {
                color: 0x00ff_ffff,
                visible: true,
            }),
        }
    }
}

/// Maps logical coordinates onto the frame rectangle and collects the ops.
struct Player {
    dc: Dc,
    saved: Vec<Dc>,
    objects: Vec<Option<GdiObject>>,
    /// Frame rectangle in device units: origin and size.
    frame: (f64, f64, f64, f64),
    path: Vec<GeometryPathCommand>,
    current: (f64, f64),
    /// True between `BEGINPATH` and the `FILLPATH`/`STROKEPATH` that ends it.
    bracketed: bool,
    /// True when unbracketed line records have queued a stroke not yet emitted.
    pending_stroke: bool,
    window_ext_set: bool,
    viewport_ext_set: bool,
    commands: usize,
    ops: Vec<MetafileOp>,
    overflowed: bool,
}

impl Player {
    fn new(frame: (f64, f64, f64, f64), handles: usize) -> Self {
        Self {
            dc: Dc::default(),
            saved: Vec::new(),
            objects: vec![None; handles.min(4096)],
            frame,
            path: Vec::new(),
            current: (0.0, 0.0),
            bracketed: false,
            pending_stroke: false,
            window_ext_set: false,
            viewport_ext_set: false,
            commands: 0,
            ops: Vec::new(),
            overflowed: false,
        }
    }

    /// Logical point to a fraction of the frame rectangle.
    fn point(&self, x: f64, y: f64) -> (f64, f64) {
        let m = self.dc.xform;
        let wx = x * m[0] + y * m[2] + m[4];
        let wy = x * m[1] + y * m[3] + m[5];
        let (sx, sy) = if self.dc.scaled {
            (
                self.dc.viewport_ext.0 / self.dc.window_ext.0,
                self.dc.viewport_ext.1 / self.dc.window_ext.1,
            )
        } else {
            (1.0, 1.0)
        };
        let dx = (wx - self.dc.window_org.0) * sx + self.dc.viewport_org.0;
        let dy = (wy - self.dc.window_org.1) * sy + self.dc.viewport_org.1;
        (
            (dx - self.frame.0) / self.frame.2,
            (dy - self.frame.1) / self.frame.3,
        )
    }

    /// Pen width as a fraction of the frame's width. A zero-width pen comes
    /// back as zero: GDI draws it one pixel wide on the output device, which
    /// only the caller knows the size of.
    fn pen_width(&self, width: f64) -> f64 {
        let scale = if self.dc.scaled {
            (self.dc.viewport_ext.0 / self.dc.window_ext.0).abs()
        } else {
            1.0
        };
        (width * scale * self.dc.xform[0].abs() / self.frame.2).abs()
    }

    fn push(&mut self, command: GeometryPathCommand) {
        self.commands += 1;
        if self.commands > MAX_COMMANDS {
            self.overflowed = true;
            return;
        }
        self.path.push(command);
    }

    fn move_to(&mut self, x: f64, y: f64) {
        self.current = (x, y);
        let (px, py) = self.point(x, y);
        self.push(GeometryPathCommand::Move { x: px, y: py });
    }

    fn line_to(&mut self, x: f64, y: f64) {
        self.current = (x, y);
        let (px, py) = self.point(x, y);
        self.push(GeometryPathCommand::Line { x: px, y: py });
    }

    fn cubic_to(&mut self, points: [(f64, f64); 3]) {
        self.current = points[2];
        let a = self.point(points[0].0, points[0].1);
        let b = self.point(points[1].0, points[1].1);
        let c = self.point(points[2].0, points[2].1);
        self.push(GeometryPathCommand::Cubic {
            cp1x: a.0,
            cp1y: a.1,
            cp2x: b.0,
            cp2y: b.1,
            x: c.0,
            y: c.1,
        });
    }

    fn brush_fill(&self) -> Option<String> {
        self.dc
            .brush
            .filter(|brush| brush.visible)
            .map(|brush| colorref_hex(brush.color))
    }

    fn pen_stroke(&self) -> Option<MetafileStroke> {
        self.dc
            .pen
            .filter(|pen| pen.visible)
            .map(|pen| MetafileStroke {
                color: colorref_hex(pen.color),
                width: self.pen_width(pen.width),
            })
    }

    fn emit(&mut self, fill: Option<String>, stroke: Option<MetafileStroke>) {
        self.pending_stroke = false;
        let path = std::mem::take(&mut self.path);
        if path.is_empty() || (fill.is_none() && stroke.is_none()) {
            return;
        }
        if self.ops.len() >= MAX_OPS {
            self.overflowed = true;
            return;
        }
        self.ops.push(MetafileOp { path, fill, stroke });
    }

    /// Draws the polyline that unbracketed `LINETO`-style records built up.
    fn flush_pending(&mut self) {
        if !self.pending_stroke {
            self.path.clear();
            return;
        }
        let stroke = self.pen_stroke();
        self.emit(None, stroke);
    }

    fn append_arc(&mut self, box_rect: (f64, f64, f64, f64), start: (f64, f64), end: (f64, f64)) {
        let (x0, y0, x1, y1) = box_rect;
        let (cx, cy) = ((x0 + x1) / 2.0, (y0 + y1) / 2.0);
        let (rx, ry) = ((x1 - x0) / 2.0, (y1 - y0) / 2.0);
        if rx == 0.0 || ry == 0.0 {
            return;
        }
        let angle = |p: (f64, f64)| ((p.1 - cy) / ry).atan2((p.0 - cx) / rx);
        let from = angle(start);
        let mut sweep = angle(end) - from;
        // GDI sweeps counter-clockwise, which is decreasing angle once the
        // y-down logical axis has been folded into `atan2`.
        while sweep >= 0.0 {
            sweep -= std::f64::consts::TAU;
        }
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

    fn append_ellipse(&mut self, box_rect: (f64, f64, f64, f64)) {
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
        self.push(GeometryPathCommand::Close);
    }

    fn append_rect(&mut self, box_rect: (f64, f64, f64, f64)) {
        let (x0, y0, x1, y1) = box_rect;
        self.move_to(x0, y0);
        self.line_to(x1, y0);
        self.line_to(x1, y1);
        self.line_to(x0, y1);
        self.push(GeometryPathCommand::Close);
    }

    fn select(&mut self, object: GdiObject) {
        match object {
            GdiObject::Pen(pen) => self.dc.pen = Some(pen),
            GdiObject::Brush(brush) => self.dc.brush = Some(brush),
        }
    }

    fn store(&mut self, index: usize, object: GdiObject) {
        if index >= self.objects.len() {
            if index >= 4096 {
                return;
            }
            self.objects.resize(index + 1, None);
        }
        self.objects[index] = Some(object);
    }

    fn finish(self) -> Option<MetafileDrawing> {
        (!self.overflowed).then_some(MetafileDrawing { ops: self.ops })
    }
}

/// The GDI stock objects a metafile may select without creating them.
fn stock_object(index: u32) -> Option<GdiObject> {
    let brush = |color| {
        Some(GdiObject::Brush(Brush {
            color,
            visible: true,
        }))
    };
    let pen = |color| {
        Some(GdiObject::Pen(Pen {
            color,
            width: 0.0,
            visible: true,
        }))
    };
    match index {
        0 => brush(0x00ff_ffff),
        1 => brush(0x00c0_c0c0),
        2 => brush(0x0080_8080),
        3 => brush(0x0040_4040),
        4 => brush(0x0000_0000),
        5 => Some(GdiObject::Brush(Brush {
            color: 0,
            visible: false,
        })),
        6 => pen(0x00ff_ffff),
        7 => pen(0x0000_0000),
        8 => Some(GdiObject::Pen(Pen {
            color: 0,
            width: 0.0,
            visible: false,
        })),
        _ => None,
    }
}

/// `BS_NULL` brushes and `PS_NULL` pens paint nothing; hatches and patterns are
/// approximated by their colour rather than dropped.
fn brush_from_style(style: u32, color: u32) -> Brush {
    Brush {
        color,
        visible: style != 1,
    }
}

fn pen_from_style(style: u32, width: f64, color: u32) -> Pen {
    Pen {
        color,
        width,
        visible: style & 0x0f != 5,
    }
}

// ---------------------------------------------------------------------------
// EMF
// ---------------------------------------------------------------------------

const EMF_SIGNATURE: u32 = 0x464D_4520; // " EMF"

/// The frame rectangle in device units. `rclFrame` is in hundredths of a
/// millimetre; scaling it by the recorded device resolution gives the rectangle
/// PowerPoint stretches onto the picture's box. Falling back to `rclBounds`
/// would zoom the ink to the box edges — visible on the MSGraph pies, whose
/// frame is a fifth wider than their ink.
fn emf_frame(bytes: &[u8]) -> Option<(f64, f64, f64, f64)> {
    let bounds = (
        i32_at(bytes, 8)? as f64,
        i32_at(bytes, 12)? as f64,
        i32_at(bytes, 16)? as f64 + 1.0,
        i32_at(bytes, 20)? as f64 + 1.0,
    );
    let frame = (
        i32_at(bytes, 24)? as f64,
        i32_at(bytes, 28)? as f64,
        i32_at(bytes, 32)? as f64,
        i32_at(bytes, 36)? as f64,
    );
    let device = (i32_at(bytes, 72)? as f64, i32_at(bytes, 76)? as f64);
    let millimetres = (i32_at(bytes, 80)? as f64, i32_at(bytes, 84)? as f64);
    let rect = if millimetres.0 > 0.0 && millimetres.1 > 0.0 {
        let per_unit = (
            device.0 / (millimetres.0 * 100.0),
            device.1 / (millimetres.1 * 100.0),
        );
        (
            frame.0 * per_unit.0,
            frame.1 * per_unit.1,
            frame.2 * per_unit.0,
            frame.3 * per_unit.1,
        )
    } else {
        bounds
    };
    let (width, height) = (rect.2 - rect.0, rect.3 - rect.1);
    if width.abs() < f64::EPSILON || height.abs() < f64::EPSILON {
        let (width, height) = (bounds.2 - bounds.0, bounds.3 - bounds.1);
        if width.abs() < f64::EPSILON || height.abs() < f64::EPSILON {
            return None;
        }
        return Some((bounds.0, bounds.1, width, height));
    }
    Some((rect.0, rect.1, width, height))
}

fn decode_emf(bytes: &[u8]) -> Option<MetafileDrawing> {
    if u32_at(bytes, 0)? != 1 || u32_at(bytes, 40)? != EMF_SIGNATURE {
        return None;
    }
    let handles = u16_at(bytes, 56)? as usize;
    let mut player = Player::new(emf_frame(bytes)?, handles + 1);
    let mut offset = 0usize;
    let mut records = 0usize;
    while offset + 8 <= bytes.len() {
        let kind = u32_at(bytes, offset)?;
        let size = u32_at(bytes, offset + 4)? as usize;
        if size < 8 || !size.is_multiple_of(4) || offset + size > bytes.len() {
            break;
        }
        records += 1;
        if records > MAX_RECORDS {
            return None;
        }
        if kind == 14 {
            break;
        }
        let body = offset + 8;
        emf_record(&mut player, bytes, kind, body);
        if player.overflowed {
            return None;
        }
        offset += size;
    }
    player.flush_pending();
    player.finish()
}

/// Reads `count` points, 16- or 32-bit, starting at `offset`.
fn read_points(bytes: &[u8], offset: usize, count: usize, small: bool) -> Option<Vec<(f64, f64)>> {
    if count > MAX_POINTS_PER_RECORD {
        return None;
    }
    let stride = if small { 4 } else { 8 };
    let mut points = Vec::with_capacity(count);
    for index in 0..count {
        let at = offset + index * stride;
        let point = if small {
            (i16_at(bytes, at)? as f64, i16_at(bytes, at + 2)? as f64)
        } else {
            (i32_at(bytes, at)? as f64, i32_at(bytes, at + 4)? as f64)
        };
        points.push(point);
    }
    Some(points)
}

fn emf_record(player: &mut Player, bytes: &[u8], kind: u32, body: usize) {
    // Poly records carry a bounds rectangle ahead of their count; Box records
    // do not.
    let poly_points = |small: bool| -> Option<Vec<(f64, f64)>> {
        let count = u32_at(bytes, body + 16)? as usize;
        read_points(bytes, body + 20, count, small)
    };
    let box_rect = || -> Option<(f64, f64, f64, f64)> {
        Some((
            i32_at(bytes, body)? as f64,
            i32_at(bytes, body + 4)? as f64,
            i32_at(bytes, body + 8)? as f64,
            i32_at(bytes, body + 12)? as f64,
        ))
    };
    match kind {
        // SETWINDOWEXTEX / SETWINDOWORGEX / SETVIEWPORTEXTEX / SETVIEWPORTORGEX
        9..=12 => {
            let Some((x, y)) = i32_at(bytes, body).zip(i32_at(bytes, body + 4)) else {
                return;
            };
            let value = (f64::from(x), f64::from(y));
            match kind {
                9 => {
                    player.dc.window_ext = value;
                    player.window_ext_set = true;
                }
                10 => player.dc.window_org = value,
                11 => {
                    player.dc.viewport_ext = value;
                    player.viewport_ext_set = true;
                }
                _ => player.dc.viewport_org = value,
            }
            player.dc.scaled = player.window_ext_set
                && player.viewport_ext_set
                && player.dc.window_ext.0 != 0.0
                && player.dc.window_ext.1 != 0.0;
        }
        33 => player.saved.push(player.dc), // SAVEDC
        34 => {
            // RESTOREDC's argument is a relative depth: -1 is the top of stack.
            let depth = i32_at(bytes, body).unwrap_or(-1);
            let back = if depth < 0 {
                (-depth) as usize
            } else {
                player.saved.len().saturating_sub(depth as usize)
            };
            let keep = player.saved.len().saturating_sub(back.max(1));
            if let Some(dc) = player.saved.get(keep).copied() {
                player.dc = dc;
                player.saved.truncate(keep);
            }
        }
        35 | 36 => {
            // SETWORLDTRANSFORM / MODIFYWORLDTRANSFORM
            let mut matrix = IDENTITY;
            for (index, slot) in matrix.iter_mut().enumerate() {
                match f32_at(bytes, body + index * 4) {
                    Some(value) => *slot = f64::from(value),
                    None => return,
                }
            }
            let mode = if kind == 36 {
                u32_at(bytes, body + 24).unwrap_or(4)
            } else {
                4
            };
            player.dc.xform = match mode {
                1 => IDENTITY,
                2 => concat(matrix, player.dc.xform),
                3 => concat(player.dc.xform, matrix),
                _ => matrix,
            };
        }
        37 => {
            // SELECTOBJECT
            let Some(handle) = u32_at(bytes, body) else {
                return;
            };
            player.flush_pending();
            let object = if handle & 0x8000_0000 != 0 {
                stock_object(handle & 0x7fff_ffff)
            } else {
                player.objects.get(handle as usize).copied().flatten()
            };
            if let Some(object) = object {
                player.select(object);
            }
        }
        38 => {
            // CREATEPEN
            let (Some(handle), Some(style), Some(width), Some(color)) = (
                u32_at(bytes, body),
                u32_at(bytes, body + 4),
                i32_at(bytes, body + 8),
                u32_at(bytes, body + 16),
            ) else {
                return;
            };
            player.store(
                handle as usize,
                GdiObject::Pen(pen_from_style(style, f64::from(width), color)),
            );
        }
        95 => {
            // EXTCREATEPEN: the log pen follows the DIB offsets.
            let (Some(handle), Some(style), Some(width), Some(color)) = (
                u32_at(bytes, body),
                u32_at(bytes, body + 20),
                u32_at(bytes, body + 24),
                u32_at(bytes, body + 32),
            ) else {
                return;
            };
            player.store(
                handle as usize,
                GdiObject::Pen(pen_from_style(style, f64::from(width), color)),
            );
        }
        39 => {
            // CREATEBRUSHINDIRECT
            let (Some(handle), Some(style), Some(color)) = (
                u32_at(bytes, body),
                u32_at(bytes, body + 4),
                u32_at(bytes, body + 8),
            ) else {
                return;
            };
            player.store(
                handle as usize,
                GdiObject::Brush(brush_from_style(style, color)),
            );
        }
        40 => {
            // DELETEOBJECT
            if let Some(handle) = u32_at(bytes, body)
                && let Some(slot) = player.objects.get_mut(handle as usize)
            {
                *slot = None;
            }
        }
        59 => {
            // BEGINPATH
            player.flush_pending();
            player.bracketed = true;
        }
        60 => player.bracketed = false,                // ENDPATH
        61 => player.push(GeometryPathCommand::Close), // CLOSEFIGURE
        62 => {
            let fill = player.brush_fill();
            player.emit(fill, None);
        }
        63 => {
            let (fill, stroke) = (player.brush_fill(), player.pen_stroke());
            player.emit(fill, stroke);
        }
        64 => {
            let stroke = player.pen_stroke();
            player.emit(None, stroke);
        }
        27 => {
            // MOVETOEX
            let Some((x, y)) = i32_at(bytes, body).zip(i32_at(bytes, body + 4)) else {
                return;
            };
            if !player.bracketed {
                player.flush_pending();
            }
            player.move_to(f64::from(x), f64::from(y));
        }
        54 => {
            // LINETO
            let Some((x, y)) = i32_at(bytes, body).zip(i32_at(bytes, body + 4)) else {
                return;
            };
            player.line_to(f64::from(x), f64::from(y));
            player.pending_stroke = !player.bracketed;
        }
        // POLYBEZIERTO / POLYLINETO and their 16-bit forms.
        5 | 6 | 88 | 89 => {
            let small = kind >= 88;
            let Some(points) = poly_points(small) else {
                return;
            };
            if matches!(kind, 5 | 88) {
                for chunk in points.as_chunks::<3>().0 {
                    player.cubic_to([chunk[0], chunk[1], chunk[2]]);
                }
            } else {
                for point in points {
                    player.line_to(point.0, point.1);
                }
            }
            player.pending_stroke = !player.bracketed;
        }
        // POLYBEZIER / POLYGON / POLYLINE and their 16-bit forms.
        2 | 3 | 4 | 85 | 86 | 87 => {
            let small = kind >= 85;
            let Some(points) = poly_points(small) else {
                return;
            };
            let Some(first) = points.first().copied() else {
                return;
            };
            if !player.bracketed {
                player.flush_pending();
            }
            player.move_to(first.0, first.1);
            if matches!(kind, 2 | 85) {
                for chunk in points[1..].as_chunks::<3>().0 {
                    player.cubic_to([chunk[0], chunk[1], chunk[2]]);
                }
            } else {
                for point in &points[1..] {
                    player.line_to(point.0, point.1);
                }
            }
            let closed = matches!(kind, 3 | 86);
            if closed {
                player.push(GeometryPathCommand::Close);
            }
            if !player.bracketed {
                let fill = closed.then(|| player.brush_fill()).flatten();
                let stroke = player.pen_stroke();
                player.emit(fill, stroke);
            }
        }
        // POLYPOLYLINE / POLYPOLYGON and their 16-bit forms.
        7 | 8 | 90 | 91 => {
            let small = kind >= 90;
            let (Some(polygons), Some(total)) =
                (u32_at(bytes, body + 16), u32_at(bytes, body + 20))
            else {
                return;
            };
            let (polygons, total) = (polygons as usize, total as usize);
            if polygons > MAX_POINTS_PER_RECORD || total > MAX_POINTS_PER_RECORD {
                return;
            }
            let mut counts = Vec::with_capacity(polygons);
            for index in 0..polygons {
                match u32_at(bytes, body + 24 + index * 4) {
                    Some(count) => counts.push(count as usize),
                    None => return,
                }
            }
            let Some(points) = read_points(bytes, body + 24 + polygons * 4, total, small) else {
                return;
            };
            if !player.bracketed {
                player.flush_pending();
            }
            let closed = matches!(kind, 8 | 91);
            let mut start = 0usize;
            for count in counts {
                let Some(run) = points.get(start..start + count) else {
                    break;
                };
                start += count;
                let Some(first) = run.first().copied() else {
                    continue;
                };
                player.move_to(first.0, first.1);
                for point in &run[1..] {
                    player.line_to(point.0, point.1);
                }
                if closed {
                    player.push(GeometryPathCommand::Close);
                }
            }
            if !player.bracketed {
                let fill = closed.then(|| player.brush_fill()).flatten();
                let stroke = player.pen_stroke();
                player.emit(fill, stroke);
            }
        }
        42 | 43 => {
            // ELLIPSE / RECTANGLE
            let Some(rect) = box_rect() else { return };
            if !player.bracketed {
                player.flush_pending();
            }
            if kind == 42 {
                player.append_ellipse(rect);
            } else {
                player.append_rect(rect);
            }
            if !player.bracketed {
                let (fill, stroke) = (player.brush_fill(), player.pen_stroke());
                player.emit(fill, stroke);
            }
        }
        47 => {
            // PIE
            let (Some(rect), Some(sx), Some(sy), Some(ex), Some(ey)) = (
                box_rect(),
                i32_at(bytes, body + 16),
                i32_at(bytes, body + 20),
                i32_at(bytes, body + 24),
                i32_at(bytes, body + 28),
            ) else {
                return;
            };
            if !player.bracketed {
                player.flush_pending();
            }
            let centre = ((rect.0 + rect.2) / 2.0, (rect.1 + rect.3) / 2.0);
            player.move_to(centre.0, centre.1);
            player.append_arc(
                rect,
                (f64::from(sx), f64::from(sy)),
                (f64::from(ex), f64::from(ey)),
            );
            player.push(GeometryPathCommand::Close);
            if !player.bracketed {
                let (fill, stroke) = (player.brush_fill(), player.pen_stroke());
                player.emit(fill, stroke);
            }
        }
        _ => {}
    }
}

// ---------------------------------------------------------------------------
// WMF
// ---------------------------------------------------------------------------

const WMF_PLACEABLE_KEY: u32 = 0x9AC6_CDD7;

/// WMF handles are slots in a table, and a create record takes the first free
/// one rather than naming an index.
fn wmf_store(player: &mut Player, object: GdiObject) {
    if let Some(slot) = player.objects.iter_mut().find(|slot| slot.is_none()) {
        *slot = Some(object);
        return;
    }
    if player.objects.len() < 4096 {
        player.objects.push(Some(object));
    }
}

fn decode_wmf(bytes: &[u8]) -> Option<MetafileDrawing> {
    let placeable = u32_at(bytes, 0)? == WMF_PLACEABLE_KEY;
    let header = if placeable { 22 } else { 0 };
    let kind = u16_at(bytes, header)?;
    if kind != 1 && kind != 2 {
        return None;
    }
    if u16_at(bytes, header + 2)? != 9 {
        return None;
    }
    let placeable_frame = placeable
        .then(|| {
            let left = f64::from(i16_at(bytes, 6)?);
            let top = f64::from(i16_at(bytes, 8)?);
            let right = f64::from(i16_at(bytes, 10)?);
            let bottom = f64::from(i16_at(bytes, 12)?);
            Some((left, top, right - left, bottom - top))
        })
        .flatten();
    let handles = u16_at(bytes, header + 10)? as usize;
    let mut player = Player::new(
        placeable_frame.unwrap_or((0.0, 0.0, 1.0, 1.0)),
        handles.min(4096),
    );
    let mut offset = header + 18;
    let mut records = 0usize;
    let mut deferred: Vec<(usize, usize)> = Vec::new();
    while offset + 6 <= bytes.len() {
        let size = u32_at(bytes, offset)? as usize * 2;
        let function = u16_at(bytes, offset + 4)?;
        if size < 6 || offset + size > bytes.len() {
            break;
        }
        records += 1;
        if records > MAX_RECORDS {
            return None;
        }
        if function == 0 {
            break;
        }
        deferred.push((function as usize, offset + 6));
        offset += size;
    }
    // WMF has no world transform and no viewport, so the window the file
    // declares *is* the frame the picture stretches onto; folding it in here
    // keeps `Player::point` an identity map for WMF.
    let mut org = None;
    let mut ext = None;
    for (function, body) in &deferred {
        match function {
            0x020B => org = i16_at(bytes, *body + 2).zip(i16_at(bytes, *body)),
            0x020C => ext = i16_at(bytes, *body + 2).zip(i16_at(bytes, *body)),
            _ => {}
        }
    }
    if let (Some(org), Some((width, height))) = (org, ext)
        && width != 0
        && height != 0
    {
        player.frame = (
            f64::from(org.0),
            f64::from(org.1),
            f64::from(width),
            f64::from(height),
        );
    }
    if player.frame.2.abs() < f64::EPSILON || player.frame.3.abs() < f64::EPSILON {
        return None;
    }
    for (function, body) in deferred {
        wmf_record(&mut player, bytes, function, body);
        if player.overflowed {
            return None;
        }
    }
    player.flush_pending();
    player.finish()
}

fn wmf_record(player: &mut Player, bytes: &[u8], function: usize, body: usize) {
    // WMF point records store y before x.
    let point_at = |offset: usize| -> Option<(f64, f64)> {
        Some((
            f64::from(i16_at(bytes, offset + 2)?),
            f64::from(i16_at(bytes, offset)?),
        ))
    };
    match function {
        0x0214 => {
            // MOVETO
            if let Some(point) = point_at(body) {
                player.flush_pending();
                player.move_to(point.0, point.1);
            }
        }
        0x0213 => {
            // LINETO
            if let Some(point) = point_at(body) {
                player.line_to(point.0, point.1);
                player.pending_stroke = true;
            }
        }
        0x02FA => {
            // CREATEPENINDIRECT
            let (Some(style), Some(width), Some(color)) = (
                u16_at(bytes, body),
                i16_at(bytes, body + 2),
                u32_at(bytes, body + 6),
            ) else {
                return;
            };
            let pen = pen_from_style(u32::from(style), f64::from(width), color);
            wmf_store(player, GdiObject::Pen(pen));
        }
        0x02FC => {
            // CREATEBRUSHINDIRECT
            let (Some(style), Some(color)) = (u16_at(bytes, body), u32_at(bytes, body + 2)) else {
                return;
            };
            let brush = brush_from_style(u32::from(style), color);
            wmf_store(player, GdiObject::Brush(brush));
        }
        0x012D => {
            // SELECTOBJECT
            let Some(index) = u16_at(bytes, body) else {
                return;
            };
            player.flush_pending();
            if let Some(object) = player.objects.get(index as usize).copied().flatten() {
                player.select(object);
            }
        }
        0x01F0 => {
            // DELETEOBJECT
            if let Some(index) = u16_at(bytes, body)
                && let Some(slot) = player.objects.get_mut(index as usize)
            {
                *slot = None;
            }
        }
        0x0324 | 0x0325 => {
            // POLYGON / POLYLINE
            let Some(count) = u16_at(bytes, body) else {
                return;
            };
            let Some(points) = read_points(bytes, body + 2, count as usize, true) else {
                return;
            };
            let Some(first) = points.first().copied() else {
                return;
            };
            player.flush_pending();
            player.move_to(first.0, first.1);
            for point in &points[1..] {
                player.line_to(point.0, point.1);
            }
            let closed = function == 0x0324;
            if closed {
                player.push(GeometryPathCommand::Close);
            }
            let fill = closed.then(|| player.brush_fill()).flatten();
            let stroke = player.pen_stroke();
            player.emit(fill, stroke);
        }
        0x0538 => {
            // POLYPOLYGON
            let Some(polygons) = u16_at(bytes, body) else {
                return;
            };
            let polygons = polygons as usize;
            if polygons > MAX_POINTS_PER_RECORD {
                return;
            }
            let mut counts = Vec::with_capacity(polygons);
            let mut total = 0usize;
            for index in 0..polygons {
                match u16_at(bytes, body + 2 + index * 2) {
                    Some(count) => {
                        total += count as usize;
                        counts.push(count as usize);
                    }
                    None => return,
                }
            }
            let Some(points) = read_points(bytes, body + 2 + polygons * 2, total, true) else {
                return;
            };
            player.flush_pending();
            let mut start = 0usize;
            for count in counts {
                let Some(run) = points.get(start..start + count) else {
                    break;
                };
                start += count;
                let Some(first) = run.first().copied() else {
                    continue;
                };
                player.move_to(first.0, first.1);
                for point in &run[1..] {
                    player.line_to(point.0, point.1);
                }
                player.push(GeometryPathCommand::Close);
            }
            let (fill, stroke) = (player.brush_fill(), player.pen_stroke());
            player.emit(fill, stroke);
        }
        0x041B | 0x0418 => {
            // RECTANGLE / ELLIPSE: bottom, right, top, left.
            let (Some(bottom), Some(right), Some(top), Some(left)) = (
                i16_at(bytes, body),
                i16_at(bytes, body + 2),
                i16_at(bytes, body + 4),
                i16_at(bytes, body + 6),
            ) else {
                return;
            };
            let rect = (
                f64::from(left),
                f64::from(top),
                f64::from(right),
                f64::from(bottom),
            );
            player.flush_pending();
            if function == 0x0418 {
                player.append_ellipse(rect);
            } else {
                player.append_rect(rect);
            }
            let (fill, stroke) = (player.brush_fill(), player.pen_stroke());
            player.emit(fill, stroke);
        }
        _ => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// An EMF header whose frame rectangle maps onto `device` pixels.
    fn emf_header(bounds: [i32; 4], frame_device: [i32; 2]) -> Vec<u8> {
        let mut header = vec![0u8; 88];
        let put = |header: &mut Vec<u8>, offset: usize, value: i32| {
            header[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
        };
        put(&mut header, 0, 1);
        put(&mut header, 4, 88);
        for (index, value) in bounds.iter().enumerate() {
            put(&mut header, 8 + index * 4, *value);
        }
        // 100 device units per millimetre keeps `rclFrame` in whole hundredths.
        for (index, value) in [0, 0, frame_device[0] * 100, frame_device[1] * 100]
            .iter()
            .enumerate()
        {
            put(&mut header, 24 + index * 4, *value);
        }
        put(&mut header, 40, EMF_SIGNATURE as i32);
        put(&mut header, 56, 16); // handles, in the low half of the u32 slot
        // One device unit per millimetre makes `rclFrame`, in hundredths of a
        // millimetre, exactly a hundred times the device rectangle.
        put(&mut header, 72, frame_device[0]);
        put(&mut header, 76, frame_device[1]);
        put(&mut header, 80, frame_device[0]);
        put(&mut header, 84, frame_device[1]);
        header
    }

    fn record(kind: u32, body: &[u8]) -> Vec<u8> {
        let size = 8 + body.len();
        let mut out = kind.to_le_bytes().to_vec();
        out.extend_from_slice(&(size as u32).to_le_bytes());
        out.extend_from_slice(body);
        out
    }

    fn i32s(values: &[i32]) -> Vec<u8> {
        values.iter().flat_map(|v| v.to_le_bytes()).collect()
    }

    fn i16s(values: &[i16]) -> Vec<u8> {
        values.iter().flat_map(|v| v.to_le_bytes()).collect()
    }

    /// Bounds rectangle, point count, then 16-bit points.
    fn poly16(points: &[(i16, i16)]) -> Vec<u8> {
        let mut body = i32s(&[0, 0, 0, 0, points.len() as i32]);
        for (x, y) in points {
            body.extend_from_slice(&i16s(&[*x, *y]));
        }
        body
    }

    fn solid_brush(handle: u32, color: u32) -> Vec<u8> {
        let mut out = record(39, &i32s(&[handle as i32, 0, color as i32, 0]));
        out.extend(record(37, &i32s(&[handle as i32])));
        out
    }

    fn emf(records: Vec<Vec<u8>>, bounds: [i32; 4], frame_device: [i32; 2]) -> Vec<u8> {
        let mut bytes = emf_header(bounds, frame_device);
        for chunk in records {
            bytes.extend(chunk);
        }
        bytes.extend(record(14, &i32s(&[0, 0, 0])));
        bytes
    }

    fn only_op(drawing: &MetafileDrawing) -> &MetafileOp {
        assert_eq!(drawing.ops.len(), 1, "expected one op");
        &drawing.ops[0]
    }

    #[test]
    fn a_bracketed_path_fill_becomes_one_op_in_frame_fractions() {
        let bytes = emf(
            vec![
                solid_brush(1, 0x0000_00ff),
                record(59, &[]),
                record(27, &i32s(&[0, 0])),
                record(89, &poly16(&[(50, 0), (0, 50)])),
                record(61, &[]),
                record(60, &[]),
                record(62, &i32s(&[0, 0, 0, 0])),
            ],
            [0, 0, 99, 99],
            [100, 100],
        );

        let drawing = decode(&bytes).expect("the path fill decodes");
        let op = only_op(&drawing);
        assert_eq!(op.fill.as_deref(), Some("#ff0000"));
        assert!(op.stroke.is_none(), "FILLPATH does not stroke");
        assert_eq!(
            op.path,
            vec![
                GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                GeometryPathCommand::Line { x: 0.5, y: 0.0 },
                GeometryPathCommand::Line { x: 0.0, y: 0.5 },
                GeometryPathCommand::Close,
            ]
        );
    }

    #[test]
    fn coordinates_scale_by_the_frame_rectangle_not_the_ink_bounds() {
        // The MSGraph pies declare a frame a fifth wider than their ink;
        // normalising by `rclBounds` would zoom them to the picture's edges.
        let records = vec![
            solid_brush(1, 0x0000_0000),
            record(86, &poly16(&[(20, 20), (60, 20), (60, 60)])),
        ];
        let bytes = emf(records, [20, 20, 60, 60], [80, 80]);

        let drawing = decode(&bytes).expect("the polygon decodes");
        let op = only_op(&drawing);
        assert_eq!(op.path[0], GeometryPathCommand::Move { x: 0.25, y: 0.25 });
        assert_eq!(op.path[1], GeometryPathCommand::Line { x: 0.75, y: 0.25 });
    }

    #[test]
    fn a_window_extent_without_a_viewport_extent_does_not_scale() {
        // `project17/image18.emf` sets a window extent of 9113 and no viewport
        // at all; applying it as a ratio collapses the pie to a point.
        let records = vec![
            record(10, &i32s(&[0, 0])),
            record(9, &i32s(&[1000, 1000])),
            solid_brush(1, 0x0000_0000),
            record(86, &poly16(&[(50, 0), (100, 0), (100, 50)])),
        ];
        let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

        let drawing = decode(&bytes).expect("the polygon decodes");
        assert_eq!(
            only_op(&drawing).path[0],
            GeometryPathCommand::Move { x: 0.5, y: 0.0 }
        );
    }

    #[test]
    fn a_window_and_viewport_extent_pair_scales_logical_units() {
        let records = vec![
            record(9, &i32s(&[200, 200])),
            record(11, &i32s(&[100, 100])),
            solid_brush(1, 0x0000_0000),
            record(86, &poly16(&[(100, 0), (200, 0), (200, 100)])),
        ];
        let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

        let drawing = decode(&bytes).expect("the polygon decodes");
        assert_eq!(
            only_op(&drawing).path[0],
            GeometryPathCommand::Move { x: 0.5, y: 0.0 }
        );
    }

    #[test]
    fn a_pie_sweeps_counter_clockwise_from_its_start_ray() {
        // GDI's default arc direction is counter-clockwise, so a wedge from
        // twelve o'clock to three o'clock is the three-quarter one. Sweeping
        // the short way instead paints every wedge of the MSGraph pies as its
        // own complement.
        let records = vec![
            solid_brush(1, 0x0000_0000),
            record(47, &i32s(&[0, 0, 100, 100, 50, 0, 100, 50])),
        ];
        let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

        let drawing = decode(&bytes).expect("the pie decodes");
        let op = only_op(&drawing);
        assert_eq!(op.path[0], GeometryPathCommand::Move { x: 0.5, y: 0.5 });
        let on_arc: Vec<_> = op
            .path
            .iter()
            .filter_map(|command| match command {
                GeometryPathCommand::Line { x, y } => Some((*x, *y)),
                _ => None,
            })
            .collect();
        assert_eq!(on_arc.first().copied(), Some((0.5, 0.0)));
        let (last_x, last_y) = on_arc.last().copied().expect("the arc has points");
        assert!((last_x - 1.0).abs() < 1e-6 && (last_y - 0.5).abs() < 1e-6);
        assert!(
            on_arc.iter().any(|(x, _)| *x < 0.01),
            "the wedge never reached nine o'clock: {on_arc:?}"
        );
        assert_eq!(op.path.last(), Some(&GeometryPathCommand::Close));
    }

    #[test]
    fn a_null_brush_leaves_a_polygon_unfilled() {
        let mut records = vec![record(39, &i32s(&[1, 1, 0x00ff_00ff, 0]))];
        records.push(record(37, &i32s(&[1])));
        records.push(record(86, &poly16(&[(0, 0), (50, 0), (50, 50)])));
        let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

        let drawing = decode(&bytes).expect("the polygon decodes");
        assert!(only_op(&drawing).fill.is_none());
    }

    #[test]
    fn bytes_that_are_not_a_metafile_are_left_to_the_image_decoder() {
        assert!(decode(b"\xff\xd8\xff\xe0 not a metafile at all").is_none());
        assert!(decode(&[]).is_none());
    }

    #[test]
    fn a_metafile_that_draws_nothing_is_not_decoded() {
        // Several parts in the corpus are 188-byte header-and-EOF stubs. They
        // stay `Image`s so the backend still counts them as skipped.
        let bytes = emf(Vec::new(), [0, 0, 99, 99], [100, 100]);
        assert!(decode(&bytes).is_none());
    }

    #[test]
    fn a_truncated_record_ends_the_replay_without_panicking() {
        let mut bytes = emf(
            vec![
                solid_brush(1, 0x0000_0000),
                record(86, &poly16(&[(0, 0), (50, 0), (50, 50)])),
            ],
            [0, 0, 99, 99],
            [100, 100],
        );
        let full = decode(&bytes).expect("the whole file decodes");
        bytes.truncate(bytes.len() - 12);
        let truncated = decode(&bytes).expect("the leading records still decode");
        assert_eq!(truncated.ops.len(), full.ops.len());

        // A point count that runs past the record must not be trusted.
        let mut lying = emf_header([0, 0, 99, 99], [100, 100]);
        lying.extend(record(86, &i32s(&[0, 0, 0, 0, 100_000])));
        assert!(decode(&lying).is_none());
    }

    #[test]
    fn a_wmf_polygon_fills_with_the_selected_brush() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&WMF_PLACEABLE_KEY.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&i16s(&[0, 0, 100, 100]));
        bytes.extend_from_slice(&1440u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        // META_HEADER: type, header words, version, size, objects, max record, members
        bytes.extend_from_slice(&1u16.to_le_bytes());
        bytes.extend_from_slice(&9u16.to_le_bytes());
        bytes.extend_from_slice(&0x0300u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&2u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());

        let wmf_record = |function: u16, body: &[u8]| {
            let words = (6 + body.len()) / 2;
            let mut out = (words as u32).to_le_bytes().to_vec();
            out.extend_from_slice(&function.to_le_bytes());
            out.extend_from_slice(body);
            out
        };
        // META_SETWINDOWORG and META_SETWINDOWEXT store y before x.
        bytes.extend(wmf_record(0x020B, &i16s(&[0, 0])));
        bytes.extend(wmf_record(0x020C, &i16s(&[100, 100])));
        let mut brush = 0u16.to_le_bytes().to_vec();
        brush.extend_from_slice(&0x0000_00ffu32.to_le_bytes());
        brush.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend(wmf_record(0x02FC, &brush));
        bytes.extend(wmf_record(0x012D, &0u16.to_le_bytes()));
        let mut polygon = 3u16.to_le_bytes().to_vec();
        polygon.extend_from_slice(&i16s(&[0, 0, 50, 0, 50, 50]));
        bytes.extend(wmf_record(0x0324, &polygon));
        bytes.extend(wmf_record(0, &[]));

        let drawing = decode(&bytes).expect("the WMF decodes");
        let op = only_op(&drawing);
        assert_eq!(op.fill.as_deref(), Some("#ff0000"));
        assert_eq!(op.path[1], GeometryPathCommand::Line { x: 0.5, y: 0.0 });
    }
}
