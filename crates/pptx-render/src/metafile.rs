//! Replays vector EMF and WMF records.

use ooxml_drawingml::GeometryPathCommand;

const MAX_RECORDS: usize = 200_000;
const MAX_OPS: usize = 4_096;
const MAX_COMMANDS: usize = 400_000;
const MAX_POINTS_PER_RECORD: usize = 65_536;
const ARC_SEGMENTS: usize = 64;

#[derive(Debug, Clone, PartialEq)]
pub struct MetafileOp {
    pub path: Vec<GeometryPathCommand>,
    pub fill: Option<String>,
    pub stroke: Option<MetafileStroke>,
    pub even_odd: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct MetafileStroke {
    pub color: String,
    pub width: f64,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct MetafileDrawing {
    pub ops: Vec<MetafileOp>,
}

pub fn decode(bytes: &[u8]) -> Option<MetafileDrawing> {
    let drawing = decode_emf(bytes).or_else(|| decode_wmf(bytes))?;
    (!drawing.ops.is_empty()).then_some(drawing)
}

pub fn is_metafile(bytes: &[u8]) -> bool {
    (u32_at(bytes, 0) == Some(1) && u32_at(bytes, 40) == Some(EMF_SIGNATURE))
        || u32_at(bytes, 0) == Some(WMF_PLACEABLE_KEY)
        || (matches!(u16_at(bytes, 0), Some(1 | 2)) && u16_at(bytes, 2) == Some(9))
}

fn u16_at(bytes: &[u8], offset: usize) -> Option<u16> {
    let slice = bytes.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([slice[0], slice[1]]))
}

fn u32_at(bytes: &[u8], offset: usize) -> Option<u32> {
    let slice = bytes.get(offset..offset.checked_add(4)?)?;
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
    scaled: bool,
    window_ext_set: bool,
    viewport_ext_set: bool,
    map_mode: i32,
    even_odd: bool,
    clockwise: bool,
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
            window_ext_set: false,
            viewport_ext_set: false,
            map_mode: 1,
            even_odd: true,
            clockwise: false,
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

struct Player {
    dc: Dc,
    saved: Vec<Dc>,
    objects: Vec<Option<GdiObject>>,
    frame: (f64, f64, f64, f64),
    path: Vec<GeometryPathCommand>,
    selected_path: Vec<GeometryPathCommand>,
    figure_start: (f64, f64),
    current: (f64, f64),
    bracketed: bool,
    pending_stroke: bool,
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
            selected_path: Vec::new(),
            figure_start: (0.0, 0.0),
            current: (0.0, 0.0),
            bracketed: false,
            pending_stroke: false,
            commands: 0,
            ops: Vec::new(),
            overflowed: false,
        }
    }

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

    fn pen_width(&self, width: f64) -> f64 {
        let scale = if self.dc.scaled {
            (self.dc.viewport_ext.0 / self.dc.window_ext.0).abs()
        } else {
            1.0
        };
        (width * scale * self.dc.xform[0].hypot(self.dc.xform[1]) / self.frame.2).abs()
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
        self.figure_start = (x, y);
        let (px, py) = self.point(x, y);
        self.push(GeometryPathCommand::Move { x: px, y: py });
    }

    fn line_to(&mut self, x: f64, y: f64) {
        self.resume_figure();
        self.current = (x, y);
        let (px, py) = self.point(x, y);
        self.push(GeometryPathCommand::Line { x: px, y: py });
    }

    fn cubic_to(&mut self, points: [(f64, f64); 3]) {
        self.resume_figure();
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

    fn resume_figure(&mut self) {
        if self.path.is_empty() || matches!(self.path.last(), Some(GeometryPathCommand::Close)) {
            self.move_to(self.current.0, self.current.1);
        }
    }

    fn close_figure(&mut self) {
        self.push(GeometryPathCommand::Close);
        self.current = self.figure_start;
    }

    fn restore(&mut self, depth: i32) -> Option<()> {
        self.flush_pending();
        let index = if depth < 0 {
            self.saved
                .len()
                .checked_sub(depth.unsigned_abs() as usize)?
        } else {
            (depth as usize).checked_sub(1)?
        };
        self.dc = *self.saved.get(index)?;
        self.saved.truncate(index);
        Some(())
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
        self.ops.push(MetafileOp {
            path,
            fill,
            stroke,
            even_odd: self.dc.even_odd,
        });
    }

    fn flush_pending(&mut self) {
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
        if self.dc.clockwise {
            while sweep <= 0.0 {
                sweep += std::f64::consts::TAU;
            }
        } else {
            while sweep >= 0.0 {
                sweep -= std::f64::consts::TAU;
            }
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

fn brush_from_style(style: u32, color: u32) -> Option<Brush> {
    if !matches!(style, 0 | 1) {
        return None;
    }
    Some(Brush {
        color,
        visible: style != 1,
    })
}

fn pen_from_style(style: u32, width: f64, color: u32) -> Option<Pen> {
    if !matches!(style & 0x0f, 0 | 5) {
        return None;
    }
    Some(Pen {
        color,
        width,
        visible: style & 0x0f != 5,
    })
}

const EMF_SIGNATURE: u32 = 0x464D_4520;

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
    if u32_at(bytes, 0)? != 1 || u32_at(bytes, 40)? != EMF_SIGNATURE || u32_at(bytes, 4)? < 88 {
        return None;
    }
    let handles = u16_at(bytes, 56)? as usize;
    let mut player = Player::new(emf_frame(bytes)?, handles + 1);
    let mut offset = 0usize;
    for _ in 0..MAX_RECORDS {
        let kind = u32_at(bytes, offset)?;
        let size = u32_at(bytes, offset + 4)? as usize;
        if size < 8 || !size.is_multiple_of(4) {
            return None;
        }
        let end = offset.checked_add(size)?;
        let record = bytes.get(offset..end)?;
        if kind == 14 {
            if size < 20 {
                return None;
            }
            player.flush_pending();
            return player.finish();
        }
        emf_record(&mut player, record, kind, 8)?;
        if player.overflowed {
            return None;
        }
        offset = end;
    }
    None
}

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

fn emf_record(player: &mut Player, bytes: &[u8], kind: u32, body: usize) -> Option<()> {
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
        9..=12 => {
            let (x, y) = i32_at(bytes, body).zip(i32_at(bytes, body + 4))?;
            let value = (f64::from(x), f64::from(y));
            if matches!(kind, 9 | 11) && (x == 0 || y == 0) {
                return None;
            }
            match kind {
                9 => {
                    player.dc.window_ext = value;
                    player.dc.window_ext_set = true;
                }
                10 => player.dc.window_org = value,
                11 => {
                    player.dc.viewport_ext = value;
                    player.dc.viewport_ext_set = true;
                }
                _ => player.dc.viewport_org = value,
            }
            player.dc.scaled = player.dc.map_mode == 8
                && player.dc.window_ext_set
                && player.dc.viewport_ext_set
                && player.dc.window_ext.0 != 0.0
                && player.dc.window_ext.1 != 0.0;
        }
        17 => {
            let mode = i32_at(bytes, body)?;
            if !matches!(mode, 1 | 8) {
                return None;
            }
            player.dc.map_mode = mode;
            player.dc.scaled = mode == 8 && player.dc.window_ext_set && player.dc.viewport_ext_set;
        }
        19 => {
            player.dc.even_odd = match u32_at(bytes, body)? {
                1 => true,
                2 => false,
                _ => return None,
            };
        }
        57 => {
            player.dc.clockwise = match u32_at(bytes, body)? {
                1 => false,
                2 => true,
                _ => return None,
            };
        }
        33 => {
            if player.saved.len() >= 1024 {
                return None;
            }
            player.saved.push(player.dc);
        }
        34 => player.restore(i32_at(bytes, body)?)?,
        35 | 36 => {
            let mut matrix = IDENTITY;
            for (index, slot) in matrix.iter_mut().enumerate() {
                match f32_at(bytes, body + index * 4) {
                    Some(value) if value.is_finite() => *slot = f64::from(value),
                    _ => return None,
                }
            }
            let mode = if kind == 36 {
                u32_at(bytes, body + 24)?
            } else {
                4
            };
            player.dc.xform = match mode {
                1 => IDENTITY,
                2 => concat(matrix, player.dc.xform),
                3 => concat(player.dc.xform, matrix),
                4 => matrix,
                _ => return None,
            };
        }
        37 => {
            let handle = u32_at(bytes, body)?;
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
            let (Some(handle), Some(style), Some(width), Some(color)) = (
                u32_at(bytes, body),
                u32_at(bytes, body + 4),
                i32_at(bytes, body + 8),
                u32_at(bytes, body + 16),
            ) else {
                return None;
            };
            player.store(
                handle as usize,
                GdiObject::Pen(pen_from_style(style, f64::from(width), color)?),
            );
        }
        95 => {
            let (Some(handle), Some(style), Some(width), Some(color)) = (
                u32_at(bytes, body),
                u32_at(bytes, body + 20),
                u32_at(bytes, body + 24),
                u32_at(bytes, body + 32),
            ) else {
                return None;
            };
            player.store(
                handle as usize,
                GdiObject::Pen(pen_from_style(style, f64::from(width), color)?),
            );
        }
        39 => {
            let (Some(handle), Some(style), Some(color)) = (
                u32_at(bytes, body),
                u32_at(bytes, body + 4),
                u32_at(bytes, body + 8),
            ) else {
                return None;
            };
            player.store(
                handle as usize,
                GdiObject::Brush(brush_from_style(style, color)?),
            );
        }
        40 => {
            if let Some(handle) = u32_at(bytes, body)
                && let Some(slot) = player.objects.get_mut(handle as usize)
            {
                *slot = None;
            }
        }
        59 => {
            player.flush_pending();
            player.path.clear();
            player.selected_path.clear();
            player.bracketed = true;
        }
        60 => {
            player.bracketed = false;
            player.selected_path = std::mem::take(&mut player.path);
        }
        61 => player.close_figure(),
        62 => {
            player.flush_pending();
            player.path = std::mem::take(&mut player.selected_path);
            let fill = player.brush_fill();
            player.emit(fill, None);
        }
        63 => {
            player.flush_pending();
            player.path = std::mem::take(&mut player.selected_path);
            let (fill, stroke) = (player.brush_fill(), player.pen_stroke());
            player.emit(fill, stroke);
        }
        64 => {
            player.flush_pending();
            player.path = std::mem::take(&mut player.selected_path);
            let stroke = player.pen_stroke();
            player.emit(None, stroke);
        }
        68 => {
            player.path.clear();
            player.selected_path.clear();
            player.bracketed = false;
            player.pending_stroke = false;
        }
        27 => {
            let (x, y) = i32_at(bytes, body).zip(i32_at(bytes, body + 4))?;
            if !player.bracketed {
                player.flush_pending();
            }
            player.move_to(f64::from(x), f64::from(y));
        }
        54 => {
            let (x, y) = i32_at(bytes, body).zip(i32_at(bytes, body + 4))?;
            player.line_to(f64::from(x), f64::from(y));
            player.pending_stroke = !player.bracketed;
        }
        5 | 6 | 88 | 89 => {
            let small = kind >= 88;
            let points = poly_points(small)?;
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
        2 | 3 | 4 | 85 | 86 | 87 => {
            let small = kind >= 85;
            let points = poly_points(small)?;
            let first = points.first().copied()?;
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
        7 | 8 | 90 | 91 => {
            let small = kind >= 90;
            let (Some(polygons), Some(total)) =
                (u32_at(bytes, body + 16), u32_at(bytes, body + 20))
            else {
                return None;
            };
            let (polygons, total) = (polygons as usize, total as usize);
            if polygons > MAX_POINTS_PER_RECORD || total > MAX_POINTS_PER_RECORD {
                return None;
            }
            let mut counts = Vec::with_capacity(polygons);
            for index in 0..polygons {
                let count = u32_at(bytes, body + 24 + index * 4)?;
                counts.push(count as usize);
            }
            if counts
                .iter()
                .try_fold(0usize, |sum, count| sum.checked_add(*count))
                != Some(total)
            {
                return None;
            }
            let points = read_points(bytes, body + 24 + polygons * 4, total, small)?;
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
            let rect = box_rect()?;
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
            let (Some(rect), Some(sx), Some(sy), Some(ex), Some(ey)) = (
                box_rect(),
                i32_at(bytes, body + 16),
                i32_at(bytes, body + 20),
                i32_at(bytes, body + 24),
                i32_at(bytes, body + 28),
            ) else {
                return None;
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
        20 if u32_at(bytes, body)? == 13 => {}
        1 | 13 | 16 | 18 | 21 | 22 | 24 | 25 | 69 | 70 | 98 => {}
        _ => return None,
    }
    Some(())
}

const WMF_PLACEABLE_KEY: u32 = 0x9AC6_CDD7;

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
    let mut eof = false;
    let mut deferred: Vec<(usize, usize, usize)> = Vec::new();
    while offset + 6 <= bytes.len() {
        let size = (u32_at(bytes, offset)? as usize).checked_mul(2)?;
        let function = u16_at(bytes, offset + 4)?;
        let end = offset.checked_add(size)?;
        if size < 6 || end > bytes.len() {
            return None;
        }
        records += 1;
        if records > MAX_RECORDS {
            return None;
        }
        if function == 0 {
            eof = true;
            break;
        }
        deferred.push((function as usize, offset, end));
        offset = end;
    }
    if !eof {
        return None;
    }
    let mut org = None;
    let mut ext = None;
    for (function, start, end) in &deferred {
        let bytes = &bytes[*start..*end];
        let body = 6;
        match function {
            0x020B => org = org.or(i16_at(bytes, body + 2).zip(i16_at(bytes, body))),
            0x020C => ext = ext.or(i16_at(bytes, body + 2).zip(i16_at(bytes, body))),
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
    player.dc.window_org = (player.frame.0, player.frame.1);
    player.dc.window_ext = (player.frame.2, player.frame.3);
    player.dc.viewport_ext = player.dc.window_ext;
    player.dc.scaled = true;
    player.frame.0 = 0.0;
    player.frame.1 = 0.0;
    for (function, start, end) in deferred {
        wmf_record(&mut player, &bytes[start..end], function, 6)?;
        if player.overflowed {
            return None;
        }
    }
    player.flush_pending();
    player.finish()
}

fn wmf_record(player: &mut Player, bytes: &[u8], function: usize, body: usize) -> Option<()> {
    let point_at = |offset: usize| -> Option<(f64, f64)> {
        Some((
            f64::from(i16_at(bytes, offset + 2)?),
            f64::from(i16_at(bytes, offset)?),
        ))
    };
    match function {
        0x020B => player.dc.window_org = point_at(body)?,
        0x020C => {
            let ext = point_at(body)?;
            if ext.0 == 0.0 || ext.1 == 0.0 {
                return None;
            }
            player.dc.window_ext = ext;
        }
        0x020D => player.dc.viewport_org = point_at(body)?,
        0x020E => player.dc.viewport_ext = point_at(body)?,
        0x0106 => {
            player.dc.even_odd = match u16_at(bytes, body)? {
                1 => true,
                2 => false,
                _ => return None,
            }
        }
        0x001E => {
            if player.saved.len() >= 1024 {
                return None;
            }
            player.saved.push(player.dc);
        }
        0x0127 => player.restore(i32::from(i16_at(bytes, body)?))?,
        0x0214 => {
            if let Some(point) = point_at(body) {
                player.flush_pending();
                player.move_to(point.0, point.1);
            }
        }
        0x0213 => {
            if let Some(point) = point_at(body) {
                player.line_to(point.0, point.1);
                player.pending_stroke = true;
            }
        }
        0x02FA => {
            let (Some(style), Some(width), Some(color)) = (
                u16_at(bytes, body),
                i16_at(bytes, body + 2),
                u32_at(bytes, body + 6),
            ) else {
                return None;
            };
            let pen = pen_from_style(u32::from(style), f64::from(width), color)?;
            wmf_store(player, GdiObject::Pen(pen));
        }
        0x02FC => {
            let (Some(style), Some(color)) = (u16_at(bytes, body), u32_at(bytes, body + 2)) else {
                return None;
            };
            let brush = brush_from_style(u32::from(style), color)?;
            wmf_store(player, GdiObject::Brush(brush));
        }
        0x012D => {
            let index = u16_at(bytes, body)?;
            player.flush_pending();
            if let Some(object) = player.objects.get(index as usize).copied().flatten() {
                player.select(object);
            }
        }
        0x01F0 => {
            if let Some(index) = u16_at(bytes, body)
                && let Some(slot) = player.objects.get_mut(index as usize)
            {
                *slot = None;
            }
        }
        0x0324 | 0x0325 => {
            let count = u16_at(bytes, body)?;
            let points = read_points(bytes, body + 2, count as usize, true)?;
            let first = points.first().copied()?;
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
            let polygons = u16_at(bytes, body)?;
            let polygons = polygons as usize;
            if polygons > MAX_POINTS_PER_RECORD {
                return None;
            }
            let mut counts = Vec::with_capacity(polygons);
            let mut total = 0usize;
            for index in 0..polygons {
                let count = u16_at(bytes, body + 2 + index * 2)?;
                total += count as usize;
                counts.push(count as usize);
            }
            let points = read_points(bytes, body + 2 + polygons * 2, total, true)?;
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
            let (Some(bottom), Some(right), Some(top), Some(left)) = (
                i16_at(bytes, body),
                i16_at(bytes, body + 2),
                i16_at(bytes, body + 4),
                i16_at(bytes, body + 6),
            ) else {
                return None;
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
        0x0103 if u16_at(bytes, body)? == 8 => {}
        0x0102 | 0x0107 | 0x0108 | 0x0201 | 0x0209 | 0x020A => {}
        _ => return None,
    }
    Some(())
}

#[cfg(test)]
mod tests {
    use super::*;

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
        for (index, value) in [0, 0, frame_device[0] * 100, frame_device[1] * 100]
            .iter()
            .enumerate()
        {
            put(&mut header, 24 + index * 4, *value);
        }
        put(&mut header, 40, EMF_SIGNATURE as i32);
        put(&mut header, 56, 16);
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
            record(17, &i32s(&[8])),
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
        assert!(decode(&bytes).is_some());
        bytes.truncate(bytes.len() - 12);
        assert!(decode(&bytes).is_none());

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
    #[test]
    fn records_cannot_read_points_from_their_successors() {
        let bytes = emf(
            vec![
                record(86, &i32s(&[0, 0, 99, 99, 3])),
                record(70, &i32s(&[12, 0, 0, 0])),
            ],
            [0, 0, 99, 99],
            [100, 100],
        );
        assert!(decode(&bytes).is_none());
        let mut bytes = vec![0u8; 18];
        bytes[0..2].copy_from_slice(&1u16.to_le_bytes());
        bytes[2..4].copy_from_slice(&9u16.to_le_bytes());
        bytes.extend(4u32.to_le_bytes());
        bytes.extend(0x0324u16.to_le_bytes());
        bytes.extend(3u16.to_le_bytes());
        bytes.extend(9u32.to_le_bytes());
        bytes.extend(0x0201u16.to_le_bytes());
        bytes.extend([0u8; 12]);
        bytes.extend(3u32.to_le_bytes());
        bytes.extend(0u16.to_le_bytes());
        assert!(decode(&bytes).is_none());
    }

    #[test]
    fn selecting_a_brush_preserves_open_and_closed_paths() {
        for before_end in [false, true] {
            let mut records = vec![
                record(59, &[]),
                record(27, &i32s(&[0, 0])),
                record(89, &poly16(&[(100, 0), (0, 100)])),
                record(61, &[]),
            ];
            if before_end {
                records.push(solid_brush(1, 0x0000ff00));
            }
            records.push(record(60, &[]));
            if !before_end {
                records.push(solid_brush(1, 0x0000ff00));
            }
            records.push(record(62, &i32s(&[0, 0, 100, 100])));
            let drawing = decode(&emf(records, [0, 0, 99, 99], [100, 100])).unwrap();
            assert_eq!(only_op(&drawing).fill.as_deref(), Some("#00ff00"));
            assert_eq!(only_op(&drawing).path.len(), 4);
        }
    }

    #[test]
    fn restore_dc_uses_absolute_levels_and_rejects_extreme_depths() {
        let bytes = emf(
            vec![
                record(33, &[]),
                record(10, &i32s(&[20, 20])),
                record(33, &[]),
                record(10, &i32s(&[40, 40])),
                record(34, &i32s(&[1])),
                record(86, &poly16(&[(0, 0), (50, 0), (50, 50)])),
            ],
            [0, 0, 99, 99],
            [100, 100],
        );
        let drawing = decode(&bytes).unwrap();
        assert_eq!(
            only_op(&drawing).path[0],
            GeometryPathCommand::Move { x: 0.0, y: 0.0 }
        );
        let bytes = emf(
            vec![record(33, &[]), record(34, &i32s(&[i32::MIN]))],
            [0, 0, 99, 99],
            [100, 100],
        );
        assert!(decode(&bytes).is_none());
    }

    #[test]
    fn world_transforms_keep_rotated_pen_widths_and_reject_nonfinite_values() {
        let matrix = |values: &[f32]| {
            values
                .iter()
                .flat_map(|v| v.to_le_bytes())
                .collect::<Vec<_>>()
        };
        let bytes = emf(
            vec![
                record(38, &i32s(&[1, 0, 10, 0, 0])),
                record(37, &i32s(&[1])),
                record(35, &matrix(&[0.0, 1.0, -1.0, 0.0, 100.0, 0.0])),
                record(43, &i32s(&[0, 0, 100, 100])),
            ],
            [0, 0, 99, 99],
            [100, 100],
        );
        let drawing = decode(&bytes).unwrap();
        assert_eq!(only_op(&drawing).stroke.as_ref().unwrap().width, 0.1);
        for invalid in [f32::NAN, f32::INFINITY, f32::NEG_INFINITY] {
            let bytes = emf(
                vec![
                    record(35, &matrix(&[invalid, 0.0, 0.0, 1.0, 0.0, 0.0])),
                    record(43, &i32s(&[0, 0, 100, 100])),
                ],
                [0, 0, 99, 99],
                [100, 100],
            );
            assert!(decode(&bytes).is_none());
        }
    }

    #[test]
    fn fill_mode_and_arc_direction_follow_the_device_context() {
        for mode in [1, 2] {
            let bytes = emf(
                vec![
                    record(19, &i32s(&[mode])),
                    record(57, &i32s(&[2])),
                    record(47, &i32s(&[0, 0, 100, 100, 50, 0, 100, 50])),
                ],
                [0, 0, 99, 99],
                [100, 100],
            );
            let drawing = decode(&bytes).unwrap();
            let op = only_op(&drawing);
            assert_eq!(op.even_odd, mode == 1);
            assert!(
                op.path
                    .iter()
                    .all(|p| !matches!(p, GeometryPathCommand::Line { x, .. } if *x < 0.49))
            );
        }
    }

    #[test]
    fn unsupported_drawing_records_do_not_produce_partial_artwork() {
        for kind in [30, 76, 84, 118] {
            let bytes = emf(
                vec![
                    record(43, &i32s(&[0, 0, 100, 100])),
                    record(kind, &[0; 100]),
                ],
                [0, 0, 99, 99],
                [100, 100],
            );
            assert!(decode(&bytes).is_none(), "record {kind}");
        }
        for mode in 1..=8u16 {
            let mut bytes = vec![0u8; 18];
            bytes[0..2].copy_from_slice(&1u16.to_le_bytes());
            bytes[2..4].copy_from_slice(&9u16.to_le_bytes());
            bytes.extend(4u32.to_le_bytes());
            bytes.extend(0x0103u16.to_le_bytes());
            bytes.extend(mode.to_le_bytes());
            bytes.extend(7u32.to_le_bytes());
            bytes.extend(0x041Bu16.to_le_bytes());
            for value in [1u16, 1, 0, 0] {
                bytes.extend(value.to_le_bytes());
            }
            bytes.extend(3u32.to_le_bytes());
            bytes.extend(0u16.to_le_bytes());
            assert_eq!(decode(&bytes).is_some(), mode == 8, "WMF mapping {mode}");
        }
    }

    #[test]
    fn replay_budgets_reject_excessive_records_operations_and_points() {
        let bytes = emf(
            vec![record(43, &i32s(&[0, 0, 100, 100])); 4097],
            [0, 0, 99, 99],
            [100, 100],
        );
        assert!(decode(&bytes).is_none());
        let points = vec![(0, 0); 65536];
        let bytes = emf(
            vec![record(86, &poly16(&points)); 7],
            [0, 0, 99, 99],
            [100, 100],
        );
        assert!(decode(&bytes).is_none());
        let mut records = vec![record(18, &i32s(&[1])); 200000];
        records.push(record(43, &i32s(&[0, 0, 100, 100])));
        let bytes = emf(records, [0, 0, 99, 99], [100, 100]);
        assert!(decode(&bytes).is_none());
    }
}
