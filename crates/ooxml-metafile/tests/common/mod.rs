//! Builders for synthetic EMF, EMF+ and WMF metafiles.
#![allow(dead_code)]

pub mod fixtures;

pub fn i32s(values: &[i32]) -> Vec<u8> {
    values.iter().flat_map(|v| v.to_le_bytes()).collect()
}

pub fn u32s(values: &[u32]) -> Vec<u8> {
    values.iter().flat_map(|v| v.to_le_bytes()).collect()
}

pub fn i16s(values: &[i16]) -> Vec<u8> {
    values.iter().flat_map(|v| v.to_le_bytes()).collect()
}

pub fn u16s(values: &[u16]) -> Vec<u8> {
    values.iter().flat_map(|v| v.to_le_bytes()).collect()
}

pub fn f32s(values: &[f32]) -> Vec<u8> {
    values.iter().flat_map(|v| v.to_le_bytes()).collect()
}

pub fn utf16(text: &str) -> Vec<u16> {
    text.encode_utf16().collect()
}

fn pad4(bytes: &mut Vec<u8>) {
    while !bytes.len().is_multiple_of(4) {
        bytes.push(0);
    }
}

/// `0x00BBGGRR` from `0xRRGGBB`.
pub const fn rgb(hex: u32) -> u32 {
    ((hex & 0xff) << 16) | (hex & 0xff00) | ((hex >> 16) & 0xff)
}

/// An EMF on a 96 DPI reference device, `width` by `height` device pixels.
pub struct Emf {
    width: i32,
    height: i32,
    body: Vec<u8>,
    records: u32,
    handles: u16,
}

impl Emf {
    pub fn new(width: i32, height: i32) -> Self {
        Self {
            width,
            height,
            body: Vec::new(),
            records: 0,
            handles: 16,
        }
    }

    pub fn rec(mut self, kind: u32, body: &[u8]) -> Self {
        let mut body = body.to_vec();
        pad4(&mut body);
        self.body.extend_from_slice(&kind.to_le_bytes());
        self.body
            .extend_from_slice(&((body.len() + 8) as u32).to_le_bytes());
        self.body.extend_from_slice(&body);
        self.records += 1;
        self
    }

    pub fn recs(self, records: Vec<(u32, Vec<u8>)>) -> Self {
        records
            .into_iter()
            .fold(self, |emf, (kind, body)| emf.rec(kind, &body))
    }

    pub fn bytes(self) -> Vec<u8> {
        let frame = |px: i32| (f64::from(px) * 2540.0 / 96.0).round() as i32;
        let mut header = i32s(&[
            0,
            0,
            self.width - 1,
            self.height - 1,
            0,
            0,
            frame(self.width),
            frame(self.height),
        ]);
        header.extend(u32s(&[0x464D_4520, 0x0001_0000, 0, self.records + 2]));
        header.extend(u16s(&[self.handles, 0]));
        header.extend(u32s(&[0, 0, 0]));
        header.extend(i32s(&[9600, 9600, 2540, 2540]));
        let mut out = u32s(&[1, (header.len() + 8) as u32]);
        out.extend(header);
        out.extend(self.body);
        out.extend(u32s(&[14, 20, 0, 16, 20]));
        let size = out.len() as u32;
        out[48..52].copy_from_slice(&size.to_le_bytes());
        out
    }
}

pub fn rect(l: i32, t: i32, r: i32, b: i32) -> (u32, Vec<u8>) {
    (43, i32s(&[l, t, r, b]))
}

pub fn ellipse(l: i32, t: i32, r: i32, b: i32) -> (u32, Vec<u8>) {
    (42, i32s(&[l, t, r, b]))
}

pub fn round_rect(l: i32, t: i32, r: i32, b: i32, w: i32, h: i32) -> (u32, Vec<u8>) {
    (44, i32s(&[l, t, r, b, w, h]))
}

pub fn arc(kind: u32, rect: [i32; 4], start: (i32, i32), end: (i32, i32)) -> (u32, Vec<u8>) {
    (
        kind,
        i32s(&[
            rect[0], rect[1], rect[2], rect[3], start.0, start.1, end.0, end.1,
        ]),
    )
}

/// `POLYGON16` (86), `POLYLINE16` (87), `POLYBEZIERTO16` (88) and friends.
pub fn poly16(kind: u32, points: &[(i16, i16)]) -> (u32, Vec<u8>) {
    let mut body = i32s(&[0, 0, 0, 0, points.len() as i32]);
    for (x, y) in points {
        body.extend(i16s(&[*x, *y]));
    }
    (kind, body)
}

pub fn move_to(x: i32, y: i32) -> (u32, Vec<u8>) {
    (27, i32s(&[x, y]))
}

pub fn line_to(x: i32, y: i32) -> (u32, Vec<u8>) {
    (54, i32s(&[x, y]))
}

pub fn bare(kind: u32) -> (u32, Vec<u8>) {
    (kind, Vec::new())
}

pub fn bounds_only(kind: u32) -> (u32, Vec<u8>) {
    (kind, i32s(&[0, 0, 0, 0]))
}

pub fn value(kind: u32, value: u32) -> (u32, Vec<u8>) {
    (kind, u32s(&[value]))
}

pub fn select(handle: u32) -> (u32, Vec<u8>) {
    (37, u32s(&[handle]))
}

pub fn stock(index: u32) -> (u32, Vec<u8>) {
    (37, u32s(&[0x8000_0000 | index]))
}

pub fn delete(handle: u32) -> (u32, Vec<u8>) {
    (40, u32s(&[handle]))
}

/// `EXTCREATEPEN`: a geometric pen `width` logical units wide.
pub fn pen(handle: u32, style: u32, width: u32, color: u32) -> (u32, Vec<u8>) {
    let mut body = u32s(&[handle, 0, 0, 0, 0]);
    body.extend(u32s(&[0x0001_0000 | style, width, 0, color, 0, 0]));
    (95, body)
}

pub fn user_pen(handle: u32, width: u32, color: u32, dashes: &[u32]) -> (u32, Vec<u8>) {
    let mut body = u32s(&[handle, 0, 0, 0, 0]);
    body.extend(u32s(&[
        0x0001_0007,
        width,
        0,
        color,
        0,
        dashes.len() as u32,
    ]));
    body.extend(u32s(dashes));
    (95, body)
}

pub fn brush(handle: u32, style: u32, color: u32, hatch: u32) -> (u32, Vec<u8>) {
    (39, u32s(&[handle, style, color, hatch]))
}

pub fn world(m: [f32; 6]) -> (u32, Vec<u8>) {
    (35, f32s(&m))
}

pub fn intersect_clip(l: i32, t: i32, r: i32, b: i32) -> (u32, Vec<u8>) {
    (30, i32s(&[l, t, r, b]))
}

pub fn exclude_clip(l: i32, t: i32, r: i32, b: i32) -> (u32, Vec<u8>) {
    (29, i32s(&[l, t, r, b]))
}

/// `EXTSELECTCLIPRGN` over device rectangles; no rectangles resets the clip.
pub fn select_clip_region(mode: u32, rects: &[[i32; 4]]) -> (u32, Vec<u8>) {
    if rects.is_empty() {
        return (75, u32s(&[0, mode]));
    }
    let mut data = u32s(&[32, 1, rects.len() as u32, (rects.len() * 16) as u32]);
    data.extend(i32s(&[0, 0, 0, 0]));
    for rect in rects {
        data.extend(i32s(rect));
    }
    let mut body = u32s(&[data.len() as u32, mode]);
    body.extend(data);
    (75, body)
}

/// `EXTCREATEFONTINDIRECTW` with a `LOGFONTW`.
pub fn font(
    handle: u32,
    height: i32,
    weight: i32,
    flags: [u8; 3],
    escapement: i32,
    charset: u8,
    face: &str,
) -> (u32, Vec<u8>) {
    let mut body = u32s(&[handle]);
    body.extend(i32s(&[height, 0, escapement, escapement, weight]));
    body.extend([flags[0], flags[1], flags[2], charset, 0, 0, 0, 0]);
    let mut name = utf16(face);
    name.resize(32, 0);
    body.extend(u16s(&name));
    (82, body)
}

/// `EXTTEXTOUTW` at `(x, y)`, with per-character advances when given.
pub fn text_out(
    x: i32,
    y: i32,
    text: &str,
    dx: Option<&[i32]>,
    options: u32,
    rect: [i32; 4],
) -> (u32, Vec<u8>) {
    let units = utf16(text);
    let string_at = 8 + 28 + 40;
    let mut string = u16s(&units);
    pad4(&mut string);
    let dx_at = if dx.is_some() {
        string_at + string.len()
    } else {
        0
    };
    let mut body = i32s(&[0, 0, -1, -1]);
    body.extend(u32s(&[1]));
    body.extend(f32s(&[1.0, 1.0]));
    body.extend(i32s(&[x, y]));
    body.extend(u32s(&[units.len() as u32, string_at as u32, options]));
    body.extend(i32s(&rect));
    body.extend(u32s(&[dx_at as u32]));
    body.extend(string);
    if let Some(dx) = dx {
        body.extend(i32s(dx));
    }
    (84, body)
}

/// A bottom-up `BITMAPINFOHEADER` DIB: `rows` from the top, one `COLORREF`
/// per pixel, stored as 24-bit BGR.
pub fn dib24(rows: &[&[u32]]) -> (Vec<u8>, Vec<u8>) {
    let (width, height) = (rows[0].len(), rows.len());
    let mut bmi = u32s(&[40, width as u32, height as u32]);
    bmi.extend(u16s(&[1, 24]));
    bmi.extend(u32s(&[0, 0, 0, 0, 0, 0]));
    let stride = (width * 3).div_ceil(4) * 4;
    let mut bits = Vec::with_capacity(stride * height);
    for row in rows.iter().rev() {
        let start = bits.len();
        for color in *row {
            bits.extend([(*color >> 16) as u8, (*color >> 8) as u8, *color as u8]);
        }
        bits.resize(start + stride, 0);
    }
    (bmi, bits)
}

/// A bottom-up 1-bit DIB over a two-entry colour table.
pub fn dib1(rows: &[&[u8]], palette: [u32; 2]) -> (Vec<u8>, Vec<u8>) {
    let (width, height) = (rows[0].len(), rows.len());
    let mut bmi = u32s(&[40, width as u32, height as u32]);
    bmi.extend(u16s(&[1, 1]));
    bmi.extend(u32s(&[0, 0, 0, 0, 2, 0]));
    for color in palette {
        bmi.extend([(color >> 16) as u8, (color >> 8) as u8, color as u8, 0]);
    }
    let stride = width.div_ceil(32) * 4;
    let mut bits = Vec::new();
    for row in rows.iter().rev() {
        let mut line = vec![0u8; stride];
        for (x, bit) in row.iter().enumerate() {
            if *bit != 0 {
                line[x / 8] |= 0x80 >> (x % 8);
            }
        }
        bits.extend(line);
    }
    (bmi, bits)
}

/// `STRETCHDIBITS` of a whole DIB into `dest` with `rop`.
pub fn stretch_dibits(dest: [i32; 4], dib: &(Vec<u8>, Vec<u8>), rop: u32) -> (u32, Vec<u8>) {
    let (bmi, bits) = dib;
    let width = i32::from_le_bytes(bmi[4..8].try_into().unwrap());
    let height = i32::from_le_bytes(bmi[8..12].try_into().unwrap()).abs();
    let bmi_at = 8 + 72;
    let mut body = i32s(&[0, 0, 0, 0, dest[0], dest[1], 0, 0, width, height]);
    body.extend(u32s(&[
        bmi_at,
        bmi.len() as u32,
        bmi_at + bmi.len() as u32,
        bits.len() as u32,
        0,
        rop,
    ]));
    body.extend(i32s(&[dest[2], dest[3]]));
    body.extend(bmi);
    body.extend(bits);
    (81, body)
}

pub fn cropped_rle_emf(side: u32) -> Vec<u8> {
    let mut bmi = u32s(&[40, side, side]);
    bmi.extend(u16s(&[1, 8]));
    bmi.extend(u32s(&[1, 4, 0, 0, 2, 0, 0, 0x00ff_ffff]));
    let (kind, mut body) = stretch_dibits([0, 0, 10, 10], &(bmi, vec![1, 1, 0, 1]), 0x00CC_0020);
    body[32..40].copy_from_slice(&i32s(&[1, 1]));
    Emf::new(10, 10).rec(kind, &body).bytes()
}

/// `ALPHABLEND` (114) or `TRANSPARENTBLT` (116) of a whole DIB into `dest`.
pub fn blend(
    kind: u32,
    dest: [i32; 4],
    dib: &(Vec<u8>, Vec<u8>),
    operation: u32,
) -> (u32, Vec<u8>) {
    let (bmi, bits) = dib;
    let width = i32::from_le_bytes(bmi[4..8].try_into().unwrap());
    let height = i32::from_le_bytes(bmi[8..12].try_into().unwrap()).abs();
    let bmi_at = 8 + 100;
    let mut body = i32s(&[0, 0, 0, 0, dest[0], dest[1], dest[2], dest[3]]);
    body.extend(u32s(&[operation]));
    body.extend(i32s(&[0, 0]));
    body.extend(f32s(&[1.0, 0.0, 0.0, 1.0, 0.0, 0.0]));
    body.extend(u32s(&[
        0,
        0,
        bmi_at,
        bmi.len() as u32,
        bmi_at + bmi.len() as u32,
        bits.len() as u32,
    ]));
    body.extend(i32s(&[width, height]));
    body.extend(bmi);
    body.extend(bits);
    (kind, body)
}

/// An EMF comment carrying EMF+ records `(type, flags, data)`.
pub fn plus(records: &[(u16, u16, Vec<u8>)]) -> (u32, Vec<u8>) {
    let mut data = u32s(&[0x2B46_4D45]);
    for (kind, flags, body) in records {
        let mut body = body.clone();
        pad4(&mut body);
        data.extend(u16s(&[*kind, *flags]));
        data.extend(u32s(&[(body.len() + 12) as u32, body.len() as u32]));
        data.extend(body);
    }
    let mut out = u32s(&[data.len() as u32]);
    out.extend(data);
    (70, out)
}

pub fn plus_header(dual: bool) -> (u16, u16, Vec<u8>) {
    (0x4001, u16::from(dual), u32s(&[0xDBC0_1002, 1, 96, 96]))
}

pub fn plus_solid_brush(id: u8, argb: u32) -> (u16, u16, Vec<u8>) {
    (
        0x4008,
        0x0100 | u16::from(id),
        u32s(&[0xDBC0_1002, 0, argb]),
    )
}

/// A pen `width` world units wide with an optional dash pattern.
pub fn plus_pen(id: u8, width: f32, argb: u32, dash: &[f32]) -> (u16, u16, Vec<u8>) {
    let flags = if dash.is_empty() { 0 } else { 0x0100 };
    let mut body = u32s(&[0xDBC0_1002, 0, flags, 0]);
    body.extend(f32s(&[width]));
    if !dash.is_empty() {
        body.extend(u32s(&[dash.len() as u32]));
        body.extend(f32s(dash));
    }
    body.extend(u32s(&[0xDBC0_1002, 0, argb]));
    (0x4008, 0x0200 | u16::from(id), body)
}

/// A path of one figure: `points` joined by lines, closed.
pub fn plus_path(id: u8, points: &[(f32, f32)]) -> (u16, u16, Vec<u8>) {
    let mut body = u32s(&[0xDBC0_1002, points.len() as u32, 0]);
    for (x, y) in points {
        body.extend(f32s(&[*x, *y]));
    }
    for index in 0..points.len() {
        body.push(match index {
            0 => 0,
            last if last == points.len() - 1 => 0x81,
            _ => 1,
        });
    }
    pad4(&mut body);
    (0x4008, 0x0300 | u16::from(id), body)
}

pub fn plus_font(id: u8, em: f32, style: i32, family: &str) -> (u16, u16, Vec<u8>) {
    let name = utf16(family);
    let mut body = u32s(&[0xDBC0_1002]);
    body.extend(f32s(&[em]));
    body.extend(u32s(&[2]));
    body.extend(i32s(&[style]));
    body.extend(u32s(&[0, name.len() as u32]));
    body.extend(u16s(&name));
    (0x4008, 0x0600 | u16::from(id), body)
}

/// A 32bpp ARGB bitmap image object, `rows` from the top.
pub fn plus_bitmap(id: u8, rows: &[&[u32]]) -> (u16, u16, Vec<u8>) {
    let (width, height) = (rows[0].len(), rows.len());
    let mut body = u32s(&[0xDBC0_1002, 1]);
    body.extend(i32s(&[width as i32, height as i32, (width * 4) as i32]));
    body.extend(u32s(&[0x0026_200A, 0]));
    for row in rows {
        for argb in *row {
            body.extend(argb.to_le_bytes());
        }
    }
    (0x4008, 0x0500 | u16::from(id), body)
}

pub fn plus_fill_rects(argb: u32, rects: &[[f32; 4]]) -> (u16, u16, Vec<u8>) {
    let mut body = u32s(&[argb, rects.len() as u32]);
    for rect in rects {
        body.extend(f32s(rect));
    }
    (0x400A, 0x8000, body)
}

pub fn plus_draw_lines(pen: u8, points: &[(f32, f32)], closed: bool) -> (u16, u16, Vec<u8>) {
    let mut body = u32s(&[points.len() as u32]);
    for (x, y) in points {
        body.extend(f32s(&[*x, *y]));
    }
    (
        0x400D,
        u16::from(pen) | if closed { 0x2000 } else { 0 },
        body,
    )
}

pub fn plus_fill_ellipse(argb: u32, rect: [f32; 4]) -> (u16, u16, Vec<u8>) {
    let mut body = u32s(&[argb]);
    body.extend(f32s(&rect));
    (0x400E, 0x8000, body)
}

pub fn plus_fill_path(path: u8, brush: u8) -> (u16, u16, Vec<u8>) {
    (0x4014, u16::from(path), u32s(&[u32::from(brush)]))
}

pub fn plus_draw_path(path: u8, pen: u8) -> (u16, u16, Vec<u8>) {
    (0x4015, u16::from(path), u32s(&[u32::from(pen)]))
}

pub fn plus_linear_brush(id: u8, rect: [f32; 4], from: u32, to: u32) -> (u16, u16, Vec<u8>) {
    let mut body = u32s(&[0xDBC0_1002, 4, 0]);
    body.extend(i32s(&[4]));
    body.extend(f32s(&rect));
    body.extend(u32s(&[from, to, 0, 0]));
    (0x4008, 0x0100 | u16::from(id), body)
}

/// `DrawDriverString` of `text` with one position per character.
pub fn plus_driver_string(
    font: u8,
    argb: u32,
    text: &str,
    origin: (f32, f32),
    advances: &[f32],
) -> (u16, u16, Vec<u8>) {
    let units = utf16(text);
    let mut body = u32s(&[argb, 1, 0, units.len() as u32]);
    body.extend(u16s(&units));
    let mut x = origin.0;
    for index in 0..units.len() {
        body.extend(f32s(&[x, origin.1]));
        x += advances.get(index).copied().unwrap_or(0.0);
    }
    (0x4036, 0x8000 | u16::from(font), body)
}

pub fn plus_draw_image(image: u8, source: [f32; 4], dest: [f32; 4]) -> (u16, u16, Vec<u8>) {
    let mut body = u32s(&[0]);
    body.extend(i32s(&[2]));
    body.extend(f32s(&source));
    body.extend(f32s(&dest));
    (0x401A, u16::from(image), body)
}

pub fn plus_clip_rect(mode: u16, rect: [f32; 4]) -> (u16, u16, Vec<u8>) {
    (0x4032, mode << 8, f32s(&rect))
}

pub fn plus_world(m: [f32; 6]) -> (u16, u16, Vec<u8>) {
    (0x402A, 0, f32s(&m))
}

pub fn plus_eof() -> (u16, u16, Vec<u8>) {
    (0x4002, 0, Vec::new())
}

/// A placeable WMF whose picture is `width` by `height` logical units at
/// `inch` units per inch.
pub struct Wmf {
    width: i16,
    height: i16,
    inch: u16,
    body: Vec<u8>,
    objects: u16,
    max_record: u32,
}

impl Wmf {
    pub fn new(width: i16, height: i16, inch: u16) -> Self {
        Self {
            width,
            height,
            inch,
            body: Vec::new(),
            objects: 16,
            max_record: 3,
        }
    }

    pub fn rec(mut self, function: u16, params: &[u8]) -> Self {
        let mut params = params.to_vec();
        if !params.len().is_multiple_of(2) {
            params.push(0);
        }
        let words = (params.len() / 2 + 3) as u32;
        self.max_record = self.max_record.max(words);
        self.body.extend(words.to_le_bytes());
        self.body.extend(function.to_le_bytes());
        self.body.extend(params);
        self
    }

    pub fn recs(self, records: Vec<(u16, Vec<u8>)>) -> Self {
        records
            .into_iter()
            .fold(self, |wmf, (function, params)| wmf.rec(function, &params))
    }

    pub fn bytes(self) -> Vec<u8> {
        let mut placeable = u32s(&[0x9AC6_CDD7]);
        placeable.extend(u16s(&[0]));
        placeable.extend(i16s(&[0, 0, self.width, self.height]));
        placeable.extend(u16s(&[self.inch]));
        placeable.extend(u32s(&[0]));
        let checksum = placeable.as_chunks::<2>().0.iter().fold(0u16, |sum, pair| {
            sum ^ u16::from_le_bytes([pair[0], pair[1]])
        });
        placeable.extend(u16s(&[checksum]));
        let total = (18 + self.body.len() + 6) / 2;
        let mut out = placeable;
        out.extend(u16s(&[1, 9, 0x0300]));
        out.extend(u32s(&[total as u32]));
        out.extend(u16s(&[self.objects]));
        out.extend(u32s(&[self.max_record]));
        out.extend(u16s(&[0]));
        out.extend(self.body);
        out.extend(u32s(&[3]));
        out.extend(u16s(&[0]));
        out
    }
}

/// WMF `PointS` pairs, `x` before `y`.
pub fn points_xy(points: &[(i16, i16)]) -> Vec<u8> {
    points.iter().flat_map(|(x, y)| i16s(&[*x, *y])).collect()
}
