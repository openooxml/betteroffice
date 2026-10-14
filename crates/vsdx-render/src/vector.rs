use ooxml_drawingml::GeometryPathCommand;

use crate::display_list::{Affine, Paint, PositionedLine, Primitive, TextParagraph};

/// Compact float formatting shared by the vector exporters.
pub(crate) fn num(value: f64) -> String {
    if value == 0.0 {
        return "0".to_owned();
    }
    let text = format!("{value:.4}");
    let text = text.trim_end_matches('0').trim_end_matches('.');
    if text == "-0" {
        "0".to_owned()
    } else {
        text.to_owned()
    }
}

/// Parses a `#rrggbb` colour into 0–1 channels.
pub(crate) fn rgb(color: &str) -> Option<(f64, f64, f64)> {
    let hex = color.strip_prefix('#')?;
    if hex.len() != 6 {
        return None;
    }
    let channel = |range: std::ops::Range<usize>| {
        u8::from_str_radix(hex.get(range)?, 16)
            .ok()
            .map(|v| f64::from(v) / 255.0)
    };
    Some((channel(0..2)?, channel(2..4)?, channel(4..6)?))
}

/// Solid colour as authored, falling back to the first gradient stop.
pub(crate) fn solid_color(paint: &Option<Paint>) -> Option<&str> {
    match paint {
        Some(Paint::Solid { color }) => rgb(color).map(|_| color.as_str()),
        Some(Paint::Gradient { stops, .. }) => stops
            .iter()
            .find(|stop| rgb(&stop.color).is_some())
            .map(|stop| stop.color.as_str()),
        None => None,
    }
}

/// Bounds of a geometry path as `(min_x, min_y, max_x, max_y)`.
pub(crate) fn path_bounds(path: &[GeometryPathCommand]) -> Option<(f32, f32, f32, f32)> {
    let (mut min_x, mut min_y) = (f32::INFINITY, f32::INFINITY);
    let (mut max_x, mut max_y) = (f32::NEG_INFINITY, f32::NEG_INFINITY);
    let mut point = |x: f64, y: f64| {
        let (x, y) = (x as f32, y as f32);
        if x.is_finite() && y.is_finite() {
            min_x = min_x.min(x);
            min_y = min_y.min(y);
            max_x = max_x.max(x);
            max_y = max_y.max(y);
        }
    };
    for command in path {
        match command {
            GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
                point(*x, *y)
            }
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                point(*cpx, *cpy);
                point(*x, *y);
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                point(*cp1x, *cp1y);
                point(*cp2x, *cp2y);
                point(*x, *y);
            }
            GeometryPathCommand::Close => {}
        }
    }
    (min_x.is_finite() && min_y.is_finite()).then_some((min_x, min_y, max_x, max_y))
}

/// One fill's linear gradient in path space: endpoints and ordered stops.
pub struct LinearGradient {
    pub start: (f32, f32),
    pub end: (f32, f32),
    pub stops: Vec<(f32, String)>,
}

/// Gradient geometry for `fill` over `path`, or `None` when it paints solid.
///
/// Shared by every exporter so the projections cannot disagree.
pub fn linear_gradient(fill: &Paint, path: &[GeometryPathCommand]) -> Option<LinearGradient> {
    let Paint::Gradient { angle_deg, stops } = fill else {
        return None;
    };
    let mut ordered: Vec<(f32, String)> = stops
        .iter()
        .filter(|stop| stop.position.is_finite() && rgb(&stop.color).is_some())
        .map(|stop| (stop.position.clamp(0.0, 1.0), stop.color.clone()))
        .collect();
    if ordered.len() < 2 {
        return None;
    }
    ordered.sort_by(|left, right| left.0.total_cmp(&right.0));
    let (min_x, min_y, max_x, max_y) = path_bounds(path)?;
    let radius = (max_x - min_x).hypot(max_y - min_y) / 2.0;
    if !radius.is_finite() || radius <= 0.0 {
        return None;
    }
    let (center_x, center_y) = ((min_x + max_x) / 2.0, (min_y + max_y) / 2.0);
    let (sin, cos) = angle_deg.unwrap_or(0.0).to_radians().sin_cos();
    Some(LinearGradient {
        start: (center_x - cos * radius, center_y - sin * radius),
        end: (center_x + cos * radius, center_y + sin * radius),
        stops: ordered,
    })
}

#[derive(Clone, Copy)]
pub(crate) struct FontClass {
    pub class: u8,
}

impl FontClass {
    pub fn of(family: &str) -> Self {
        let family = family.to_lowercase();
        let class = if ["times", "serif", "georgia", "garamond", "palatino"]
            .iter()
            .any(|marker| family.contains(marker))
        {
            1
        } else if ["courier", "mono", "consol", "menlo"]
            .iter()
            .any(|marker| family.contains(marker))
        {
            2
        } else {
            0
        };
        Self { class }
    }

    pub fn generic(&self) -> &'static str {
        match self.class {
            1 => "serif",
            2 => "monospace",
            _ => "sans-serif",
        }
    }
}

/// Width and height of a JPEG stream without decoding it.
pub(crate) fn jpeg_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8 {
        return None;
    }
    let mut index = 2;
    while index + 4 <= bytes.len() {
        if bytes[index] != 0xFF {
            return None;
        }
        let marker = bytes[index + 1];
        index += 2;
        if marker == 0xD9 {
            return None;
        }
        if marker == 0x01 || (0xD0..=0xD7).contains(&marker) {
            continue;
        }
        if index + 2 > bytes.len() {
            return None;
        }
        let length = u16::from_be_bytes([bytes[index], bytes[index + 1]]) as usize;
        if length < 2 || index + length > bytes.len() {
            return None;
        }
        if matches!(marker, 0xC0..=0xC2) && length >= 7 {
            let height = u16::from_be_bytes([bytes[index + 3], bytes[index + 4]]) as u32;
            let width = u16::from_be_bytes([bytes[index + 5], bytes[index + 6]]) as u32;
            if width == 0 || height == 0 {
                return None;
            }
            return Some((width, height));
        }
        index += length;
    }
    None
}

/// Width and height of a PNG stream from its IHDR chunk.
pub(crate) fn png_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 24
        || bytes[0..8] != [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]
        || &bytes[12..16] != b"IHDR"
    {
        return None;
    }
    let width = u32::from_be_bytes([bytes[16], bytes[17], bytes[18], bytes[19]]);
    let height = u32::from_be_bytes([bytes[20], bytes[21], bytes[22], bytes[23]]);
    if width == 0 || height == 0 {
        return None;
    }
    Some((width, height))
}

pub fn z_order(primitive: &Primitive) -> u32 {
    match primitive {
        Primitive::Shape { z_order, .. }
        | Primitive::Image { z_order, .. }
        | Primitive::TextBox { z_order, .. }
        | Primitive::Placeholder { z_order, .. }
        | Primitive::Group { z_order, .. } => *z_order,
    }
}

/// Flattens groups depth-first, composing transforms, ordered back-to-front.
pub fn collect_ordered<'b>(primitive: &'b Primitive, out: &mut Vec<(&'b Primitive, Affine)>) {
    match primitive {
        Primitive::Group {
            primitives,
            transform,
            ..
        } => {
            let mut children: Vec<&Primitive> = primitives.iter().collect();
            children.sort_by_key(|child| z_order(child));
            for child in children {
                let before = out.len();
                collect_ordered(child, out);
                for (_, matrix) in &mut out[before..] {
                    *matrix = transform.compose(*matrix);
                }
            }
            out.push((primitive, Affine::identity()));
        }
        Primitive::Shape { transform, .. }
        | Primitive::Image { transform, .. }
        | Primitive::TextBox { transform, .. } => out.push((primitive, *transform)),
        Primitive::Placeholder { .. } => out.push((primitive, Affine::identity())),
    }
}

pub(crate) fn stop_x(line: &PositionedLine, position: u32) -> Option<f32> {
    line.caret_stops
        .iter()
        .find(|stop| stop.position == position)
        .map(|stop| stop.x)
}

/// One uniformly styled slice of a laid-out line.
pub(crate) struct Fragment {
    pub text: String,
    pub x: f32,
    pub line_y: f32,
    pub family: String,
    pub generic: &'static str,
    pub size_in: f32,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub letter_spacing: f32,
    pub color: String,
}

/// Owned text slice for exporters outside this crate.
pub struct TextFragment {
    pub text: String,
    pub x: f32,
    pub line_y: f32,
    pub family: String,
    pub size_in: f32,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub letter_spacing: f32,
    pub color: String,
}

/// Splits laid-out lines at run boundaries for exporters.
pub fn text_fragments(paragraphs: &[TextParagraph], lines: &[PositionedLine]) -> Vec<TextFragment> {
    fragments(paragraphs, lines)
        .into_iter()
        .map(|fragment| TextFragment {
            text: fragment.text,
            x: fragment.x,
            line_y: fragment.line_y,
            family: fragment.family,
            size_in: fragment.size_in,
            bold: fragment.bold,
            italic: fragment.italic,
            underline: fragment.underline,
            letter_spacing: fragment.letter_spacing,
            color: fragment.color,
        })
        .collect()
}

/// Splits laid-out lines at run boundaries, dropping control characters.
pub(crate) fn fragments(paragraphs: &[TextParagraph], lines: &[PositionedLine]) -> Vec<Fragment> {
    let mut offset = 0u32;
    let runs: Vec<(&crate::display_list::TextRun, u32, u32)> = paragraphs
        .iter()
        .flat_map(|paragraph| paragraph.runs.iter())
        .map(|run| {
            let start = offset;
            offset += run.text.len() as u32;
            (run, start, offset)
        })
        .collect();
    let mut out = Vec::new();
    for line in lines {
        for (run, run_start, run_end) in &runs {
            let start = (*run_start).max(line.start);
            let end = (*run_end).min(line.end);
            if start >= end {
                continue;
            }
            let fragment = &run.text[(start - run_start) as usize..(end - run_start) as usize];
            let visible: String = fragment
                .chars()
                .filter(|ch| !matches!(ch, '\t' | '\n' | '\r'))
                .collect();
            if visible.is_empty() || !run.size_in.is_finite() || run.size_in <= 0.0 {
                continue;
            }
            let class = FontClass::of(&run.family);
            out.push(Fragment {
                text: visible,
                x: stop_x(line, start).unwrap_or(line.x),
                line_y: line.y,
                family: run.family.clone(),
                generic: class.generic(),
                size_in: run.size_in,
                bold: run.bold,
                italic: run.italic,
                underline: run.underline,
                letter_spacing: run.letter_spacing,
                color: rgb(&run.color)
                    .map(|_| run.color.clone())
                    .unwrap_or_else(|| "#000000".into()),
            });
        }
    }
    out
}

/// Escapes text for SVG content and attribute values.
pub(crate) fn escape(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(ch),
        }
    }
    out
}

/// Formats an affine as an SVG transform attribute.
pub(crate) fn matrix(matrix: Affine) -> String {
    format!(
        "matrix({} {} {} {} {} {})",
        num(f64::from(matrix.a)),
        num(f64::from(matrix.b)),
        num(f64::from(matrix.c)),
        num(f64::from(matrix.d)),
        num(f64::from(matrix.e)),
        num(f64::from(matrix.f))
    )
}
