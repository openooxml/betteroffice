//! A replayed drawing as one self-contained SVG document.
//!
//! The output references nothing outside itself: no scripts, no styles, no
//! external URLs, and bitmaps only as embedded PNG or JPEG `data:` URLs.

use std::collections::HashMap;
use std::fmt::Write as _;
use std::sync::Arc;

use base64::Engine as _;

use crate::drawing::{
    Bitmap, ClipChain, Drawing, Hatch, Image, LineCap, LineJoin, Omission, Op, Paint, PathCommand,
    Pixels, Refusal, Rgba, Shape, Spread, Stroke, Text, TextAnchor,
};

/// Largest SVG document [`to_svg`] writes.
pub const MAX_SVG_BYTES: usize = 16 * 1024 * 1024;

#[derive(Debug, Clone, PartialEq)]
pub struct Svg {
    pub markup: String,
    pub width: f64,
    pub height: f64,
    pub omissions: Vec<Omission>,
}

/// Replays `bytes` and writes the drawing as SVG.
pub fn to_svg(bytes: &[u8]) -> Result<Svg, Refusal> {
    to_svg_with_budget(bytes, &mut crate::ReplayBudget::default())
}

/// Writes SVG while spending cumulative replay allowances.
pub fn to_svg_with_budget(bytes: &[u8], budget: &mut crate::ReplayBudget) -> Result<Svg, Refusal> {
    let drawing = crate::replay_with_budget(bytes, budget)?;
    let markup = write(&drawing)?;
    Ok(Svg {
        markup,
        width: drawing.width,
        height: drawing.height,
        omissions: drawing.omissions,
    })
}

/// A neutral box with a picture glyph, standing in for a picture that could
/// not be drawn.
pub fn placeholder_svg(width: f64, height: f64) -> String {
    let (width, height) = (
        clean(width).clamp(1.0, 16_384.0),
        clean(height).clamp(1.0, 16_384.0),
    );
    let scale = (width.min(height) / 48.0).clamp(0.25, 4.0);
    format!(
        concat!(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}" preserveAspectRatio="none">"#,
            r##"<rect x="0.5" y="0.5" width="{iw}" height="{ih}" fill="#f1f3f4" stroke="#c4c7c5"/>"##,
            r##"<g transform="translate({cx} {cy}) scale({s})" fill="none" stroke="#9aa0a6" stroke-width="1.5" stroke-linejoin="round">"##,
            r#"<rect x="-12" y="-9" width="24" height="18" rx="2"/><circle cx="-5" cy="-3.5" r="2"/>"#,
            r#"<path d="M-10 7L-3 0L2 5L5 2L10 7"/></g></svg>"#
        ),
        w = num(width),
        h = num(height),
        iw = num((width - 1.0).max(0.0)),
        ih = num((height - 1.0).max(0.0)),
        cx = num(width / 2.0),
        cy = num(height / 2.0),
        s = num(scale),
    )
}

fn clean(value: f64) -> f64 {
    if value.is_finite() { value } else { 1.0 }
}

struct Writer {
    out: String,
    clips: HashMap<*const ClipChain, usize>,
    paints: HashMap<String, usize>,
    /// Each bitmap's `data:` URL, encoded once however often it is drawn.
    bitmaps: HashMap<*const Bitmap, Result<Arc<str>, Oversize>>,
    /// Set when a write stopped at the size limit.
    full: bool,
    next: usize,
    width: f64,
    height: f64,
}

pub(crate) fn write(drawing: &Drawing) -> Result<String, Refusal> {
    let mut writer = Writer {
        out: String::new(),
        clips: HashMap::new(),
        paints: HashMap::new(),
        bitmaps: HashMap::new(),
        full: false,
        next: 0,
        width: drawing.width,
        height: drawing.height,
    };
    let (w, h) = (num(drawing.width), num(drawing.height));
    let _ = write!(
        writer.out,
        r#"<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}" preserveAspectRatio="none">"#
    );
    let mut open: Option<*const ClipChain> = None;
    for op in &drawing.ops {
        let clip = match op {
            Op::Shape(shape) => &shape.clip,
            Op::Text(text) => &text.clip,
            Op::Image(image) => &image.clip,
        };
        let key = clip.as_ref().map(Arc::as_ptr);
        if key != open {
            if open.is_some() {
                writer.out.push_str("</g>");
            }
            if let Some(chain) = clip {
                let id = writer.clip(chain);
                let _ = write!(writer.out, r#"<g clip-path="url(#c{id})">"#);
            }
            open = key;
        }
        match op {
            Op::Shape(shape) => writer.shape(shape),
            Op::Text(text) => writer.text(text),
            Op::Image(image) => writer.image(image),
        }
        if writer.full || writer.out.len() > MAX_SVG_BYTES {
            return Err(Refusal(
                "the metafile's SVG would exceed the display size limit".to_owned(),
            ));
        }
    }
    if open.is_some() {
        writer.out.push_str("</g>");
    }
    writer.out.push_str("</svg>");
    if writer.full || writer.out.len() > MAX_SVG_BYTES {
        return Err(Refusal(
            "the metafile's SVG would exceed the display size limit".to_owned(),
        ));
    }
    Ok(writer.out)
}

impl Writer {
    fn id(&mut self) -> usize {
        self.next += 1;
        self.next
    }

    /// Defines `chain` and its ancestors once each, returning its id.
    fn clip(&mut self, chain: &Arc<ClipChain>) -> usize {
        if let Some(id) = self.clips.get(&Arc::as_ptr(chain)) {
            return *id;
        }
        let parent = chain.parent.as_ref().map(|parent| self.clip(parent));
        let id = self.id();
        self.clips.insert(Arc::as_ptr(chain), id);
        let _ = write!(self.out, r#"<clipPath id="c{id}""#);
        if let Some(parent) = parent {
            let _ = write!(self.out, r#" clip-path="url(#c{parent})""#);
        }
        self.out.push_str(r#"><path d=""#);
        if chain.region.exclude {
            let (w, h) = (self.width, self.height);
            path_data(
                &mut self.out,
                &crate::player::rect_path([-10.0 * w, -10.0 * h, 11.0 * w, 11.0 * h]),
            );
        }
        path_data(&mut self.out, &chain.region.path);
        self.out.push('"');
        if chain.region.exclude || chain.region.even_odd {
            self.out.push_str(r#" clip-rule="evenodd""#);
        }
        self.out.push_str("/></clipPath>");
        id
    }

    /// The SVG paint for `paint`, defining a pattern or gradient when it needs one.
    fn paint(&mut self, paint: &Paint) -> (String, Option<f64>) {
        match paint {
            Paint::Solid(color) => (color.hex(), opacity(*color)),
            Paint::Hatch {
                style,
                color,
                background,
                cell,
            } => {
                let key = format!("h{style:?}{color:?}{background:?}{cell}");
                let id = self.define(key, |writer, id| {
                    let c = num(*cell);
                    let _ = write!(
                        writer.out,
                        r#"<pattern id="p{id}" patternUnits="userSpaceOnUse" width="{c}" height="{c}">"#
                    );
                    if let Some(background) = background {
                        let _ = write!(
                            writer.out,
                            r#"<rect width="{c}" height="{c}" fill="{}"/>"#,
                            background.hex()
                        );
                    }
                    let (half, full) = (num(cell / 2.0), c.clone());
                    let lines = match style {
                        Hatch::Horizontal => format!("M0 {half}H{full}"),
                        Hatch::Vertical => format!("M{half} 0V{full}"),
                        Hatch::ForwardDiagonal => format!("M0 0L{full} {full}"),
                        Hatch::BackwardDiagonal => format!("M0 {full}L{full} 0"),
                        Hatch::Cross => format!("M0 {half}H{full}M{half} 0V{full}"),
                        Hatch::DiagonalCross => {
                            format!("M0 0L{full} {full}M0 {full}L{full} 0")
                        }
                    };
                    let _ = write!(
                        writer.out,
                        r#"<path d="{lines}" stroke="{}" stroke-width="{}"/></pattern>"#,
                        color.hex(),
                        num(cell / 8.0)
                    );
                });
                (format!("url(#p{id})"), None)
            }
            Paint::Pattern {
                tile,
                width,
                height,
            } => {
                let key = format!("t{:p}{width}{height}", Arc::as_ptr(tile));
                let id = self.define(key, |writer, id| {
                    let href = writer.bitmap_url(tile).unwrap_or_else(|| Arc::from(""));
                    let (w, h) = (num(*width), num(*height));
                    let _ = write!(
                        writer.out,
                        r#"<pattern id="p{id}" patternUnits="userSpaceOnUse" width="{w}" height="{h}"><image width="{w}" height="{h}" preserveAspectRatio="none" href="{href}"/></pattern>"#,
                    );
                });
                (format!("url(#p{id})"), None)
            }
            Paint::Linear(gradient) => {
                let key = format!("g{gradient:?}");
                let id = self.define(key, |writer, id| {
                    let spread = match gradient.spread {
                        Spread::Pad => "pad",
                        Spread::Repeat => "repeat",
                        Spread::Reflect => "reflect",
                    };
                    let _ = write!(
                        writer.out,
                        r#"<linearGradient id="p{id}" gradientUnits="userSpaceOnUse" x1="{}" y1="{}" x2="{}" y2="{}" spreadMethod="{spread}">"#,
                        num(gradient.start.0),
                        num(gradient.start.1),
                        num(gradient.end.0),
                        num(gradient.end.1)
                    );
                    for (offset, color) in gradient.stops.iter() {
                        let _ = write!(
                            writer.out,
                            r#"<stop offset="{}" stop-color="{}""#,
                            num(offset.clamp(0.0, 1.0)),
                            color.hex()
                        );
                        if let Some(alpha) = opacity(*color) {
                            let _ = write!(writer.out, r#" stop-opacity="{}""#, num(alpha));
                        }
                        writer.out.push_str("/>");
                    }
                    writer.out.push_str("</linearGradient>");
                });
                (format!("url(#p{id})"), None)
            }
        }
    }

    fn define(&mut self, key: String, body: impl FnOnce(&mut Self, usize)) -> usize {
        if let Some(id) = self.paints.get(&key) {
            return *id;
        }
        let id = self.id();
        self.paints.insert(key, id);
        self.out.push_str("<defs>");
        body(self, id);
        self.out.push_str("</defs>");
        id
    }

    fn shape(&mut self, shape: &Shape) {
        let fill = shape.fill.as_ref().map(|paint| self.paint(paint));
        let stroke = shape
            .stroke
            .as_ref()
            .map(|stroke| (self.paint(&stroke.paint), stroke));
        self.out.push_str(r#"<path d=""#);
        path_data(&mut self.out, &shape.path);
        self.out.push('"');
        match fill {
            Some((paint, alpha)) => {
                let _ = write!(self.out, r#" fill="{paint}""#);
                if let Some(alpha) = alpha {
                    let _ = write!(self.out, r#" fill-opacity="{}""#, num(alpha));
                }
                if shape.even_odd {
                    self.out.push_str(r#" fill-rule="evenodd""#);
                }
            }
            None => self.out.push_str(r#" fill="none""#),
        }
        if let Some(((paint, alpha), stroke)) = stroke {
            self.stroke(&paint, alpha, stroke);
        }
        self.out.push_str("/>");
    }

    fn stroke(&mut self, paint: &str, alpha: Option<f64>, stroke: &Stroke) {
        let _ = write!(
            self.out,
            r#" stroke="{paint}" stroke-width="{}""#,
            num(stroke.width)
        );
        if let Some(alpha) = alpha {
            let _ = write!(self.out, r#" stroke-opacity="{}""#, num(alpha));
        }
        match stroke.cap {
            LineCap::Butt => {}
            LineCap::Round => self.out.push_str(r#" stroke-linecap="round""#),
            LineCap::Square => self.out.push_str(r#" stroke-linecap="square""#),
        }
        match stroke.join {
            LineJoin::Miter => {
                if stroke.miter_limit != 4.0 {
                    let _ = write!(
                        self.out,
                        r#" stroke-miterlimit="{}""#,
                        num(stroke.miter_limit.max(1.0))
                    );
                }
            }
            LineJoin::Round => self.out.push_str(r#" stroke-linejoin="round""#),
            LineJoin::Bevel => self.out.push_str(r#" stroke-linejoin="bevel""#),
        }
        if let Some(dash) = stroke.dash.as_ref().filter(|dash| {
            dash.iter()
                .all(|length| length.is_finite() && *length >= 0.0)
                && dash.iter().any(|length| *length > 0.0)
        }) {
            self.out.push_str(r#" stroke-dasharray=""#);
            for (index, length) in dash.iter().enumerate() {
                if index > 0 {
                    self.out.push(' ');
                }
                self.out.push_str(&num(*length));
            }
            self.out.push('"');
        }
    }

    fn text(&mut self, text: &Text) {
        let content: String = text.text.chars().filter(|c| xml_char(*c)).collect();
        if content.trim().is_empty() {
            return;
        }
        let m = text.transform;
        let scale = (m[0] * m[3] - m[1] * m[2]).abs().sqrt();
        if !scale.is_normal() || !m.iter().all(|value| value.is_finite()) {
            return;
        }
        let (fill, alpha) = self.paint(&text.fill);
        let font = &text.font;
        let _ = write!(
            self.out,
            r#"<text transform="matrix({} {} {} {} {} {})" font-family="{}" font-size="{}""#,
            ratio(m[0] / scale),
            ratio(m[1] / scale),
            ratio(m[2] / scale),
            ratio(m[3] / scale),
            num(m[4]),
            num(m[5]),
            font_family(&font.family),
            num(font.size * scale)
        );
        if font.weight >= 600 {
            self.out.push_str(r#" font-weight="bold""#);
        } else if font.weight != 400 && font.weight != 0 {
            let _ = write!(
                self.out,
                r#" font-weight="{}""#,
                font.weight.clamp(100, 900)
            );
        }
        if font.italic {
            self.out.push_str(r#" font-style="italic""#);
        }
        match (font.underline, font.strike) {
            (true, true) => self
                .out
                .push_str(r#" text-decoration="underline line-through""#),
            (true, false) => self.out.push_str(r#" text-decoration="underline""#),
            (false, true) => self.out.push_str(r#" text-decoration="line-through""#),
            (false, false) => {}
        }
        let _ = write!(self.out, r#" fill="{fill}""#);
        if let Some(alpha) = alpha {
            let _ = write!(self.out, r#" fill-opacity="{}""#, num(alpha));
        }
        match text.anchor {
            TextAnchor::Start => {}
            TextAnchor::Middle => self.out.push_str(r#" text-anchor="middle""#),
            TextAnchor::End => self.out.push_str(r#" text-anchor="end""#),
        }
        self.out.push_str(r#" xml:space="preserve""#);
        let positions = text
            .positions
            .as_ref()
            .filter(|positions| positions.len() == text.text.chars().count());
        match positions {
            Some(positions) => {
                let kept: Vec<(f64, f64)> = text
                    .text
                    .chars()
                    .zip(positions)
                    .filter(|(c, _)| xml_char(*c))
                    .map(|(_, position)| *position)
                    .collect();
                self.out.push_str(r#" x=""#);
                list(&mut self.out, kept.iter().map(|(x, _)| x * scale));
                self.out.push_str(r#"" y=""#);
                if kept.iter().all(|(_, y)| *y == kept[0].1) {
                    self.out.push_str(&num(kept[0].1 * scale));
                } else {
                    list(&mut self.out, kept.iter().map(|(_, y)| y * scale));
                }
                self.out.push('"');
            }
            None => self.out.push_str(r#" x="0" y="0""#),
        }
        self.out.push('>');
        escape(&mut self.out, &content);
        self.out.push_str("</text>");
    }

    /// `bitmap` as a `data:` URL, encoded on first use; `None` when it
    /// cannot be encoded or would not fit the size limit.
    fn bitmap_url(&mut self, bitmap: &Arc<Bitmap>) -> Option<Arc<str>> {
        let href = match self
            .bitmaps
            .entry(Arc::as_ptr(bitmap))
            .or_insert_with(|| data_url(bitmap).map(Arc::from))
            .clone()
        {
            Ok(href) => href,
            Err(Oversize(oversize)) => {
                self.full |= oversize;
                return None;
            }
        };
        if self.out.len().saturating_add(href.len()) > MAX_SVG_BYTES {
            self.full = true;
            return None;
        }
        Some(href)
    }

    fn image(&mut self, image: &Image) {
        let Some(href) = self.bitmap_url(&image.bitmap) else {
            return;
        };
        let m = image.transform;
        if !m.iter().all(|value| value.is_finite()) {
            return;
        }
        let _ = write!(
            self.out,
            r#"<image width="{}" height="{}" preserveAspectRatio="none" transform="matrix({} {} {} {} {} {})""#,
            image.bitmap.width,
            image.bitmap.height,
            ratio(m[0]),
            ratio(m[1]),
            ratio(m[2]),
            ratio(m[3]),
            num(m[4]),
            num(m[5])
        );
        if image.opacity < 1.0 {
            let _ = write!(self.out, r#" opacity="{}""#, num(image.opacity.max(0.0)));
        }
        let _ = write!(self.out, r#" href="{href}"/>"#);
    }
}

fn opacity(color: Rgba) -> Option<f64> {
    (color.a < 255).then(|| f64::from(color.a) / 255.0)
}

/// A bitmap as a `data:` URL, PNG-encoding raw pixels.
/// Why a bitmap has no `data:` URL: `true` when it would not fit the size
/// limit, `false` when it could not be encoded.
#[derive(Clone, Copy)]
struct Oversize(bool);

fn data_url(bitmap: &Bitmap) -> Result<String, Oversize> {
    let (mime, bytes) = match &bitmap.pixels {
        Pixels::Encoded { mime, bytes } => (*mime, bytes.clone()),
        Pixels::Rgba(rgba) => {
            let pixels = rgba.as_chunks::<4>().0;
            let opaque = pixels.iter().all(|pixel| pixel[3] == 255);
            let encoded = if opaque {
                let rgb: Vec<u8> = pixels
                    .iter()
                    .flat_map(|pixel| [pixel[0], pixel[1], pixel[2]])
                    .collect();
                ooxml_drawingml::png_encode::encode_rgb8(&rgb, bitmap.width, bitmap.height)
            } else {
                ooxml_drawingml::png_encode::encode_rgba8(rgba, bitmap.width, bitmap.height)
            };
            ("image/png", encoded.map_err(|_| Oversize(false))?)
        }
    };
    if bytes.len() / 3 * 4 > MAX_SVG_BYTES {
        return Err(Oversize(true));
    }
    Ok(format!(
        "data:{mime};base64,{}",
        base64::engine::general_purpose::STANDARD.encode(bytes)
    ))
}

/// Writes `path`, stopping once `out` passes the size limit.
fn path_data(out: &mut String, path: &[PathCommand]) {
    for command in path {
        if out.len() > MAX_SVG_BYTES {
            return;
        }
        match command {
            PathCommand::Move { x, y } => {
                let _ = write!(out, "M{} {}", num(*x), num(*y));
            }
            PathCommand::Line { x, y } => {
                let _ = write!(out, "L{} {}", num(*x), num(*y));
            }
            PathCommand::Quad { cpx, cpy, x, y } => {
                let _ = write!(out, "Q{} {} {} {}", num(*cpx), num(*cpy), num(*x), num(*y));
            }
            PathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                let _ = write!(
                    out,
                    "C{} {} {} {} {} {}",
                    num(*cp1x),
                    num(*cp1y),
                    num(*cp2x),
                    num(*cp2y),
                    num(*x),
                    num(*y)
                );
            }
            PathCommand::Close => out.push('Z'),
        }
    }
}

fn list(out: &mut String, values: impl Iterator<Item = f64>) {
    for (index, value) in values.enumerate() {
        if out.len() > MAX_SVG_BYTES {
            return;
        }
        if index > 0 {
            out.push(' ');
        }
        out.push_str(&num(value));
    }
}

/// A coordinate to three decimals, without trailing zeros.
pub(crate) fn num(value: f64) -> String {
    fixed(value, 1_000.0)
}

/// A transform coefficient to six decimals.
fn ratio(value: f64) -> String {
    fixed(value, 1_000_000.0)
}

fn fixed(value: f64, scale: f64) -> String {
    if !value.is_finite() {
        return "0".to_owned();
    }
    let rounded = (value * scale).round() / scale;
    if rounded == 0.0 {
        return "0".to_owned();
    }
    let mut text = format!("{rounded:.6}");
    while text.ends_with('0') {
        text.pop();
    }
    if text.ends_with('.') {
        text.pop();
    }
    text
}

fn xml_char(c: char) -> bool {
    matches!(c, '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..)
}

fn escape(out: &mut String, text: &str) {
    for c in text.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\t' | '\n' | '\r' => out.push(' '),
            c => out.push(c),
        }
    }
}

/// The face, then metric-compatible substitutes and a generic family.
fn font_family(face: &str) -> String {
    let face: String = face
        .chars()
        .filter(|c| c.is_alphanumeric() || matches!(c, ' ' | '-' | '_' | '.'))
        .take(64)
        .collect();
    let face = face.trim();
    let lower = face.to_ascii_lowercase();
    let (alternates, generic): (&[&str], &str) = match lower.as_str() {
        "arial"
        | "helvetica"
        | "arial unicode ms"
        | "system"
        | "ms sans serif"
        | "microsoft sans serif" => (
            &["Liberation Sans", "Arimo", "Helvetica", "Arial"],
            "sans-serif",
        ),
        "times new roman" | "times" | "ms serif" => {
            (&["Liberation Serif", "Tinos", "Times"], "serif")
        }
        "courier new" | "courier" => (&["Liberation Mono", "Cousine", "Courier"], "monospace"),
        "calibri" => (&["Carlito"], "sans-serif"),
        "cambria" | "cambria math" => (&["Caladea"], "serif"),
        "symbol" => (&["Times New Roman", "Liberation Serif"], "serif"),
        _ if lower.contains("mono") || lower.contains("courier") || lower.contains("consol") => {
            (&[], "monospace")
        }
        _ if lower.contains("times")
            || lower.contains("roman")
            || lower.contains("serif") && !lower.contains("sans")
            || lower.contains("georgia")
            || lower.contains("garamond")
            || lower.contains("book") =>
        {
            (&[], "serif")
        }
        _ => (&[], "sans-serif"),
    };
    let mut family = String::new();
    if !face.is_empty() {
        let _ = write!(family, "'{face}'");
    }
    for alternate in alternates {
        if !alternate.eq_ignore_ascii_case(face) {
            let _ = write!(
                family,
                "{}'{alternate}'",
                if family.is_empty() { "" } else { ", " }
            );
        }
    }
    let _ = write!(
        family,
        "{}{generic}",
        if family.is_empty() { "" } else { ", " }
    );
    family
}
