//! A replayed drawing as one self-contained SVG document.
//!
//! The output references nothing outside itself: no scripts, no styles, no
//! external URLs, and bitmaps only as embedded PNG or JPEG `data:` URLs.

use std::collections::HashMap;
use std::fmt::Write as _;
use std::rc::Rc;
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
    let mut output_bytes = MAX_SVG_BYTES;
    to_svg_with_limits(bytes, budget, &mut output_bytes)
}

/// Spends replay work and output bytes, including writes before refusal.
pub fn to_svg_with_limits(
    bytes: &[u8],
    budget: &mut crate::ReplayBudget,
    output_bytes: &mut usize,
) -> Result<Svg, Refusal> {
    if *output_bytes == 0 {
        return Err(output_refusal());
    }
    crate::with_budget(budget, |shared| {
        let drawing = crate::play_nested(bytes, 0, Rc::clone(&shared)).map_err(Refusal)?;
        let markup = write_with_limits(&drawing, shared, output_bytes)?;
        Ok(Svg {
            markup,
            width: drawing.width,
            height: drawing.height,
            omissions: drawing.omissions,
        })
    })
}

fn output_refusal() -> Refusal {
    Refusal("the metafile's SVG would exceed the display size limit".to_owned())
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

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
struct GradientKey {
    stops: *const (f64, Rgba),
    start: [u64; 2],
    end: [u64; 2],
    spread: u8,
}

struct Writer {
    out: Output,
    clips: HashMap<*const ClipChain, usize>,
    paints: HashMap<String, usize>,
    gradients: HashMap<GradientKey, usize>,
    /// Each bitmap's `data:` URL, encoded once however often it is drawn.
    bitmaps: HashMap<*const Bitmap, Result<Arc<str>, Oversize>>,
    next: usize,
    width: f64,
    height: f64,
}

struct Output {
    markup: String,
    limit: usize,
    full: bool,
    budget: Rc<crate::player::SharedBudget>,
}

impl std::fmt::Write for Output {
    fn write_str(&mut self, text: &str) -> std::fmt::Result {
        if self.failed() {
            return Err(std::fmt::Error);
        }
        let end = self.markup.len().saturating_add(text.len());
        if end > self.limit {
            self.full = true;
            return Err(std::fmt::Error);
        }
        let work = end.div_ceil(64) - self.markup.len().div_ceil(64);
        if !self.budget.spend(work as u64, 0) {
            return Err(std::fmt::Error);
        }
        self.markup.push_str(text);
        Ok(())
    }
}

impl Output {
    fn failed(&self) -> bool {
        self.full || self.budget.exceeded.get()
    }

    fn push_str(&mut self, text: &str) {
        let _ = self.write_str(text);
    }

    fn push(&mut self, c: char) {
        let _ = self.write_char(c);
    }
}

fn write_with_limits(
    drawing: &Drawing,
    budget: Rc<crate::player::SharedBudget>,
    output_bytes: &mut usize,
) -> Result<String, Refusal> {
    let mut writer = Writer {
        out: Output {
            markup: String::new(),
            limit: (*output_bytes).min(MAX_SVG_BYTES),
            full: false,
            budget,
        },
        clips: HashMap::new(),
        paints: HashMap::new(),
        gradients: HashMap::new(),
        bitmaps: HashMap::new(),
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
        if writer.out.failed() || !writer.out.budget.spend(1, 0) {
            break;
        }
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
        if writer.out.failed() {
            break;
        }
        match op {
            Op::Shape(shape) => writer.shape(shape),
            Op::Text(text) => writer.text(text),
            Op::Image(image) => writer.image(image),
        }
    }
    if open.is_some() {
        writer.out.push_str("</g>");
    }
    writer.out.push_str("</svg>");
    if writer.out.full && *output_bytes <= MAX_SVG_BYTES {
        *output_bytes = 0;
    } else {
        *output_bytes -= writer.out.markup.len();
    }
    if writer.out.failed() {
        return Err(output_refusal());
    }
    Ok(writer.out.markup)
}

impl Writer {
    fn id(&mut self) -> usize {
        self.next += 1;
        self.next
    }

    /// Defines `chain` and its ancestors once each, returning its id.
    fn clip(&mut self, chain: &Arc<ClipChain>) -> usize {
        if self.out.failed() || !self.out.budget.spend(1, 0) {
            return 0;
        }
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
        if self.out.failed() {
            return (String::new(), None);
        }
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
                let key = GradientKey {
                    stops: Arc::as_ptr(&gradient.stops).cast(),
                    start: [gradient.start.0.to_bits(), gradient.start.1.to_bits()],
                    end: [gradient.end.0.to_bits(), gradient.end.1.to_bits()],
                    spread: gradient.spread as u8,
                };
                if let Some(id) = self.gradients.get(&key) {
                    return (format!("url(#p{id})"), None);
                }
                if !self.out.budget.spend(gradient.stops.len() as u64, 0) {
                    return (String::new(), None);
                }
                let id = self.define(format!("g{gradient:?}"), |writer, id| {
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
                        if writer.out.failed() {
                            break;
                        }
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
                self.gradients.insert(key, id);
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
                if self.out.failed() {
                    break;
                }
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
        if self.out.failed() {
            return None;
        }
        if !self.bitmaps.contains_key(&Arc::as_ptr(bitmap)) {
            let bytes = match &bitmap.pixels {
                Pixels::Encoded { bytes, .. } => bytes.len(),
                Pixels::Rgba(rgba) => rgba.len(),
            };
            if !self.out.budget.spend(bytes.div_ceil(64) as u64, 0) {
                return None;
            }
        }
        let limit = self.out.limit.saturating_sub(self.out.markup.len());
        let href = match self
            .bitmaps
            .entry(Arc::as_ptr(bitmap))
            .or_insert_with(|| data_url(bitmap, limit).map(Arc::from))
            .clone()
        {
            Ok(href) => href,
            Err(Oversize(oversize)) => {
                self.out.full |= oversize;
                return None;
            }
        };
        if self.out.markup.len().saturating_add(href.len()) > self.out.limit {
            self.out.full = true;
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

fn data_url(bitmap: &Bitmap, limit: usize) -> Result<String, Oversize> {
    let (mime, bytes) = match &bitmap.pixels {
        Pixels::Encoded { mime, bytes } => {
            if bytes
                .len()
                .div_ceil(3)
                .saturating_mul(4)
                .saturating_add(mime.len() + 13)
                > limit
            {
                return Err(Oversize(true));
            }
            (*mime, bytes.clone())
        }
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
    if bytes
        .len()
        .div_ceil(3)
        .saturating_mul(4)
        .saturating_add(mime.len() + 13)
        > limit
    {
        return Err(Oversize(true));
    }
    Ok(format!(
        "data:{mime};base64,{}",
        base64::engine::general_purpose::STANDARD.encode(bytes)
    ))
}

/// Writes `path` within the output allowance.
fn path_data(out: &mut Output, path: &[PathCommand]) {
    for command in path {
        if out.failed() {
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

fn list(out: &mut Output, values: impl Iterator<Item = f64>) {
    for (index, value) in values.enumerate() {
        if out.failed() {
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

fn escape(out: &mut Output, text: &str) {
    for c in text.chars() {
        if out.failed() {
            break;
        }
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::drawing::LinearGradient;
    use crate::test_records::*;

    fn write(drawing: &Drawing) -> Result<String, Refusal> {
        crate::with_budget(&mut crate::ReplayBudget::default(), |shared| {
            let mut output_bytes = MAX_SVG_BYTES;
            write_with_limits(drawing, shared, &mut output_bytes)
        })
    }

    fn gradient_fills(count: usize, translated: bool) -> Vec<u8> {
        let mut brush = plus_linear_brush(1, [0.0, 0.0, 10.0, 10.0], 0xff00_0000, 0xffff_ffff);
        brush.2[8..12].copy_from_slice(&4u32.to_le_bytes());
        brush.2.extend(u32s(&[256]));
        brush
            .2
            .extend((0..256).flat_map(|index| (index as f32 / 255.0).to_le_bytes()));
        brush.2.extend(u32s(&[0xff00_0000; 256]));
        let mut records = vec![plus_header(false), brush];
        for index in 0..count {
            if translated {
                records.push(plus_world([1.0, 0.0, 0.0, 1.0, index as f32, 0.0]));
            }
            records.push((
                0x400A,
                0,
                [u32s(&[1, 1]), f32s(&[0.0, 0.0, 10.0, 10.0])].concat(),
            ));
        }
        records.push(plus_eof());
        let (kind, body) = plus(&records);
        Emf::new(100, 100).rec(kind, &body).bytes()
    }

    #[test]
    fn repeated_linear_fills_define_one_gradient() {
        let drawing = crate::replay(&gradient_fills(32, false)).unwrap();
        assert_eq!(drawing.ops.len(), 32);
        let svg = write(&drawing).unwrap();
        assert_eq!(svg.matches("<linearGradient").count(), 1);
        assert_eq!(svg.matches("<stop ").count(), 256);
        assert_eq!(svg.matches("fill=\"url(#p1)\"").count(), 32);
    }

    #[test]
    fn translated_linear_fills_charge_stops_on_every_cache_miss() {
        for translated in [false, true] {
            let bytes = gradient_fills(8, translated);
            let mut replay_budget = crate::ReplayBudget::default();
            crate::replay_with_budget(&bytes, &mut replay_budget).unwrap();
            let mut budget = crate::ReplayBudget::default();
            let svg = to_svg_with_budget(&bytes, &mut budget).unwrap();
            let misses = if translated { 8 } else { 1 };
            assert_eq!(svg.markup.matches("<stop ").count(), misses * 256);
            assert_eq!(
                replay_budget.work - budget.work,
                (misses * 256 + 8 + svg.markup.len().div_ceil(64)) as u64
            );
            let mut budget = crate::ReplayBudget {
                work: crate::ReplayBudget::default().work - replay_budget.work + 256,
                pixels: 0,
            };
            assert!(to_svg_with_budget(&bytes, &mut budget).is_err());
            assert_eq!(budget.work, 0);
        }
    }

    #[test]
    fn svg_writes_stop_at_the_remaining_output_allowance() {
        let bytes = gradient_fills(8, true);
        let svg = to_svg(&bytes).unwrap();
        let mut budget = crate::ReplayBudget::default();
        let mut remaining = svg.markup.len();
        assert_eq!(
            to_svg_with_limits(&bytes, &mut budget, &mut remaining).unwrap(),
            svg
        );
        assert_eq!(remaining, 0);
        let before = budget;
        assert!(to_svg_with_limits(&bytes, &mut budget, &mut remaining).is_err());
        assert_eq!(budget, before);

        let mut budget = crate::ReplayBudget::default();
        let mut replay_budget = budget;
        crate::replay_with_budget(&bytes, &mut replay_budget).unwrap();
        let mut remaining = 512;
        assert!(to_svg_with_limits(&bytes, &mut budget, &mut remaining).is_err());
        assert_eq!(remaining, 0);
        assert!(replay_budget.work - budget.work <= 256 + 1 + 512u64.div_ceil(64));
    }

    #[test]
    fn output_appends_never_exceed_the_byte_or_work_allowance() {
        let budget = Rc::new(crate::player::SharedBudget {
            remaining: std::cell::Cell::new(crate::ReplayBudget { work: 2, pixels: 0 }),
            exceeded: std::cell::Cell::new(false),
        });
        let mut out = Output {
            markup: String::new(),
            limit: 80,
            full: false,
            budget: Rc::clone(&budget),
        };
        out.push_str(&"x".repeat(79));
        assert_eq!(out.markup.len(), 79);
        assert_eq!(budget.remaining.get().work, 0);
        out.push('é');
        assert!(out.full);
        assert_eq!(out.markup.len(), 79);
        out.push_str("discarded");
        assert_eq!(out.markup.len(), 79);

        out.full = false;
        out.limit = 1_000;
        out.push_str(&"x".repeat(50));
        assert!(budget.exceeded.get());
        assert_eq!(out.markup.len(), 79);
    }

    #[test]
    fn bitmap_encoding_charges_bytes_before_allocating_the_url() {
        let bitmap = Arc::new(Bitmap {
            width: 1,
            height: 1,
            pixels: Pixels::Encoded {
                mime: "image/png",
                bytes: vec![0; 1_024],
            },
        });
        let drawing = Drawing {
            width: 10.0,
            height: 10.0,
            ops: vec![Op::Image(Image {
                transform: crate::player::IDENTITY,
                bitmap,
                opacity: 1.0,
                clip: None,
            })],
            omissions: Vec::new(),
        };
        let mut budget = crate::ReplayBudget { work: 4, pixels: 0 };
        let mut remaining = MAX_SVG_BYTES;
        assert!(
            crate::with_budget(&mut budget, |shared| {
                write_with_limits(&drawing, shared, &mut remaining)
            })
            .is_err()
        );
        assert_eq!(budget.work, 0);
        assert!(MAX_SVG_BYTES - remaining < 256);
    }

    #[test]
    fn linear_gradient_markup_and_content_deduplication_stay_identical() {
        let gradient = LinearGradient {
            start: (0.0, 0.0),
            end: (10.0, 0.0),
            stops: vec![(0.0, Rgba::BLACK), (1.0, Rgba::WHITE)].into(),
            spread: Spread::Pad,
        };
        let shape = |gradient| {
            Op::Shape(Shape {
                path: crate::player::rect_path([0.0, 0.0, 10.0, 10.0]),
                fill: Some(Paint::Linear(Arc::new(gradient))),
                stroke: None,
                even_odd: false,
                clip: None,
            })
        };
        let mut drawing = Drawing {
            width: 10.0,
            height: 10.0,
            ops: vec![shape(gradient.clone())],
            omissions: Vec::new(),
        };
        let svg = write(&drawing).unwrap();
        assert_eq!(
            svg,
            concat!(
                r#"<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="0 0 10 10" preserveAspectRatio="none"><defs>"#,
                r#"<linearGradient id="p1" gradientUnits="userSpaceOnUse" x1="0" y1="0" x2="10" y2="0" spreadMethod="pad">"#,
                r##"<stop offset="0" stop-color="#000000"/><stop offset="1" stop-color="#ffffff"/></linearGradient></defs>"##,
                r##"<path d="M0 0L10 0L10 10L0 10Z" fill="url(#p1)"/></svg>"##,
            )
        );
        drawing.ops.push(shape(LinearGradient {
            stops: gradient.stops.to_vec().into(),
            ..gradient.clone()
        }));
        assert_eq!(
            write(&drawing).unwrap().matches("<linearGradient").count(),
            1
        );
        for changed in [
            LinearGradient {
                start: (1.0, 0.0),
                ..gradient.clone()
            },
            LinearGradient {
                end: (20.0, 0.0),
                ..gradient.clone()
            },
            LinearGradient {
                spread: Spread::Repeat,
                ..gradient
            },
        ] {
            drawing.ops.push(shape(changed));
        }
        assert_eq!(
            write(&drawing).unwrap().matches("<linearGradient").count(),
            4
        );
    }
}
