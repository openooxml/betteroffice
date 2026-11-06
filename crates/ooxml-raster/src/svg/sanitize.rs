//! Re-serialises a document before `usvg` sees it, so it can only read what
//! the audit read. Only allowlisted elements are written, and on them only
//! allowlisted attributes in no namespace, each value re-emitted from the
//! parse `usvg` makes of it: numbers, lengths, paints, transforms, path data
//! and points in canonical form, CSS as the declarations `simplecss` yields.
//! A value `usvg` would fail to parse is written as `x`, which it fails to
//! parse the same way, so presence and fallback are kept.

use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;
use std::str::FromStr;

use resvg::usvg::roxmltree::{Document, Node};
use svgtypes::{
    Align, AspectRatio, Color, FuncIRI, Length, LengthListParser, LengthUnit, Number, Paint,
    PaintFallback, PaintOrder, PaintOrderKind, PathParser, PathSegment, PointsParser, Transform,
    ViewBox,
};

use super::audit::{ALLOWED, SVG_NS};
use super::style::{self, StyleSheet};
use super::{
    MAX_SVG_SANITIZED_BYTES, MAX_SVG_STYLE_WORK, MAX_SVG_VALUE_BYTES, SvgRefusal, reference,
};

const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";

/// Elements whose content `usvg` never draws: written out of the document,
/// though their elements are still held to the allowlist.
const UNDRAWN: [&str; 5] = ["title", "desc", "text", "tspan", "image"];

/// How `usvg` reads a value, which decides how it is re-emitted.
#[derive(Clone, Copy)]
enum Kind {
    Length,
    Number,
    Paint,
    Color,
    StopColor,
    Link,
    Keyword,
    Transform,
    Dashes,
    PaintOrder,
    ViewBox,
    AspectRatio,
    FontSize,
    Path,
    Points,
}

/// A presentation property, as an attribute or in CSS.
fn property(name: &str) -> Option<Kind> {
    Some(match name {
        "fill" | "stroke" => Kind::Paint,
        "fill-opacity" | "stroke-opacity" | "opacity" | "stop-opacity" | "stroke-width"
        | "stroke-dashoffset" => Kind::Length,
        "stroke-miterlimit" => Kind::Number,
        "stroke-dasharray" => Kind::Dashes,
        "fill-rule" | "clip-rule" | "stroke-linecap" | "stroke-linejoin" | "display"
        | "visibility" | "overflow" | "shape-rendering" => Kind::Keyword,
        "clip-path" | "mask" | "marker-start" | "marker-mid" | "marker-end" | "filter" => {
            Kind::Link
        }
        "color" => Kind::Color,
        "stop-color" => Kind::StopColor,
        "paint-order" => Kind::PaintOrder,
        "transform" => Kind::Transform,
        "font-size" => Kind::FontSize,
        _ => return None,
    })
}

/// A property `usvg` takes from CSS: the presentation ones, the blend
/// properties it reads only there, and the `marker` shorthand.
fn declaration(name: &str) -> Option<Kind> {
    match name {
        "mix-blend-mode" | "isolation" => Some(Kind::Keyword),
        "marker" => Some(Kind::Link),
        _ => property(name),
    }
}

/// An attribute of `element`, besides `id`, `class`, `style` and `href`.
fn attribute_kind(element: &str, name: &str) -> Option<Kind> {
    Some(match name {
        "x" | "y" | "width" | "height" | "rx" | "ry" | "cx" | "cy" | "r" | "x1" | "y1" | "x2"
        | "y2" | "fx" | "fy" | "fr" | "offset" => Kind::Length,
        "viewBox" => Kind::ViewBox,
        "preserveAspectRatio" => Kind::AspectRatio,
        "gradientUnits" | "clipPathUnits" | "spreadMethod" => Kind::Keyword,
        "gradientTransform" => Kind::Transform,
        "d" if element == "path" => Kind::Path,
        "points" if matches!(element, "polyline" | "polygon") => Kind::Points,
        _ => return property(name),
    })
}

/// What each id names, for the references written into the document.
#[derive(Default)]
struct Ids<'a> {
    /// Ids of elements not written: a reference to one is refused, as `usvg`
    /// would have reached content the audit never read.
    dropped: HashSet<&'a str>,
    /// The kinds of element each written id is carried by, one bit per name
    /// in [`ALLOWED`], so a check costs the same however many carry it.
    written: HashMap<&'a str, u32>,
}

/// The bit [`Ids::written`] records a kind of element under.
fn kind_bit(name: &str) -> u32 {
    ALLOWED
        .iter()
        .position(|allowed| *allowed == name)
        .map_or(0, |index| 1 << index)
}

struct Sanitizer<'a> {
    out: String,
    ids: Ids<'a>,
    /// Style work spent tokenizing CSS text here.
    work: u64,
}

/// The elements a reference from `property` may name. `usvg` walks whatever it
/// names looking for a cycle before it checks the kind, so a reference to any
/// other element that exists is refused. `None` for a `use`, which takes any.
fn expected(property: &str) -> Option<&'static [&'static str]> {
    Some(match property {
        "fill" | "stroke" | "gradient" => &["linearGradient", "radialGradient"],
        "clip-path" => &["clipPath"],
        "mask" => &["mask"],
        "marker" | "marker-start" | "marker-mid" | "marker-end" => &["marker"],
        "filter" => &["filter"],
        _ => return None,
    })
}

/// The document as `usvg` may read it, or a refusal.
pub(super) fn sanitize(document: &Document<'_>) -> Result<String, SvgRefusal> {
    let root = document.root_element();
    if !matches!(root.tag_name().namespace(), None | Some(SVG_NS)) {
        return Err(SvgRefusal::NotSvg);
    }
    let mut sanitizer = Sanitizer {
        out: String::with_capacity(document.input_text().len().min(MAX_SVG_SANITIZED_BYTES)),
        ids: Ids::default(),
        work: 0,
    };
    sanitizer.collect_ids(root);
    sanitizer.write_tree(root, document)?;
    Ok(sanitizer.out)
}

/// What becomes of an element and everything inside it.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Fate {
    Written,
    /// Not written, but its elements are still held to the allowlist.
    Undrawn,
    /// Not written or read: metadata and foreign markup.
    Skipped,
}

/// An element's fate, from its parent's. A gradient keeps only its stops, and
/// a stop nothing: `usvg` scans a gradient's children again for every shape it
/// paints, and remembers nothing for one that fails.
fn fate(node: Node<'_, '_>, parent: Fate) -> Result<Fate, SvgRefusal> {
    if parent == Fate::Skipped || !matches!(node.tag_name().namespace(), None | Some(SVG_NS)) {
        return Ok(Fate::Skipped);
    }
    let name = node.tag_name().name();
    if name == "metadata" {
        return Ok(Fate::Skipped);
    }
    if !ALLOWED.contains(&name) {
        return Err(SvgRefusal::UnsupportedElement);
    }
    let within = node.parent_element().map(|parent| parent.tag_name().name());
    let unpainted = match within {
        Some("linearGradient" | "radialGradient") => name != "stop",
        Some("stop") => true,
        _ => false,
    };
    if parent == Fate::Undrawn || UNDRAWN.contains(&name) || name == "style" || unpainted {
        return Ok(Fate::Undrawn);
    }
    Ok(Fate::Written)
}

impl<'a> Sanitizer<'a> {
    fn collect_ids(&mut self, root: Node<'a, 'a>) {
        let mut stack = vec![(root, Fate::Written)];
        while let Some((node, parent)) = stack.pop() {
            let fate = fate(node, parent).unwrap_or(Fate::Skipped);
            for attribute in node
                .attributes()
                .filter(|attribute| attribute.name() == "id")
            {
                if fate == Fate::Written && attribute.namespace().is_none() {
                    let kinds = self.ids.written.entry(attribute.value()).or_default();
                    *kinds |= kind_bit(node.tag_name().name());
                } else {
                    self.ids.dropped.insert(attribute.value());
                }
            }
            for child in node.children().filter(|child| child.is_element()) {
                stack.push((child, fate));
            }
        }
    }

    fn write_tree(
        &mut self,
        root: Node<'a, 'a>,
        document: &'a Document<'a>,
    ) -> Result<(), SvgRefusal> {
        enum Step<'n> {
            Open(Node<'n, 'n>, Fate),
            Close(&'n str),
        }
        let mut stack = vec![Step::Open(root, Fate::Written)];
        let mut first = true;
        while let Some(step) = stack.pop() {
            let (node, parent) = match step {
                Step::Open(node, parent) => (node, parent),
                Step::Close(name) => {
                    self.out.push_str("</");
                    self.out.push_str(name);
                    self.out.push('>');
                    continue;
                }
            };
            let fate = fate(node, parent)?;
            let children: Vec<Node<'a, 'a>> =
                node.children().filter(|child| child.is_element()).collect();
            if fate != Fate::Written {
                for child in children.into_iter().rev() {
                    stack.push(Step::Open(child, fate));
                }
                continue;
            }
            let name = node.tag_name().name();
            self.out.push('<');
            self.out.push_str(name);
            if first {
                self.out.push_str(" xmlns=\"");
                self.out.push_str(SVG_NS);
                self.out.push('"');
            }
            self.write_attributes(node)?;
            if first {
                self.out.push('>');
                self.write_styles(document)?;
                stack.push(Step::Close(name));
            } else if children.is_empty() {
                self.out.push_str("/>");
            } else {
                self.out.push('>');
                stack.push(Step::Close(name));
            }
            first = false;
            for child in children.into_iter().rev() {
                stack.push(Step::Open(child, fate));
            }
            if self.out.len() > MAX_SVG_SANITIZED_BYTES {
                return Err(SvgRefusal::DocumentTooLarge);
            }
        }
        Ok(())
    }

    /// Every stylesheet `usvg` would apply, as one `<style>` of the rules
    /// `simplecss` would read from them.
    fn write_styles(&mut self, document: &'a Document<'a>) -> Result<(), SvgRefusal> {
        let sheet = StyleSheet::collect(document)?;
        let mut rules = String::new();
        let ids = &self.ids;
        let mut work = self.work;
        sheet.write(&mut rules, |text, out| {
            declarations(text, out, ids, &mut work)
        })?;
        self.work = work;
        if !rules.is_empty() {
            self.out.push_str("<style>");
            escape(&mut self.out, &rules);
            self.out.push_str("</style>");
        }
        Ok(())
    }

    fn write_attributes(&mut self, node: Node<'a, 'a>) -> Result<(), SvgRefusal> {
        let element = node.tag_name().name();
        let (mut href, mut xlink_href) = (None, None);
        let mut value = String::new();
        for attribute in node.attributes() {
            let (name, raw) = (attribute.name(), attribute.value());
            match (attribute.namespace(), name) {
                (None, _) => {}
                (Some(XLINK_NS), "href") => {
                    xlink_href = xlink_href.or(Some(raw));
                    continue;
                }
                (Some(XLINK_NS), "title") | (Some(XML_NS), "space" | "lang") => continue,
                (Some(SVG_NS | XLINK_NS | XML_NS), _) => {
                    return Err(SvgRefusal::UnsupportedElement);
                }
                _ => continue,
            }
            value.clear();
            match name {
                "href" => {
                    href = href.or(Some(raw));
                    continue;
                }
                "id" => {
                    if !safe(raw) {
                        continue;
                    }
                    value.push_str(raw);
                }
                "class" => {
                    if raw.len() > MAX_SVG_VALUE_BYTES {
                        return Err(SvgRefusal::ExpansionTooLarge);
                    }
                    value.push_str(raw);
                }
                "style" => {
                    style::screen(raw)?;
                    let mut found = Vec::new();
                    reference::css(raw, &mut found)?;
                    declarations(raw, &mut value, &self.ids, &mut self.work)?;
                    if value.is_empty() {
                        continue;
                    }
                }
                _ => {
                    let Some(kind) = attribute_kind(element, name) else {
                        continue;
                    };
                    if font_relative(element, name, raw) {
                        return Err(SvgRefusal::ExpansionTooLarge);
                    }
                    canonical(kind, name, raw, &mut value, &self.ids)?;
                }
            }
            self.write_attribute(name, &value);
        }
        let followed = matches!(element, "use" | "linearGradient" | "radialGradient");
        if let (true, Some(raw)) = (followed, href.or(xlink_href))
            && let Some(target) = reference::href(raw)?
        {
            let property = if element == "use" { "use" } else { "gradient" };
            check_link(target, property, &self.ids)?;
            value.clear();
            value.push('#');
            value.push_str(target);
            self.write_attribute("href", &value);
        }
        Ok(())
    }

    fn write_attribute(&mut self, name: &str, value: &str) {
        self.out.push(' ');
        self.out.push_str(name);
        self.out.push_str("=\"");
        escape(&mut self.out, value);
        self.out.push('"');
    }
}

/// CSS declarations as `simplecss` tokenizes them for `usvg`, re-emitted for
/// the properties `usvg` takes. The tokenizer's rescans are charged first.
fn declarations(
    text: &str,
    out: &mut String,
    ids: &Ids<'_>,
    work: &mut u64,
) -> Result<(), SvgRefusal> {
    *work = work.saturating_add(style::rescans(text.len()));
    if *work > MAX_SVG_STYLE_WORK {
        return Err(SvgRefusal::ExpansionTooLarge);
    }
    let mut value = String::new();
    for token in simplecss::DeclarationTokenizer::from(text) {
        let Some(kind) = declaration(token.name) else {
            continue;
        };
        value.clear();
        canonical(kind, token.name, token.value, &mut value, ids)?;
        if !out.is_empty() && !out.ends_with('{') {
            out.push(';');
        }
        out.push_str(token.name);
        out.push(':');
        out.push_str(&value);
        if token.important {
            out.push_str(" !important");
        }
    }
    Ok(())
}

/// Writes `raw`, the value of `name` read as `kind`, in canonical form.
fn canonical(
    kind: Kind,
    name: &str,
    raw: &str,
    out: &mut String,
    ids: &Ids<'_>,
) -> Result<(), SvgRefusal> {
    match kind {
        Kind::Path => return path(raw, out),
        Kind::Points => return points(raw, out),
        _ => {}
    }
    match kind {
        Kind::Paint => {
            reference::paint(raw)?;
        }
        Kind::Link => {
            reference::func_iri(raw)?;
        }
        _ => {}
    }
    if raw.len() > MAX_SVG_VALUE_BYTES {
        return Err(SvgRefusal::ExpansionTooLarge);
    }
    if raw == "inherit" {
        out.push_str(raw);
        return Ok(());
    }
    match kind {
        Kind::Length => match Length::from_str(raw) {
            Ok(length) => self::length(length, out),
            Err(_) => out.push('x'),
        },
        Kind::Number => match Number::from_str(raw) {
            Ok(number) => num(number.0, out),
            Err(_) => out.push('x'),
        },
        Kind::Paint => match Paint::from_str(raw) {
            Ok(Paint::None | Paint::Inherit) => out.push_str("none"),
            Ok(Paint::CurrentColor) => out.push_str("currentColor"),
            Ok(Paint::ContextFill | Paint::ContextStroke) => {
                return Err(SvgRefusal::UnsupportedStyle);
            }
            Ok(Paint::Color(value)) => color(value, out),
            Ok(Paint::FuncIRI(target, fallback)) => {
                check_link(target, name, ids)?;
                let _ = write!(out, "url(#{target})");
                match fallback {
                    None => {}
                    Some(PaintFallback::None) => out.push_str(" none"),
                    Some(PaintFallback::CurrentColor) => out.push_str(" currentColor"),
                    Some(PaintFallback::Color(value)) => {
                        out.push(' ');
                        color(value, out);
                    }
                }
            }
            Err(_) => out.push('x'),
        },
        Kind::Color => match Color::from_str(raw) {
            Ok(value) => color(value, out),
            Err(_) => out.push('x'),
        },
        Kind::StopColor => match raw {
            "currentColor" => return Err(SvgRefusal::UnsupportedStyle),
            _ => match Color::from_str(raw) {
                Ok(value) => color(value, out),
                Err(_) => out.push('x'),
            },
        },
        Kind::Link => match raw {
            "none" => out.push_str(raw),
            _ => match FuncIRI::from_str(raw) {
                Ok(FuncIRI(target)) => {
                    check_link(target, name, ids)?;
                    let _ = write!(out, "url(#{target})");
                }
                Err(_) => out.push('x'),
            },
        },
        Kind::Keyword => keyword(raw, out),
        Kind::Transform => match Transform::from_str(raw) {
            Ok(value) => {
                let row = [value.a, value.b, value.c, value.d, value.e, value.f];
                out.push_str("matrix(");
                for (index, number) in row.into_iter().enumerate() {
                    if index > 0 {
                        out.push(' ');
                    }
                    num(number, out);
                }
                out.push(')');
            }
            Err(_) => out.push('x'),
        },
        Kind::Dashes => {
            let start = out.len();
            for entry in LengthListParser::from(raw) {
                let Ok(entry) = entry else {
                    break;
                };
                if out.len() > start {
                    out.push(' ');
                }
                length(entry, out);
            }
            if out.len() == start {
                out.push_str("none");
            }
        }
        Kind::PaintOrder => match PaintOrder::from_str(raw) {
            Ok(value) => {
                for (index, kind) in value.order.into_iter().enumerate() {
                    if index > 0 {
                        out.push(' ');
                    }
                    out.push_str(match kind {
                        PaintOrderKind::Fill => "fill",
                        PaintOrderKind::Stroke => "stroke",
                        PaintOrderKind::Markers => "markers",
                    });
                }
            }
            Err(_) => out.push('x'),
        },
        Kind::ViewBox => match ViewBox::from_str(raw) {
            Ok(value) => {
                for (index, number) in [value.x, value.y, value.w, value.h].into_iter().enumerate()
                {
                    if index > 0 {
                        out.push(' ');
                    }
                    num(number, out);
                }
            }
            Err(_) => out.push('x'),
        },
        Kind::AspectRatio => match AspectRatio::from_str(raw) {
            Ok(value) => {
                if value.defer {
                    out.push_str("defer ");
                }
                out.push_str(match value.align {
                    Align::None => "none",
                    Align::XMinYMin => "xMinYMin",
                    Align::XMidYMin => "xMidYMin",
                    Align::XMaxYMin => "xMaxYMin",
                    Align::XMinYMid => "xMinYMid",
                    Align::XMidYMid => "xMidYMid",
                    Align::XMaxYMid => "xMaxYMid",
                    Align::XMinYMax => "xMinYMax",
                    Align::XMidYMax => "xMidYMax",
                    Align::XMaxYMax => "xMaxYMax",
                });
                if value.slice {
                    out.push_str(" slice");
                }
            }
            Err(_) => out.push('x'),
        },
        Kind::FontSize => match Length::from_str(raw) {
            Ok(value) => length(value, out),
            Err(_) => keyword(raw, out),
        },
        Kind::Path | Kind::Points => unreachable!(),
    }
    Ok(())
}

/// Whether a gradient coordinate or radius is in `em` or `ex`: `usvg` walks
/// and re-reads the gradient's ancestors for a font size every time a shape
/// resolves the gradient, and keeps nothing when the result is a solid colour.
fn font_relative(element: &str, name: &str, raw: &str) -> bool {
    matches!(element, "linearGradient" | "radialGradient")
        && matches!(
            name,
            "x1" | "y1" | "x2" | "y2" | "cx" | "cy" | "r" | "fx" | "fy" | "fr"
        )
        && Length::from_str(raw)
            .is_ok_and(|length| matches!(length.unit, LengthUnit::Em | LengthUnit::Ex))
}

/// Path data as `svgtypes` parses it for `usvg`, up to its first error, each
/// segment with its command letter.
fn path(raw: &str, out: &mut String) -> Result<(), SvgRefusal> {
    for segment in PathParser::from(raw) {
        let Ok(segment) = segment else {
            break;
        };
        let mut numbers = [0.0; 7];
        let (upper, count) = match segment {
            PathSegment::MoveTo { x, y, .. } => {
                numbers[..2].copy_from_slice(&[x, y]);
                ('M', 2)
            }
            PathSegment::LineTo { x, y, .. } => {
                numbers[..2].copy_from_slice(&[x, y]);
                ('L', 2)
            }
            PathSegment::HorizontalLineTo { x, .. } => {
                numbers[0] = x;
                ('H', 1)
            }
            PathSegment::VerticalLineTo { y, .. } => {
                numbers[0] = y;
                ('V', 1)
            }
            PathSegment::CurveTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
                ..
            } => {
                numbers[..6].copy_from_slice(&[x1, y1, x2, y2, x, y]);
                ('C', 6)
            }
            PathSegment::SmoothCurveTo { x2, y2, x, y, .. } => {
                numbers[..4].copy_from_slice(&[x2, y2, x, y]);
                ('S', 4)
            }
            PathSegment::Quadratic { x1, y1, x, y, .. } => {
                numbers[..4].copy_from_slice(&[x1, y1, x, y]);
                ('Q', 4)
            }
            PathSegment::SmoothQuadratic { x, y, .. } => {
                numbers[..2].copy_from_slice(&[x, y]);
                ('T', 2)
            }
            PathSegment::EllipticalArc {
                rx,
                ry,
                x_axis_rotation,
                large_arc,
                sweep,
                x,
                y,
                ..
            } => {
                let flags = [f64::from(u8::from(large_arc)), f64::from(u8::from(sweep))];
                numbers.copy_from_slice(&[rx, ry, x_axis_rotation, flags[0], flags[1], x, y]);
                ('A', 7)
            }
            PathSegment::ClosePath { .. } => ('Z', 0),
        };
        out.push(if segment.is_abs() {
            upper
        } else {
            upper.to_ascii_lowercase()
        });
        for (index, number) in numbers[..count].iter().enumerate() {
            if index > 0 {
                out.push(' ');
            }
            num(*number, out);
        }
    }
    Ok(())
}

/// Points as `svgtypes` parses them for `usvg`, pair by pair.
fn points(raw: &str, out: &mut String) -> Result<(), SvgRefusal> {
    for (index, (x, y)) in PointsParser::from(raw).enumerate() {
        if index > 0 {
            out.push(' ');
        }
        num(x, out);
        out.push(',');
        num(y, out);
    }
    Ok(())
}

/// A keyword `usvg` compares whole, written as is when it could be one.
fn keyword(raw: &str, out: &mut String) {
    let word = raw.len() <= 32
        && raw
            .bytes()
            .all(|byte| byte.is_ascii_alphabetic() || byte == b'-');
    out.push_str(if word && !raw.is_empty() { raw } else { "x" });
}

fn length(value: Length, out: &mut String) {
    num(value.number, out);
    out.push_str(match value.unit {
        LengthUnit::None => "",
        LengthUnit::Em => "em",
        LengthUnit::Ex => "ex",
        LengthUnit::Px => "px",
        LengthUnit::In => "in",
        LengthUnit::Cm => "cm",
        LengthUnit::Mm => "mm",
        LengthUnit::Pt => "pt",
        LengthUnit::Pc => "pc",
        LengthUnit::Percent => "%",
    });
}

fn color(value: Color, out: &mut String) {
    let _ = write!(
        out,
        "#{:02x}{:02x}{:02x}",
        value.red, value.green, value.blue
    );
    if value.alpha != 255 {
        let _ = write!(out, "{:02x}", value.alpha);
    }
}

/// The shortest decimal that parses back to `value` exactly, in plain or
/// exponent form. `svgtypes` reads numbers with `f64::from_str`.
fn num(value: f64, out: &mut String) {
    let magnitude = value.abs();
    if value.fract() == 0.0 && magnitude < 1e15 && !(value == 0.0 && value.is_sign_negative()) {
        let _ = write!(out, "{}", value as i64);
    } else if (1e-5..1e16).contains(&magnitude) {
        let _ = write!(out, "{value}");
    } else {
        let _ = write!(out, "{value:e}");
    }
}

/// Whether an id can be written into a reference and read back unchanged.
fn safe(id: &str) -> bool {
    !id.is_empty()
        && id.len() <= 256
        && id
            .chars()
            .all(|c| !c.is_whitespace() && !c.is_control() && !"()'\"\\;,<>&{}".contains(c))
}

/// Refuses a reference from `property` that could not be written back
/// unchanged, that names content left out of the document, or that names an
/// element of another kind than `property` converts.
fn check_link(target: &str, property: &str, ids: &Ids<'_>) -> Result<(), SvgRefusal> {
    if !safe(target) || ids.dropped.contains(target) {
        return Err(SvgRefusal::UnsupportedElement);
    }
    if let (Some(expected), Some(kinds)) = (expected(property), ids.written.get(target)) {
        let allowed = expected.iter().fold(0, |mask, name| mask | kind_bit(name));
        if kinds & !allowed != 0 {
            return Err(SvgRefusal::UnsupportedElement);
        }
    }
    Ok(())
}

fn escape(out: &mut String, value: &str) {
    for c in value.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\t' => out.push_str("&#9;"),
            '\n' => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            _ => out.push(c),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sanitized(source: &str) -> Result<String, SvgRefusal> {
        let document = Document::parse(source).expect("xml");
        sanitize(&document)
    }

    #[test]
    fn only_allowlisted_attributes_in_no_namespace_are_written() {
        let source = concat!(
            r#"<svg xmlns="http://www.w3.org/2000/svg" xmlns:f="urn:f" xmlns:xlink="http://www.w3.org/1999/xlink" "#,
            r#"viewBox=" 0 0  10 10" version="1.1" xml:space="preserve">"#,
            r##"<rect f:style="stroke-width:100000" data-x="1" width=" 5" fill="  red" class="a  b"/>"##,
            r##"<use xlink:href="#r" href="#q"/><g id="q"/><g id="r"/></svg>"##
        );
        assert_eq!(
            sanitized(source).unwrap(),
            concat!(
                r#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">"#,
                r##"<rect width="5" fill="#ff0000" class="a  b"/><use href="#q"/><g id="q"/><g id="r"/></svg>"##
            )
        );
    }

    #[test]
    fn values_are_written_as_usvg_parses_them() {
        let source = concat!(
            r##"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0,0 1 1 .5.5a1 1 0 0 1 2 2z" "##,
            r##"transform="translate(1) scale(2)" stroke="url(#g) #0f08" stroke-dasharray="1, 2 bad 3" "##,
            r##"fill="nonsense" fill-rule=" evenodd" stroke-width="1e30"/><linearGradient id="g"/></svg>"##
        );
        assert_eq!(
            sanitized(source).unwrap(),
            concat!(
                r##"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0L1 1L0.5 0.5a1 1 0 0 1 2 2z" "##,
                r##"transform="matrix(2 0 0 2 1 0)" stroke="url(#g) #00ff0088" stroke-dasharray="1 2" "##,
                r##"fill="x" fill-rule="x" stroke-width="1e30"/><linearGradient id="g"/></svg>"##
            )
        );
    }

    #[test]
    fn css_is_written_as_the_declarations_simplecss_reads() {
        let source = concat!(
            r##"<svg xmlns="http://www.w3.org/2000/svg"><style>.a, rect{ *fill : red !important; font:x }</style>"##,
            r##"<x:style xmlns:x="urn:x">g{stroke:blue}</x:style>"##,
            r##"<rect class="a" style="stroke-width:/* c */ 2;junk;opacity:1"/></svg>"##
        );
        assert_eq!(
            sanitized(source).unwrap(),
            concat!(
                r##"<svg xmlns="http://www.w3.org/2000/svg"><style>.a,rect{fill:#ff0000 !important}g{stroke:#0000ff}</style>"##,
                r##"<rect class="a" style="stroke-width:2"/></svg>"##
            )
        );
    }

    #[test]
    fn a_gradient_is_written_with_its_stops_alone() {
        let source = concat!(
            r##"<svg xmlns="http://www.w3.org/2000/svg"><linearGradient id="g"><g/><stop offset="0"><g/></stop>"##,
            r##"<rect/></linearGradient></svg>"##
        );
        assert_eq!(
            sanitized(source).unwrap(),
            r##"<svg xmlns="http://www.w3.org/2000/svg"><linearGradient id="g"><stop offset="0"></stop></linearGradient></svg>"##
        );
    }

    #[test]
    fn a_reference_to_another_kind_of_element_is_refused() {
        for body in [
            r##"<g id="g"/><rect clip-path="url(#g)"/>"##,
            r##"<g id="g"/><rect fill="url(#g) red"/>"##,
            r##"<g id="g"/><rect style="stroke:url(#g)"/>"##,
            r##"<clipPath id="c"/><rect mask="url(#c)"/>"##,
            r##"<rect id="r"/><linearGradient href="#r"/>"##,
        ] {
            let source = format!(r#"<svg xmlns="http://www.w3.org/2000/svg">{body}</svg>"#);
            assert_eq!(
                sanitized(&source).err(),
                Some(SvgRefusal::UnsupportedElement),
                "{body}"
            );
        }
        let fine = concat!(
            r##"<svg xmlns="http://www.w3.org/2000/svg"><clipPath id="c"/><linearGradient id="g"/>"##,
            r##"<rect clip-path="url(#c)" fill="url(#g)" stroke="url(#missing) red"/><use href="#c"/></svg>"##
        );
        assert!(sanitized(fine).is_ok());
    }

    #[test]
    fn a_gradient_placed_in_font_units_is_refused() {
        for gradient in [
            r#"<radialGradient id="g" gradientUnits="userSpaceOnUse" r="0em"/>"#,
            r#"<radialGradient id="g" fx=" 1ex"/>"#,
            r#"<linearGradient id="g" x2="2em"/>"#,
        ] {
            let source = format!(r#"<svg xmlns="http://www.w3.org/2000/svg">{gradient}</svg>"#);
            assert_eq!(
                sanitized(&source).err(),
                Some(SvgRefusal::ExpansionTooLarge),
                "{gradient}"
            );
        }
        let fine = concat!(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><radialGradient id="g" r="50%" cx="1in"/>"#,
            r#"<rect x="1em" width="2ex" style="x1:1em"/></svg>"#
        );
        assert!(sanitized(fine).is_ok());
    }

    #[test]
    fn a_stop_coloured_from_its_ancestors_is_refused() {
        for stop in [
            r#"<stop offset="0" stop-color="currentColor"/>"#,
            r#"<stop offset="0" style="stop-color:currentColor"/>"#,
        ] {
            let source = format!(
                r#"<svg xmlns="http://www.w3.org/2000/svg"><linearGradient id="g">{stop}</linearGradient></svg>"#
            );
            assert_eq!(
                sanitized(&source).err(),
                Some(SvgRefusal::UnsupportedStyle),
                "{stop}"
            );
        }
        let sheet = concat!(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><style>stop{stop-color:currentColor}</style>"#,
            r#"<linearGradient id="g"><stop offset="0"/></linearGradient></svg>"#
        );
        assert_eq!(sanitized(sheet).err(), Some(SvgRefusal::UnsupportedStyle));
    }

    #[test]
    fn context_paint_is_refused() {
        for paint in ["fill=\"context-fill\"", "style=\"stroke: context-stroke\""] {
            let source =
                format!(r#"<svg xmlns="http://www.w3.org/2000/svg"><rect {paint}/></svg>"#);
            assert_eq!(sanitized(&source).err(), Some(SvgRefusal::UnsupportedStyle));
        }
    }

    #[test]
    fn a_reference_into_content_left_out_is_refused() {
        for body in [
            r##"<metadata><g id="m"/></metadata><use href="#m"/>"##,
            r##"<desc><g id="d"/></desc><rect fill="url(#d)"/>"##,
            r##"<rect fill="url(#a;b)"/>"##,
        ] {
            let source = format!(r#"<svg xmlns="http://www.w3.org/2000/svg">{body}</svg>"#);
            assert_eq!(
                sanitized(&source).err(),
                Some(SvgRefusal::UnsupportedElement),
                "{body}"
            );
        }
    }
}
