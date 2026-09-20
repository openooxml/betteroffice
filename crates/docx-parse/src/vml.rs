//! Legacy VML image and watermark parsing.

use base64::Engine as _;
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::drawingml::{ShapeOutline, Transform2D};
use crate::image::{
    Image, ImageCrop, ImageEffect, ImageEffects, ImagePosition, ImageSize, ImageWrap, PositionAxis,
    placeholder_image,
};
use crate::media::{MediaMap, resolve_image_data};
use crate::relationships::RelationshipMap;
use crate::scalars::ColorValue;
use crate::xml::{XmlElement, namespaces};

const EMU_PER_PIXEL: f64 = 9_525.0;
const MAX_STYLE_BYTES: usize = 65_536;
const MAX_STYLE_DECLARATIONS: usize = 512;
const MAX_STYLE_KEY_BYTES: usize = 128;
const MAX_STYLE_VALUE_BYTES: usize = 4_096;
const MAX_VML_SHAPES: usize = 10_000;
const MAX_VML_DEPTH: usize = 256;
const MAX_SAFE_EMU: f64 = 1_000_000_000_000.0;

pub type VmlStyle = IndexMap<String, String>;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HorizontalRule {
    pub width: Option<f64>,
    pub width_percent: Option<f64>,
    pub height: f64,
    pub alignment: String,
    pub no_shade: bool,
    pub color: String,
    pub xml: String,
}

pub fn parse_horizontal_rule(picture: &XmlElement) -> Option<HorizontalRule> {
    let mut shapes = Vec::new();
    collect_horizontal_rule_shapes(picture, &IndexMap::new(), 0, &mut shapes);
    let [(shape, bindings)] = shapes.as_slice() else {
        return None;
    };
    if shape.local_name() != "rect"
        || direct_image_data(shape).is_some()
        || !matches!(
            office_attribute(shape, bindings, "hr"),
            Some("t" | "true" | "1")
        )
    {
        return None;
    }
    let style = parse_style_attr(shape.attribute(None, "style"));
    let standard = matches!(
        office_attribute(shape, bindings, "hrstd"),
        Some("t" | "true" | "1")
    );
    let width = (!standard)
        .then(|| css_length_to_px(style.get("width").map(String::as_str)))
        .flatten()
        .filter(|value| *value > 0.0)
        .and_then(pixels_to_emu);
    if width.is_some() {
        return None;
    }
    let width_percent = office_attribute(shape, bindings, "hrpct")
        .and_then(js_number)
        .filter(|value| *value > 0.0 && *value <= 1000.0)
        .map(|value| value / 10.0);
    if !standard && width_percent.is_some() {
        return None;
    }
    let height = css_length_to_px(style.get("height").map(String::as_str))
        .filter(|value| *value > 0.0)
        .and_then(pixels_to_emu)
        .unwrap_or(19_050.0);
    Some(HorizontalRule {
        width,
        width_percent,
        height,
        alignment: match (standard, office_attribute(shape, bindings, "hralign")) {
            (true, _) => "left",
            (_, Some("left")) => "left",
            (_, Some("right")) => "right",
            _ => "center",
        }
        .to_owned(),
        no_shade: matches!(
            office_attribute(shape, bindings, "hrnoshade"),
            Some("t" | "true" | "1")
        ),
        color: normalize_color(shape.attribute(None, "fillcolor")),
        xml: picture.to_raw_inline_xml(),
    })
}

fn collect_horizontal_rule_shapes<'a>(
    element: &'a XmlElement,
    inherited: &IndexMap<&'a str, &'a str>,
    depth: usize,
    output: &mut Vec<(&'a XmlElement, IndexMap<&'a str, &'a str>)>,
) {
    if depth > MAX_VML_DEPTH || output.len() >= MAX_VML_SHAPES {
        return;
    }
    let mut bindings = inherited.clone();
    for (name, value) in &element.attributes {
        if name == "xmlns" {
            bindings.insert("", value);
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            bindings.insert(prefix, value);
        }
    }
    let prefix = element.namespace_prefix().unwrap_or("");
    if (prefix == "v" && !bindings.contains_key(prefix)
        || bindings.get(prefix).copied() == Some(namespaces::V))
        && matches!(
            element.local_name(),
            "shape" | "rect" | "roundrect" | "oval"
        )
    {
        output.push((element, bindings.clone()));
    }
    for child in element.child_elements() {
        collect_horizontal_rule_shapes(child, &bindings, depth + 1, output);
        if output.len() >= MAX_VML_SHAPES {
            break;
        }
    }
}

fn office_attribute<'a>(
    element: &'a XmlElement,
    bindings: &IndexMap<&str, &str>,
    local: &str,
) -> Option<&'a str> {
    element.attributes.iter().find_map(|(name, value)| {
        if name == local {
            return Some(value.as_str());
        }
        let (prefix, name) = name.split_once(':')?;
        (name == local
            && (prefix == "o" && !bindings.contains_key(prefix)
                || bindings.get(prefix).copied() == Some(namespaces::O)))
        .then_some(value.as_str())
    })
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "lowercase")]
pub enum Watermark {
    Text {
        text: String,
        font: String,
        color: String,
        semitransparent: bool,
        layout: String,
        #[serde(rename = "fontSize", skip_serializing_if = "Option::is_none")]
        font_size: Option<f64>,
    },
    Picture {
        #[serde(rename = "relId", skip_serializing_if = "Option::is_none")]
        relationship_id: Option<String>,
        #[serde(rename = "mediaPath", skip_serializing_if = "Option::is_none")]
        media_path: Option<String>,
        #[serde(rename = "contentType", skip_serializing_if = "Option::is_none")]
        content_type: Option<String>,
        #[serde(rename = "dataUrl", skip_serializing_if = "Option::is_none")]
        data_url: Option<String>,
        scale: f64,
        washout: bool,
        #[serde(rename = "widthEmu", skip_serializing_if = "Option::is_none")]
        width_emu: Option<f64>,
        #[serde(rename = "heightEmu", skip_serializing_if = "Option::is_none")]
        height_emu: Option<f64>,
    },
}

/// Parses bounded style declarations with stable first-insertion order.
pub fn parse_style_attr(style: Option<&str>) -> VmlStyle {
    let mut output = VmlStyle::new();
    let Some(style) = style else { return output };
    let style = truncate_utf8(style, MAX_STYLE_BYTES);
    for declaration in style.split(';').take(MAX_STYLE_DECLARATIONS) {
        let Some((key, value)) = declaration.split_once(':') else {
            continue;
        };
        let key = key.trim().to_ascii_lowercase();
        if key.is_empty()
            || key.len() > MAX_STYLE_KEY_BYTES
            || matches!(key.as_str(), "__proto__" | "prototype" | "constructor")
        {
            continue;
        }
        output.insert(
            key,
            truncate_utf8(value.trim(), MAX_STYLE_VALUE_BYTES).to_owned(),
        );
    }
    output
}

pub fn is_watermark_shape(shape: &XmlElement, id_lower: &str) -> bool {
    if id_lower.contains("watermark") {
        return true;
    }
    if shape
        .attribute(None, "type")
        .unwrap_or_default()
        .contains("_t136")
    {
        return true;
    }
    let image_data = direct_image_data(shape);
    image_data.is_some_and(|image| {
        image.attribute(None, "gain").is_some()
            || image.attribute(Some("o"), "gain").is_some()
            || image.attribute(None, "blacklevel").is_some()
    })
}

/// Return the ordinary VML image contained by `w:pict`/`w:object`. Watermark
/// shapes are deliberately left to [`extract_watermark`].
pub fn parse_vml_image_content(
    picture: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
) -> Option<Image> {
    let mut shapes = Vec::new();
    collect_vml_shapes(picture, 0, &mut shapes);
    for shape in shapes.into_iter().take(MAX_VML_SHAPES) {
        let Some(image_data) = direct_image_data(shape) else {
            continue;
        };
        let relationship_id = image_data
            .attribute(Some("r"), "id")
            .or_else(|| image_data.attribute(Some("r"), "embed"))
            .or_else(|| image_data.attribute(None, "id"))
            .unwrap_or_default();
        if relationship_id.is_empty() {
            continue;
        }
        let id_lower = shape
            .attribute(None, "id")
            .unwrap_or_default()
            .to_ascii_lowercase();
        if is_watermark_shape(shape, &id_lower) {
            continue;
        }

        let resolved = resolve_image_data(relationship_id, relationships, media);
        let style = parse_style_attr(shape.attribute(None, "style"));
        let mut width = css_length_to_px(style.get("width").map(String::as_str));
        let mut height = css_length_to_px(style.get("height").map(String::as_str));
        if width.is_none() || height.is_none() {
            if let Some((intrinsic_width, intrinsic_height)) =
                intrinsic_size_px(bytes_from_image_src(resolved.src.as_deref()).as_deref())
                && intrinsic_width > 0.0
                && intrinsic_height > 0.0
            {
                match (width, height) {
                    (None, None) => {
                        width = Some(intrinsic_width);
                        height = Some(intrinsic_height);
                    }
                    (None, Some(value)) => width = Some(value * intrinsic_width / intrinsic_height),
                    (Some(value), None) => {
                        height = Some(value * intrinsic_height / intrinsic_width)
                    }
                    (Some(_), Some(_)) => {}
                }
            }
        }
        let positioned = matches!(
            style.get("position").map(String::as_str),
            Some("absolute" | "relative")
        );
        let wrap_type = shape
            .child_elements()
            .find(|child| child.local_name() == "wrap")
            .and_then(|wrap| wrap.attribute(None, "type"));
        let mut image = placeholder_image(relationship_id);
        image.size = ImageSize {
            width: width.and_then(pixels_to_emu).unwrap_or(0.0),
            height: height.and_then(pixels_to_emu).unwrap_or(0.0),
        };
        image.wrap = blank_wrap(if !positioned {
            "inline"
        } else {
            match wrap_type {
                Some("topAndBottom") => "topAndBottom",
                Some("none") => "inFront",
                Some("tight") => "tight",
                Some("through") => "through",
                _ => "square",
            }
        });
        image.src = resolved.src;
        image.mime_type = resolved.mime_type;
        image.filename = resolved.filename;
        image.title = image_data
            .attribute(Some("o"), "title")
            .filter(|value| !value.is_empty())
            .map(str::to_owned);
        image.alt = shape
            .attribute(None, "alt")
            .or_else(|| shape.attribute(Some("o"), "alt"))
            .filter(|value| !value.is_empty())
            .map(|value| truncate_utf8(value, 4_096).to_owned());

        if positioned {
            // `left`/`top` are group-child coordinates, never anchor offsets.
            let anchor_left = css_length_to_px(style.get("margin-left").map(String::as_str))
                .and_then(pixels_to_emu);
            let anchor_top = css_length_to_px(style.get("margin-top").map(String::as_str))
                .and_then(pixels_to_emu);
            let left = anchor_left.or_else(|| {
                css_length_to_px(style.get("left").map(String::as_str)).and_then(pixels_to_emu)
            });
            let top = anchor_top.or_else(|| {
                css_length_to_px(style.get("top").map(String::as_str)).and_then(pixels_to_emu)
            });
            image.position = Some(ImagePosition {
                use_simple_pos: None,
                simple_pos: None,
                relative_height: style
                    .get("z-index")
                    .and_then(|value| js_number(value))
                    .filter(safe_number),
                behind_doc: None,
                hidden: matches!(style.get("visibility").map(String::as_str), Some("hidden"))
                    .then_some(true)
                    .or_else(|| {
                        matches!(style.get("mso-hide").map(String::as_str), Some("all"))
                            .then_some(true)
                    }),
                locked: None,
                horizontal: PositionAxis {
                    relative_to: vml_horizontal_relative_to(
                        style
                            .get("mso-position-horizontal-relative")
                            .map(String::as_str),
                    )
                    .to_owned(),
                    alignment: None,
                    pos_offset: anchor_left,
                    offset: left,
                },
                vertical: PositionAxis {
                    relative_to: vml_vertical_relative_to(
                        style
                            .get("mso-position-vertical-relative")
                            .map(String::as_str),
                    )
                    .to_owned(),
                    alignment: None,
                    pos_offset: anchor_top,
                    offset: top,
                },
            });
        }

        let rotation = style
            .get("rotation")
            .and_then(|value| js_number(value))
            .filter(safe_number);
        let flip = style
            .get("flip")
            .map(|value| value.to_ascii_lowercase())
            .unwrap_or_default();
        if rotation.is_some() || flip.contains('x') || flip.contains('y') {
            image.transform = Some(Transform2D {
                rotation,
                flip_h: flip.contains('x').then_some(true),
                flip_v: flip.contains('y').then_some(true),
            });
        }

        let crop = ImageCrop {
            left: vml_fraction(image_data.attribute(None, "cropleft")),
            top: vml_fraction(image_data.attribute(None, "croptop")),
            right: vml_fraction(image_data.attribute(None, "cropright")),
            bottom: vml_fraction(image_data.attribute(None, "cropbottom")),
        };
        if crop.left.is_some()
            || crop.top.is_some()
            || crop.right.is_some()
            || crop.bottom.is_some()
        {
            image.crop = Some(crop);
        }

        let stroke_enabled = shape.attribute(None, "stroked");
        if !matches!(stroke_enabled, Some("f" | "false" | "0")) {
            let stroke_color = shape.attribute(None, "strokecolor");
            let stroke_width =
                css_length_to_px(shape.attribute(None, "strokeweight")).and_then(pixels_to_emu);
            if stroke_color.is_some() || stroke_width.is_some() {
                image.outline = Some(ShapeOutline {
                    width: stroke_width,
                    color: stroke_color.map(|value| ColorValue {
                        rgb: Some(value.trim_start_matches('#').to_owned()),
                        ..ColorValue::default()
                    }),
                    ..ShapeOutline::default()
                });
            }
        }

        let gain = image_data
            .attribute(None, "gain")
            .and_then(parse_javascript_float_prefix);
        let blacklevel = image_data
            .attribute(None, "blacklevel")
            .and_then(parse_javascript_float_prefix);
        if gain.is_some() || blacklevel.is_some() {
            image.effects = Some(ImageEffects {
                brightness: blacklevel,
                contrast: gain,
                saturation: None,
                ordered: Some(vec![ImageEffect {
                    kind: Some("brightnessContrast".to_owned()),
                    amount: blacklevel,
                    threshold: gain,
                    colors: None,
                    raw_name: None,
                }]),
            });
        }
        return Some(image);
    }
    None
}

pub fn extract_watermark(
    header_root: Option<&XmlElement>,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
) -> Option<Watermark> {
    let root = header_root?;
    let mut shapes = Vec::new();
    collect_named(root, "v:shape", 0, &mut shapes, MAX_VML_SHAPES);
    for shape in shapes {
        let id_lower = shape
            .attribute(None, "id")
            .unwrap_or_default()
            .to_ascii_lowercase();
        let text_path = shape
            .child_elements()
            .find(|child| child.local_name() == "textpath");
        let image_data = direct_image_data(shape);
        if !is_watermark_shape(shape, &id_lower) && text_path.is_none() {
            continue;
        }
        let style = parse_style_attr(shape.attribute(None, "style"));
        let rotation = style
            .get("rotation")
            .and_then(|value| parse_javascript_float_prefix(value))
            .unwrap_or(0.0);
        let diagonal = rotation.abs() > 5.0;
        if let Some(text_path) = text_path {
            let text = truncate_utf8(
                text_path.attribute(None, "string").unwrap_or_default(),
                32_767,
            )
            .to_owned();
            let text_style = parse_style_attr(text_path.attribute(None, "style"));
            let raw_font = text_style
                .get("font-family")
                .map(String::as_str)
                .unwrap_or("Calibri");
            let font = sanitize_font(raw_font);
            let fill = shape
                .child_elements()
                .find(|child| child.local_name() == "fill");
            let semitransparent = fill
                .and_then(|element| element.attribute(None, "opacity"))
                .and_then(parse_javascript_float_prefix)
                .is_some_and(|value| value < 1.0);
            return Some(Watermark::Text {
                text,
                font,
                color: normalize_color(shape.attribute(None, "fillcolor")),
                semitransparent,
                layout: if diagonal { "diagonal" } else { "horizontal" }.to_owned(),
                font_size: None,
            });
        }
        if let Some(image_data) = image_data {
            let relationship_id = image_data
                .attribute(Some("r"), "id")
                .or_else(|| image_data.attribute(Some("r"), "embed"))
                .or_else(|| image_data.attribute(None, "id"))
                .unwrap_or_default();
            let resolved = resolve_watermark_image(relationship_id, relationships, media);
            let washout = image_data.attribute(None, "gain").is_some()
                || image_data.attribute(None, "blacklevel").is_some();
            let width =
                vml_length_to_px(style.get("width").map(String::as_str)).and_then(pixels_to_emu);
            let height =
                vml_length_to_px(style.get("height").map(String::as_str)).and_then(pixels_to_emu);
            return Some(Watermark::Picture {
                relationship_id: (!relationship_id.is_empty()).then(|| relationship_id.to_owned()),
                media_path: resolved.media_path,
                content_type: resolved.content_type,
                data_url: resolved.data_url,
                scale: 1.0,
                washout,
                width_emu: width,
                height_emu: height,
            });
        }
    }
    None
}

#[derive(Default)]
struct ResolvedWatermarkImage {
    data_url: Option<String>,
    media_path: Option<String>,
    content_type: Option<String>,
}

fn resolve_watermark_image(
    relationship_id: &str,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
) -> ResolvedWatermarkImage {
    let Some(relationship) = relationships.and_then(|value| value.get(relationship_id)) else {
        return ResolvedWatermarkImage::default();
    };
    if relationship.target.is_empty() {
        return ResolvedWatermarkImage::default();
    }
    let target = &relationship.target;
    let filename = target.rsplit('/').next().unwrap_or(target);
    let trimmed = target.trim_start_matches('/');
    let candidates = [
        target.to_owned(),
        trimmed.to_owned(),
        format!("word/{trimmed}"),
        format!("word/media/{filename}"),
        format!("media/{filename}"),
    ];
    let Some(media) = media else {
        return ResolvedWatermarkImage::default();
    };
    for candidate in candidates {
        if let Some(file) = media.get_case_insensitive(&candidate) {
            return ResolvedWatermarkImage {
                data_url: Some(if file.data_url.is_empty() {
                    file.base64().to_owned()
                } else {
                    file.data_url.clone()
                }),
                media_path: Some(file.path.clone()),
                content_type: Some(file.mime_type.clone()),
            };
        }
    }
    ResolvedWatermarkImage::default()
}

fn collect_vml_shapes<'a>(element: &'a XmlElement, depth: usize, output: &mut Vec<&'a XmlElement>) {
    if depth > MAX_VML_DEPTH || output.len() >= MAX_VML_SHAPES {
        return;
    }
    if matches!(
        element.name.as_str(),
        "v:shape" | "v:rect" | "v:roundrect" | "v:oval"
    ) {
        output.push(element);
    }
    for child in element.child_elements() {
        collect_vml_shapes(child, depth + 1, output);
        if output.len() >= MAX_VML_SHAPES {
            break;
        }
    }
}

fn collect_named<'a>(
    element: &'a XmlElement,
    name: &str,
    depth: usize,
    output: &mut Vec<&'a XmlElement>,
    max: usize,
) {
    if depth > MAX_VML_DEPTH || output.len() >= max {
        return;
    }
    if element.name == name {
        output.push(element);
    }
    for child in element.child_elements() {
        collect_named(child, name, depth + 1, output, max);
    }
}

fn direct_image_data(shape: &XmlElement) -> Option<&XmlElement> {
    shape
        .child_elements()
        .find(|child| child.local_name() == "imagedata")
}

fn css_length_to_px(raw: Option<&str>) -> Option<f64> {
    parse_length(raw, false)
}

fn vml_length_to_px(raw: Option<&str>) -> Option<f64> {
    parse_length(raw, true)
}

fn parse_length(raw: Option<&str>, default_points: bool) -> Option<f64> {
    let value = raw?.trim();
    if value.is_empty() || value.bytes().any(|byte| byte.is_ascii_whitespace()) {
        return None;
    }
    let split = value
        .find(|character: char| character.is_ascii_alphabetic())
        .unwrap_or(value.len());
    let number = &value[..split];
    let unit = &value[split..];
    if number.is_empty()
        || !number
            .bytes()
            .all(|byte| byte.is_ascii_digit() || matches!(byte, b'-' | b'.'))
    {
        return None;
    }
    let number = number.parse::<f64>().ok()?;
    if !number.is_finite() {
        return None;
    }
    let pixels = match unit {
        "pt" => number / 72.0 * 96.0,
        "in" => number * 96.0,
        "cm" => number / 2.54 * 96.0,
        "mm" => number / 25.4 * 96.0,
        "pc" if !default_points => number / 6.0 * 96.0,
        "px" => number,
        "" if default_points => number * 96.0 / 72.0,
        "" => number,
        _ => return None,
    };
    safe_number(&pixels).then_some(pixels)
}

fn vml_fraction(raw: Option<&str>) -> Option<f64> {
    let raw = raw?;
    if raw.is_empty() {
        return None;
    }
    let value = raw.trim();
    let fixed = value.ends_with('f') || value.ends_with('F');
    let stripped = if fixed {
        &value[..value.len() - 1]
    } else {
        value
    };
    let parsed = parse_javascript_float_prefix(stripped)?;
    Some((if fixed { parsed / 65_536.0 } else { parsed }).clamp(0.0, 1.0))
}

fn vml_horizontal_relative_to(raw: Option<&str>) -> &'static str {
    match raw {
        Some("page") => "page",
        Some("margin") => "margin",
        Some("text" | "character") => "character",
        _ => "column",
    }
}

fn vml_vertical_relative_to(raw: Option<&str>) -> &'static str {
    match raw {
        Some("page") => "page",
        Some("margin") => "margin",
        Some("line") => "line",
        _ => "paragraph",
    }
}

fn bytes_from_image_src(source: Option<&str>) -> Option<Vec<u8>> {
    let source = source?;
    let encoded = if let Some(source) = source.strip_prefix("data:") {
        source.split_once(',')?.1
    } else if !source.contains(':')
        && source
            .bytes()
            .next()
            .is_some_and(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'/'))
    {
        source
    } else {
        return None;
    };
    base64::engine::general_purpose::STANDARD
        .decode(encoded)
        .ok()
}

fn intrinsic_size_px(bytes: Option<&[u8]>) -> Option<(f64, f64)> {
    let bytes = bytes?;
    if bytes.len() < 24 {
        return None;
    }
    if bytes.starts_with(&[0x89, b'P', b'N', b'G']) {
        let width = u32::from_be_bytes(bytes[16..20].try_into().ok()?);
        let height = u32::from_be_bytes(bytes[20..24].try_into().ok()?);
        return Some((width as f64, height as f64));
    }
    if bytes.starts_with(b"GIF") {
        let width = u16::from_le_bytes(bytes[6..8].try_into().ok()?);
        let height = u16::from_le_bytes(bytes[8..10].try_into().ok()?);
        return Some((width as f64, height as f64));
    }
    if bytes.starts_with(b"BM") && bytes.len() >= 26 {
        let width = i32::from_le_bytes(bytes[18..22].try_into().ok()?);
        let height = i32::from_le_bytes(bytes[22..26].try_into().ok()?).abs();
        return Some((width as f64, height as f64));
    }
    if bytes.starts_with(&[0xff, 0xd8]) {
        let mut offset = 2;
        while offset + 9 < bytes.len() {
            if bytes[offset] != 0xff {
                offset += 1;
                continue;
            }
            let marker = bytes[offset + 1];
            if (0xc0..=0xcf).contains(&marker) && !matches!(marker, 0xc4 | 0xc8 | 0xcc) {
                let height = u16::from_be_bytes([bytes[offset + 5], bytes[offset + 6]]);
                let width = u16::from_be_bytes([bytes[offset + 7], bytes[offset + 8]]);
                return Some((width as f64, height as f64));
            }
            let length = u16::from_be_bytes([bytes[offset + 2], bytes[offset + 3]]) as usize;
            if length < 2 {
                break;
            }
            offset = offset.saturating_add(2 + length);
        }
    }
    None
}

fn normalize_color(raw: Option<&str>) -> String {
    let Some(raw) = raw else {
        return "#C0C0C0".to_owned();
    };
    let value = raw.trim();
    let hex = value.strip_prefix('#').unwrap_or(value);
    if matches!(hex.len(), 3 | 6) && hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return if value.starts_with('#') {
            value.to_owned()
        } else if hex.len() == 6 {
            format!("#{hex}")
        } else {
            "#C0C0C0".to_owned()
        };
    }
    let lower = value.to_ascii_lowercase();
    if matches!(
        lower.as_str(),
        "black"
            | "white"
            | "silver"
            | "gray"
            | "red"
            | "maroon"
            | "yellow"
            | "olive"
            | "lime"
            | "green"
            | "aqua"
            | "teal"
            | "blue"
            | "navy"
            | "fuchsia"
            | "purple"
    ) {
        lower
    } else {
        "#C0C0C0".to_owned()
    }
}

fn sanitize_font(raw: &str) -> String {
    let without_quotes = raw
        .replace(['\"', '\''], "")
        .trim_start_matches("&quot;")
        .trim_end_matches("&quot;")
        .chars()
        .filter(|character| !character.is_ascii_control())
        .collect::<String>();
    let trimmed = truncate_utf8(without_quotes.trim(), 255);
    if trimmed.is_empty() {
        "Calibri".to_owned()
    } else {
        trimmed.to_owned()
    }
}

fn parse_javascript_float_prefix(value: &str) -> Option<f64> {
    let value = value.trim_start();
    let bytes = value.as_bytes();
    let mut index = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let integer_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    let mut has_digits = index > integer_start;
    if bytes.get(index) == Some(&b'.') {
        index += 1;
        let fraction_start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        has_digits |= index > fraction_start;
    }
    if !has_digits {
        return None;
    }
    if matches!(bytes.get(index), Some(b'e' | b'E')) {
        let exponent = index;
        index += 1;
        if matches!(bytes.get(index), Some(b'+' | b'-')) {
            index += 1;
        }
        let digits = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        if index == digits {
            index = exponent;
        }
    }
    value[..index]
        .parse::<f64>()
        .ok()
        .filter(|number| number.is_finite())
}

fn js_number(value: &str) -> Option<f64> {
    let value = value.trim();
    if value.is_empty() {
        return Some(0.0);
    }
    value
        .parse::<f64>()
        .ok()
        .filter(|number| number.is_finite())
}

fn pixels_to_emu(pixels: f64) -> Option<f64> {
    let value = js_round(pixels * EMU_PER_PIXEL);
    safe_number(&value).then_some(value)
}

fn js_round(value: f64) -> f64 {
    (value + 0.5).floor()
}

fn safe_number(value: &f64) -> bool {
    value.is_finite() && value.abs() <= MAX_SAFE_EMU
}

fn blank_wrap(kind: &str) -> ImageWrap {
    ImageWrap {
        wrap_type: kind.to_owned(),
        wrap_text: None,
        dist_t: None,
        dist_b: None,
        dist_l: None,
        dist_r: None,
        polygon: None,
    }
}

fn truncate_utf8(value: &str, max_bytes: usize) -> &str {
    if value.len() <= max_bytes {
        return value;
    }
    let mut end = max_bytes;
    while !value.is_char_boundary(end) {
        end -= 1;
    }
    &value[..end]
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::media::build_media_map;
    use crate::relationships::Relationship;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        parse_xml(xml.as_bytes(), "vml", &mut budget)
            .unwrap()
            .root()
            .unwrap()
            .clone()
    }

    #[test]
    fn horizontal_rule_retains_relative_width_and_raw_vml() {
        let xml = r##"<w:pict xmlns:w="w" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect style="width:0pt;height:1.5pt" o:hr="t" o:hrstd="t" o:hralign="center" fillcolor="#A0A0A0" stroked="f"/></w:pict>"##;
        let rule = parse_horizontal_rule(&root(xml)).unwrap();
        assert_eq!(rule.width, None);
        assert_eq!(rule.width_percent, None);
        assert_eq!(rule.height, 19_050.0);
        assert_eq!(rule.alignment, "left");
        assert!(!rule.no_shade);
        assert_eq!(rule.color, "#A0A0A0");
        let run = crate::inline::Run {
            node_type: crate::inline::RunType::Run,
            formatting: None,
            property_changes: None,
            content: vec![crate::inline::RunContent::HorizontalRule {
                rule: Box::new(rule.clone()),
            }],
        };
        let mut context =
            crate::serializer::SerializerContext::new(&crate::serializer::SerializerDeterminism {
                seed: "0".repeat(64),
                now: "2000-01-01T00:00:00.000Z".to_owned(),
            })
            .unwrap();
        let saved = crate::serializer::serialize_run(&run, &mut context).unwrap();
        assert_eq!(parse_horizontal_rule(&root(&saved)).unwrap().width, None);
        assert!(saved.contains(&rule.xml));
        assert_eq!(parse_horizontal_rule(&root(&rule.xml)), Some(rule));
    }

    #[test]
    fn horizontal_rule_distinguishes_rectangles_and_authored_sizing() {
        let xml = r#"<w:pict xmlns:w="w" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect style="width:72pt;height:3pt" o:hr="0"/></w:pict>"#;
        assert!(parse_horizontal_rule(&root(xml)).is_none());
        let xml = xml.replace(
            "o:hr=\"0\"",
            "o:hr=\"1\" o:hrpct=\"50\" o:hralign=\"right\" o:hrnoshade=\"true\"",
        );
        assert!(parse_horizontal_rule(&root(&xml)).is_none());
        let xml = xml
            .replace("width:72pt", "width:0pt")
            .replace("o:hrpct=\"50\" ", "");
        let rule = parse_horizontal_rule(&root(&xml)).unwrap();
        assert_eq!(rule.width, None);
        assert_eq!(rule.width_percent, None);
        assert_eq!(rule.height, 38_100.0);
        assert_eq!(rule.alignment, "right");
        assert!(rule.no_shade);
        let xml = xml.replace("o:hr=\"1\"", "o:hr=\"1\" o:hrstd=\"t\" o:hrpct=\"50\"");
        assert_eq!(
            parse_horizontal_rule(&root(&xml)).unwrap().width_percent,
            Some(5.0)
        );
    }

    #[test]
    fn horizontal_rule_resolves_inherited_aliases_and_preserves_replay_bindings() {
        let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wx="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:legacy="urn:schemas-microsoft-com:vml" xmlns:office="urn:schemas-microsoft-com:office:office" xmlns:ext="urn:extension" xmlns:unused="urn:unused" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body><w:p><w:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wx:pict mc:Ignorable="unused"><legacy:rect style="width:0pt;height:1.5pt" office:hr="t" office:hrstd="t" office:hrpct="500" office:hrnoshade="t" ext:value="kept" a:value="kept" pic:value="kept"/></wx:pict></w:r></w:p></w:body></w:document>"#;
        let parsed = root(xml);
        let picture = parsed.find_deep("w", "pict").unwrap();
        let rule = parse_horizontal_rule(picture).unwrap();
        assert_eq!(rule.width_percent, Some(50.0));
        assert!(rule.no_shade);
        for prefix in ["wx", "legacy", "office", "ext", "unused", "a", "pic"] {
            assert!(rule.xml.contains(&format!("xmlns:{prefix}=")));
        }
        assert_eq!(parse_horizontal_rule(&root(&rule.xml)), Some(rule.clone()));
        let run = crate::inline::Run {
            node_type: crate::inline::RunType::Run,
            formatting: None,
            property_changes: None,
            content: vec![crate::inline::RunContent::HorizontalRule {
                rule: Box::new(rule.clone()),
            }],
        };
        let mut context =
            crate::serializer::SerializerContext::new(&crate::serializer::SerializerDeterminism {
                seed: "0".repeat(64),
                now: "2000-01-01T00:00:00.000Z".to_owned(),
            })
            .unwrap();
        let saved = crate::serializer::serialize_run(&run, &mut context).unwrap();
        assert_eq!(
            parse_horizontal_rule(&root(&saved)).unwrap().width_percent,
            Some(50.0)
        );
        let paragraph = parsed.find_deep("w", "p").unwrap();
        assert!(paragraph.attributes.is_empty());
    }

    #[test]
    fn horizontal_rule_keeps_canonical_picture_attributes_unchanged() {
        let picture = r#"<w:pict><v:rect o:hr="t" o:hrstd="t"/></w:pict>"#;
        let xml = format!(
            r#"<w:document xmlns:w="{}" xmlns:v="{}" xmlns:o="{}"><w:p><w:r>{picture}</w:r></w:p></w:document>"#,
            namespaces::W,
            namespaces::V,
            namespaces::O
        );
        let parsed = root(&xml);
        let parsed_picture = parsed.find_deep("w", "pict").unwrap();
        assert!(parsed_picture.attributes.is_empty());
        assert_eq!(parsed_picture.to_raw_inline_xml(), picture);
    }

    #[test]
    fn horizontal_rule_rejects_foreign_namespaces_and_shadowed_bindings() {
        for (vml, office, rect, hr) in [
            (namespaces::V, "urn:foreign", "legacy", "office"),
            ("urn:foreign", namespaces::O, "legacy", "office"),
            (namespaces::V, "urn:foreign", "v", "o"),
            ("urn:foreign", namespaces::O, "v", "o"),
        ] {
            let xml = format!(
                r#"<w:document xmlns:w="w" xmlns:{rect}="{vml}" xmlns:{hr}="{office}"><w:p><w:r><w:pict><{rect}:rect {hr}:hr="t"/></w:pict></w:r></w:p></w:document>"#
            );
            assert!(parse_horizontal_rule(root(&xml).find_deep("w", "pict").unwrap()).is_none());
        }
        let xml = format!(
            r#"<w:document xmlns:w="w" xmlns:legacy="{}" xmlns:office="{}"><w:p xmlns:legacy="urn:foreign"><w:r><w:pict><legacy:rect office:hr="t"/></w:pict></w:r></w:p></w:document>"#,
            namespaces::V,
            namespaces::O
        );
        assert!(parse_horizontal_rule(root(&xml).find_deep("w", "pict").unwrap()).is_none());
        let xml = format!(
            r#"<w:pict xmlns:w="w" xmlns:legacy="{}" xmlns:office="{}"><legacy:rect office:hr="t"><legacy:imagedata/></legacy:rect></w:pict>"#,
            namespaces::V,
            namespaces::O
        );
        assert!(parse_horizontal_rule(&root(&xml)).is_none());
    }

    #[test]
    fn horizontal_rule_keeps_mixed_vml_picture_fallback() {
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="r"><v:rect o:hr="t"/><v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rId7"/></v:shape></w:pict>"#,
        );
        assert!(parse_horizontal_rule(&picture).is_none());
        let image = parse_vml_image_content(&picture, None, None).unwrap();
        assert_eq!(image.relationship_id, "rId7");
        assert_eq!(image.size.width, 914_400.0);
    }

    #[test]
    fn parses_positioned_vml_image_and_mirrors_the_legacy_offset() {
        let media = build_media_map(&[("word/media/logo.png".to_owned(), vec![0; 24])]);
        let relationships = RelationshipMap::from([(
            "rId7".to_owned(),
            Relationship {
                id: "rId7".to_owned(),
                relationship_type: "image".to_owned(),
                target: "media/logo.png".to_owned(),
                target_mode: None,
            },
        )]);
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r" xmlns:o="o"><v:shape id="Picture 1" style="position:absolute;width:72pt;height:1in;margin-left:2px;rotation:45;flip:x" alt="logo"><v:imagedata r:id="rId7" o:title="Logo" cropleft="32768f"/></v:shape></w:pict>"#,
        );
        let image = parse_vml_image_content(&picture, Some(&relationships), Some(&media)).unwrap();
        assert_eq!(image.size.width, 914_400.0);
        assert_eq!(image.size.height, 914_400.0);
        let horizontal = image.position.unwrap().horizontal;
        assert_eq!(horizontal.pos_offset, Some(19_050.0));
        assert_eq!(horizontal.offset, Some(19_050.0));
        assert_eq!(image.crop.unwrap().left, Some(0.5));
        assert_eq!(image.transform.unwrap().rotation, Some(45.0));
    }

    #[test]
    fn header_vml_shape_keeps_its_word_offsets() {
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r" xmlns:o="o"><v:shape id="_x0000_s1025" style="position:absolute;margin-left:-10.55pt;margin-top:42.75pt;width:505pt;height:133pt;mso-position-horizontal-relative:text;mso-position-vertical-relative:text"><v:imagedata r:id="rId2"/></v:shape></w:pict>"#,
        );
        let position = parse_vml_image_content(&picture, None, None)
            .unwrap()
            .position
            .unwrap();
        assert_eq!(position.horizontal.relative_to, "character");
        assert_eq!(position.horizontal.pos_offset, Some(-133_985.0));
        assert_eq!(position.vertical.relative_to, "paragraph");
        assert_eq!(position.vertical.pos_offset, Some(542_925.0));
    }

    #[test]
    fn grouped_vml_child_coordinates_are_not_anchor_offsets() {
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r" xmlns:o="o"><v:group style="position:absolute;margin-left:36pt;margin-top:770.6pt;width:540pt;height:1.5pt"><v:shape id="Picture 41" style="position:absolute;left:2476;top:4038;width:6952;height:3368"><v:imagedata r:id="rId2"/></v:shape></v:group></w:pict>"#,
        );
        let position = parse_vml_image_content(&picture, None, None)
            .unwrap()
            .position
            .unwrap();
        assert_eq!(position.horizontal.pos_offset, None);
        assert_eq!(position.vertical.pos_offset, None);
        assert_eq!(position.horizontal.offset, Some(23_583_900.0));
    }

    #[test]
    fn text_watermark_is_claimed_and_sanitized() {
        let header = root(
            r##"<w:hdr xmlns:w="w" xmlns:v="v"><v:shape id="PowerPlusWaterMarkObject1" type="#_x0000_t136" style="rotation:315" fillcolor="C0C0C0"><v:fill opacity=".5"/><v:textpath string="CONFIDENTIAL" style="font-family:'Calibri'"/></v:shape></w:hdr>"##,
        );
        assert_eq!(
            extract_watermark(Some(&header), None, None),
            Some(Watermark::Text {
                text: "CONFIDENTIAL".to_owned(),
                font: "Calibri".to_owned(),
                color: "#C0C0C0".to_owned(),
                semitransparent: true,
                layout: "diagonal".to_owned(),
                font_size: None,
            })
        );
        assert!(parse_vml_image_content(&header, None, None).is_none());
    }

    #[test]
    fn style_and_huge_dimensions_are_bounded() {
        let style = parse_style_attr(Some("__proto__:x;width:1e999px;A:1;A:2"));
        assert!(!style.contains_key("__proto__"));
        assert_eq!(style.get("a").map(String::as_str), Some("2"));
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r"><v:shape style="width:999999999999px;height:2px"><v:imagedata r:id="x"/></v:shape></w:pict>"#,
        );
        let image = parse_vml_image_content(&picture, None, None).unwrap();
        assert_eq!(image.size.width, 0.0);
        assert_eq!(image.size.height, 19_050.0);
    }
}
