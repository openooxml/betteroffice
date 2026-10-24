//! Legacy VML image and watermark parsing.

use base64::Engine as _;
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::drawingml::{ShapeOutline, Transform2D};
use crate::image::{
    Image, ImageCrop, ImageEffect, ImageEffects, ImagePosition, ImageSize, ImageWrap, PositionAxis,
    placeholder_image,
};
use crate::media::{MediaMap, media_token_index, resolve_image_data};
use crate::relationships::RelationshipMap;
use crate::scalars::ColorValue;
use crate::xml::{ParseBudget, XmlElement, namespaces};

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
    vml_image_content(picture, relationships, media, None)
}

/// [`parse_vml_image_content`] reading `media:{n}` sources through `budget`'s
/// media table, where a failed read fails the parse.
pub(crate) fn vml_image_content(
    picture: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
    budget: Option<&mut ParseBudget<'_>>,
) -> Option<Image> {
    vml_images(picture, relationships, media, budget, 1)
        .into_iter()
        .next()
}

pub(crate) fn vml_image_contents(
    picture: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
    budget: Option<&mut ParseBudget<'_>>,
) -> Vec<Image> {
    vml_images(picture, relationships, media, budget, MAX_VML_SHAPES)
}

pub(crate) fn vml_story_image_contents(
    picture: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
    budget: Option<&mut ParseBudget<'_>>,
) -> Vec<Image> {
    vml_shape_images(picture, false, relationships, media, budget, MAX_VML_SHAPES)
}

fn vml_images(
    picture: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
    mut budget: Option<&mut ParseBudget<'_>>,
    limit: usize,
) -> Vec<Image> {
    let images = vml_shape_images(
        picture,
        false,
        relationships,
        media,
        budget.as_deref_mut(),
        limit,
    );
    if !images.is_empty() {
        return images;
    }
    vml_shape_images(picture, true, relationships, media, budget, 1)
}

fn vml_shape_images(
    picture: &XmlElement,
    text_boxes: bool,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
    mut budget: Option<&mut ParseBudget<'_>>,
    limit: usize,
) -> Vec<Image> {
    let mut shapes = Vec::new();
    collect_vml_shapes(picture, None, text_boxes, 0, &mut shapes);
    let mut bound = Vec::new();
    collect_horizontal_rule_shapes(picture, &IndexMap::new(), 0, &mut bound);
    let mut images = Vec::new();
    for (shape, group) in shapes {
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
        if let Some(group) = group {
            width = group_length_to_px(style.get("width").map(String::as_str), group.scale_x);
            height = group_length_to_px(style.get("height").map(String::as_str), group.scale_y);
        }
        if width.is_none() || height.is_none() {
            if let Some((intrinsic_width, intrinsic_height)) = intrinsic_size_px(
                image_src_bytes(resolved.src.as_deref(), budget.as_deref_mut()).as_deref(),
            ) && intrinsic_width > 0.0
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
        let group_style = group.map(|group| parse_style_attr(group.root.attribute(None, "style")));
        let placement_style = group_style.as_ref().unwrap_or(&style);
        let positioned = matches!(
            placement_style.get("position").map(String::as_str),
            Some("absolute" | "relative")
        );
        let wrap_type = group
            .map_or(shape, |group| group.root)
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
                None if group.is_some() => "inFront",
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
        image.horizontal_rule = (!positioned
            && group.is_none()
            && matches!(
                bound
                    .iter()
                    .find(|(candidate, _)| std::ptr::eq(*candidate, shape))
                    .and_then(|(_, bindings)| office_attribute(shape, bindings, "hr")),
                Some("t" | "true" | "1")
            ))
        .then_some(true);

        if positioned {
            let anchor_left = css_length_to_px(style.get("margin-left").map(String::as_str))
                .and_then(pixels_to_emu);
            let anchor_top = css_length_to_px(style.get("margin-top").map(String::as_str))
                .and_then(pixels_to_emu);
            let mut left = anchor_left.or_else(|| {
                css_length_to_px(style.get("left").map(String::as_str)).and_then(pixels_to_emu)
            });
            let mut top = anchor_top.or_else(|| {
                css_length_to_px(style.get("top").map(String::as_str)).and_then(pixels_to_emu)
            });
            if let Some(group) = group {
                left = pixels_to_emu(
                    group.offset_x
                        + group_length_to_px(
                            style
                                .get("left")
                                .or_else(|| style.get("margin-left"))
                                .map(String::as_str),
                            group.scale_x,
                        )
                        .unwrap_or(0.0),
                );
                top = pixels_to_emu(
                    group.offset_y
                        + group_length_to_px(
                            style
                                .get("top")
                                .or_else(|| style.get("margin-top"))
                                .map(String::as_str),
                            group.scale_y,
                        )
                        .unwrap_or(0.0),
                );
            }
            let z_index = placement_style
                .get("z-index")
                .or_else(|| style.get("z-index"))
                .and_then(|value| js_number(value))
                .filter(safe_number);
            let behind_doc = z_index.map(|value| value < 0.0);
            if behind_doc == Some(true) && matches!(wrap_type, None | Some("none")) {
                image.wrap = blank_wrap("behind");
            }
            image.position = Some(ImagePosition {
                use_simple_pos: None,
                simple_pos: None,
                relative_height: z_index.map(|value| value.max(0.0)),
                behind_doc,
                hidden: (matches!(style.get("visibility").map(String::as_str), Some("hidden"))
                    || matches!(
                        placement_style.get("visibility").map(String::as_str),
                        Some("hidden")
                    ))
                .then_some(true)
                .or_else(|| {
                    (matches!(style.get("mso-hide").map(String::as_str), Some("all"))
                        || matches!(
                            placement_style.get("mso-hide").map(String::as_str),
                            Some("all")
                        ))
                    .then_some(true)
                }),
                locked: None,
                horizontal: PositionAxis {
                    relative_to: vml_horizontal_relative_to(
                        placement_style
                            .get("mso-position-horizontal-relative")
                            .map(String::as_str),
                    )
                    .to_owned(),
                    alignment: None,
                    pos_offset: if group.is_some() { left } else { anchor_left },
                    offset: left,
                },
                vertical: PositionAxis {
                    relative_to: vml_vertical_relative_to(
                        placement_style
                            .get("mso-position-vertical-relative")
                            .map(String::as_str),
                    )
                    .to_owned(),
                    alignment: None,
                    pos_offset: if group.is_some() { top } else { anchor_top },
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
        images.push(image);
        if images.len() >= limit {
            break;
        }
    }
    images
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
        if let Some(file) = media
            .iter()
            .find(|(key, _)| key.eq_ignore_ascii_case(&candidate))
            .map(|(_, file)| file)
        {
            return ResolvedWatermarkImage {
                data_url: Some(if file.data_url.is_empty() {
                    file.base64.clone()
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

#[derive(Clone, Copy)]
struct VmlGroupTransform<'a> {
    root: &'a XmlElement,
    scale_x: f64,
    scale_y: f64,
    offset_x: f64,
    offset_y: f64,
}

fn vml_coordinate_pair(raw: Option<&str>, default: (f64, f64)) -> (f64, f64) {
    raw.and_then(|raw| {
        let mut values = raw
            .split(|character: char| character == ',' || character.is_whitespace())
            .filter(|value| !value.is_empty());
        let x = js_number(values.next()?).filter(safe_number)?;
        let y = js_number(values.next()?).filter(safe_number)?;
        values.next().is_none().then_some((x, y))
    })
    .unwrap_or(default)
}

fn group_length_to_px(raw: Option<&str>, scale: f64) -> Option<f64> {
    let raw = raw?.trim();
    if raw.is_empty() {
        return None;
    }
    let coordinate = js_number(raw).or_else(|| {
        ["pt", "px", "in", "cm", "mm", "pc"]
            .into_iter()
            .find_map(|unit| {
                raw.strip_suffix(unit)
                    .filter(|value| !value.trim().is_empty())
                    .and_then(js_number)
            })
    })?;
    let pixels = coordinate * scale;
    safe_number(&pixels).then_some(pixels)
}

fn group_transform<'a>(
    element: &'a XmlElement,
    parent: Option<VmlGroupTransform<'a>>,
) -> VmlGroupTransform<'a> {
    let style = parse_style_attr(element.attribute(None, "style"));
    let (origin_x, origin_y) =
        vml_coordinate_pair(element.attribute(None, "coordorigin"), (0.0, 0.0));
    let (coord_width, coord_height) =
        vml_coordinate_pair(element.attribute(None, "coordsize"), (1000.0, 1000.0));
    let coord_width = if coord_width > 0.0 {
        coord_width
    } else {
        1000.0
    };
    let coord_height = if coord_height > 0.0 {
        coord_height
    } else {
        1000.0
    };
    let parent_scale_x = parent.map_or(1.0, |parent| parent.scale_x);
    let parent_scale_y = parent.map_or(1.0, |parent| parent.scale_y);
    let length = |raw, scale| {
        if parent.is_some() {
            group_length_to_px(raw, scale)
        } else {
            css_length_to_px(raw)
        }
    };
    let width = length(style.get("width").map(String::as_str), parent_scale_x)
        .unwrap_or(coord_width * parent_scale_x);
    let height = length(style.get("height").map(String::as_str), parent_scale_y)
        .unwrap_or(coord_height * parent_scale_y);
    let offset = |coordinate, margin| {
        if parent.is_some() {
            style.get(coordinate).or_else(|| style.get(margin))
        } else {
            style.get(margin).or_else(|| style.get(coordinate))
        }
        .map(String::as_str)
    };
    let left = length(offset("left", "margin-left"), parent_scale_x).unwrap_or(0.0);
    let top = length(offset("top", "margin-top"), parent_scale_y).unwrap_or(0.0);
    let scale_x = width / coord_width;
    let scale_y = height / coord_height;
    VmlGroupTransform {
        root: parent.map_or(element, |parent| parent.root),
        scale_x,
        scale_y,
        offset_x: parent.map_or(0.0, |parent| parent.offset_x) + left - origin_x * scale_x,
        offset_y: parent.map_or(0.0, |parent| parent.offset_y) + top - origin_y * scale_y,
    }
}

fn collect_vml_shapes<'a>(
    element: &'a XmlElement,
    group: Option<VmlGroupTransform<'a>>,
    text_boxes: bool,
    depth: usize,
    output: &mut Vec<(&'a XmlElement, Option<VmlGroupTransform<'a>>)>,
) {
    let text_box = matches!(element.name.as_str(), "v:textbox" | "w:txbxContent");
    if depth > MAX_VML_DEPTH || output.len() >= MAX_VML_SHAPES || (text_box && !text_boxes) {
        return;
    }
    let group = if element.name == "v:group" {
        Some(group_transform(element, group))
    } else if text_box {
        None
    } else {
        group
    };
    if matches!(
        element.name.as_str(),
        "v:shape" | "v:rect" | "v:roundrect" | "v:oval"
    ) {
        output.push((element, group));
    }
    let mut children: Vec<_> = element
        .child_elements()
        .map(|child| {
            let z_index = if element.name == "v:group" {
                parse_style_attr(child.attribute(None, "style"))
                    .get("z-index")
                    .and_then(|value| js_number(value))
                    .filter(safe_number)
                    .unwrap_or(0.0)
            } else {
                0.0
            };
            (z_index, child)
        })
        .collect();
    if element.name == "v:group" {
        children.sort_by(|(left, _), (right, _)| {
            left.partial_cmp(right).unwrap_or(std::cmp::Ordering::Equal)
        });
    }
    for (_, child) in children {
        collect_vml_shapes(child, group, text_boxes, depth + 1, output);
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

/// Whether `shape`'s style positions it from the page or margin.
pub(crate) fn placed_off_the_text(shape: &XmlElement) -> bool {
    style_placed_off_the_text(shape.attribute(None, "style"))
}

pub(crate) fn style_placed_off_the_text(style: Option<&str>) -> bool {
    let Some(style) = style else { return false };
    let mut position = None;
    let mut vertical = None;
    for declaration in truncate_utf8(style, MAX_STYLE_BYTES)
        .split(';')
        .take(MAX_STYLE_DECLARATIONS)
    {
        let Some((key, value)) = declaration.split_once(':') else {
            continue;
        };
        let key = key.trim();
        let value = truncate_utf8(value.trim(), MAX_STYLE_VALUE_BYTES);
        if key.eq_ignore_ascii_case("position") {
            position = Some(value);
        } else if key.eq_ignore_ascii_case("mso-position-vertical-relative") {
            vertical = Some(value);
        }
    }
    matches!(position, Some("absolute" | "relative"))
        && matches!(vml_vertical_relative_to(vertical), "page" | "margin")
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

fn image_src_bytes(source: Option<&str>, budget: Option<&mut ParseBudget<'_>>) -> Option<Vec<u8>> {
    let index = source.and_then(media_token_index);
    if let (Some(budget), Some(index)) = (budget, index)
        && let Some(table) = budget.media_table()
    {
        return match table.bytes(index) {
            Ok(bytes) => Some(bytes.into_owned()),
            Err(error) => {
                budget.fail_media_read(error);
                None
            }
        };
    }
    bytes_from_image_src(source)
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
    fn grouped_vml_child_coordinates_use_the_group_anchor_and_default_coordsize() {
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r" xmlns:o="o"><v:group style="position:absolute;margin-left:36pt;margin-top:770.6pt;width:540pt;height:1.5pt"><v:shape id="Picture 41" style="position:absolute;left:2476;top:4038;width:6952;height:3368"><v:imagedata r:id="rId2"/></v:shape></v:group></w:pict>"#,
        );
        let position = parse_vml_image_content(&picture, None, None)
            .unwrap()
            .position
            .unwrap();
        assert_eq!(position.horizontal.pos_offset, Some(17_437_608.0));
        assert_eq!(position.vertical.pos_offset, Some(9_863_544.0));
        assert_eq!(position.horizontal.offset, position.horizontal.pos_offset);
        assert_eq!(position.vertical.offset, position.vertical.pos_offset);
    }

    #[test]
    fn nested_vml_groups_compose_origins_scales_and_offsets() {
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r"><v:group coordorigin="10,20" coordsize="1000,500" style="position:absolute;margin-left:8px;margin-top:12px;width:200px;height:100px;mso-position-horizontal-relative:page;mso-position-vertical-relative:page"><v:group coordorigin="50,60" coordsize="200,100" style="left:110;top:70;width:400;height:200"><v:shape style="left:75;top:85;width:50;height:25"><v:imagedata r:id="rId1"/></v:shape></v:group><v:shape style="left:10;top:20;width:100;height:50"><v:imagedata r:id="rId2"/></v:shape></v:group></w:pict>"#,
        );
        let images = vml_image_contents(&picture, None, None, None);
        assert_eq!(images.len(), 2);
        for (image, (id, x, y, width, height)) in images.iter().zip([
            ("rId1", 38.0, 32.0, 20.0, 10.0),
            ("rId2", 8.0, 12.0, 20.0, 10.0),
        ]) {
            assert_eq!(image.relationship_id, id);
            assert_eq!(image.size.width, width * EMU_PER_PIXEL);
            assert_eq!(image.size.height, height * EMU_PER_PIXEL);
            let position = image.position.as_ref().unwrap();
            assert_eq!(position.horizontal.relative_to, "page");
            assert_eq!(position.vertical.relative_to, "page");
            assert_eq!(position.horizontal.pos_offset, Some(x * EMU_PER_PIXEL));
            assert_eq!(position.vertical.pos_offset, Some(y * EMU_PER_PIXEL));
        }
    }

    #[test]
    fn grouped_vml_placement_uses_root_position_and_explicit_wrap() {
        for (position, wrap, expected) in [
            ("", "", "inline"),
            ("position:absolute;", "", "inFront"),
            (
                "position:relative;",
                r#"<w10:wrap type="square"/>"#,
                "square",
            ),
        ] {
            let picture = root(&format!(
                r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r" xmlns:w10="w10"><v:group coordsize="1000,1000" style="{position}width:200px;height:100px"><v:shape style="position:absolute;left:100;top:200;width:300;height:400"><v:imagedata r:id="rId1"/></v:shape><v:shape style="width:200;height:100"><v:imagedata r:id="rId2"/></v:shape>{wrap}</v:group></w:pict>"#
            ));
            let images = vml_image_contents(&picture, None, None, None);
            assert_eq!(images.len(), 2);
            for image in images {
                assert_eq!(image.wrap.wrap_type, expected);
                assert_eq!(image.position.is_some(), !position.is_empty());
            }
        }
    }

    #[test]
    fn grouped_vml_lengths_scale_coordinate_suffixes_and_empty_values() {
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r"><v:group coordorigin="10,20" coordsize="1000,500" style="position:absolute;left:8;top:12;width:200;height:100"><v:group coordorigin="50,60" coordsize="200,100" style="left:110pt;top:70px;width:400pt;height:200px"><v:shape style="left:75pt;top:85px;width:50pt;height:25px"><v:imagedata r:id="rId1"/></v:shape></v:group><v:group coordsize="1000,500" style="width:;height:100"><v:shape style="width:50;height:25"><v:imagedata r:id="rId2"/></v:shape></v:group></v:group></w:pict>"#,
        );
        let images = vml_image_contents(&picture, None, None, None);
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].size.width, 20.0 * EMU_PER_PIXEL);
        assert_eq!(images[0].size.height, 10.0 * EMU_PER_PIXEL);
        let position = images[0].position.as_ref().unwrap();
        assert_eq!(position.horizontal.pos_offset, Some(38.0 * EMU_PER_PIXEL));
        assert_eq!(position.vertical.pos_offset, Some(32.0 * EMU_PER_PIXEL));
        assert_eq!(images[1].size.width, 10.0 * EMU_PER_PIXEL);
        assert_eq!(images[1].size.height, EMU_PER_PIXEL);
        assert_eq!(group_length_to_px(Some(""), 0.2), None);
        assert_eq!(group_length_to_px(Some(" "), 0.2), None);
    }

    #[test]
    fn grouped_vml_empty_shape_dimensions_use_intrinsic_size() {
        let mut png = vec![0; 24];
        png[..4].copy_from_slice(&[137, b'P', b'N', b'G']);
        png[16..20].copy_from_slice(&40_u32.to_be_bytes());
        png[20..24].copy_from_slice(&20_u32.to_be_bytes());
        let media = build_media_map(&[("word/media/picture.png".to_owned(), png)]);
        let relationships = RelationshipMap::from([(
            "rId1".to_owned(),
            Relationship {
                id: "rId1".to_owned(),
                relationship_type: "image".to_owned(),
                target: "media/picture.png".to_owned(),
                target_mode: None,
            },
        )]);
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r"><v:group style="width:100px;height:100px"><v:shape style="width:;height:"><v:imagedata r:id="rId1"/></v:shape></v:group></w:pict>"#,
        );
        let image = vml_image_content(&picture, Some(&relationships), Some(&media), None).unwrap();
        assert_eq!(image.size.width, 40.0 * EMU_PER_PIXEL);
        assert_eq!(image.size.height, 20.0 * EMU_PER_PIXEL);
    }

    #[test]
    fn grouped_vml_offsets_prefer_coordinates_and_fall_back_to_margins() {
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r"><v:group coordsize="100,100" style="position:absolute;left:99px;top:99px;margin-left:8px;margin-top:12px;width:200px;height:100px"><v:group coordsize="100,100" style="left:10;top:20;margin-left:90;margin-top:90;width:50;height:50"><v:shape style="left:5;top:6;margin-left:80;margin-top:80;width:10;height:10"><v:imagedata r:id="rId1"/></v:shape></v:group><v:group coordsize="100,100" style="margin-left:10;margin-top:20;width:50;height:50"><v:shape style="margin-left:5;margin-top:6;width:10;height:10"><v:imagedata r:id="rId2"/></v:shape></v:group></v:group></w:pict>"#,
        );
        let images = vml_image_contents(&picture, None, None, None);
        assert_eq!(images.len(), 2);
        for image in images {
            let position = image.position.unwrap();
            assert_eq!(position.horizontal.pos_offset, Some(33.0 * EMU_PER_PIXEL));
            assert_eq!(position.vertical.pos_offset, Some(35.0 * EMU_PER_PIXEL));
        }
    }

    #[test]
    fn grouped_vml_negative_root_z_index_stays_behind_text() {
        for (root_z, child_z, expected_z, behind) in [
            ("z-index:-10;", "z-index:5;", 0.0, true),
            ("z-index:10;", "z-index:-5;", 10.0, false),
            ("", "z-index:-5;", 0.0, true),
        ] {
            let picture = root(&format!(
                r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r"><v:group style="position:absolute;{root_z}width:200px;height:100px"><v:shape style="{child_z}width:500;height:500"><v:imagedata r:id="rId1"/></v:shape></v:group></w:pict>"#
            ));
            let image = vml_image_content(&picture, None, None, None).unwrap();
            assert_eq!(
                image.wrap.wrap_type,
                if behind { "behind" } else { "inFront" }
            );
            let position = image.position.unwrap();
            assert_eq!(position.relative_height, Some(expected_z));
            assert_eq!(position.behind_doc, Some(behind));
        }
    }

    #[test]
    fn grouped_vml_pictures_follow_child_z_order_and_keep_root_stacking() {
        for root_z in [-10.0_f64, 10.0] {
            let picture = root(&format!(
                r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r"><v:group style="position:absolute;z-index:{root_z};width:100px;height:100px"><v:shape style="z-index:2;width:500;height:500"><v:imagedata r:id="rIdHigh"/></v:shape><v:shape style="z-index:1;width:500;height:500"><v:imagedata r:id="rIdLow"/></v:shape><v:group style="z-index:1;width:500;height:500"><v:shape style="z-index:2;width:500;height:500"><v:imagedata r:id="rIdNestedHigh"/></v:shape><v:shape style="z-index:1;width:500;height:500"><v:imagedata r:id="rIdNestedLow"/></v:shape></v:group></v:group></w:pict>"#
            ));
            let images = vml_image_contents(&picture, None, None, None);
            assert_eq!(
                images
                    .iter()
                    .map(|image| image.relationship_id.as_str())
                    .collect::<Vec<_>>(),
                ["rIdLow", "rIdNestedLow", "rIdNestedHigh", "rIdHigh"]
            );
            for image in images {
                let behind = root_z < 0.0;
                assert_eq!(
                    image.wrap.wrap_type,
                    if behind { "behind" } else { "inFront" }
                );
                let position = image.position.unwrap();
                assert_eq!(position.relative_height, Some(root_z.max(0.0)));
                assert_eq!(position.behind_doc, Some(behind));
            }
        }
    }

    #[test]
    fn vml_picture_does_not_claim_images_inside_a_textbox_story() {
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r"><v:shape style="width:200px;height:100px"><v:imagedata r:id="rIdOuter"/><v:textbox><w:txbxContent><w:p><w:r><w:pict><v:shape style="width:20px;height:10px"><v:imagedata r:id="rId1"/></v:shape></w:pict></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>"#,
        );
        let outer = vml_image_contents(&picture, None, None, None);
        assert_eq!(outer.len(), 1);
        assert_eq!(outer[0].relationship_id, "rIdOuter");
        assert_eq!(outer[0].size.width, 200.0 * EMU_PER_PIXEL);
        let inner = picture
            .find_deep("w", "txbxContent")
            .unwrap()
            .find_deep("w", "pict")
            .unwrap();
        let images = vml_image_contents(inner, None, None, None);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].size.width, 20.0 * EMU_PER_PIXEL);
    }

    #[test]
    fn vml_picture_only_inside_a_textbox_story_stays_visible() {
        let picture = root(
            r#"<w:pict xmlns:w="w" xmlns:v="v" xmlns:r="r"><v:group style="width:400px;height:200px" coordsize="100,100"><v:shape style="width:100;height:100"><v:textbox><w:txbxContent><w:p><w:r><w:pict><v:shape style="width:20px;height:10px"><v:imagedata r:id="rId1"/></v:shape><v:shape style="width:30px;height:15px"><v:imagedata r:id="rId2"/></v:shape></w:pict></w:r></w:p></w:txbxContent></v:textbox></v:shape></v:group></w:pict>"#,
        );
        let images = vml_image_contents(&picture, None, None, None);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].relationship_id, "rId1");
        assert_eq!(images[0].size.width, 20.0 * EMU_PER_PIXEL);
        assert_eq!(images[0].size.height, 10.0 * EMU_PER_PIXEL);
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
