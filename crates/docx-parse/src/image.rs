//! DrawingML picture parsing and WordprocessingML anchor/wrap helpers.

use serde::{Deserialize, Serialize};

use crate::drawingml::{ShapeOutline, Transform2D, parse_outline, rot_to_degrees};
use crate::media::{MediaMap, resolve_image_data};
use crate::relationships::{RelationshipMap, resolve_target};
use crate::xml::{XmlElement, parse_javascript_integer_prefix};

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ImageSize {
    pub width: f64,
    pub height: f64,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageWrap {
    #[serde(rename = "type")]
    pub wrap_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_t: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_b: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_l: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_r: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub polygon: Option<Vec<ImageWrapPoint>>,
}

impl ImageWrap {
    fn named(wrap_type: &str) -> Self {
        Self {
            wrap_type: wrap_type.to_owned(),
            wrap_text: None,
            dist_t: None,
            dist_b: None,
            dist_l: None,
            dist_r: None,
            polygon: None,
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ImageWrapPoint {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub y: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImagePosition {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub use_simple_pos: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub simple_pos: Option<OptionalPoint>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relative_height: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub behind_doc: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub locked: Option<bool>,
    pub horizontal: PositionAxis,
    pub vertical: PositionAxis,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct OptionalPoint {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub y: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PositionAxis {
    pub relative_to: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pos_offset: Option<f64>,
    // Pinned VML typo: that parser emits `offset`, not `posOffset`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset: Option<f64>,
}

impl PositionAxis {
    fn relative_to(value: &str) -> Self {
        Self {
            relative_to: value.to_owned(),
            alignment: None,
            pos_offset: None,
            offset: None,
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ImagePadding {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ImageCrop {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageEffect {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub amount: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub threshold: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub colors: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub raw_name: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ImageEffects {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub brightness: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub contrast: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub saturation: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ordered: Option<Vec<ImageEffect>>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageRotationBounds {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset_x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset_y: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ImageHyperlink {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub href: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub target: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub history: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_location: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Image {
    #[serde(rename = "type")]
    pub image_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(rename = "rId")]
    pub relationship_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub src: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mime_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filename: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alt: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    pub size: ImageSize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub original_size: Option<ImageSize>,
    pub wrap: ImageWrap,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ImagePosition>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub transform: Option<Transform2D>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub padding: Option<ImagePadding>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crop: Option<ImageCrop>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub decorative: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub layout_in_cell: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub allow_overlap: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hlink_href: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline: Option<ShapeOutline>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub effects: Option<ImageEffects>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation_bounds: Option<ImageRotationBounds>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<ImageHyperlink>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub effect_extent: Option<ImagePadding>,
}

pub fn parse_drawing(
    drawing: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
) -> Option<Image> {
    if is_text_box_drawing(drawing) {
        return None;
    }
    drawing
        .child_elements()
        .find_map(|container| match container.name.as_str() {
            "wp:inline" => Some(parse_inline(container, relationships, media)),
            "wp:anchor" => Some(parse_anchor(container, relationships, media)),
            _ => None,
        })
}

fn parse_inline(
    container: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
) -> Image {
    let size = parse_extent(container.child_by_full_name("wp:extent"));
    let padding = parse_effect_extent(container.child_by_full_name("wp:effectExtent"));
    let properties = parse_doc_properties(container.child_by_full_name("wp:docPr"));
    let (blip_fill, blip) = find_blip_chain(container);
    let relationship_id = extract_blip_id(blip);
    let resolved = resolve_image_data(&relationship_id, relationships, media);
    let transform = parse_picture_transform(container);
    let mut wrap = ImageWrap::named("inline");
    wrap.dist_t = numeric_attr(Some(container), "distT");
    wrap.dist_b = numeric_attr(Some(container), "distB");
    wrap.dist_l = numeric_attr(Some(container), "distL");
    wrap.dist_r = numeric_attr(Some(container), "distR");
    let mut image = image_base(relationship_id, size, wrap);
    apply_common_image_fields(
        &mut image,
        container,
        blip_fill,
        blip,
        properties,
        resolved,
        padding,
        transform,
        relationships,
    );
    image
}

fn parse_anchor(
    container: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
) -> Image {
    let size = parse_extent(container.child_by_full_name("wp:extent"));
    let padding = parse_effect_extent(container.child_by_full_name("wp:effectExtent"));
    let properties = parse_doc_properties(container.child_by_full_name("wp:docPr"));
    let behind_doc = container.attribute(None, "behindDoc") == Some("1");
    let anchor_hidden = container.attribute(None, "hidden") == Some("1");
    let layout_in_cell = container
        .attribute(None, "layoutInCell")
        .map(|value| value == "1");
    let allow_overlap = container
        .attribute(None, "allowOverlap")
        .map(|value| value == "1");
    let relative_height = container
        .attribute(None, "relativeHeight")
        .and_then(js_number)
        .filter(|value| value.is_finite());
    let simple_pos_flag = container.attribute(None, "simplePos") == Some("1");
    let simple_pos = container.child_by_full_name("wp:simplePos");
    // Incumbent direct-child lookup means nested graphicFrameLocks is ignored.
    let locked = container
        .child_by_full_name("a:graphicFrameLocks")
        .is_some();
    let distances = Distances {
        top: numeric_attr(Some(container), "distT"),
        bottom: numeric_attr(Some(container), "distB"),
        left: numeric_attr(Some(container), "distL"),
        right: numeric_attr(Some(container), "distR"),
    };
    let wrap_element = container
        .child_elements()
        .find(|element| is_wrap_name(&element.name));
    let wrap = parse_wrap_element(wrap_element, behind_doc, Some(&distances));
    let horizontal = parse_position_h(container.child_by_full_name("wp:positionH"));
    let vertical = parse_position_v(container.child_by_full_name("wp:positionV"));
    let mut position = if horizontal.is_some() || vertical.is_some() {
        Some(ImagePosition {
            use_simple_pos: None,
            simple_pos: None,
            relative_height: None,
            behind_doc: None,
            hidden: None,
            locked: None,
            horizontal: horizontal.unwrap_or_else(|| PositionAxis::relative_to("column")),
            vertical: vertical.unwrap_or_else(|| PositionAxis::relative_to("paragraph")),
        })
    } else if simple_pos_flag
        || relative_height.is_some()
        || properties.hidden
        || anchor_hidden
        || locked
    {
        Some(default_position())
    } else {
        None
    };
    if let Some(position) = &mut position {
        position.use_simple_pos = simple_pos_flag.then_some(true);
        if let Some(simple_pos) = simple_pos {
            position.simple_pos = Some(OptionalPoint {
                x: numeric_attr(Some(simple_pos), "x"),
                y: numeric_attr(Some(simple_pos), "y"),
            });
        }
        position.relative_height = relative_height;
        position.behind_doc = behind_doc.then_some(true);
        position.hidden = (properties.hidden || anchor_hidden).then_some(true);
        position.locked = locked.then_some(true);
    }
    let (blip_fill, blip) = find_blip_chain(container);
    let relationship_id = extract_blip_id(blip);
    let resolved = resolve_image_data(&relationship_id, relationships, media);
    let transform = parse_picture_transform(container);
    let mut image = image_base(relationship_id, size, wrap);
    image.position = position;
    image.layout_in_cell = layout_in_cell;
    image.allow_overlap = allow_overlap;
    apply_common_image_fields(
        &mut image,
        container,
        blip_fill,
        blip,
        properties,
        resolved,
        padding,
        transform,
        relationships,
    );
    image
}

fn image_base(relationship_id: String, size: ImageSize, wrap: ImageWrap) -> Image {
    Image {
        image_type: "image".into(),
        id: None,
        relationship_id,
        src: None,
        mime_type: None,
        filename: None,
        alt: None,
        title: None,
        size,
        original_size: None,
        wrap,
        position: None,
        transform: None,
        padding: None,
        crop: None,
        opacity: None,
        decorative: None,
        layout_in_cell: None,
        allow_overlap: None,
        hlink_href: None,
        outline: None,
        effects: None,
        rotation_bounds: None,
        hyperlink: None,
        effect_extent: None,
    }
}

pub fn placeholder_image(relationship_id: &str) -> Image {
    image_base(
        relationship_id.to_owned(),
        ImageSize::default(),
        ImageWrap::named("inline"),
    )
}

#[allow(clippy::too_many_arguments)]
fn apply_common_image_fields(
    image: &mut Image,
    container: &XmlElement,
    blip_fill: Option<&XmlElement>,
    blip: Option<&XmlElement>,
    properties: DocProperties,
    resolved: crate::media::ResolvedImageData,
    padding: Option<ImagePadding>,
    transform: Option<Transform2D>,
    relationships: Option<&RelationshipMap>,
) {
    image.id = properties.id;
    image.alt = properties.alt;
    image.title = properties.title;
    image.decorative = properties.decorative.then_some(true);
    image.src = resolved.src;
    image.mime_type = resolved.mime_type;
    image.filename = resolved.filename;
    image.padding = padding.clone();
    image.transform = transform;
    image.crop = parse_image_crop(blip_fill);
    image.opacity = parse_image_opacity(blip);
    if let Some(ordered) = parse_blip_effects(blip) {
        image
            .effects
            .get_or_insert_with(ImageEffects::default)
            .ordered = Some(ordered);
    }
    image.outline = parse_picture_outline(container);
    image.rotation_bounds = rotation_bounds(&image.size, image.transform.as_ref());
    if padding.is_some() {
        image.effect_extent = padding;
    }
    if let (Some(relationship_id), Some(relationships)) = (properties.hyperlink_id, relationships) {
        if let Some(href) = resolve_target(relationships, &relationship_id).and_then(sanitize_href)
        {
            image.hlink_href = Some(href.to_owned());
            image.hyperlink = Some(ImageHyperlink {
                href: Some(href.to_owned()),
                tooltip: properties.hyperlink_tooltip,
                target: properties.hyperlink_target,
                history: properties.hyperlink_history,
                doc_location: None,
            });
        }
    }
}

#[derive(Default)]
struct DocProperties {
    id: Option<String>,
    alt: Option<String>,
    title: Option<String>,
    decorative: bool,
    hyperlink_id: Option<String>,
    hyperlink_tooltip: Option<String>,
    hyperlink_target: Option<String>,
    hyperlink_history: Option<bool>,
    hidden: bool,
}

fn parse_doc_properties(element: Option<&XmlElement>) -> DocProperties {
    let Some(element) = element else {
        return DocProperties::default();
    };
    let hyperlink = element.child_by_full_name("a:hlinkClick");
    DocProperties {
        id: element
            .attribute(None, "id")
            .filter(|value| !value.is_empty())
            .map(str::to_owned),
        alt: element
            .attribute(None, "descr")
            .filter(|value| !value.is_empty())
            .map(str::to_owned),
        title: element
            .attribute(None, "title")
            .filter(|value| !value.is_empty())
            .map(str::to_owned),
        decorative: element.attribute(None, "decorative") == Some("1")
            || element.child_by_full_name("adec:decorative").is_some(),
        hyperlink_id: hyperlink
            .and_then(|value| value.attribute(Some("r"), "id"))
            .filter(|value| !value.is_empty())
            .map(str::to_owned),
        hyperlink_tooltip: hyperlink
            .and_then(|value| value.attribute(None, "tooltip"))
            .map(str::to_owned),
        hyperlink_target: hyperlink
            .and_then(|value| value.attribute(None, "tgtFrame"))
            .map(str::to_owned),
        hyperlink_history: hyperlink.map(|value| value.attribute(None, "history") != Some("0")),
        hidden: element.attribute(None, "hidden") == Some("1"),
    }
}

fn parse_extent(element: Option<&XmlElement>) -> ImageSize {
    ImageSize {
        width: numeric_attr(element, "cx").unwrap_or(0.0),
        height: numeric_attr(element, "cy").unwrap_or(0.0),
    }
}

fn parse_effect_extent(element: Option<&XmlElement>) -> Option<ImagePadding> {
    let element = element?;
    let padding = ImagePadding {
        left: Some(numeric_attr(Some(element), "l").unwrap_or(0.0)),
        top: Some(numeric_attr(Some(element), "t").unwrap_or(0.0)),
        right: Some(numeric_attr(Some(element), "r").unwrap_or(0.0)),
        bottom: Some(numeric_attr(Some(element), "b").unwrap_or(0.0)),
    };
    let all_zero = [padding.left, padding.top, padding.right, padding.bottom]
        .into_iter()
        .flatten()
        .all(|value| value == 0.0);
    (!all_zero).then_some(padding)
}

fn find_blip_chain(container: &XmlElement) -> (Option<&XmlElement>, Option<&XmlElement>) {
    let Some(graphic) = container.child_by_full_name("a:graphic") else {
        return (None, None);
    };
    let Some(data) = graphic.child_by_full_name("a:graphicData") else {
        return (None, None);
    };
    let Some(picture) = data.child_by_full_name("pic:pic") else {
        return (None, None);
    };
    let Some(fill) = picture.child_by_full_name("pic:blipFill") else {
        return (None, None);
    };
    (Some(fill), fill.child_by_full_name("a:blip"))
}

fn extract_blip_id(blip: Option<&XmlElement>) -> String {
    blip.and_then(|element| {
        element
            .attribute(Some("r"), "embed")
            .or_else(|| element.attribute(None, "embed"))
            .or_else(|| element.attribute(Some("r"), "link"))
    })
    .unwrap_or_default()
    .to_owned()
}

fn parse_picture_transform(container: &XmlElement) -> Option<Transform2D> {
    let graphic = container.child_by_full_name("a:graphic")?;
    let data = graphic.child_by_full_name("a:graphicData")?;
    let picture = data.child_by_full_name("pic:pic")?;
    let properties = picture.child_by_full_name("pic:spPr")?;
    let transform = properties.child_by_full_name("a:xfrm")?;
    let rotation = rot_to_degrees(transform.attribute(None, "rot"));
    let flip_h = (transform.attribute(None, "flipH") == Some("1")).then_some(true);
    let flip_v = (transform.attribute(None, "flipV") == Some("1")).then_some(true);
    (rotation.is_some() || flip_h.is_some() || flip_v.is_some()).then_some(Transform2D {
        rotation,
        flip_h,
        flip_v,
    })
}

fn parse_image_crop(fill: Option<&XmlElement>) -> Option<ImageCrop> {
    let rect = fill?.child_by_full_name("a:srcRect")?;
    let fraction = |name| {
        numeric_attr(Some(rect), name)
            .filter(|value| *value != 0.0)
            .map(|value| value / 100_000.0)
    };
    let crop = ImageCrop {
        left: fraction("l"),
        top: fraction("t"),
        right: fraction("r"),
        bottom: fraction("b"),
    };
    (crop.left.is_some() || crop.top.is_some() || crop.right.is_some() || crop.bottom.is_some())
        .then_some(crop)
}

fn parse_image_opacity(blip: Option<&XmlElement>) -> Option<f64> {
    let amount = numeric_attr(blip?.child_by_full_name("a:alphaModFix"), "amt")?;
    (amount < 100_000.0).then(|| (amount / 100_000.0).clamp(0.0, 1.0))
}

fn parse_blip_effects(blip: Option<&XmlElement>) -> Option<Vec<ImageEffect>> {
    let mut effects = Vec::new();
    for child in blip?.child_elements().take(256) {
        let name = child.local_name();
        if name == "alphaModFix" {
            continue;
        }
        let mut effect = ImageEffect::default();
        match name {
            "grayscl" => effect.kind = Some("grayscale".into()),
            "biLevel" => {
                effect.kind = Some("biLevel".into());
                effect.threshold = numeric_attr(Some(child), "thresh");
            }
            "lum" => {
                effect.kind = Some("brightnessContrast".into());
                effect.amount = numeric_attr(Some(child), "bright");
                effect.threshold = numeric_attr(Some(child), "contrast");
            }
            "satMod" | "satOff" => {
                effect.kind = Some("saturation".into());
                effect.amount = numeric_attr(Some(child), "val");
            }
            "clrChange" => {
                effect.kind = Some("colorChange".into());
                effect.raw_name = Some(name.into());
            }
            "duotone" => {
                effect.kind = Some("duotone".into());
                effect.raw_name = Some(name.into());
            }
            "blur" => {
                effect.kind = Some("blur".into());
                effect.amount = numeric_attr(Some(child), "rad");
            }
            value if value.to_ascii_lowercase().contains("alpha") => {
                effect.kind = Some("alpha".into());
                effect.amount = numeric_attr(Some(child), "amt");
            }
            _ => {
                effect.kind = Some("unknown".into());
                effect.raw_name = Some(name.chars().take(64).collect());
            }
        }
        effects.push(effect);
    }
    (!effects.is_empty()).then_some(effects)
}

fn parse_picture_outline(container: &XmlElement) -> Option<ShapeOutline> {
    let graphic = container.child_by_full_name("a:graphic")?;
    let data = graphic.child_by_full_name("a:graphicData")?;
    let picture = data.child_by_full_name("pic:pic")?;
    parse_outline(picture.child_by_full_name("pic:spPr"))
}

fn rotation_bounds(
    size: &ImageSize,
    transform: Option<&Transform2D>,
) -> Option<ImageRotationBounds> {
    let rotation = transform?.rotation.filter(|value| *value != 0.0)?;
    let radians = rotation * std::f64::consts::PI / 180.0;
    let width = (size.width * radians.cos()).abs() + (size.height * radians.sin()).abs();
    let height = (size.width * radians.sin()).abs() + (size.height * radians.cos()).abs();
    Some(ImageRotationBounds {
        width: Some(width),
        height: Some(height),
        offset_x: Some((size.width - width) / 2.0),
        offset_y: Some((size.height - height) / 2.0),
    })
}

pub fn parse_position_h(element: Option<&XmlElement>) -> Option<PositionAxis> {
    let element = element?;
    let mut position =
        PositionAxis::relative_to(element.attribute(None, "relativeFrom").unwrap_or("column"));
    if let Some(align) = element.child_by_full_name("wp:align") {
        position.alignment = Some(align.text_content());
    } else if let Some(offset) = element.child_by_full_name("wp:posOffset") {
        position.pos_offset = Some(parse_integer_text(&offset.text_content()).unwrap_or(0.0));
    }
    Some(position)
}

pub fn parse_position_v(element: Option<&XmlElement>) -> Option<PositionAxis> {
    let element = element?;
    let mut position = PositionAxis::relative_to(
        element
            .attribute(None, "relativeFrom")
            .unwrap_or("paragraph"),
    );
    if let Some(align) = element.child_by_full_name("wp:align") {
        position.alignment = Some(align.text_content());
    } else if let Some(offset) = element.child_by_full_name("wp:posOffset") {
        position.pos_offset = Some(parse_integer_text(&offset.text_content()).unwrap_or(0.0));
    }
    Some(position)
}

pub fn parse_anchor_position(anchor: &XmlElement) -> Option<ImagePosition> {
    let horizontal_element = anchor.child_by_full_name("wp:positionH");
    let vertical_element = anchor.child_by_full_name("wp:positionV");
    let simple_pos = anchor.child_by_full_name("wp:simplePos");
    let use_simple_pos = anchor.attribute(None, "simplePos") == Some("1");
    let relative_height = numeric_attr(Some(anchor), "relativeHeight");
    let behind_doc = anchor.attribute(None, "behindDoc") == Some("1");
    let hidden_raw = anchor.attribute(None, "hidden");
    let locked_raw = anchor.attribute(None, "locked");
    if horizontal_element.is_none()
        && vertical_element.is_none()
        && simple_pos.is_none()
        && !use_simple_pos
        && relative_height.is_none()
        && !behind_doc
        && hidden_raw.is_none()
        && locked_raw.is_none()
    {
        return None;
    }
    Some(ImagePosition {
        use_simple_pos: use_simple_pos.then_some(true),
        simple_pos: (simple_pos.is_some() && use_simple_pos).then(|| OptionalPoint {
            x: numeric_attr(simple_pos, "x"),
            y: numeric_attr(simple_pos, "y"),
        }),
        relative_height: relative_height.filter(|value| value.is_finite()),
        behind_doc: behind_doc.then_some(true),
        hidden: matches!(hidden_raw, Some("1" | "true")).then_some(true),
        locked: matches!(locked_raw, Some("1" | "true")).then_some(true),
        horizontal: parse_position_h(horizontal_element)
            .unwrap_or_else(|| PositionAxis::relative_to("column")),
        vertical: parse_position_v(vertical_element)
            .unwrap_or_else(|| PositionAxis::relative_to("paragraph")),
    })
}

#[derive(Default)]
struct Distances {
    top: Option<f64>,
    bottom: Option<f64>,
    left: Option<f64>,
    right: Option<f64>,
}

fn parse_wrap_element(
    element: Option<&XmlElement>,
    behind_doc: bool,
    distances: Option<&Distances>,
) -> ImageWrap {
    let Some(element) = element else {
        let mut wrap = ImageWrap::named(if behind_doc { "behind" } else { "inFront" });
        apply_distances(&mut wrap, distances);
        return wrap;
    };
    let wrap_type = match element.local_name() {
        "wrapNone" => {
            if behind_doc {
                "behind"
            } else {
                "inFront"
            }
        }
        "wrapSquare" => "square",
        "wrapTight" => "tight",
        "wrapThrough" => "through",
        "wrapTopAndBottom" => "topAndBottom",
        _ => "square",
    };
    let mut wrap = ImageWrap::named(wrap_type);
    wrap.wrap_text = element.attribute(None, "wrapText").map(str::to_owned);
    wrap.dist_t =
        numeric_attr(Some(element), "distT").or_else(|| distances.and_then(|value| value.top));
    wrap.dist_b =
        numeric_attr(Some(element), "distB").or_else(|| distances.and_then(|value| value.bottom));
    wrap.dist_l =
        numeric_attr(Some(element), "distL").or_else(|| distances.and_then(|value| value.left));
    wrap.dist_r =
        numeric_attr(Some(element), "distR").or_else(|| distances.and_then(|value| value.right));
    if let Some(polygon) = element.child_by_full_name("wp:wrapPolygon") {
        let points = polygon.children_by_local_name("start").chain(polygon.children_by_local_name("lineTo")).take(2_048).map(|point| ImageWrapPoint { x: numeric_attr(Some(point), "x"), y: numeric_attr(Some(point), "y") }).filter(|point| matches!((point.x,point.y),(Some(x),Some(y)) if x.is_finite() && y.is_finite() && x.abs()<=1_000_000_000.0 && y.abs()<=1_000_000_000.0)).collect::<Vec<_>>();
        if points.len() > 1 {
            wrap.polygon = Some(points);
        }
    }
    wrap
}

/// Chart drawings intentionally ignore anchor-level distance fallbacks, matching
/// the incumbent `parseWrapElement(wrap, behindDoc)` call site.
pub fn parse_wrap_element_without_distances(
    element: Option<&XmlElement>,
    behind_doc: bool,
) -> ImageWrap {
    parse_wrap_element(element, behind_doc, None)
}

pub fn parse_anchor_wrap(anchor: &XmlElement) -> ImageWrap {
    let distances = Distances {
        top: numeric_attr(Some(anchor), "distT"),
        bottom: numeric_attr(Some(anchor), "distB"),
        left: numeric_attr(Some(anchor), "distL"),
        right: numeric_attr(Some(anchor), "distR"),
    };
    let element = anchor
        .child_elements()
        .find(|element| is_wrap_name(&element.name));
    parse_wrap_element(
        element,
        anchor.attribute(None, "behindDoc") == Some("1"),
        Some(&distances),
    )
}

fn apply_distances(wrap: &mut ImageWrap, distances: Option<&Distances>) {
    if let Some(d) = distances {
        wrap.dist_t = d.top;
        wrap.dist_b = d.bottom;
        wrap.dist_l = d.left;
        wrap.dist_r = d.right;
    }
}
fn is_wrap_name(name: &str) -> bool {
    matches!(
        name,
        "wp:wrapNone" | "wp:wrapSquare" | "wp:wrapTight" | "wp:wrapThrough" | "wp:wrapTopAndBottom"
    )
}
fn default_position() -> ImagePosition {
    ImagePosition {
        use_simple_pos: None,
        simple_pos: None,
        relative_height: None,
        behind_doc: None,
        hidden: None,
        locked: None,
        horizontal: PositionAxis::relative_to("column"),
        vertical: PositionAxis::relative_to("paragraph"),
    }
}

pub fn is_text_box_drawing(drawing: &XmlElement) -> bool {
    let Some(container) = drawing
        .child_elements()
        .find(|element| matches!(element.name.as_str(), "wp:inline" | "wp:anchor"))
    else {
        return false;
    };
    let Some(graphic) = container.child_by_full_name("a:graphic") else {
        return false;
    };
    let Some(data) = graphic.child_by_full_name("a:graphicData") else {
        return false;
    };
    let Some(shape) = data.child_by_full_name("wps:wsp") else {
        return false;
    };
    shape.child_by_full_name("wps:txbx").is_some()
}

pub fn sanitize_href(value: &str) -> Option<&str> {
    if value.is_empty() {
        return None;
    }
    let probe = value
        .chars()
        .filter(|c| !matches!(c, '\t' | '\n' | '\r'))
        .collect::<String>();
    let probe = probe.trim_start_matches(|c: char| (c as u32) <= 0x20);
    if probe.is_empty() {
        return None;
    }
    let Some(colon) = probe.find(':') else {
        return Some(value);
    };
    let scheme = &probe[..colon];
    let mut chars = scheme.chars();
    if !chars.next().is_some_and(|c| c.is_ascii_alphabetic())
        || !chars.all(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '.' | '-'))
    {
        return Some(value);
    }
    matches!(
        scheme.to_ascii_lowercase().as_str(),
        "http" | "https" | "mailto" | "tel" | "ftp"
    )
    .then_some(value)
}

fn numeric_attr(element: Option<&XmlElement>, name: &str) -> Option<f64> {
    parse_javascript_integer_prefix(element?.attribute(None, name)?)
}
fn parse_integer_text(value: &str) -> Option<f64> {
    parse_javascript_integer_prefix(value)
}
fn js_number(value: &str) -> Option<f64> {
    let value = value.trim();
    if value.is_empty() {
        Some(0.0)
    } else {
        value.parse().ok()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::media::build_media_map;
    use crate::relationships::{Relationship, TargetMode};
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        parse_xml(
            xml.as_bytes(),
            "drawing.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap()
        .root()
        .unwrap()
        .clone()
    }

    #[test]
    fn parses_inline_picture_crop_opacity_effects_and_embedded_bytes() {
        let drawing = root(
            r#"<w:drawing><wp:inline distT="1" distR="4"><wp:extent cx="1000000" cy="500000"/><wp:effectExtent l="3" t="1" r="4" b="2"/><wp:docPr id="1" descr="alt"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId1"><a:alphaModFix amt="50000"/><a:grayscl/></a:blip><a:srcRect l="10000"/></pic:blipFill><pic:spPr><a:xfrm rot="1800000"/><a:ln w="9525"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        );
        let relationships = RelationshipMap::from([(
            "rId1".into(),
            Relationship {
                id: "rId1".into(),
                relationship_type: "image".into(),
                target: "media/a.png".into(),
                target_mode: Some(TargetMode::Internal),
            },
        )]);
        let media = build_media_map(&[("word/media/a.png".into(), vec![1, 2, 3])]);
        let image = parse_drawing(&drawing, Some(&relationships), Some(&media)).unwrap();
        assert_eq!(image.crop.unwrap().left, Some(0.1));
        assert_eq!(image.opacity, Some(0.5));
        assert_eq!(image.wrap.dist_t, Some(1.0));
        assert_eq!(image.src.as_deref(), Some("data:image/png;base64,AQID"));
        assert_eq!(
            image.effects.unwrap().ordered.unwrap()[0].kind.as_deref(),
            Some("grayscale")
        );
        assert!(image.rotation_bounds.is_some());
    }

    #[test]
    fn anchor_defaults_wrap_polygon_and_url_sanitization_are_pinned() {
        let drawing = root(
            r#"<w:drawing><wp:anchor behindDoc="1" simplePos="1" relativeHeight="7" layoutInCell="0" allowOverlap="1"><wp:simplePos x="9" y="10"/><wp:positionH relativeFrom="page"><wp:posOffset>bad</wp:posOffset></wp:positionH><wp:extent cx="1" cy="2"/><wp:wrapTight wrapText="left"><wp:wrapPolygon><wp:start x="0" y="0"/><wp:lineTo x="1" y="1"/></wp:wrapPolygon></wp:wrapTight><wp:docPr id="2"><a:hlinkClick r:id="link"/></wp:docPr><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="missing"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>"#,
        );
        let relationships = RelationshipMap::from([(
            "link".into(),
            Relationship {
                id: "link".into(),
                relationship_type: "hyperlink".into(),
                target: "java\tscript:alert(1)".into(),
                target_mode: Some(TargetMode::External),
            },
        )]);
        let image = parse_drawing(&drawing, Some(&relationships), None).unwrap();
        assert_eq!(image.wrap.wrap_type, "tight");
        assert_eq!(image.wrap.polygon.unwrap().len(), 2);
        assert_eq!(image.position.unwrap().horizontal.pos_offset, Some(0.0));
        assert_eq!(image.layout_in_cell, Some(false));
        assert!(image.hlink_href.is_none());
        assert_eq!(sanitize_href(" ../relative "), Some(" ../relative "));
        assert_eq!(
            sanitize_href("https://example.test"),
            Some("https://example.test")
        );
        assert_eq!(sanitize_href("data:text/html,x"), None);
    }

    #[test]
    fn text_box_shapes_are_not_misparsed_as_images() {
        let drawing = root(
            "<w:drawing><wp:inline><a:graphic><a:graphicData><wps:wsp><wps:txbx/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>",
        );
        assert!(parse_drawing(&drawing, None, None).is_none());
    }
}
