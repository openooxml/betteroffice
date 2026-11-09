//! Text-box drawing metadata and story-dispatch boundary.

use serde::{Deserialize, Serialize};

use crate::drawingml::{
    ShapeFill, ShapeOutline, parse_fill, parse_outline, resolve_color_value_to_hex,
};
use crate::image::{
    ImagePadding, ImagePosition, ImageSize, ImageWrap, parse_anchor_position, parse_anchor_wrap,
};
use crate::shape::{RelativeRect, Shape, ShapeTextBody, ShapeTextBodyProperties};
use crate::xml::XmlElement;

const DEFAULT_MARGIN_EMU: f64 = 91_440.0;
const MAX_SAFE_DRAWING_NUMBER: f64 = 1_000_000_000_000.0;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBox {
    #[serde(rename = "type")]
    pub content_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub size: ImageSize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ImagePosition>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap: Option<ImageWrap>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFill>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline: Option<ShapeOutline>,
    /// Story nodes owned by the block parser.
    pub content: Vec<serde_json::Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margins: Option<ImagePadding>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_properties: Option<TextBoxBodyProperties>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBoxBodyProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub upright: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_center: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub columns: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_spacing: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal_overflow: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_overflow: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margins: Option<ImagePadding>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub auto_fit: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_scale: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_spacing_reduction: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub from_word_art: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub preset_text_warp: Option<String>,
}

/// Ordered text-box block references for the recursive story dispatcher.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextBoxBlockElement<'a> {
    Paragraph(&'a XmlElement),
    Table(&'a XmlElement),
}

pub fn extract_text_box_content_elements(
    content: Option<&XmlElement>,
) -> (Vec<&XmlElement>, Vec<&XmlElement>) {
    let Some(content) = content else {
        return (Vec::new(), Vec::new());
    };
    (
        content.children_by_local_name("p").collect(),
        content.children_by_local_name("tbl").collect(),
    )
}

pub fn text_box_block_elements(content: Option<&XmlElement>) -> Vec<TextBoxBlockElement<'_>> {
    let Some(content) = content else {
        return Vec::new();
    };
    content
        .child_elements()
        .filter_map(|child| match child.local_name() {
            "p" => Some(TextBoxBlockElement::Paragraph(child)),
            "tbl" => Some(TextBoxBlockElement::Table(child)),
            _ => None,
        })
        .collect()
}

pub fn is_shape_text_box(shape: &XmlElement) -> bool {
    shape.child_by_full_name("wps:txbx").is_some()
}

fn text_box_elements(drawing: &XmlElement) -> Option<(&XmlElement, &XmlElement)> {
    let container = drawing
        .child_elements()
        .find(|child| matches!(child.name.as_str(), "wp:inline" | "wp:anchor"))?;
    let graphic = container.child_by_full_name("a:graphic")?;
    let graphic_data = graphic.child_by_full_name("a:graphicData")?;
    let shape = graphic_data.child_by_full_name("wps:wsp")?;
    shape.child_by_full_name("wps:txbx")?;
    Some((container, shape))
}

pub fn parse_text_box(drawing: &XmlElement) -> Option<TextBox> {
    let (container, shape) = text_box_elements(drawing)?;
    let properties = shape.child_by_full_name("wps:spPr");
    let body = shape.child_by_full_name("wps:bodyPr");
    let extent = container.child_by_full_name("wp:extent");
    let size = ImageSize {
        width: safe_numeric_attribute(extent, "cx").unwrap_or(0.0),
        height: safe_numeric_attribute(extent, "cy").unwrap_or(0.0),
    };
    let doc_properties = container.child_by_full_name("wp:docPr");
    let id = doc_properties
        .and_then(|properties| properties.attribute(None, "id"))
        .filter(|value| !value.is_empty())
        .map(str::to_owned);
    let name = doc_properties
        .and_then(|properties| properties.attribute(None, "name"))
        .filter(|value| !value.is_empty())
        .map(str::to_owned);
    let (margins, body_properties) = parse_body_properties(body);
    let anchored = container.name == "wp:anchor";
    Some(TextBox {
        content_type: "textBox".to_owned(),
        id,
        name,
        size,
        position: anchored.then(|| parse_anchor_position(container)).flatten(),
        wrap: anchored.then(|| parse_anchor_wrap(container)),
        fill: parse_fill(properties),
        outline: parse_outline(properties),
        content: Vec::new(),
        margins,
        body_properties,
    })
}

pub fn get_text_box_content_element(shape: &XmlElement) -> Option<&XmlElement> {
    shape
        .child_by_full_name("wps:txbx")?
        .child_by_full_name("w:txbxContent")
}

/// The model of a standalone text box drawing, with `content` as its text.
pub(crate) fn text_box_shape(
    drawing: &XmlElement,
    content: Vec<serde_json::Value>,
) -> Option<Shape> {
    let (_, element) = text_box_elements(drawing)?;
    let mut shape = crate::shape::parse_shape_from_drawing(drawing)?;
    let body = shape_body_properties(element.child_by_full_name("wps:bodyPr"));
    shape.text_body = Some(ShapeTextBody {
        vertical: body.as_ref().and_then(|body| {
            body.vertical
                .as_deref()
                .filter(|value| *value != "horizontal")
                .map(|_| true)
        }),
        rotation: body.as_ref().and_then(|body| body.rotation),
        anchor: body.as_ref().and_then(|body| body.anchor.clone()),
        anchor_center: body.as_ref().and_then(|body| body.anchor_center),
        auto_fit: body.as_ref().and_then(|body| body.auto_fit.clone()),
        margins: body.as_ref().and_then(|body| body_margins(&body.margins)),
        content,
    });
    shape.text_body_properties = shape.text_body_properties.or(body);
    Some(shape)
}

pub(crate) fn capture_bookmark_source(shape: &mut Shape, source: &XmlElement) {
    if !has_bookmarked_text_box(source) {
        return;
    }
    shape.drawing_xml = Some(source.to_bookmarked_raw_inline_xml());
    shape.drawing_run_xml = None;
    shape.drawing_model = Some(serde_json::json!({
        "frame": crate::smart_art::frame_model(shape),
        "otherSha256": crate::smart_art::other_model_sha256(shape),
        "textBoxContentSha256": shape_text_bodies(shape).iter().map(|body| content_sha256(&body.content)).collect::<Vec<_>>()
    }));
}

/// The group model without its frame, text-box text and media sources, which a replay
/// updates in place or which differ only by how the media was resolved.
pub(crate) fn group_model_sha256(shape: &Shape) -> String {
    fn clear(value: &mut serde_json::Value) {
        match value {
            serde_json::Value::Object(map) => {
                if let Some(body) = map
                    .get_mut("textBody")
                    .and_then(serde_json::Value::as_object_mut)
                {
                    body.insert("content".to_owned(), serde_json::Value::Array(Vec::new()));
                }
                map.retain(|_, value| {
                    !value
                        .as_str()
                        .is_some_and(|text| text.starts_with("media:") || text.starts_with("data:"))
                });
                map.values_mut().for_each(clear);
            }
            serde_json::Value::Array(items) => items.iter_mut().for_each(clear),
            _ => {}
        }
    }
    let mut value = serde_json::to_value(shape).expect("shape model");
    clear(&mut value);
    crate::smart_art::other_value_sha256(value)
}

pub(crate) fn content_sha256(content: &[serde_json::Value]) -> String {
    use sha2::{Digest, Sha256};

    let model = crate::smart_art::wire_model(serde_json::Value::Array(content.to_vec()));
    format!(
        "{:x}",
        Sha256::digest(serde_json::to_vec(&model).expect("text box content"))
    )
}

/// An `mc:AlternateContent` holding a bookmarked text box next to another drawing in one
/// branch: no single shape can own its source, so each drawing keeps its own.
pub(crate) fn shares_a_bookmarked_branch(element: &XmlElement) -> bool {
    has_bookmarked_text_box(element)
        && element.child_elements().any(|branch| {
            branch
                .child_elements()
                .filter(|child| matches!(child.local_name(), "drawing" | "pict" | "object"))
                .count()
                > 1
        })
}

pub(crate) fn contains_text_box(element: &XmlElement) -> bool {
    element.name == "w:txbxContent" || element.child_elements().any(contains_text_box)
}

pub(crate) fn has_bookmarked_text_box(element: &XmlElement) -> bool {
    has_block_bookmarks(element, false)
}

fn has_block_bookmarks(element: &XmlElement, in_text_box: bool) -> bool {
    let in_text_box = in_text_box || element.name == "w:txbxContent";
    if in_text_box && element.name == "w:p" {
        return contains_bookmark_boundary(element);
    }
    (in_text_box && crate::block::has_direct_block_bookmark(element))
        || element
            .child_elements()
            .any(|child| has_block_bookmarks(child, in_text_box))
}

fn contains_bookmark_boundary(element: &XmlElement) -> bool {
    element.child_elements().any(|child| {
        matches!(child.local_name(), "bookmarkStart" | "bookmarkEnd")
            || contains_bookmark_boundary(child)
    })
}

pub(crate) fn source_xml(source: &XmlElement) -> String {
    if has_bookmarked_text_box(source) {
        source.to_bookmarked_raw_inline_xml()
    } else {
        source.to_raw_inline_xml()
    }
}

pub(crate) fn shape_text_bodies(shape: &Shape) -> Vec<&ShapeTextBody> {
    fn visit<'a>(shape: &'a Shape, bodies: &mut Vec<&'a ShapeTextBody>) {
        bodies.extend(shape.text_body.iter());
        for child in shape.children.iter().flatten() {
            visit(child, bodies);
        }
    }
    let mut bodies = Vec::new();
    visit(shape, &mut bodies);
    bodies
}

pub fn parse_text_box_from_shape(
    shape: &XmlElement,
    size: ImageSize,
    position: Option<ImagePosition>,
    wrap: Option<ImageWrap>,
) -> Option<TextBox> {
    shape.child_by_full_name("wps:txbx")?;
    let properties = shape.child_by_full_name("wps:spPr");
    let body = shape.child_by_full_name("wps:bodyPr");
    let non_visual = shape.child_by_full_name("wps:cNvPr");
    let id = non_visual
        .and_then(|properties| properties.attribute(None, "id"))
        .filter(|value| !value.is_empty())
        .map(str::to_owned);
    let name = non_visual
        .and_then(|properties| properties.attribute(None, "name"))
        .filter(|value| !value.is_empty())
        .map(str::to_owned);
    let (margins, body_properties) = parse_body_properties(body);
    Some(TextBox {
        content_type: "textBox".to_owned(),
        id,
        name,
        size,
        position,
        wrap,
        fill: parse_fill(properties),
        outline: parse_outline(properties),
        content: Vec::new(),
        margins,
        body_properties,
    })
}

fn parse_body_properties(
    body: Option<&XmlElement>,
) -> (Option<ImagePadding>, Option<TextBoxBodyProperties>) {
    let Some(body) = shape_body_properties(body) else {
        return (None, None);
    };
    let margins = body_margins(&body.margins).map(|margins| ImagePadding {
        left: margins.left,
        top: margins.top,
        right: margins.right,
        bottom: margins.bottom,
    });
    let properties = TextBoxBodyProperties {
        vertical: body.vertical,
        rotation: body.rotation,
        upright: body.upright,
        anchor: body.anchor,
        anchor_center: body.anchor_center,
        columns: body.columns,
        column_spacing: body.column_spacing,
        wrap: body.wrap,
        horizontal_overflow: body.horizontal_overflow,
        vertical_overflow: body.vertical_overflow,
        margins: margins.clone(),
        auto_fit: body.auto_fit,
        font_scale: body.font_scale,
        line_spacing_reduction: body.line_spacing_reduction,
        from_word_art: body.from_word_art,
        preset_text_warp: body.preset_text_warp,
    };
    (margins, Some(properties))
}

fn shape_body_properties(body: Option<&XmlElement>) -> Option<ShapeTextBodyProperties> {
    let mut properties = crate::shape::parse_text_body_properties(body?);
    properties.rotation = properties.rotation.filter(safe_number);
    properties.upright = properties.upright.filter(|value| *value);
    properties.anchor_center = properties.anchor_center.filter(|value| *value);
    properties.from_word_art = properties.from_word_art.filter(|value| *value);
    (properties != ShapeTextBodyProperties::default()).then_some(properties)
}

fn body_margins(margins: &RelativeRect) -> Option<RelativeRect> {
    (margins.left.is_some()
        || margins.top.is_some()
        || margins.right.is_some()
        || margins.bottom.is_some())
    .then(|| margins.clone())
}

pub fn emu_to_pixels(emu: f64) -> f64 {
    js_round(emu * 96.0 / 914_400.0)
}

pub fn get_text_box_dimensions_px(text_box: &TextBox) -> ImageSize {
    ImageSize {
        width: emu_to_pixels(text_box.size.width),
        height: emu_to_pixels(text_box.size.height),
    }
}

pub fn get_text_box_margins_px(text_box: &TextBox) -> ImagePadding {
    let margins = text_box.margins.as_ref();
    ImagePadding {
        top: Some(emu_to_pixels(
            margins
                .and_then(|value| value.top)
                .unwrap_or(DEFAULT_MARGIN_EMU),
        )),
        bottom: Some(emu_to_pixels(
            margins
                .and_then(|value| value.bottom)
                .unwrap_or(DEFAULT_MARGIN_EMU),
        )),
        left: Some(emu_to_pixels(
            margins
                .and_then(|value| value.left)
                .unwrap_or(DEFAULT_MARGIN_EMU),
        )),
        right: Some(emu_to_pixels(
            margins
                .and_then(|value| value.right)
                .unwrap_or(DEFAULT_MARGIN_EMU),
        )),
    }
}

pub fn is_floating_text_box(text_box: &TextBox) -> bool {
    text_box.position.is_some() || text_box.wrap.is_some()
}

pub fn has_text_box_fill(text_box: &TextBox) -> bool {
    text_box
        .fill
        .as_ref()
        .is_some_and(|fill| fill.fill_type != "none")
}

pub fn has_text_box_outline(text_box: &TextBox) -> bool {
    text_box.outline.is_some()
}

pub fn resolve_text_box_fill_color(text_box: &TextBox) -> Option<String> {
    let fill = text_box.fill.as_ref()?;
    (fill.fill_type == "solid")
        .then(|| resolve_color_value_to_hex(fill.color.as_ref()))
        .flatten()
}

pub fn resolve_text_box_outline_color(text_box: &TextBox) -> Option<String> {
    resolve_color_value_to_hex(text_box.outline.as_ref()?.color.as_ref())
}

pub fn get_text_box_outline_width_px(text_box: &TextBox) -> f64 {
    text_box
        .outline
        .as_ref()
        .and_then(|outline| outline.width)
        .filter(|width| *width != 0.0)
        .map(emu_to_pixels)
        .unwrap_or(0.0)
}

fn safe_numeric_attribute(element: Option<&XmlElement>, name: &str) -> Option<f64> {
    element?
        .parse_numeric_attribute(None, name, 1.0)
        .filter(safe_number)
}

fn safe_number(value: &f64) -> bool {
    value.is_finite() && value.abs() <= MAX_SAFE_DRAWING_NUMBER
}

fn js_round(value: f64) -> f64 {
    (value + 0.5).floor()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        parse_xml(xml.as_bytes(), "textbox", &mut budget)
            .unwrap()
            .root()
            .unwrap()
            .clone()
    }

    #[test]
    fn parses_metadata_body_properties_and_ordered_story_boundary() {
        let drawing = root(
            r#"<w:drawing xmlns:w="w" xmlns:wp="wp" xmlns:a="a" xmlns:wps="wps"><wp:anchor relativeHeight="3"><wp:extent cx="914400" cy="457200"/><wp:docPr id="9"/><wp:positionH relativeFrom="page"/><wp:wrapSquare/><a:graphic><a:graphicData><wps:wsp><wps:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr><wps:txbx><w:txbxContent><w:p/><w:tbl/><w:p/></w:txbxContent></wps:txbx><wps:bodyPr lIns="91440" vert="vert270" anchor="ctr"><a:normAutofit fontScale="80000"/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>"#,
        );
        assert!(crate::image::is_text_box_drawing(&drawing));
        let text_box = parse_text_box(&drawing).unwrap();
        assert_eq!(text_box.id.as_deref(), Some("9"));
        assert_eq!(text_box.size.width, 914_400.0);
        assert_eq!(
            text_box
                .body_properties
                .as_ref()
                .unwrap()
                .vertical
                .as_deref(),
            Some("vertical270")
        );
        assert_eq!(
            text_box
                .body_properties
                .as_ref()
                .unwrap()
                .auto_fit
                .as_deref(),
            Some("normal")
        );
        assert_eq!(get_text_box_dimensions_px(&text_box).width, 96.0);
        let content = get_text_box_content_element(
            drawing
                .child_by_full_name("wp:anchor")
                .unwrap()
                .child_by_full_name("a:graphic")
                .unwrap()
                .child_by_full_name("a:graphicData")
                .unwrap()
                .child_by_full_name("wps:wsp")
                .unwrap(),
        );
        let blocks = text_box_block_elements(content);
        assert!(matches!(
            blocks.as_slice(),
            [
                TextBoxBlockElement::Paragraph(_),
                TextBoxBlockElement::Table(_),
                TextBoxBlockElement::Paragraph(_)
            ]
        ));
    }

    #[test]
    fn direct_shape_defaults_and_huge_dimensions_are_safe() {
        let shape = root(r#"<wps:wsp xmlns:wps="wps"><wps:cNvPr id="x"/><wps:txbx/></wps:wsp>"#);
        let text_box = parse_text_box_from_shape(
            &shape,
            ImageSize {
                width: 1.0e99,
                height: 0.0,
            },
            None,
            None,
        )
        .unwrap();
        assert_eq!(text_box.id.as_deref(), Some("x"));
        assert!(!is_floating_text_box(&text_box));
        assert_eq!(get_text_box_margins_px(&text_box).left, Some(10.0));
    }

    #[test]
    fn shared_body_parser_preserves_text_box_defaults_and_numeric_limits() {
        for (body, expected) in [
            ("", serde_json::Value::Null),
            ("<wps:bodyPr/>", serde_json::Value::Null),
            (
                r#"<wps:bodyPr upright="0" anchorCtr="false" fromWordArt="off" rot="6000000000000000000" numCol="1000000000001" lIns="1000000000001"/>"#,
                serde_json::Value::Null,
            ),
            (
                r#"<wps:bodyPr vert="vert270" rot="58432" upright="1" anchor="ctr" anchorCtr="true" numCol="2tail" spcCol="123" wrap="none" horzOverflow="clip" vertOverflow="ellipsis" lIns="0" tIns="91440" fromWordArt="on"><a:normAutofit fontScale="80000" lnSpcReduction="1000000000001"/><a:prstTxWarp prst="textWave1"/></wps:bodyPr>"#,
                serde_json::json!({
                    "vertical": "vertical270",
                    "rotation": 58432.0 / 60000.0,
                    "upright": true,
                    "anchor": "middle",
                    "anchorCenter": true,
                    "columns": 2.0,
                    "columnSpacing": 123.0,
                    "wrap": "none",
                    "horizontalOverflow": "clip",
                    "verticalOverflow": "ellipsis",
                    "margins": {"left": 0.0, "top": 91440.0},
                    "autoFit": "normal",
                    "fontScale": 80000.0,
                    "fromWordArt": true,
                    "presetTextWarp": "textWave1",
                }),
            ),
        ] {
            let shape = root(&format!(
                r#"<wps:wsp xmlns:wps="wps" xmlns:a="a"><wps:txbx/>{body}</wps:wsp>"#
            ));
            let text_box = parse_text_box_from_shape(
                &shape,
                ImageSize {
                    width: 1.0,
                    height: 1.0,
                },
                None,
                None,
            )
            .unwrap();
            assert_eq!(
                serde_json::to_value(&text_box.body_properties).unwrap(),
                expected,
                "{body}"
            );
            assert_eq!(
                serde_json::to_value(&text_box.margins).unwrap(),
                expected.get("margins").cloned().unwrap_or_default(),
                "{body}"
            );
        }
    }
}
