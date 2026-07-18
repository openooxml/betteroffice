//! Text-box drawing metadata and story-dispatch boundary.

use serde::{Deserialize, Serialize};

use crate::drawingml::{
    ShapeFill, ShapeOutline, parse_fill, parse_outline, resolve_color_value_to_hex, rot_to_degrees,
};
use crate::image::{
    ImagePadding, ImagePosition, ImageSize, ImageWrap, parse_anchor_position, parse_anchor_wrap,
};
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
    pub size: ImageSize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ImagePosition>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap: Option<ImageWrap>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFill>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline: Option<ShapeOutline>,
    /// S4 intentionally leaves story nodes to the owning block parser.
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

/// Ordered text-box block references consumed later by the single recursive
/// story dispatcher. This replaces the TS paragraph/table callback cycle.
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

pub fn parse_text_box(drawing: &XmlElement) -> Option<TextBox> {
    let container = drawing
        .child_elements()
        .find(|child| matches!(child.name.as_str(), "wp:inline" | "wp:anchor"))?;
    let graphic = container.child_by_full_name("a:graphic")?;
    let graphic_data = graphic.child_by_full_name("a:graphicData")?;
    let shape = graphic_data.child_by_full_name("wps:wsp")?;
    shape.child_by_full_name("wps:txbx")?;
    let properties = shape.child_by_full_name("wps:spPr");
    let body = shape.child_by_full_name("wps:bodyPr");
    let extent = container.child_by_full_name("wp:extent");
    let size = ImageSize {
        width: safe_numeric_attribute(extent, "cx").unwrap_or(0.0),
        height: safe_numeric_attribute(extent, "cy").unwrap_or(0.0),
    };
    let id = container
        .child_by_full_name("wp:docPr")
        .and_then(|properties| properties.attribute(None, "id"))
        .filter(|value| !value.is_empty())
        .map(str::to_owned);
    let (margins, body_properties) = parse_body_properties(body);
    let anchored = container.name == "wp:anchor";
    Some(TextBox {
        content_type: "textBox".to_owned(),
        id,
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

pub fn parse_text_box_from_shape(
    shape: &XmlElement,
    size: ImageSize,
    position: Option<ImagePosition>,
    wrap: Option<ImageWrap>,
) -> Option<TextBox> {
    shape.child_by_full_name("wps:txbx")?;
    let properties = shape.child_by_full_name("wps:spPr");
    let body = shape.child_by_full_name("wps:bodyPr");
    let id = shape
        .child_by_full_name("wps:cNvPr")
        .and_then(|properties| properties.attribute(None, "id"))
        .filter(|value| !value.is_empty())
        .map(str::to_owned);
    let (margins, body_properties) = parse_body_properties(body);
    Some(TextBox {
        content_type: "textBox".to_owned(),
        id,
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
    let Some(body) = body else {
        return (None, None);
    };
    let margins = ImagePadding {
        left: safe_numeric_attribute(Some(body), "lIns"),
        right: safe_numeric_attribute(Some(body), "rIns"),
        top: safe_numeric_attribute(Some(body), "tIns"),
        bottom: safe_numeric_attribute(Some(body), "bIns"),
    };
    let margins = (margins.left.is_some()
        || margins.right.is_some()
        || margins.top.is_some()
        || margins.bottom.is_some())
    .then_some(margins);
    let child = |name| {
        body.child_elements()
            .find(|element| element.local_name() == name)
    };
    let normal = child("normAutofit");
    let properties = TextBoxBodyProperties {
        vertical: match body.attribute(None, "vert") {
            Some("horz") => Some("horizontal"),
            Some("vert") => Some("vertical"),
            Some("vert270") => Some("vertical270"),
            Some("wordArtVert" | "wordArtVertRtl") => Some("wordArtVertical"),
            Some("eaVert") => Some("eastAsianVertical"),
            Some("mongolianVert") => Some("mongolianVertical"),
            _ => None,
        }
        .map(str::to_owned),
        rotation: rot_to_degrees(body.attribute(None, "rot")).filter(safe_number),
        upright: (body.attribute(None, "upright") == Some("1")).then_some(true),
        anchor: match body.attribute(None, "anchor") {
            Some("t") => Some("top"),
            Some("ctr") => Some("middle"),
            Some("b") => Some("bottom"),
            Some("dist") => Some("distributed"),
            Some("just") => Some("justified"),
            _ => None,
        }
        .map(str::to_owned),
        anchor_center: (body.attribute(None, "anchorCtr") == Some("1")).then_some(true),
        columns: safe_numeric_attribute(Some(body), "numCol"),
        column_spacing: safe_numeric_attribute(Some(body), "spcCol"),
        wrap: match body.attribute(None, "wrap") {
            Some("none") => Some("none"),
            Some("square") => Some("square"),
            _ => None,
        }
        .map(str::to_owned),
        horizontal_overflow: match body.attribute(None, "horzOverflow") {
            Some("clip") => Some("clip"),
            Some("overflow") => Some("overflow"),
            _ => None,
        }
        .map(str::to_owned),
        vertical_overflow: match body.attribute(None, "vertOverflow") {
            Some("clip") => Some("clip"),
            Some("ellipsis") => Some("ellipsis"),
            Some("overflow") => Some("overflow"),
            _ => None,
        }
        .map(str::to_owned),
        margins: margins.clone(),
        auto_fit: if child("noAutofit").is_some() {
            Some("none".to_owned())
        } else if normal.is_some() {
            Some("normal".to_owned())
        } else if child("spAutoFit").is_some() {
            Some("shape".to_owned())
        } else {
            None
        },
        font_scale: safe_numeric_attribute(normal, "fontScale"),
        line_spacing_reduction: safe_numeric_attribute(normal, "lnSpcReduction"),
        from_word_art: (body.attribute(None, "fromWordArt") == Some("1")).then_some(true),
        preset_text_warp: child("prstTxWarp")
            .and_then(|element| element.attribute(None, "prst"))
            .map(str::to_owned),
    };
    let has_properties = properties.vertical.is_some()
        || properties.rotation.is_some()
        || properties.upright.is_some()
        || properties.anchor.is_some()
        || properties.anchor_center.is_some()
        || properties.columns.is_some()
        || properties.column_spacing.is_some()
        || properties.wrap.is_some()
        || properties.horizontal_overflow.is_some()
        || properties.vertical_overflow.is_some()
        || properties.margins.is_some()
        || properties.auto_fit.is_some()
        || properties.font_scale.is_some()
        || properties.line_spacing_reduction.is_some()
        || properties.from_word_art.is_some()
        || properties.preset_text_warp.is_some();
    (margins, has_properties.then_some(properties))
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
}
