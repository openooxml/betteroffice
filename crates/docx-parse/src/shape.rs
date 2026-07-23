//! DrawingML shapes, detailed paint, and bounded group/canvas scenes.

use serde::{Deserialize, Serialize};

use crate::drawingml::{
    GeometryPathCommand, ShapeFill, ShapeOutline, Transform2D, parse_color_element,
    parse_custom_geometry_path, parse_fill as parse_basic_fill, parse_gradient_fill,
    parse_line_end, parse_preset_geometry_path, parse_shape_type, parse_transform, rot_to_degrees,
};
use crate::image::{
    Image, ImagePadding, ImagePosition, ImageSize, ImageWrap, parse_anchor_position,
    parse_anchor_wrap, placeholder_image,
};
use crate::media::{MediaMap, resolve_image_data};
use crate::relationships::{RelationshipMap, TargetMode};
use crate::scalars::ColorValue;
use crate::xml::{XmlElement, parse_javascript_integer_prefix};

const MAX_SCENE_DEPTH: usize = 32;
const MAX_SCENE_NODES: usize = 4_096;
const MAX_SAFE_DRAWING_NUMBER: f64 = 1_000_000_000_000.0;

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RelativeRect {
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
pub struct ShapeFillPaint {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gradient_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub angle: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stops: Option<Vec<PaintStop>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path_shape: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub focus_rect: Option<RelativeRect>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotate_with_shape: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pattern_preset: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub foreground_color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub background_color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub picture: Option<Image>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_ref_index: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub src_rect: Option<RelativeRect>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_mode: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tile: Option<PictureTile>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stretch_rect: Option<RelativeRect>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub picture_opacity: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct PaintStop {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PictureTile {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset_x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset_y: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub scale_x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub scale_y: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub flip: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeStrokePaint {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFillPaint>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub custom_dash: Option<Vec<f64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub compound: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cap: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub join: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub miter_limit: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub head_end: Option<StrokeEnd>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tail_end: Option<StrokeEnd>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_ref_index: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct StrokeEnd {
    #[serde(rename = "type", skip_serializing_if = "Option::is_none")]
    pub end_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub length: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeEffect {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub blur_radius: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub distance: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub direction: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub raw_name: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeTextBodyProperties {
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
    pub margins: RelativeRect,
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

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeTextBody {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_center: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub auto_fit: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margins: Option<RelativeRect>,
    pub content: Vec<serde_json::Value>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SceneTransform {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset_x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset_y: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub child_offset_x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub child_offset_y: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub child_width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub child_height: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub flip_h: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub flip_v: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DrawingSceneNode {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shape: Option<Box<Shape>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub children: Option<Vec<DrawingSceneNode>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub transform: Option<SceneTransform>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFillPaint>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub effects: Option<Vec<ShapeEffect>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct DrawingScene {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub version: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub root: Option<DrawingSceneNode>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub decorative: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Shape {
    #[serde(rename = "type")]
    pub shape_kind: String,
    pub shape_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub size: ImageSize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset: Option<ShapeOffset>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ImagePosition>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap: Option<ImageWrap>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFill>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline: Option<ShapeOutline>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub transform: Option<Transform2D>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_body: Option<ShapeTextBody>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub geometry_path: Option<Vec<GeometryPathCommand>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub children: Option<Vec<Shape>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub custom_geometry: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub scene: Option<Box<DrawingScene>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_paint: Option<ShapeFillPaint>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_paint: Option<ShapeStrokePaint>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub effects: Option<Vec<ShapeEffect>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub effect_extent: Option<ImagePadding>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_body_properties: Option<ShapeTextBodyProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub decorative: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relative_height: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ShapeOffset {
    pub x: f64,
    pub y: f64,
}

impl Shape {
    pub(crate) fn empty(shape_type: String, size: ImageSize) -> Self {
        Self {
            shape_kind: "shape".into(),
            shape_type,
            id: None,
            name: None,
            size,
            offset: None,
            position: None,
            wrap: None,
            fill: None,
            outline: None,
            transform: None,
            text_body: None,
            geometry_path: None,
            children: None,
            custom_geometry: None,
            scene: None,
            fill_paint: None,
            stroke_paint: None,
            effects: None,
            effect_extent: None,
            text_body_properties: None,
            title: None,
            description: None,
            decorative: None,
            hidden: None,
            relative_height: None,
        }
    }
}

pub fn parse_shape(node: &XmlElement) -> Shape {
    let non_visual = first_descendant(Some(node), "cNvPr", 0);
    let properties = direct_child(Some(node), "spPr");
    let style = direct_child(Some(node), "style");
    let text_box = direct_child(Some(node), "txbx");
    let text_content = text_box.and_then(|element| element.child_by_full_name("w:txbxContent"));
    let body_properties = direct_child(Some(node), "bodyPr")
        .or_else(|| first_descendant(direct_child(Some(node), "txBody"), "bodyPr", 0));
    let shape_type = parse_shape_type(properties);
    let custom_path = parse_custom_geometry_path(
        properties.and_then(|element| element.child_by_full_name("a:custGeom")),
    );
    let preset = properties.and_then(|element| element.child_by_full_name("a:prstGeom"));
    let has_adjustments = direct_child(preset, "avLst")
        .is_some_and(|list| list.children_by_local_name("gd").next().is_some());
    let transform =
        parse_transform(properties.and_then(|element| element.child_by_full_name("a:xfrm")));
    let preset_path = has_adjustments
        .then(|| {
            parse_preset_geometry_path(properties, transform.size.width / transform.size.height)
        })
        .flatten();
    let fill = parse_legacy_fill(properties, style);
    let outline = parse_legacy_outline(properties, style);
    let fill_paint = parse_fill_paint(properties, style);
    let stroke_paint = parse_stroke_paint(properties, style);
    let effects = parse_effects(properties, style);
    let detailed_body = body_properties.map(parse_text_body_properties);
    let mut shape = Shape::empty(
        shape_type,
        ImageSize {
            width: transform.size.width,
            height: transform.size.height,
        },
    );
    if let Some(non_visual) = non_visual {
        shape.id = non_visual.attribute(None, "id").map(str::to_owned);
        shape.name = non_visual.attribute(None, "name").map(str::to_owned);
        shape.title = non_visual.attribute(None, "title").map(str::to_owned);
        shape.description = non_visual.attribute(None, "descr").map(str::to_owned);
        shape.hidden = bool_attribute(Some(non_visual), "hidden");
        shape.decorative = bool_attribute(Some(non_visual), "decorative");
    }
    shape.fill = fill;
    shape.outline = outline;
    shape.transform = transform.transform;
    shape.geometry_path = custom_path.or(preset_path);
    if let Some(paint) = fill_paint {
        if matches!(paint.kind.as_deref(), Some("pattern" | "picture" | "theme"))
            || paint.focus_rect.is_some()
            || paint.rotate_with_shape.is_some()
        {
            shape.fill_paint = Some(paint);
        }
    }
    if let Some(paint) = stroke_paint {
        let detailed = paint.custom_dash.is_some()
            || paint.compound.is_some()
            || paint.alignment.is_some()
            || paint.miter_limit.is_some()
            || paint.theme_ref_index.is_some()
            || paint
                .fill
                .as_ref()
                .and_then(|fill| fill.kind.as_deref())
                .is_some_and(|kind| kind != "solid");
        if detailed {
            shape.stroke_paint = Some(paint);
        }
    }
    shape.effects = effects;
    shape.text_body_properties = detailed_body.clone();
    if text_content.is_some() || body_properties.is_some() {
        let content = text_content
            .map(|element| {
                element
                    .children_by_local_name("p")
                    .map(|_| serde_json::json!({"type":"paragraph","formatting":{},"content":[]}))
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        let detailed = detailed_body.unwrap_or_default();
        let margins = relative_rect_nonempty(&detailed.margins).then_some(detailed.margins);
        shape.text_body = Some(ShapeTextBody {
            vertical: detailed
                .vertical
                .as_deref()
                .filter(|value| *value != "horizontal")
                .map(|_| true),
            rotation: detailed.rotation,
            anchor: detailed.anchor,
            anchor_center: detailed.anchor_center,
            auto_fit: detailed.auto_fit,
            margins,
            content,
        });
    }
    shape
}

fn parse_legacy_fill(
    properties: Option<&XmlElement>,
    style: Option<&XmlElement>,
) -> Option<ShapeFill> {
    if let Some(mut fill) = parse_basic_fill(properties) {
        if fill.fill_type == "gradient" {
            if let Some(gradient) =
                properties.and_then(|element| element.child_by_full_name("a:gradFill"))
            {
                fill = parse_gradient_fill(gradient);
            }
        }
        return Some(fill);
    }
    if let Some(properties) = properties {
        if properties.child_by_full_name("a:pattFill").is_some() {
            return Some(ShapeFill {
                fill_type: "pattern".into(),
                color: None,
                gradient: None,
            });
        }
        if properties.child_by_full_name("a:blipFill").is_some() {
            return Some(ShapeFill {
                fill_type: "picture".into(),
                color: None,
                gradient: None,
            });
        }
    }
    let reference = style.and_then(|element| element.child_by_full_name("a:fillRef"))?;
    if reference.attribute(None, "idx") == Some("0") {
        return Some(ShapeFill {
            fill_type: "none".into(),
            color: None,
            gradient: None,
        });
    }
    parse_color_element(Some(reference)).map(|color| ShapeFill {
        fill_type: "solid".into(),
        color: Some(color),
        gradient: None,
    })
}

fn parse_fill_paint(
    properties: Option<&XmlElement>,
    style: Option<&XmlElement>,
) -> Option<ShapeFillPaint> {
    if properties.is_none() && style.is_none() {
        return None;
    }
    if direct_child(properties, "noFill").is_some() {
        return Some(ShapeFillPaint {
            kind: Some("none".into()),
            ..ShapeFillPaint::default()
        });
    }
    if let Some(solid) = direct_child(properties, "solidFill") {
        return Some(ShapeFillPaint {
            kind: Some("solid".into()),
            color: parse_color_element(Some(solid)),
            ..ShapeFillPaint::default()
        });
    }
    if let Some(gradient) = direct_child(properties, "gradFill") {
        let legacy = parse_gradient_fill(gradient).gradient;
        let path = direct_child(Some(gradient), "path");
        let raw_path = path.and_then(|value| value.attribute(None, "path"));
        let focus_rect = direct_child(path, "fillToRect").map(parse_relative_rect_value);
        return Some(ShapeFillPaint {
            kind: Some("gradient".into()),
            gradient_type: Some(
                match raw_path {
                    Some("circle") => "radial",
                    Some("rect") => "rectangular",
                    Some(_) => "path",
                    None => "linear",
                }
                .into(),
            ),
            angle: legacy.as_ref().and_then(|value| value.angle),
            stops: legacy.map(|value| {
                value
                    .stops
                    .into_iter()
                    .map(|stop| PaintStop {
                        position: Some(stop.position),
                        color: Some(stop.color),
                    })
                    .collect()
            }),
            path_shape: raw_path
                .filter(|value| matches!(*value, "circle" | "rect" | "shape"))
                .map(str::to_owned),
            focus_rect,
            rotate_with_shape: bool_attribute(Some(gradient), "rotWithShape"),
            ..ShapeFillPaint::default()
        });
    }
    if let Some(pattern) = direct_child(properties, "pattFill") {
        return Some(ShapeFillPaint {
            kind: Some("pattern".into()),
            pattern_preset: pattern.attribute(None, "prst").map(str::to_owned),
            foreground_color: direct_child(Some(pattern), "fgClr")
                .and_then(|value| parse_color_element(Some(value))),
            background_color: direct_child(Some(pattern), "bgClr")
                .and_then(|value| parse_color_element(Some(value))),
            ..ShapeFillPaint::default()
        });
    }
    if let Some(picture) = direct_child(properties, "blipFill") {
        let blip = first_descendant(Some(picture), "blip", 0);
        let relationship_id = blip.and_then(|value| {
            value
                .attribute(Some("r"), "embed")
                .or_else(|| value.attribute(None, "embed"))
        });
        let mut paint = ShapeFillPaint {
            kind: Some("picture".into()),
            picture: relationship_id.map(placeholder_image),
            rotate_with_shape: bool_attribute(Some(picture), "rotWithShape"),
            src_rect: direct_child(Some(picture), "srcRect").map(parse_relative_rect_value),
            picture_opacity: parse_blip_alpha(blip),
            ..ShapeFillPaint::default()
        };
        parse_picture_fill_mode(picture, &mut paint);
        return Some(paint);
    }
    let reference = style.and_then(|value| first_descendant(Some(value), "fillRef", 0))?;
    let index = finite_attribute(Some(reference), "idx");
    Some(ShapeFillPaint {
        kind: Some(if index == Some(0.0) { "none" } else { "theme" }.into()),
        color: (index != Some(0.0))
            .then(|| parse_color_element(Some(reference)))
            .flatten(),
        theme_ref_index: index,
        ..ShapeFillPaint::default()
    })
}

fn parse_picture_fill_mode(picture: &XmlElement, paint: &mut ShapeFillPaint) {
    if let Some(tile) = direct_child(Some(picture), "tile") {
        let scale = |name| {
            finite_attribute(Some(tile), name).map(|value| (value / 100_000.0).clamp(0.01, 100.0))
        };
        let offset = |name| {
            finite_attribute(Some(tile), name)
                .map(|value| js_round(value * 96.0 / 914_400.0).clamp(-100_000.0, 100_000.0))
        };
        let flip = tile
            .attribute(None, "flip")
            .filter(|value| matches!(*value, "x" | "y" | "xy"))
            .map(str::to_owned);
        paint.fill_mode = Some("tile".into());
        paint.tile = Some(PictureTile {
            offset_x: offset("tx"),
            offset_y: offset("ty"),
            scale_x: scale("sx"),
            scale_y: scale("sy"),
            alignment: tile.attribute(None, "algn").map(str::to_owned),
            flip,
        });
    } else if let Some(stretch) = direct_child(Some(picture), "stretch") {
        paint.fill_mode = Some("stretch".into());
        paint.stretch_rect = direct_child(Some(stretch), "fillRect").map(parse_relative_rect_value);
    }
}

fn parse_blip_alpha(blip: Option<&XmlElement>) -> Option<f64> {
    finite_attribute(direct_child(blip, "alphaModFix"), "amt")
        .map(|value| (value / 100_000.0).clamp(0.0, 1.0))
}

fn parse_relative_rect_value(element: &XmlElement) -> RelativeRect {
    let edge = |name| {
        finite_attribute(Some(element), name).map(|value| (value / 100_000.0).clamp(-10.0, 10.0))
    };
    RelativeRect {
        left: edge("l"),
        top: edge("t"),
        right: edge("r"),
        bottom: edge("b"),
    }
}
fn relative_rect_nonempty(rect: &RelativeRect) -> bool {
    rect.left.is_some() || rect.top.is_some() || rect.right.is_some() || rect.bottom.is_some()
}

fn parse_legacy_outline(
    properties: Option<&XmlElement>,
    style: Option<&XmlElement>,
) -> Option<ShapeOutline> {
    let line = properties.and_then(|value| value.child_by_full_name("a:ln"));
    let Some(line) = line else {
        let reference = style.and_then(|value| value.child_by_full_name("a:lnRef"))?;
        if reference.attribute(None, "idx") == Some("0") {
            return None;
        }
        return parse_color_element(Some(reference)).map(|color| ShapeOutline {
            width: Some(9_525.0),
            color: Some(color),
            ..ShapeOutline::default()
        });
    };
    if line.child_by_full_name("a:noFill").is_some() {
        return None;
    }
    let mut outline = ShapeOutline {
        width: finite_attribute(Some(line), "w"),
        color: line
            .child_by_full_name("a:solidFill")
            .and_then(|value| parse_color_element(Some(value))),
        style: line
            .child_by_full_name("a:prstDash")
            .and_then(|value| value.attribute(None, "val"))
            .map(str::to_owned),
        ..ShapeOutline::default()
    };
    outline.cap = match line.attribute(None, "cap") {
        Some("flat") => Some("flat".into()),
        Some("rnd") => Some("round".into()),
        Some("sq") => Some("square".into()),
        _ => None,
    };
    outline.join = if line.child_by_full_name("a:bevel").is_some() {
        Some("bevel".into())
    } else if line.child_by_full_name("a:round").is_some() {
        Some("round".into())
    } else if line.child_by_full_name("a:miter").is_some() {
        Some("miter".into())
    } else {
        None
    };
    outline.head_end = line.child_by_full_name("a:headEnd").map(parse_line_end);
    outline.tail_end = line.child_by_full_name("a:tailEnd").map(parse_line_end);
    Some(outline)
}

fn parse_stroke_paint(
    properties: Option<&XmlElement>,
    style: Option<&XmlElement>,
) -> Option<ShapeStrokePaint> {
    let line = direct_child(properties, "ln");
    let reference = style.and_then(|value| first_descendant(Some(value), "lnRef", 0));
    if line.is_none() && reference.is_none() {
        return None;
    }
    if direct_child(line, "noFill").is_some() {
        return None;
    }
    let custom_dash = direct_child(line, "custDash")
        .map(|list| {
            list.children_by_local_name("ds")
                .flat_map(|dash| {
                    match (
                        finite_attribute(Some(dash), "d"),
                        finite_attribute(Some(dash), "sp"),
                    ) {
                        (Some(d), Some(sp)) => vec![d, sp],
                        _ => Vec::new(),
                    }
                })
                .collect::<Vec<_>>()
        })
        .filter(|values| !values.is_empty());
    let compound = match line.and_then(|value| value.attribute(None, "cmpd")) {
        Some("sng") => Some("single"),
        Some("dbl") => Some("double"),
        Some("thickThin") => Some("thickThin"),
        Some("thinThick") => Some("thinThick"),
        Some("tri") => Some("triple"),
        _ => None,
    }
    .map(str::to_owned);
    let alignment = match line.and_then(|value| value.attribute(None, "algn")) {
        Some("ctr") => Some("center"),
        Some("in") => Some("inset"),
        _ => None,
    }
    .map(str::to_owned);
    let cap = match line.and_then(|value| value.attribute(None, "cap")) {
        Some("flat") => Some("flat"),
        Some("rnd") => Some("round"),
        Some("sq") => Some("square"),
        _ => None,
    }
    .map(str::to_owned);
    let miter = direct_child(line, "miter");
    let dash = direct_child(line, "prstDash");
    let fill = line.and_then(|value| parse_fill_paint(Some(value), None));
    let theme_ref_index = finite_attribute(reference, "idx");
    let theme_fill = reference.map(|value| ShapeFillPaint {
        kind: Some("theme".into()),
        color: parse_color_element(Some(value)),
        theme_ref_index,
        ..ShapeFillPaint::default()
    });
    Some(ShapeStrokePaint {
        fill: fill.or(theme_fill),
        width: finite_attribute(line, "w"),
        dash: dash
            .and_then(|value| value.attribute(None, "val"))
            .map(str::to_owned),
        custom_dash,
        compound,
        alignment,
        cap,
        join: if direct_child(line, "bevel").is_some() {
            Some("bevel".into())
        } else if direct_child(line, "round").is_some() {
            Some("round".into())
        } else if miter.is_some() {
            Some("miter".into())
        } else {
            None
        },
        miter_limit: finite_attribute(miter, "lim"),
        head_end: direct_child(line, "headEnd").map(stroke_end),
        tail_end: direct_child(line, "tailEnd").map(stroke_end),
        theme_ref_index,
    })
}

fn stroke_end(element: &XmlElement) -> StrokeEnd {
    StrokeEnd {
        end_type: element.attribute(None, "type").map(str::to_owned),
        width: element.attribute(None, "w").map(str::to_owned),
        length: element.attribute(None, "len").map(str::to_owned),
    }
}

fn parse_effects(
    properties: Option<&XmlElement>,
    style: Option<&XmlElement>,
) -> Option<Vec<ShapeEffect>> {
    let mut effects = Vec::new();
    if let Some(list) = direct_child(properties, "effectLst") {
        for effect in list.child_elements().take(64) {
            let name = effect.local_name();
            let color = parse_color_element(Some(effect));
            let mut value = ShapeEffect {
                kind: Some(
                    match name {
                        "outerShdw" | "innerShdw" => "shadow",
                        "glow" => "glow",
                        "reflection" => "reflection",
                        "softEdge" => "softEdge",
                        "blur" => "blur",
                        _ => "unknown",
                    }
                    .into(),
                ),
                color: None,
                ..ShapeEffect::default()
            };
            match name {
                "outerShdw" | "innerShdw" => {
                    value.color = color;
                    value.blur_radius = finite_attribute(Some(effect), "blurRad");
                    value.distance = finite_attribute(Some(effect), "dist");
                    value.direction =
                        finite_attribute(Some(effect), "dir").map(|number| number / 60_000.0);
                    value.size = finite_attribute(Some(effect), "sx");
                    value.raw_name = Some(name.into())
                }
                "glow" => {
                    value.color = color;
                    value.blur_radius = finite_attribute(Some(effect), "rad")
                }
                "reflection" => {
                    value.blur_radius = finite_attribute(Some(effect), "blurRad");
                    value.distance = finite_attribute(Some(effect), "dist");
                    value.direction =
                        finite_attribute(Some(effect), "dir").map(|number| number / 60_000.0);
                    value.size = finite_attribute(Some(effect), "sy")
                }
                "softEdge" | "blur" => value.blur_radius = finite_attribute(Some(effect), "rad"),
                _ => value.raw_name = Some(name.into()),
            }
            effects.push(value)
        }
    }
    if direct_child(properties, "effectDag").is_some() {
        effects.push(ShapeEffect {
            kind: Some("unknown".into()),
            raw_name: Some("effectDag".into()),
            ..ShapeEffect::default()
        })
    }
    if let Some(reference) = style.and_then(|value| first_descendant(Some(value), "effectRef", 0)) {
        effects.push(ShapeEffect {
            kind: Some("unknown".into()),
            color: parse_color_element(Some(reference)),
            raw_name: Some(format!(
                "effectRef:{}",
                reference.attribute(None, "idx").unwrap_or_default()
            )),
            ..ShapeEffect::default()
        })
    }
    (!effects.is_empty()).then_some(effects)
}

fn parse_text_body_properties(element: &XmlElement) -> ShapeTextBodyProperties {
    let vertical = match element.attribute(None, "vert") {
        Some("horz") => Some("horizontal"),
        Some("vert") => Some("vertical"),
        Some("vert270") => Some("vertical270"),
        Some("wordArtVert" | "wordArtVertRtl") => Some("wordArtVertical"),
        Some("eaVert") => Some("eastAsianVertical"),
        Some("mongolianVert") => Some("mongolianVertical"),
        _ => None,
    }
    .map(str::to_owned);
    let anchor = match element.attribute(None, "anchor") {
        Some("t") => Some("top"),
        Some("ctr") => Some("middle"),
        Some("b") => Some("bottom"),
        Some("dist") => Some("distributed"),
        Some("just") => Some("justified"),
        _ => None,
    }
    .map(str::to_owned);
    let normal = direct_child(Some(element), "normAutofit");
    let preset = direct_child(Some(element), "prstTxWarp");
    ShapeTextBodyProperties {
        vertical,
        rotation: rot_to_degrees(element.attribute(None, "rot")),
        upright: bool_attribute(Some(element), "upright"),
        anchor,
        anchor_center: bool_attribute(Some(element), "anchorCtr"),
        columns: finite_attribute(Some(element), "numCol"),
        column_spacing: finite_attribute(Some(element), "spcCol"),
        wrap: match element.attribute(None, "wrap") {
            Some("none") => Some("none".into()),
            Some("square") => Some("square".into()),
            _ => None,
        },
        horizontal_overflow: match element.attribute(None, "horzOverflow") {
            Some("clip" | "overflow") => element.attribute(None, "horzOverflow").map(str::to_owned),
            _ => None,
        },
        vertical_overflow: match element.attribute(None, "vertOverflow") {
            Some("clip" | "ellipsis" | "overflow") => {
                element.attribute(None, "vertOverflow").map(str::to_owned)
            }
            _ => None,
        },
        margins: RelativeRect {
            left: finite_attribute(Some(element), "lIns"),
            right: finite_attribute(Some(element), "rIns"),
            top: finite_attribute(Some(element), "tIns"),
            bottom: finite_attribute(Some(element), "bIns"),
        },
        auto_fit: if direct_child(Some(element), "noAutofit").is_some() {
            Some("none".into())
        } else if normal.is_some() {
            Some("normal".into())
        } else if direct_child(Some(element), "spAutoFit").is_some() {
            Some("shape".into())
        } else {
            None
        },
        font_scale: finite_attribute(normal, "fontScale"),
        line_spacing_reduction: finite_attribute(normal, "lnSpcReduction"),
        from_word_art: bool_attribute(Some(element), "fromWordArt"),
        preset_text_warp: preset
            .and_then(|value| value.attribute(None, "prst"))
            .map(str::to_owned),
    }
}

pub fn parse_shape_from_drawing(drawing: &XmlElement) -> Option<Shape> {
    let container = drawing
        .child_elements()
        .find(|value| matches!(value.name.as_str(), "wp:inline" | "wp:anchor"))?;
    let is_anchor = container.name == "wp:anchor";
    let graphic = container.child_by_full_name("a:graphic")?;
    let data = graphic.child_by_full_name("a:graphicData")?;
    let root = data.child_elements().find(|value| {
        is_scene_shape(value.local_name())
            || is_scene_group(value.local_name())
            || value.local_name() == "wpc"
    })?;
    let mut scene = None;
    let mut shape = if is_scene_shape(root.local_name()) {
        parse_shape(root)
    } else {
        let parsed = parse_drawing_scene_tree(root);
        let legacy = scene_node_to_legacy(parsed.root.as_ref()?);
        scene = Some(parsed);
        legacy?
    };
    if let Some(extent) = container.child_by_full_name("wp:extent") {
        shape.size = ImageSize {
            width: numeric_attribute(Some(extent), "cx").unwrap_or(0.0),
            height: numeric_attribute(Some(extent), "cy").unwrap_or(0.0),
        }
    }
    if is_anchor {
        shape.position = parse_anchor_position(container);
        shape.wrap = Some(parse_anchor_wrap(container))
    }
    if let Some(properties) = container.child_by_full_name("wp:docPr") {
        let title = properties.attribute(None, "title").map(str::to_owned);
        let description = properties.attribute(None, "descr").map(str::to_owned);
        let hidden = bool_attribute(Some(properties), "hidden");
        let decorative = bool_attribute(Some(properties), "decorative");
        if let Some(value) = properties
            .attribute(None, "id")
            .filter(|value| !value.is_empty())
        {
            shape.id = Some(value.into())
        }
        if let Some(value) = properties
            .attribute(None, "name")
            .filter(|value| !value.is_empty())
        {
            shape.name = Some(value.into())
        }
        if title.as_deref().is_some_and(|value| !value.is_empty()) {
            shape.title = title.clone()
        }
        if description
            .as_deref()
            .is_some_and(|value| !value.is_empty())
        {
            shape.description = description.clone()
        }
        if hidden.is_some() {
            shape.hidden = hidden
        }
        if decorative.is_some() {
            shape.decorative = decorative
        }
        if let Some(scene) = &mut scene {
            scene.title = title.or(scene.title.take());
            scene.description = description.or(scene.description.take());
            scene.hidden = hidden.or(scene.hidden);
            scene.decorative = decorative.or(scene.decorative)
        }
    }
    if let Some(extent) = container.child_by_full_name("wp:effectExtent") {
        shape.effect_extent = Some(ImagePadding {
            left: finite_attribute(Some(extent), "l"),
            top: finite_attribute(Some(extent), "t"),
            right: finite_attribute(Some(extent), "r"),
            bottom: finite_attribute(Some(extent), "b"),
        })
    }
    if let Some(position) = &shape.position {
        shape.relative_height = position.relative_height;
        shape.hidden = position.hidden.or(shape.hidden)
    }
    if let Some(scene) = scene {
        shape.scene = Some(Box::new(scene))
    }
    Some(shape)
}

pub fn is_shape_drawing(drawing: &XmlElement) -> bool {
    let Some(container) = drawing
        .child_elements()
        .find(|value| matches!(value.name.as_str(), "wp:inline" | "wp:anchor"))
    else {
        return false;
    };
    let Some(graphic) = container.child_by_full_name("a:graphic") else {
        return false;
    };
    let Some(data) = graphic.child_by_full_name("a:graphicData") else {
        return false;
    };
    data.child_elements().any(|value| {
        matches!(
            value.local_name(),
            "wsp" | "wgp" | "wpc" | "spTree" | "grpSp"
        )
    })
}

pub fn resolve_shape_fill_pictures(
    shape: &mut Shape,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
) {
    let (Some(relationships), Some(media)) = (relationships, media) else {
        return;
    };
    resolve_shape_pictures(shape, relationships, media, 0)
}
fn resolve_shape_pictures(
    shape: &mut Shape,
    relationships: &RelationshipMap,
    media: &MediaMap,
    depth: usize,
) {
    if depth > MAX_SCENE_DEPTH {
        return;
    }
    resolve_paint_picture(shape.fill_paint.as_mut(), relationships, media);
    if let Some(stroke) = &mut shape.stroke_paint {
        resolve_paint_picture(stroke.fill.as_mut(), relationships, media)
    }
    if let Some(children) = &mut shape.children {
        for child in children {
            resolve_shape_pictures(child, relationships, media, depth + 1)
        }
    }
    if let Some(root) = shape.scene.as_mut().and_then(|scene| scene.root.as_mut()) {
        resolve_scene_pictures(root, relationships, media, depth + 1)
    }
}
fn resolve_scene_pictures(
    node: &mut DrawingSceneNode,
    relationships: &RelationshipMap,
    media: &MediaMap,
    depth: usize,
) {
    if depth > MAX_SCENE_DEPTH {
        return;
    }
    resolve_paint_picture(node.fill.as_mut(), relationships, media);
    if let Some(shape) = &mut node.shape {
        resolve_shape_pictures(shape, relationships, media, depth + 1)
    }
    if let Some(children) = &mut node.children {
        for child in children {
            resolve_scene_pictures(child, relationships, media, depth + 1)
        }
    }
}
fn resolve_paint_picture(
    paint: Option<&mut ShapeFillPaint>,
    relationships: &RelationshipMap,
    media: &MediaMap,
) {
    let Some(image) = paint.and_then(|value| value.picture.as_mut()) else {
        return;
    };
    if image.relationship_id.is_empty() || image.src.is_some() {
        return;
    }
    let Some(relationship) = relationships.get(&image.relationship_id) else {
        return;
    };
    if relationship.target_mode == Some(TargetMode::External) {
        return;
    }
    let resolved = resolve_image_data(&image.relationship_id, Some(relationships), Some(media));
    if resolved
        .src
        .as_deref()
        .is_some_and(|value| value.starts_with("data:") || value.starts_with("blob:"))
    {
        image.src = resolved.src
    }
}

pub fn parse_drawing_scene_tree(root: &XmlElement) -> DrawingScene {
    let mut count = 0;
    let parsed = if is_scene_shape(root.local_name())
        || is_scene_group(root.local_name())
        || root.local_name() == "wpc"
    {
        parse_scene_node(root, Matrix::IDENTITY, &mut count, 0)
    } else {
        let children = root
            .child_elements()
            .filter_map(|child| parse_scene_node(child, Matrix::IDENTITY, &mut count, 0))
            .collect::<Vec<_>>();
        if children.len() == 1 {
            children.into_iter().next()
        } else {
            Some(DrawingSceneNode {
                kind: Some("canvas".into()),
                children: Some(children),
                ..DrawingSceneNode::default()
            })
        }
    };
    DrawingScene {
        version: Some(1.0),
        root: Some(parsed.unwrap_or_else(|| DrawingSceneNode {
            kind: Some("canvas".into()),
            children: Some(Vec::new()),
            ..DrawingSceneNode::default()
        })),
        ..DrawingScene::default()
    }
}

fn parse_scene_node(
    element: &XmlElement,
    parent: Matrix,
    count: &mut usize,
    depth: usize,
) -> Option<DrawingSceneNode> {
    if depth > MAX_SCENE_DEPTH || *count >= MAX_SCENE_NODES {
        return None;
    }
    *count += 1;
    let local = element.local_name();
    let (id, name) = non_visual_properties(element);
    if is_scene_shape(local) {
        let mut shape = parse_shape(element);
        let transform_element = scene_transform(element);
        let off = direct_child(transform_element, "off");
        shape.offset = Some(ShapeOffset {
            x: finite_attribute(off, "x").unwrap_or(0.0),
            y: finite_attribute(off, "y").unwrap_or(0.0),
        });
        let matrix = local_shape_matrix(parent, &shape);
        let mut transform =
            matrix_contract(matrix, Some(shape.size.width), Some(shape.size.height));
        let parent_rotation = parent.b.atan2(parent.a) * 180.0 / std::f64::consts::PI;
        let parent_flipped = parent.a * parent.d - parent.b * parent.c < 0.0;
        shape.offset = Some(ShapeOffset {
            x: transform.offset_x.unwrap_or(0.0),
            y: transform.offset_y.unwrap_or(0.0),
        });
        shape.size = ImageSize {
            width: transform.width.unwrap_or(shape.size.width),
            height: transform.height.unwrap_or(shape.size.height),
        };
        let own = shape.transform.take().unwrap_or_default();
        shape.transform = Some(Transform2D {
            rotation: Some(parent_rotation + own.rotation.unwrap_or(0.0)),
            flip_h: (parent_flipped != own.flip_h.unwrap_or(false)).then_some(true),
            flip_v: own.flip_v.filter(|value| *value),
        });
        transform.rotation = shape.transform.as_ref().and_then(|value| value.rotation);
        transform.flip_h = shape.transform.as_ref().and_then(|value| value.flip_h);
        transform.flip_v = shape.transform.as_ref().and_then(|value| value.flip_v);
        return Some(DrawingSceneNode {
            kind: Some("shape".into()),
            id,
            name,
            shape: Some(Box::new(shape)),
            transform: Some(transform),
            ..DrawingSceneNode::default()
        });
    }
    if is_scene_group(local) || local == "wpc" {
        let transform_element = scene_transform(element);
        let matrix = group_matrix(parent, transform_element);
        let extent = direct_child(transform_element, "ext");
        let child_offset = direct_child(transform_element, "chOff");
        let child_extent = direct_child(transform_element, "chExt");
        let properties =
            direct_child(Some(element), "grpSpPr").or_else(|| direct_child(Some(element), "spPr"));
        let off = direct_child(transform_element, "off");
        let parent_offset = parent.apply(
            finite_attribute(off, "x").unwrap_or(0.0),
            finite_attribute(off, "y").unwrap_or(0.0),
        );
        let scale_x = parent.a.hypot(parent.b);
        let scale_y = parent.c.hypot(parent.d);
        let parent_rotation = parent.b.atan2(parent.a) * 180.0 / std::f64::consts::PI;
        let authored_flip =
            transform_element.and_then(|value| value.attribute(None, "flipH")) == Some("1");
        let transform = SceneTransform {
            offset_x: Some(parent_offset.0),
            offset_y: Some(parent_offset.1),
            width: finite_attribute(extent, "cx").map(|value| value * scale_x),
            height: finite_attribute(extent, "cy").map(|value| value * scale_y),
            rotation: Some(
                parent_rotation
                    + rot_to_degrees(
                        transform_element.and_then(|value| value.attribute(None, "rot")),
                    )
                    .unwrap_or(0.0),
            ),
            flip_h: ((parent.a * parent.d - parent.b * parent.c < 0.0) != authored_flip)
                .then_some(true),
            flip_v: (transform_element.and_then(|value| value.attribute(None, "flipV"))
                == Some("1"))
            .then_some(true),
            child_offset_x: finite_attribute(child_offset, "x"),
            child_offset_y: finite_attribute(child_offset, "y"),
            child_width: finite_attribute(child_extent, "cx"),
            child_height: finite_attribute(child_extent, "cy"),
        };
        let children = element
            .child_elements()
            .filter_map(|child| parse_scene_node(child, matrix, count, depth + 1))
            .collect();
        return Some(DrawingSceneNode {
            kind: Some(if local == "wpc" { "canvas" } else { "group" }.into()),
            id,
            name,
            transform: Some(transform),
            fill: parse_fill_paint(properties, direct_child(Some(element), "style")),
            effects: parse_effects(properties, direct_child(Some(element), "style")),
            children: Some(children),
            ..DrawingSceneNode::default()
        });
    }
    if local == "pic" {
        return Some(DrawingSceneNode {
            kind: Some("picture".into()),
            id,
            name,
            relationship_id: relationship_id(element),
            ..DrawingSceneNode::default()
        });
    }
    if local == "graphicFrame" {
        return Some(DrawingSceneNode {
            kind: Some(
                if first_descendant(Some(element), "chart", 0).is_some() {
                    "chart"
                } else {
                    "graphicFrame"
                }
                .into(),
            ),
            id,
            name,
            relationship_id: relationship_id(element),
            transform: Some(matrix_contract(
                group_matrix(parent, scene_transform(element)),
                None,
                None,
            )),
            ..DrawingSceneNode::default()
        });
    }
    if local == "contentPart" {
        return Some(DrawingSceneNode {
            kind: Some("contentPart".into()),
            id,
            name,
            relationship_id: relationship_id(element),
            ..DrawingSceneNode::default()
        });
    }
    None
}

fn scene_node_to_legacy(node: &DrawingSceneNode) -> Option<Shape> {
    if node.kind.as_deref() == Some("shape") {
        return node.shape.as_deref().cloned();
    }
    if !matches!(node.kind.as_deref(), Some("group" | "canvas")) {
        return None;
    }
    let transform = node.transform.as_ref();
    let mut shape = Shape::empty(
        "rect".into(),
        ImageSize {
            width: transform.and_then(|value| value.width).unwrap_or(0.0),
            height: transform.and_then(|value| value.height).unwrap_or(0.0),
        },
    );
    shape.id = node.id.clone();
    shape.name = node.name.clone();
    shape.offset = Some(ShapeOffset {
        x: transform.and_then(|value| value.offset_x).unwrap_or(0.0),
        y: transform.and_then(|value| value.offset_y).unwrap_or(0.0),
    });
    shape.transform = Some(Transform2D {
        rotation: transform.and_then(|value| value.rotation),
        flip_h: transform.and_then(|value| value.flip_h),
        flip_v: transform.and_then(|value| value.flip_v),
    });
    shape.children = node
        .children
        .as_ref()
        .map(|values| values.iter().filter_map(scene_node_to_legacy).collect());
    shape.fill_paint = node.fill.clone();
    shape.effects = node.effects.clone();
    Some(shape)
}

#[derive(Clone, Copy)]
struct Matrix {
    a: f64,
    b: f64,
    c: f64,
    d: f64,
    e: f64,
    f: f64,
}
impl Matrix {
    const IDENTITY: Self = Self {
        a: 1.0,
        b: 0.0,
        c: 0.0,
        d: 1.0,
        e: 0.0,
        f: 0.0,
    };
    fn multiply(self, right: Self) -> Self {
        Self {
            a: self.a * right.a + self.c * right.b,
            b: self.b * right.a + self.d * right.b,
            c: self.a * right.c + self.c * right.d,
            d: self.b * right.c + self.d * right.d,
            e: self.a * right.e + self.c * right.f + self.e,
            f: self.b * right.e + self.d * right.f + self.f,
        }
    }
    fn translate(x: f64, y: f64) -> Self {
        Self {
            e: x,
            f: y,
            ..Self::IDENTITY
        }
    }
    fn scale(x: f64, y: f64) -> Self {
        Self {
            a: x,
            d: y,
            b: 0.0,
            c: 0.0,
            e: 0.0,
            f: 0.0,
        }
    }
    fn rotate(degrees: f64) -> Self {
        let radians = degrees * std::f64::consts::PI / 180.0;
        Self {
            a: radians.cos(),
            b: radians.sin(),
            c: -radians.sin(),
            d: radians.cos(),
            e: 0.0,
            f: 0.0,
        }
    }
    fn apply(self, x: f64, y: f64) -> (f64, f64) {
        (
            self.a * x + self.c * y + self.e,
            self.b * x + self.d * y + self.f,
        )
    }
}

fn scene_transform(element: &XmlElement) -> Option<&XmlElement> {
    let properties = direct_child(Some(element), "grpSpPr")
        .or_else(|| direct_child(Some(element), "spPr"))
        .or_else(|| direct_child(Some(element), "xfrm"));
    if properties.is_some_and(|value| value.local_name() == "xfrm") {
        properties
    } else {
        direct_child(properties, "xfrm")
    }
}
fn group_matrix(parent: Matrix, transform: Option<&XmlElement>) -> Matrix {
    let Some(transform) = transform else {
        return parent;
    };
    let off = direct_child(Some(transform), "off");
    let extent = direct_child(Some(transform), "ext");
    let child_off = direct_child(Some(transform), "chOff");
    let child_extent = direct_child(Some(transform), "chExt");
    let x = finite_attribute(off, "x").unwrap_or(0.0);
    let y = finite_attribute(off, "y").unwrap_or(0.0);
    let width = finite_attribute(extent, "cx").unwrap_or(1.0);
    let height = finite_attribute(extent, "cy").unwrap_or(1.0);
    let child_x = finite_attribute(child_off, "x").unwrap_or(0.0);
    let child_y = finite_attribute(child_off, "y").unwrap_or(0.0);
    let child_width = nonzero(finite_attribute(child_extent, "cx").unwrap_or(width));
    let child_height = nonzero(finite_attribute(child_extent, "cy").unwrap_or(height));
    let rotation = rot_to_degrees(transform.attribute(None, "rot")).unwrap_or(0.0);
    let flip_x = if transform.attribute(None, "flipH") == Some("1") {
        -1.0
    } else {
        1.0
    };
    let flip_y = if transform.attribute(None, "flipV") == Some("1") {
        -1.0
    } else {
        1.0
    };
    let center_x = width / 2.0;
    let center_y = height / 2.0;
    let mut local = Matrix::translate(-child_x, -child_y);
    local = Matrix::scale(width / child_width, height / child_height).multiply(local);
    local = Matrix::translate(-center_x, -center_y).multiply(local);
    local = Matrix::scale(flip_x, flip_y).multiply(local);
    local = Matrix::rotate(rotation).multiply(local);
    local = Matrix::translate(x + center_x, y + center_y).multiply(local);
    parent.multiply(local)
}
fn matrix_contract(matrix: Matrix, width: Option<f64>, height: Option<f64>) -> SceneTransform {
    let origin = matrix.apply(0.0, 0.0);
    let scale_x = matrix.a.hypot(matrix.b);
    let scale_y = matrix.c.hypot(matrix.d);
    SceneTransform {
        offset_x: Some(origin.0),
        offset_y: Some(origin.1),
        width: width.map(|value| (value * scale_x).abs()),
        height: height.map(|value| (value * scale_y).abs()),
        rotation: Some(matrix.b.atan2(matrix.a) * 180.0 / std::f64::consts::PI),
        flip_h: (matrix.a * matrix.d - matrix.b * matrix.c < 0.0).then_some(true),
        ..SceneTransform::default()
    }
}
fn local_shape_matrix(parent: Matrix, shape: &Shape) -> Matrix {
    let offset_x = shape.offset.as_ref().map(|value| value.x).unwrap_or(0.0);
    let offset_y = shape.offset.as_ref().map(|value| value.y).unwrap_or(0.0);
    let center_x = shape.size.width / 2.0;
    let center_y = shape.size.height / 2.0;
    let transform = shape.transform.as_ref();
    let mut local = Matrix::translate(-center_x, -center_y);
    local = Matrix::scale(
        if transform.and_then(|value| value.flip_h) == Some(true) {
            -1.0
        } else {
            1.0
        },
        if transform.and_then(|value| value.flip_v) == Some(true) {
            -1.0
        } else {
            1.0
        },
    )
    .multiply(local);
    local =
        Matrix::rotate(transform.and_then(|value| value.rotation).unwrap_or(0.0)).multiply(local);
    local = Matrix::translate(offset_x + center_x, offset_y + center_y).multiply(local);
    parent.multiply(local)
}

fn non_visual_properties(element: &XmlElement) -> (Option<String>, Option<String>) {
    let container = element.child_elements().find(|value| {
        matches!(
            value.local_name(),
            "nvSpPr" | "cNvSpPr" | "nvGrpSpPr" | "cNvGrpSpPr" | "nvPicPr" | "nvGraphicFramePr"
        )
    });
    let properties =
        direct_child(Some(element), "cNvPr").or_else(|| first_descendant(container, "cNvPr", 0));
    (
        properties
            .and_then(|value| value.attribute(None, "id"))
            .map(str::to_owned),
        properties
            .and_then(|value| value.attribute(None, "name"))
            .map(str::to_owned),
    )
}
fn relationship_id(element: &XmlElement) -> Option<String> {
    let target = first_descendant(Some(element), "chart", 0)
        .or_else(|| first_descendant(Some(element), "blip", 0))
        .or_else(|| first_descendant(Some(element), "contentPart", 0))?;
    target
        .attribute(Some("r"), "id")
        .or_else(|| target.attribute(Some("r"), "embed"))
        .or_else(|| target.attribute(None, "id"))
        .or_else(|| target.attribute(None, "embed"))
        .map(str::to_owned)
}
fn is_scene_shape(name: &str) -> bool {
    matches!(name, "wsp" | "sp" | "cxnSp")
}
fn is_scene_group(name: &str) -> bool {
    matches!(name, "wgp" | "grpSp" | "spTree")
}
fn direct_child<'a>(parent: Option<&'a XmlElement>, name: &str) -> Option<&'a XmlElement> {
    parent?
        .child_elements()
        .find(|value| value.local_name() == name)
}
fn first_descendant<'a>(
    parent: Option<&'a XmlElement>,
    name: &str,
    depth: usize,
) -> Option<&'a XmlElement> {
    if depth > MAX_SCENE_DEPTH {
        return None;
    }
    for child in parent?.child_elements() {
        if child.local_name() == name {
            return Some(child);
        }
        if let Some(found) = first_descendant(Some(child), name, depth + 1) {
            return Some(found);
        }
    }
    None
}
fn finite_attribute(element: Option<&XmlElement>, name: &str) -> Option<f64> {
    numeric_attribute(element, name)
        .filter(|value| value.is_finite() && value.abs() <= MAX_SAFE_DRAWING_NUMBER)
}
fn numeric_attribute(element: Option<&XmlElement>, name: &str) -> Option<f64> {
    parse_javascript_integer_prefix(element?.attribute(None, name)?)
}
fn bool_attribute(element: Option<&XmlElement>, name: &str) -> Option<bool> {
    match element?.attribute(None, name)? {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    }
}
fn nonzero(value: f64) -> f64 {
    if value == 0.0 { 1.0 } else { value }
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
        parse_xml(xml.as_bytes(), "shape.xml", &mut ParseBudget::new(&limits))
            .unwrap()
            .root()
            .unwrap()
            .clone()
    }
    #[test]
    fn parses_shape_details_and_placeholder_text() {
        let shape = parse_shape(&root(
            r#"<wps:wsp><wps:cNvPr id="7" name="Shape"/><wps:spPr><a:xfrm rot="2700000" flipH="1"><a:off x="10" y="20"/><a:ext cx="100" cy="50"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 20000"/></a:avLst></a:prstGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill><a:ln w="12700" cap="rnd"><a:headEnd type="triangle"/></a:ln><a:effectLst><a:glow rad="5"><a:srgbClr val="00FF00"/></a:glow></a:effectLst></wps:spPr><wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx><wps:bodyPr vert="vert"/></wps:wsp>"#,
        ));
        assert_eq!(shape.shape_type, "roundRect");
        assert_eq!(shape.transform.as_ref().unwrap().rotation, Some(45.0));
        assert!(shape.geometry_path.is_some());
        assert_eq!(shape.text_body.unwrap().content.len(), 1);
        assert_eq!(shape.effects.unwrap()[0].kind.as_deref(), Some("glow"));
        assert_eq!(shape.outline.unwrap().head_end.unwrap().width, None)
    }
    #[test]
    fn drawing_extent_anchor_and_scene_limits_are_applied() {
        let drawing = root(
            r#"<w:drawing><wp:anchor relativeHeight="9"><wp:extent cx="1000" cy="500"/><wp:positionH relativeFrom="page"><wp:posOffset>2</wp:posOffset></wp:positionH><wp:positionV/><wp:wrapSquare/><wp:docPr id="3" name="Outer" hidden="1"/><a:graphic><a:graphicData><wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></wpg:grpSpPr><wps:wsp><wps:cNvPr id="4"/><wps:spPr><a:xfrm><a:off x="10" y="10"/><a:ext cx="20" cy="20"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing>"#,
        );
        let shape = parse_shape_from_drawing(&drawing).unwrap();
        assert_eq!(
            shape.size,
            ImageSize {
                width: 1000.0,
                height: 500.0
            }
        );
        assert_eq!(shape.relative_height, Some(9.0));
        assert_eq!(shape.children.as_ref().unwrap().len(), 1);
        assert_eq!(shape.scene.as_ref().unwrap().version, Some(1.0))
    }
    #[test]
    fn huge_group_numbers_are_dropped_without_non_finite_output() {
        let scene = parse_drawing_scene_tree(&root(
            "<wpg:wgp><wpg:grpSpPr><a:xfrm><a:ext cx=\"1e999\" cy=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></wpg:grpSpPr></wpg:wgp>",
        ));
        let json = serde_json::to_string(&scene).unwrap();
        assert!(!json.contains("null"));
    }
}
