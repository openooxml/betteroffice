//! SmartArt cached-scene parsing and basic-layout fallback.

use std::collections::HashMap;

use indexmap::IndexMap;

use crate::drawingml::{ShapeFill, ShapeOutline, preset_geometry_to_path};
use crate::image::{ImagePadding, ImageSize, ImageWrap, parse_anchor_position, parse_anchor_wrap};
use crate::media::MediaMap;
use crate::relationships::{
    RelationshipMap, TargetMode, parse_relationships, relationship_part_path, resolve_relative_path,
};
use crate::scalars::ColorValue;
use crate::shape::{
    DrawingScene, DrawingSceneNode, Shape, ShapeOffset, ShapeTextBody,
    parse_drawing_scene_with_sources, placed_scene_node, render_scene_connector_paths,
    resolve_shape_fill_pictures,
};
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

const MAX_DIAGRAM_DEPTH: usize = 32;
const MAX_DIAGRAM_SHAPES: usize = 4_096;
const MAX_SAFE_DRAWING_NUMBER: f64 = 1_000_000_000_000.0;

#[derive(Clone, Debug, Default)]
pub struct SmartArtContext {
    pub parts: IndexMap<String, Vec<u8>>,
    pub warnings: Vec<String>,
    pub(crate) drawing_relationships: HashMap<String, Option<RelationshipMap>>,
}

pub fn create_smart_art_context<S: AsRef<[u8]>>(all_xml: &IndexMap<String, S>) -> SmartArtContext {
    let mut parts = IndexMap::new();
    for (path, xml) in all_xml {
        if path.to_ascii_lowercase().starts_with("word/diagrams/") {
            parts.insert(normalize_package_path(path), xml.as_ref().to_vec());
        }
    }
    SmartArtContext {
        parts,
        warnings: Vec::new(),
        drawing_relationships: HashMap::new(),
    }
}

pub fn is_smart_art_drawing(drawing: &XmlElement) -> bool {
    find_descendant(drawing, "relIds", 0).is_some()
}

pub fn parse_smart_art_from_drawing(
    drawing: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
    context: Option<&mut SmartArtContext>,
    budget: &mut ParseBudget<'_>,
) -> Result<Option<Shape>, ParseError> {
    parse_smart_art(drawing, relationships, media, context, budget, false)
}

pub fn smart_art_for_rendering(
    drawing: &XmlElement,
    relationships: &RelationshipMap,
    media: &MediaMap,
    context: &mut SmartArtContext,
    budget: &mut ParseBudget<'_>,
) -> Result<Option<Shape>, ParseError> {
    parse_smart_art(
        drawing,
        Some(relationships),
        Some(media),
        Some(context),
        budget,
        true,
    )
}

fn parse_smart_art(
    drawing: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
    context: Option<&mut SmartArtContext>,
    budget: &mut ParseBudget<'_>,
    render_connectors: bool,
) -> Result<Option<Shape>, ParseError> {
    let (Some(relationships), Some(context)) = (relationships, context) else {
        return Ok(None);
    };
    let Some(relationship_ids) = find_descendant(drawing, "relIds", 0) else {
        return Ok(None);
    };
    let data_path = relationship_ids
        .attribute(Some("r"), "dm")
        .and_then(|id| relationships.get(id))
        .and_then(|relationship| resolve_internal_target(Some(&relationship.target)));
    let layout_path = relationship_ids
        .attribute(Some("r"), "lo")
        .and_then(|id| relationships.get(id))
        .and_then(|relationship| resolve_internal_target(Some(&relationship.target)));
    let size = get_drawing_extent(drawing);

    if let Some((drawing_path, drawing_xml)) = data_path
        .as_deref()
        .and_then(|path| find_pre_rendered_drawing(path, context))
    {
        let colors_path = relationship_ids
            .attribute(Some("r"), "cs")
            .and_then(|id| relationships.get(id))
            .and_then(|relationship| resolve_internal_target(Some(&relationship.target)));
        let text_colors = match (
            data_path
                .as_deref()
                .and_then(|path| context.parts.get(path)),
            colors_path
                .as_deref()
                .and_then(|path| context.parts.get(path)),
        ) {
            (Some(data), Some(colors)) => style_text_colors(data, colors, drawing_xml, budget),
            _ => HashMap::new(),
        };
        let parsed = parse_styled_drawing_scene(
            drawing_xml,
            "SmartArt drawing",
            &text_colors,
            budget,
            render_connectors,
        )?;
        if !parsed.shapes.is_empty() {
            let mut shape = diagram_group(size, parsed.shapes, parsed.scene);
            if memchr::memmem::find(drawing_xml, b"blip").is_some() {
                let path = drawing_path.to_owned();
                let own = drawing_relationships(&path, context, budget);
                resolve_shape_fill_pictures(&mut shape, own, media);
            }
            if render_connectors {
                apply_frame_metadata(&mut shape, drawing);
            } else {
                retain_drawing_source(&mut shape, drawing);
            }
            return Ok(Some(shape));
        }
    }

    if render_connectors {
        return Ok(None);
    }
    let data_xml = data_path
        .as_deref()
        .and_then(|path| context.parts.get(path))
        .cloned();
    let layout_xml = layout_path
        .as_deref()
        .and_then(|path| context.parts.get(path))
        .cloned();
    let Some(data_xml) = data_xml else {
        context
            .warnings
            .push("SmartArt diagram was skipped: no diagram data part was found.".to_owned());
        return Ok(None);
    };
    let fallback = layout_basic_diagram(&data_xml, layout_xml.as_deref(), &size, context, budget)?;
    Ok((!fallback.is_empty()).then(|| {
        let mut shape = diagram_group(size, fallback, None);
        retain_drawing_source(&mut shape, drawing);
        shape
    }))
}

pub(crate) const FRAME_FIELDS: &[&str] = &[
    "size",
    "position",
    "wrap",
    "effectExtent",
    "id",
    "name",
    "title",
    "description",
    "decorative",
    "hidden",
    "relativeHeight",
];

pub(crate) fn wire_model(value: serde_json::Value) -> serde_json::Value {
    use serde_json::Value;
    match value {
        Value::Number(number) => number.as_f64().map_or(Value::Null, |number| {
            serde_json::from_str(ryu_js::Buffer::new().format(number)).expect("finite model number")
        }),
        Value::Object(object) => Value::Object(
            object
                .into_iter()
                .map(|(key, value)| (key, wire_model(value)))
                .collect(),
        ),
        Value::Array(values) => Value::Array(values.into_iter().map(wire_model).collect()),
        value => value,
    }
}

pub(crate) fn frame_model(shape: &Shape) -> serde_json::Value {
    wire_model(serde_json::json!({
        "size": &shape.size, "position": &shape.position, "wrap": &shape.wrap,
        "effectExtent": &shape.effect_extent, "id": &shape.id, "name": &shape.name,
        "title": &shape.title, "description": &shape.description, "decorative": &shape.decorative,
        "hidden": &shape.hidden, "relativeHeight": &shape.relative_height,
    }))
}

pub(crate) fn other_model_sha256(shape: &Shape) -> String {
    other_value_sha256(serde_json::to_value(shape).expect("shape model"))
}

pub(crate) fn other_value_sha256(mut model: serde_json::Value) -> String {
    use sha2::{Digest, Sha256};
    let object = model.as_object_mut().expect("shape object");
    for key in FRAME_FIELDS
        .iter()
        .copied()
        .chain(["drawingXml", "drawingRunXml", "drawingModel"])
    {
        object.remove(key);
    }
    if let Some(scene) = object
        .get_mut("scene")
        .and_then(serde_json::Value::as_object_mut)
    {
        for key in ["title", "description", "decorative", "hidden"] {
            scene.remove(key);
        }
    }
    format!(
        "{:x}",
        Sha256::digest(crate::paragraph_patch::canonical_json_v1(&wire_model(model)).as_bytes())
    )
}

pub(crate) fn apply_frame_metadata(shape: &mut Shape, drawing: &XmlElement) {
    let Some(container) = drawing
        .child_elements()
        .find(|node| matches!(node.local_name(), "inline" | "anchor"))
    else {
        return;
    };
    shape.size = get_drawing_extent(drawing);
    shape.position = if container.local_name() == "anchor" {
        parse_anchor_position(container)
    } else {
        None
    };
    shape.wrap = Some(if container.local_name() == "anchor" {
        parse_anchor_wrap(container)
    } else {
        ImageWrap {
            wrap_type: "inline".to_owned(),
            wrap_text: None,
            polygon: None,
            dist_t: container.parse_numeric_attribute(None, "distT", 1.0),
            dist_b: container.parse_numeric_attribute(None, "distB", 1.0),
            dist_l: container.parse_numeric_attribute(None, "distL", 1.0),
            dist_r: container.parse_numeric_attribute(None, "distR", 1.0),
        }
    });
    shape.effect_extent = container
        .child_by_local_name("effectExtent")
        .map(|node| ImagePadding {
            left: node.parse_numeric_attribute(None, "l", 1.0),
            top: node.parse_numeric_attribute(None, "t", 1.0),
            right: node.parse_numeric_attribute(None, "r", 1.0),
            bottom: node.parse_numeric_attribute(None, "b", 1.0),
        });
    if let Some(properties) = container.child_by_local_name("docPr") {
        shape.id = properties.attribute(None, "id").map(str::to_owned);
        shape.name = properties.attribute(None, "name").map(str::to_owned);
    }
    crate::paragraph::apply_shape_metadata(shape, drawing);
}

fn retain_drawing_source(shape: &mut Shape, drawing: &XmlElement) {
    apply_frame_metadata(shape, drawing);
    shape.drawing_xml = Some(drawing.to_raw_inline_xml());
    shape.drawing_model = Some(
        serde_json::json!({"frame": frame_model(shape), "otherSha256": other_model_sha256(shape)}),
    );
}

pub struct DiagramDrawing {
    pub scene: Option<DrawingScene>,
    pub shapes: Vec<Shape>,
}

pub fn parse_diagram_drawing_scene(
    xml: &[u8],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<DiagramDrawing, ParseError> {
    parse_styled_drawing_scene(xml, part, &HashMap::new(), budget, false)
}

fn parse_styled_drawing_scene(
    xml: &[u8],
    part: &str,
    text_colors: &HashMap<String, ColorValue>,
    budget: &mut ParseBudget<'_>,
    render_connectors: bool,
) -> Result<DiagramDrawing, ParseError> {
    let document = parse_xml(xml, part, budget)?;
    let Some(root) = document.root() else {
        return Ok(DiagramDrawing {
            scene: None,
            shapes: Vec::new(),
        });
    };
    let source = if root.local_name() == "spTree" {
        root
    } else {
        root.child_by_local_name("spTree").unwrap_or(root)
    };
    let (mut scene, sources) = parse_drawing_scene_with_sources(source);
    if render_connectors {
        render_scene_connector_paths(scene.root.as_mut(), &mut sources.iter().copied());
    }
    let mut shapes = Vec::new();
    if let Some(frame) = collect_scene_shapes(
        scene.root.as_mut(),
        &mut shapes,
        &mut sources.into_iter(),
        text_colors,
    ) {
        scene.root = Some(DrawingSceneNode {
            kind: Some("canvas".into()),
            children: Some(scene.root.take().into_iter().chain([frame]).collect()),
            ..DrawingSceneNode::default()
        });
    }
    Ok(DiagramDrawing {
        scene: Some(scene),
        shapes,
    })
}

pub fn parse_diagram_drawing_shapes(
    xml: &[u8],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<Shape>, ParseError> {
    Ok(parse_diagram_drawing_scene(xml, part, budget)?.shapes)
}

/// Collects the scene's shapes in order, returning the text frame node to
/// place after `node` when its text sits in a box of its own.
fn collect_scene_shapes<'a>(
    node: Option<&mut DrawingSceneNode>,
    shapes: &mut Vec<Shape>,
    sources: &mut impl Iterator<Item = &'a XmlElement>,
    text_colors: &HashMap<String, ColorValue>,
) -> Option<DrawingSceneNode> {
    let node = node?;
    if node.kind.as_deref() == Some("shape")
        && let Some(shape) = node.shape.as_deref_mut()
    {
        if let Some(source) = sources.next() {
            if source.local_name() == "cxnSp" && find_descendant(source, "prstGeom", 0).is_none() {
                shape.shape_type = "straightConnector1".to_owned();
                shape.geometry_path =
                    preset_geometry_to_path("straightConnector1", &HashMap::new(), 1.0);
            }
            let font_color = source
                .attribute(None, "modelId")
                .and_then(|id| text_colors.get(id))
                .cloned()
                .or_else(|| {
                    source
                        .child_by_local_name("style")
                        .and_then(|style| style.child_by_local_name("fontRef"))
                        .and_then(|reference| {
                            crate::drawingml::parse_color_element(Some(reference))
                        })
                });
            let paragraphs =
                parse_drawing_text_body(source.child_by_local_name("txBody"), font_color.as_ref());
            if !paragraphs.is_empty() {
                let body = ShapeTextBody {
                    vertical: None,
                    rotation: None,
                    anchor: None,
                    anchor_center: None,
                    auto_fit: None,
                    margins: None,
                    content: paragraphs,
                };
                match text_frame(source, shape) {
                    Some(mut frame) => {
                        shape.text_body = None;
                        frame.text_body = Some(body);
                        let node = placed_scene_node(frame);
                        shapes.push(shape.clone());
                        shapes.extend(node.shape.as_deref().cloned());
                        return Some(node);
                    }
                    None => shape.text_body = Some(body),
                }
            }
        }
        shapes.push(shape.clone());
    }
    if let Some(children) = &mut node.children {
        let mut index = 0;
        while index < children.len() {
            if let Some(frame) =
                collect_scene_shapes(children.get_mut(index), shapes, sources, text_colors)
            {
                index += 1;
                children.insert(index, frame);
            }
            index += 1;
        }
    }
    None
}

/// The text box `dsp:txXfrm` places apart from its shape. Its rotation adds to
/// the shape's, so a layout-flipped shape keeps upright text.
fn text_frame(source: &XmlElement, shape: &Shape) -> Option<Shape> {
    let frame = source.child_by_local_name("txXfrm")?;
    let number = |element: Option<&XmlElement>, name| {
        element
            .and_then(|element| element.parse_numeric_attribute(None, name, 1.0))
            .filter(safe_number)
    };
    let offset = frame.child_by_local_name("off");
    let extent = frame.child_by_local_name("ext");
    let mut text = Shape::empty(
        "rect".to_owned(),
        ImageSize {
            width: number(extent, "cx")?,
            height: number(extent, "cy")?,
        },
    );
    text.offset = Some(ShapeOffset {
        x: number(offset, "x")?,
        y: number(offset, "y")?,
    });
    let rotation = shape
        .transform
        .as_ref()
        .and_then(|transform| transform.rotation)
        .unwrap_or(0.0)
        + frame
            .parse_numeric_attribute(None, "rot", 1.0)
            .filter(safe_number)
            .map_or(0.0, |rot| rot / 60_000.0);
    text.transform = Some(crate::drawingml::Transform2D {
        rotation: Some(rotation.rem_euclid(360.0)),
        flip_h: None,
        flip_v: None,
    });
    text.geometry_path = preset_geometry_to_path("rect", &HashMap::new(), 1.0);
    text.fill = Some(ShapeFill {
        fill_type: "none".to_owned(),
        color: None,
        gradient: None,
    });
    text.outline = Some(ShapeOutline {
        width: Some(0.0),
        ..ShapeOutline::default()
    });
    text.text_body_properties = shape.text_body_properties.clone();
    Some(text)
}

/// Each drawn shape's text colour from the colour definition's
/// `txFillClrLst`, which Word applies over the cached drawing's `fontRef`.
/// A shape showing the text of an undrawn text-only point, found through the
/// data node its text starts with, takes that point's colour.
fn style_text_colors(
    data_xml: &[u8],
    colors_xml: &[u8],
    drawing_xml: &[u8],
    budget: &mut ParseBudget<'_>,
) -> HashMap<String, ColorValue> {
    let mut colors = HashMap::new();
    let (Ok(data), Ok(definitions), Ok(drawing)) = (
        parse_xml(data_xml, "SmartArt data", budget),
        parse_xml(colors_xml, "SmartArt colors", budget),
        parse_xml(drawing_xml, "SmartArt drawing", budget),
    ) else {
        return colors;
    };
    let (Some(data), Some(definitions), Some(drawing)) =
        (data.root(), definitions.root(), drawing.root())
    else {
        return colors;
    };
    let mut drawn = Vec::new();
    collect_model_ids(drawing, &mut drawn, 0);
    let labels: HashMap<&str, &XmlElement> = definitions
        .children_by_local_name("styleLbl")
        .filter_map(|label| Some((label.attribute(None, "name")?, label)))
        .collect();
    let points: HashMap<&str, &XmlElement> = data
        .child_by_local_name("ptLst")
        .into_iter()
        .flat_map(|list| list.children_by_local_name("pt"))
        .filter_map(|point| {
            Some((
                point.attribute(None, "modelId")?,
                point.child_by_local_name("prSet")?,
            ))
        })
        .collect();
    let mut shown_by: HashMap<&str, Vec<&str>> = HashMap::new();
    for connection in data
        .child_by_local_name("cxnLst")
        .into_iter()
        .flat_map(|list| list.children_by_local_name("cxn"))
        .filter(|connection| connection.attribute(None, "type") == Some("presOf"))
    {
        if let (Some(node), Some(point)) = (
            connection.attribute(None, "srcId"),
            connection.attribute(None, "destId"),
        ) {
            shown_by.entry(node).or_default().push(point);
        }
    }
    let color = |id: &str| {
        let properties = points.get(id)?;
        let list = labels
            .get(properties.attribute(None, "presStyleLbl")?)?
            .child_by_local_name("txFillClrLst")?;
        let entries: Vec<ColorValue> = list
            .child_elements()
            .filter_map(crate::drawingml::parse_color)
            .collect();
        let index = properties
            .parse_numeric_attribute(None, "presStyleIdx", 1.0)
            .filter(|value| value.is_finite() && *value >= 0.0)
            .unwrap_or(0.0) as usize;
        let count = properties
            .parse_numeric_attribute(None, "presStyleCnt", 1.0)
            .filter(|value| value.is_finite() && *value >= 1.0)
            .unwrap_or(1.0) as usize;
        let last = entries.len().checked_sub(1)?;
        let pick = match list.attribute(None, "meth") {
            Some("cycle") => index % entries.len(),
            Some("repeat") => index.min(last),
            _ if count > 1 => (index.min(count - 1) * last + (count - 1) / 2) / (count - 1),
            _ => 0,
        };
        Some(entries[pick].clone())
    };
    let mut nodes_by_text: HashMap<String, Vec<&str>> = HashMap::new();
    for point in data
        .child_by_local_name("ptLst")
        .into_iter()
        .flat_map(|list| list.children_by_local_name("pt"))
        .filter(|point| matches!(point.attribute(None, "type"), None | Some("node")))
    {
        if let (Some(id), Some(text)) = (
            point.attribute(None, "modelId"),
            point
                .child_by_local_name("t")
                .and_then(first_paragraph_text),
        ) {
            nodes_by_text.entry(text).or_default().push(id);
        }
    }
    let drawn_ids: Vec<&str> = drawn.iter().map(|(id, _)| *id).collect();
    for (id, text) in &drawn {
        let nodes = text.as_ref().and_then(|text| nodes_by_text.get(text));
        let node = nodes.and_then(|nodes| {
            nodes
                .iter()
                .find(|node| shown_by.get(**node).is_some_and(|shown| shown.contains(id)))
                .or_else(|| nodes.first())
        });
        let text_point = node.and_then(|node| shown_by.get(node)).and_then(|shown| {
            shown
                .iter()
                .find(|point| **point != *id && !drawn_ids.contains(point))
        });
        if let Some(value) = text_point
            .and_then(|point| color(point))
            .or_else(|| color(id))
        {
            colors.insert((*id).to_owned(), value);
        }
    }
    colors
}

/// Each drawn shape's model id with the text of its first paragraph.
fn collect_model_ids<'a>(
    element: &'a XmlElement,
    ids: &mut Vec<(&'a str, Option<String>)>,
    depth: usize,
) {
    if depth > MAX_DIAGRAM_DEPTH || ids.len() >= MAX_DIAGRAM_SHAPES {
        return;
    }
    if element.local_name() == "sp"
        && let Some(id) = element.attribute(None, "modelId")
    {
        let text = element
            .child_by_local_name("txBody")
            .and_then(first_paragraph_text);
        ids.push((id, text));
        return;
    }
    for child in element.child_elements() {
        collect_model_ids(child, ids, depth + 1);
    }
}

fn first_paragraph_text(body: &XmlElement) -> Option<String> {
    body.children_by_local_name("p")
        .map(|paragraph| collect_drawing_text(paragraph, 0))
        .find(|text| !text.is_empty())
}

fn parse_drawing_text_body(
    text_body: Option<&XmlElement>,
    font_color: Option<&ColorValue>,
) -> Vec<serde_json::Value> {
    let Some(text_body) = text_body else {
        return Vec::new();
    };
    let mut numbering = HashMap::new();
    let mut paragraphs: Vec<serde_json::Value> = text_body
        .child_elements()
        .filter(|child| child.local_name() == "p")
        .filter_map(|paragraph| parse_drawing_text_paragraph(paragraph, font_color, &mut numbering))
        .collect();
    let outer_spacing = text_body
        .child_by_local_name("bodyPr")
        .and_then(|properties| properties.parse_boolean_attribute(None, "spcFirstLastPara"))
        == Some(true);
    if !outer_spacing {
        if let Some(first) = paragraphs.first_mut() {
            first["formatting"]["spaceBefore"] = serde_json::json!(0.0);
        }
        if let Some(last) = paragraphs.last_mut() {
            last["formatting"]["spaceAfter"] = serde_json::json!(0.0);
        }
    }
    paragraphs
}

/// EMU per twip.
const EMU_PER_TWIP: f64 = 635.0;
/// DrawingML text's default size, in hundredths of a point.
const DEFAULT_TEXT_SIZE: f64 = 1800.0;

fn parse_drawing_text_paragraph(
    paragraph: &XmlElement,
    font_color: Option<&ColorValue>,
    numbering: &mut HashMap<u32, usize>,
) -> Option<serde_json::Value> {
    let properties = paragraph.child_by_local_name("pPr");
    let defaults = properties.and_then(|properties| properties.child_by_local_name("defRPr"));
    let mut runs = Vec::new();
    let mut first_run = None;
    for child in paragraph.child_elements() {
        match child.local_name() {
            "r" | "fld" => {
                let text = collect_drawing_text(child, 0);
                if !text.is_empty() {
                    first_run.get_or_insert(child);
                    runs.push(drawing_run(
                        &text,
                        parse_drawing_run_formatting(child, font_color, defaults),
                    ));
                }
            }
            "br" => runs.push(drawing_run("\n", None)),
            _ => {}
        }
    }
    if runs.is_empty() {
        return None;
    }
    let attribute = |name| properties.and_then(|properties| properties.attribute(None, name));
    let number = |name| {
        properties
            .and_then(|properties| properties.parse_numeric_attribute(None, name, 1.0))
            .filter(safe_number)
    };
    let alignment = match attribute("algn") {
        Some("ctr") => "center",
        Some("r") => "right",
        Some("just" | "dist") => "both",
        _ => "left",
    };
    let size = first_run
        .and_then(|run| run.child_by_local_name("rPr"))
        .and_then(|properties| properties.parse_numeric_attribute(None, "sz", 1.0))
        .filter(safe_number)
        .unwrap_or(DEFAULT_TEXT_SIZE);
    let spacing = |name| {
        let spacing = properties?.child_by_local_name(name)?;
        if let Some(points) = spacing.child_by_local_name("spcPts") {
            return points
                .parse_numeric_attribute(None, "val", 1.0)
                .filter(safe_number)
                .map(|value| value / 5.0);
        }
        spacing
            .child_by_local_name("spcPct")?
            .parse_numeric_attribute(None, "val", 1.0)
            .filter(safe_number)
            .map(|value| value / 100_000.0 * size * 1.2 / 5.0)
    };
    let (line, line_rule) = properties
        .and_then(|properties| properties.child_by_local_name("lnSpc"))
        .and_then(|spacing| {
            if let Some(points) = spacing.child_by_local_name("spcPts") {
                let value = points
                    .parse_numeric_attribute(None, "val", 1.0)
                    .filter(safe_number)?;
                return Some((value / 5.0, "exact"));
            }
            let value = spacing
                .child_by_local_name("spcPct")?
                .parse_numeric_attribute(None, "val", 1.0)
                .filter(safe_number)?;
            Some((value / 100_000.0 * 240.0, "auto"))
        })
        .unwrap_or((240.0, "auto"));
    let left = number("marL").unwrap_or(0.0) / EMU_PER_TWIP;
    let first_line = number("indent").unwrap_or(0.0) / EMU_PER_TWIP;
    let formatting = serde_json::json!({
        "alignment": alignment,
        "spaceBefore": spacing("spcBef").unwrap_or(0.0),
        "spaceAfter": spacing("spcAft").unwrap_or(0.0),
        "lineSpacing": line,
        "lineSpacingRule": line_rule,
        "indentLeft": left,
        "indentFirstLine": first_line.abs(),
        "hangingIndent": first_line < 0.0,
    });
    let level = number("lvl").unwrap_or(0.0).clamp(0.0, 8.0) as u32;
    numbering.retain(|deeper, _| *deeper <= level);
    let mut paragraph_json = serde_json::json!({
        "type": "paragraph",
        "formatting": formatting,
        "content": runs,
    });
    if let Some(text) = bullet_text(properties, level, numbering) {
        let text_formatting = first_run
            .and_then(|run| parse_drawing_run_formatting(run, font_color, defaults))
            .unwrap_or_default();
        let bullet_property = |name| properties.and_then(|value| value.child_by_local_name(name));
        let mut bullet = serde_json::json!({ "text": text });
        if let Some(color) = bullet_property("buClr")
            .and_then(|color| crate::drawingml::parse_color_element(Some(color)))
            .and_then(|color| serde_json::to_value(color).ok())
            .or_else(|| text_formatting.get("color").cloned())
        {
            bullet["color"] = color;
        }
        let percent = bullet_property("buSzPct")
            .and_then(|value| value.parse_numeric_attribute(None, "val", 1.0))
            .filter(safe_number)
            .map_or(1.0, |value| value / 100_000.0);
        bullet["size"] = serde_json::json!(size / 100.0 * percent);
        if let Some(typeface) = bullet_property("buFont")
            .and_then(|font| font.attribute(None, "typeface"))
            .filter(|typeface| !typeface.starts_with('+'))
        {
            bullet["fontFamily"] = serde_json::json!(typeface);
        }
        paragraph_json["bullet"] = bullet;
    } else {
        numbering.remove(&level);
    }
    Some(paragraph_json)
}

/// The marker a `buChar` or `buAutoNum` paragraph starts with.
fn bullet_text(
    properties: Option<&XmlElement>,
    level: u32,
    numbering: &mut HashMap<u32, usize>,
) -> Option<String> {
    let properties = properties?;
    if let Some(character) = properties.child_by_local_name("buChar") {
        return character.attribute(None, "char").map(str::to_owned);
    }
    let scheme = properties.child_by_local_name("buAutoNum")?;
    let start = scheme
        .parse_numeric_attribute(None, "startAt", 1.0)
        .filter(|value| value.is_finite())
        .unwrap_or(1.0)
        .clamp(1.0, 32767.0) as usize;
    let count = numbering.entry(level).or_insert(0);
    *count += 1;
    let value = start.saturating_add(*count - 1);
    let kind = scheme.attribute(None, "type").unwrap_or("arabicPeriod");
    let body = ["alphaLc", "alphaUc", "romanLc", "romanUc", "arabic"]
        .into_iter()
        .find(|prefix| kind.starts_with(prefix))
        .unwrap_or("");
    let suffix = kind.strip_prefix(body).unwrap_or(kind);
    let label = match body {
        "alphaLc" => alphabetic(value).to_ascii_lowercase(),
        "alphaUc" => alphabetic(value),
        "romanLc" => roman(value).to_ascii_lowercase(),
        "romanUc" => roman(value),
        _ => value.to_string(),
    };
    Some(match suffix {
        "ParenBoth" => format!("({label})"),
        "ParenR" => format!("{label})"),
        "Plain" => label,
        _ => format!("{label}."),
    })
}

fn alphabetic(mut value: usize) -> String {
    let mut label = Vec::new();
    while value > 0 {
        value -= 1;
        label.push(b'A' + (value % 26) as u8);
        value /= 26;
    }
    label.reverse();
    String::from_utf8(label).unwrap_or_default()
}

fn roman(mut value: usize) -> String {
    const NUMERALS: [(usize, &str); 13] = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut label = String::new();
    for (amount, numeral) in NUMERALS {
        while value >= amount {
            label.push_str(numeral);
            value -= amount;
        }
    }
    label
}

fn drawing_run(text: &str, formatting: Option<serde_json::Value>) -> serde_json::Value {
    let mut run = serde_json::Map::new();
    run.insert("type".to_owned(), serde_json::json!("run"));
    run.insert(
        "content".to_owned(),
        serde_json::json!([{ "type": "text", "text": text }]),
    );
    if let Some(formatting) = formatting {
        run.insert("formatting".to_owned(), formatting);
    }
    serde_json::Value::Object(run)
}

fn collect_drawing_text(parent: &XmlElement, depth: usize) -> String {
    if depth > MAX_DIAGRAM_DEPTH {
        return String::new();
    }
    let mut text = String::new();
    for child in parent.child_elements() {
        if child.local_name() == "t" {
            text.push_str(&child.text_content());
        } else {
            text.push_str(&collect_drawing_text(child, depth + 1));
        }
    }
    text
}

fn parse_drawing_run_formatting(
    run: &XmlElement,
    font_color: Option<&ColorValue>,
    defaults: Option<&XmlElement>,
) -> Option<serde_json::Value> {
    let mut formatting = serde_json::Map::new();
    if let Some(color) = font_color.and_then(|color| serde_json::to_value(color).ok()) {
        formatting.insert("color".to_owned(), color);
    }
    for properties in defaults.into_iter().chain(run.child_by_local_name("rPr")) {
        apply_run_properties(properties, &mut formatting);
    }
    (!formatting.is_empty()).then_some(serde_json::Value::Object(formatting))
}

/// Lays `a:rPr` or `a:defRPr` over `formatting`.
fn apply_run_properties(
    properties: &XmlElement,
    formatting: &mut serde_json::Map<String, serde_json::Value>,
) {
    if let Some(bold) = properties.parse_boolean_attribute(None, "b") {
        formatting.insert("bold".to_owned(), serde_json::json!(bold));
    }
    if let Some(italic) = properties.parse_boolean_attribute(None, "i") {
        formatting.insert("italic".to_owned(), serde_json::json!(italic));
    }
    if properties
        .attribute(None, "u")
        .is_some_and(|underline| underline != "none")
    {
        formatting.insert(
            "underline".to_owned(),
            serde_json::json!({"style": "single"}),
        );
    }
    if properties
        .attribute(None, "strike")
        .is_some_and(|strike| strike != "noStrike")
    {
        formatting.insert("strike".to_owned(), serde_json::json!(true));
    }
    let kerning = properties
        .parse_numeric_attribute(None, "kern", 1.0)
        .filter(safe_number);
    let spacing = properties
        .parse_numeric_attribute(None, "spc", 1.0)
        .filter(safe_number);
    if kerning.is_some() || spacing.is_some() {
        let mut character = serde_json::Map::new();
        if let Some(kerning) = kerning.filter(|value| *value > 0.0) {
            character.insert("kerning".to_owned(), serde_json::json!(kerning / 50.0));
        }
        if let Some(spacing) = spacing.filter(|value| *value != 0.0) {
            character.insert("spacing".to_owned(), serde_json::json!(spacing / 5.0));
        }
        if !character.is_empty() {
            formatting.insert(
                "characterSpacing".to_owned(),
                serde_json::Value::Object(character),
            );
        }
    }
    match properties.attribute(None, "cap") {
        Some("all") => {
            formatting.insert("allCaps".to_owned(), serde_json::json!(true));
        }
        Some("small") => {
            formatting.insert("smallCaps".to_owned(), serde_json::json!(true));
        }
        _ => {}
    }
    if let Some(size) = properties
        .parse_numeric_attribute(None, "sz", 1.0)
        .filter(safe_number)
    {
        formatting.insert("fontSize".to_owned(), serde_json::json!(size / 50.0));
    }
    let latin = properties
        .child_by_local_name("latin")
        .and_then(|element| element.attribute(None, "typeface"));
    let east_asia = properties
        .child_by_local_name("ea")
        .and_then(|element| element.attribute(None, "typeface"));
    let complex = properties
        .child_by_local_name("cs")
        .and_then(|element| element.attribute(None, "typeface"));
    if latin.is_some() || east_asia.is_some() || complex.is_some() {
        let mut family = serde_json::Map::new();
        if let Some(latin) = latin {
            family.insert("ascii".to_owned(), serde_json::json!(latin));
            family.insert("hAnsi".to_owned(), serde_json::json!(latin));
        }
        if let Some(east_asia) = east_asia {
            family.insert("eastAsia".to_owned(), serde_json::json!(east_asia));
        }
        if let Some(complex) = complex {
            family.insert("cs".to_owned(), serde_json::json!(complex));
        }
        formatting.insert("fontFamily".to_owned(), serde_json::Value::Object(family));
    }
    if let Some(language) = properties.attribute(None, "lang") {
        formatting.insert(
            "language".to_owned(),
            serde_json::json!({"latin": language}),
        );
    }
    if let Some(solid) = properties.child_by_full_name("a:solidFill") {
        if let Some(color) = crate::drawingml::parse_color_element(Some(solid)) {
            formatting.insert(
                "color".to_owned(),
                serde_json::to_string(&color)
                    .and_then(|s| serde_json::from_str::<serde_json::Value>(&s))
                    .unwrap(),
            );
        }
    }
}

fn get_drawing_extent(drawing: &XmlElement) -> ImageSize {
    let extent = find_descendant(drawing, "extent", 0);
    ImageSize {
        width: extent
            .and_then(|element| element.parse_numeric_attribute(None, "cx", 1.0))
            .filter(safe_number)
            .unwrap_or(0.0),
        height: extent
            .and_then(|element| element.parse_numeric_attribute(None, "cy", 1.0))
            .filter(safe_number)
            .unwrap_or(0.0),
    }
}

fn diagram_group(size: ImageSize, children: Vec<Shape>, scene: Option<DrawingScene>) -> Shape {
    let mut shape = Shape::empty("rect".to_owned(), size);
    shape.fill = Some(ShapeFill {
        fill_type: "none".to_owned(),
        color: None,
        gradient: None,
    });
    shape.outline = Some(ShapeOutline {
        width: Some(0.0),
        ..ShapeOutline::default()
    });
    shape.children = Some(children);
    shape.scene = scene.map(Box::new);
    shape
}

fn normalize_package_path(path: &str) -> String {
    let mut parts: Vec<&str> = Vec::new();
    for part in path.trim_start_matches('/').split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
        }
    }
    parts.join("/")
}

fn resolve_internal_target(target: Option<&str>) -> Option<String> {
    let target = target?;
    if has_uri_scheme(target) {
        return None;
    }
    let normalized = normalize_package_path(target);
    if normalized.starts_with("word/") {
        Some(normalized)
    } else {
        Some(normalize_package_path(&format!("word/{normalized}")))
    }
}

fn has_uri_scheme(target: &str) -> bool {
    let Some((scheme, _)) = target.split_once(':') else {
        return false;
    };
    !scheme.is_empty()
        && scheme.as_bytes()[0].is_ascii_alphabetic()
        && scheme
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'.' | b'-'))
}

fn find_pre_rendered_drawing<'a>(
    data_path: &str,
    context: &'a SmartArtContext,
) -> Option<(&'a str, &'a [u8])> {
    let (directory, filename) = data_path.rsplit_once('/')?;
    let number = data_file_number(filename)?;
    let direct = format!("{directory}/drawing{number}.xml");
    if let Some((path, xml)) = context.parts.get_key_value(&direct) {
        return Some((path.as_str(), xml.as_slice()));
    }
    let suffix = format!("/drawing{number}.xml").to_ascii_lowercase();
    context
        .parts
        .iter()
        .find(|(path, _)| path.to_ascii_lowercase().ends_with(&suffix))
        .map(|(path, xml)| (path.as_str(), xml.as_slice()))
}

/// The relationships of a diagram drawing part, its pictures' references, with
/// each internal target made a package path: they are relative to that part. A
/// part is read once, on a copy of `budget`: a picture lookup must not spend
/// what the rest of the package may need, and a part it cannot read costs the
/// pictures only, with a warning.
fn drawing_relationships<'c>(
    part: &str,
    context: &'c mut SmartArtContext,
    budget: &ParseBudget<'_>,
) -> Option<&'c RelationshipMap> {
    if !context.drawing_relationships.contains_key(part) {
        let path = relationship_part_path(part);
        let xml = context
            .parts
            .iter()
            .find(|(candidate, _)| candidate.eq_ignore_ascii_case(&path))
            .map(|(_, xml)| xml);
        let parsed = match xml.map(|xml| parse_relationships(xml, &path, &mut budget.clone())) {
            None => None,
            Some(Ok(mut relationships)) => {
                relationships.retain(|_, relationship| {
                    relationship.target_mode == Some(TargetMode::External)
                        || resolve_relative_path(part, &relationship.target)
                            .map(|target| relationship.target = target)
                            .is_ok()
                });
                Some(relationships)
            }
            Some(Err(_)) => {
                context.warnings.push(format!(
                    "SmartArt pictures were left unresolved: {path} could not be read."
                ));
                None
            }
        };
        context
            .drawing_relationships
            .insert(part.to_owned(), parsed);
    }
    context.drawing_relationships.get(part)?.as_ref()
}

fn data_file_number(filename: &str) -> Option<&str> {
    let lower = filename.to_ascii_lowercase();
    let number = lower.strip_prefix("data")?.strip_suffix(".xml")?;
    (!number.is_empty() && number.bytes().all(|byte| byte.is_ascii_digit()))
        .then_some(&filename[4..filename.len() - 4])
}

fn find_descendant<'a>(
    parent: &'a XmlElement,
    local_name: &str,
    depth: usize,
) -> Option<&'a XmlElement> {
    if depth > MAX_DIAGRAM_DEPTH {
        return None;
    }
    for child in parent.child_elements() {
        if child.local_name() == local_name {
            return Some(child);
        }
        if let Some(nested) = find_descendant(child, local_name, depth + 1) {
            return Some(nested);
        }
    }
    None
}

#[derive(Clone, Copy)]
enum BasicLayout {
    List,
    Process,
    Hierarchy,
}

fn classify_layout(xml: Option<&[u8]>) -> Option<BasicLayout> {
    let lower = String::from_utf8_lossy(xml?).to_ascii_lowercase();
    if [
        "cycle", "matrix", "pyramid", "radial", "venn", "target", "gear",
    ]
    .iter()
    .any(|needle| lower.contains(needle))
    {
        return None;
    }
    if ["hierarchy", "organization chart", "org chart", "orgchart"]
        .iter()
        .any(|needle| lower.contains(needle))
    {
        return Some(BasicLayout::Hierarchy);
    }
    if ["chevron", "process"]
        .iter()
        .any(|needle| lower.contains(needle))
    {
        return Some(BasicLayout::Process);
    }
    if ["basic list", "basiclist", "list"]
        .iter()
        .any(|needle| lower.contains(needle))
    {
        return Some(BasicLayout::List);
    }
    None
}

fn layout_basic_diagram(
    data_xml: &[u8],
    layout_xml: Option<&[u8]>,
    size: &ImageSize,
    context: &mut SmartArtContext,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<Shape>, ParseError> {
    let Some(kind) = classify_layout(layout_xml) else {
        context.warnings.push("SmartArt diagram fallback was skipped: layout is not one of basic list, process, or hierarchy.".to_owned());
        return Ok(Vec::new());
    };
    let points = parse_data_model_points(data_xml, budget)?;
    if points.is_empty() {
        return Ok(Vec::new());
    }
    Ok(match kind {
        BasicLayout::Hierarchy => layout_hierarchy(&points, size),
        BasicLayout::List | BasicLayout::Process => layout_linear(&points, size, kind),
    })
}

struct DiagramPoint {
    #[allow(dead_code)]
    id: String,
    text: String,
}

fn parse_data_model_points(
    xml: &[u8],
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<DiagramPoint>, ParseError> {
    let document = parse_xml(xml, "SmartArt data", budget)?;
    let Some(root) = document.root() else {
        return Ok(Vec::new());
    };
    let mut points = Vec::new();
    collect_data_points(root, &mut points, 0);
    Ok(points)
}

fn collect_data_points(element: &XmlElement, points: &mut Vec<DiagramPoint>, depth: usize) {
    if depth > MAX_DIAGRAM_DEPTH || points.len() >= MAX_DIAGRAM_SHAPES {
        return;
    }
    if element.local_name() == "pt" {
        let text = collect_point_text(element, 0).trim().to_owned();
        if !text.is_empty() && element.attribute(None, "type") != Some("doc") {
            points.push(DiagramPoint {
                id: element
                    .attribute(None, "modelId")
                    .map(str::to_owned)
                    .unwrap_or_else(|| points.len().to_string()),
                text,
            });
        }
    }
    for child in element.child_elements() {
        collect_data_points(child, points, depth + 1);
    }
}

fn collect_point_text(element: &XmlElement, depth: usize) -> String {
    if depth > MAX_DIAGRAM_DEPTH {
        return String::new();
    }
    let mut text = String::new();
    if element.local_name() == "t" && element.child_elements().next().is_none() {
        text.push_str(&element.text_content());
    }
    for child in element.child_elements() {
        text.push_str(&collect_point_text(child, depth + 1));
    }
    text
}

fn layout_linear(points: &[DiagramPoint], size: &ImageSize, kind: BasicLayout) -> Vec<Shape> {
    let count = points.len() as f64;
    let gap = (size.width * 0.04).min(182_880.0);
    let width = ((size.width - gap * (count + 1.0)) / count).max(1.0);
    let height = (size.height * 0.45).max(1.0);
    let y = ((size.height - height) / 2.0).max(0.0);
    points
        .iter()
        .enumerate()
        .map(|(index, point)| {
            fallback_shape(
                &point.text,
                if matches!(kind, BasicLayout::Process) {
                    "chevron"
                } else {
                    "roundRect"
                },
                gap + index as f64 * (width + gap),
                y,
                width,
                height,
            )
        })
        .collect()
}

fn layout_hierarchy(points: &[DiagramPoint], size: &ImageSize) -> Vec<Shape> {
    let mut shapes = Vec::new();
    let top_width = size.width * 0.38;
    let height = size.height * 0.22;
    shapes.push(fallback_shape(
        &points[0].text,
        "roundRect",
        (size.width - top_width) / 2.0,
        size.height * 0.08,
        top_width,
        height,
    ));
    let children = &points[1..];
    if children.is_empty() {
        return shapes;
    }
    let gap = (size.width * 0.04).min(182_880.0);
    let child_width =
        ((size.width - gap * (children.len() as f64 + 1.0)) / children.len() as f64).max(1.0);
    let y = size.height * 0.58;
    for (index, child) in children.iter().enumerate() {
        shapes.push(fallback_shape(
            &child.text,
            "roundRect",
            gap + index as f64 * (child_width + gap),
            y,
            child_width,
            height,
        ));
    }
    shapes
}

fn fallback_shape(text: &str, shape_type: &str, x: f64, y: f64, width: f64, height: f64) -> Shape {
    let mut shape = Shape::empty(shape_type.to_owned(), ImageSize { width, height });
    shape.offset = Some(ShapeOffset { x, y });
    shape.fill = Some(ShapeFill {
        fill_type: "solid".to_owned(),
        color: Some(ColorValue {
            theme_color: Some("accent1".to_owned()),
            ..ColorValue::default()
        }),
        gradient: None,
    });
    shape.outline = Some(ShapeOutline {
        width: Some(9_525.0),
        color: Some(ColorValue {
            theme_color: Some("dk1".to_owned()),
            ..ColorValue::default()
        }),
        ..ShapeOutline::default()
    });
    shape.text_body = Some(ShapeTextBody {
        vertical: None,
        rotation: None,
        anchor: None,
        anchor_center: None,
        auto_fit: None,
        margins: None,
        content: vec![serde_json::json!({
            "type": "paragraph",
            "formatting": { "alignment": "center" },
            "content": [{
                "type": "run",
                "content": [{ "type": "text", "text": text }]
            }]
        })],
    });
    shape
}

fn safe_number(value: &f64) -> bool {
    value.is_finite() && value.abs() <= MAX_SAFE_DRAWING_NUMBER
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::relationships::Relationship;
    use crate::xml::ParseLimits;

    fn drawing() -> XmlElement {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        parse_xml(
            br#"<w:drawing xmlns:w="w" xmlns:wp="wp" xmlns:dgm="dgm" xmlns:r="r"><wp:inline><wp:extent cx="1000" cy="500"/><dgm:relIds r:dm="rId1" r:lo="rId2"/></wp:inline></w:drawing>"#,
            "drawing",
            &mut budget,
        )
        .unwrap()
        .root()
        .unwrap()
        .clone()
    }

    fn auto_number(attributes: &str) -> Option<String> {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let xml = format!(r#"<a:pPr xmlns:a="a"><a:buAutoNum {attributes}/></a:pPr>"#);
        let document = parse_xml(xml.as_bytes(), "pPr", &mut budget).unwrap();
        bullet_text(document.root(), 0, &mut HashMap::new())
    }

    #[test]
    fn alpha_auto_number_uses_letters() {
        let label = auto_number(r#"type="alphaLcParenR" startAt="2""#);
        assert_eq!(label.as_deref(), Some("b)"));
    }

    #[test]
    fn roman_auto_number_uses_numerals() {
        let label = auto_number(r#"type="romanUcPeriod" startAt="4""#);
        assert_eq!(label.as_deref(), Some("IV."));
    }

    #[test]
    fn huge_start_value_is_clamped() {
        let label = auto_number(r#"type="arabicPlain" startAt="99999999999999999999999""#);
        assert_eq!(label.as_deref(), Some("32767"));
    }

    #[test]
    fn run_text_defaults_to_the_shape_style_font_color() {
        let mut context = SmartArtContext::default();
        context.parts.insert(
            "word/diagrams/drawing1.xml".to_owned(),
            br#"<dsp:drawing xmlns:dsp="dsp" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp modelId="1"><dsp:nvSpPr><dsp:cNvPr id="1" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm><a:prstGeom prst="rect"/></dsp:spPr><dsp:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="2300"/><a:t>Styled</a:t></a:r><a:r><a:t>Bare</a:t></a:r><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Own</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>"#.to_vec(),
        );
        let relationships = RelationshipMap::from([(
            "rId1".to_owned(),
            Relationship {
                id: "rId1".to_owned(),
                relationship_type: "diagramData".to_owned(),
                target: "diagrams/data1.xml".to_owned(),
                target_mode: None,
            },
        )]);
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let shape = parse_smart_art_from_drawing(
            &drawing(),
            Some(&relationships),
            Some(&MediaMap::new()),
            Some(&mut context),
            &mut budget,
        )
        .unwrap()
        .unwrap();
        let children = shape.children.unwrap();
        let text_body = children[0].text_body.as_ref().unwrap();
        let colors: Vec<_> = text_body.content[0]["content"]
            .as_array()
            .unwrap()
            .iter()
            .map(|run| run["formatting"]["color"].clone())
            .collect();
        assert_eq!(
            colors,
            [
                serde_json::json!({"themeColor": "lt1"}),
                serde_json::json!({"themeColor": "lt1"}),
                serde_json::json!({"rgb": "FF0000"}),
            ]
        );
    }

    #[test]
    fn cached_text_takes_its_frame_paragraphs_and_text_point_color() {
        let paragraphs = r#"<a:p><a:pPr algn="ctr"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="10000"/></a:spcBef><a:spcAft><a:spcPct val="35000"/></a:spcAft><a:buNone/></a:pPr><a:r><a:rPr sz="2800" u="sng"/><a:t>Alpha</a:t></a:r></a:p><a:p><a:pPr marL="228600" lvl="1" indent="-228600"><a:spcAft><a:spcPct val="15000"/></a:spcAft><a:buChar char="•"/><a:defRPr b="1" i="1"/></a:pPr><a:r><a:rPr sz="2200" i="0"/><a:t>one</a:t></a:r></a:p>"#;
        let drawing = format!(
            r#"<dsp:drawing xmlns:dsp="dsp" xmlns:a="a"><dsp:spTree><dsp:grpSpPr/><dsp:sp modelId="shape"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm rot="10800000"><a:off x="0" y="0"/><a:ext cx="4000" cy="2000"/></a:xfrm><a:prstGeom prst="trapezoid"/></dsp:spPr><dsp:style><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr anchor="ctr"/>{paragraphs}</dsp:txBody><dsp:txXfrm rot="10800000"><a:off x="1000" y="500"/><a:ext cx="2000" cy="1000"/></dsp:txXfrm></dsp:sp></dsp:spTree></dsp:drawing>"#
        );
        let mut context = SmartArtContext::default();
        context.parts.insert(
            "word/diagrams/drawing1.xml".to_owned(),
            drawing.into_bytes(),
        );
        context.parts.insert(
            "word/diagrams/data1.xml".to_owned(),
            br#"<dgm:dataModel xmlns:dgm="d" xmlns:a="a"><dgm:ptLst><dgm:pt modelId="node"><dgm:prSet/><dgm:t><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="shape" type="pres"><dgm:prSet presAssocID="node" presStyleLbl="node1" presStyleIdx="0" presStyleCnt="1"/></dgm:pt><dgm:pt modelId="text" type="pres"><dgm:prSet presAssocID="node" presStyleLbl="revTx" presStyleIdx="0" presStyleCnt="1"/></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="c1" type="presOf" srcId="node" destId="shape"/><dgm:cxn modelId="c2" type="presOf" srcId="node" destId="text"/></dgm:cxnLst></dgm:dataModel>"#.to_vec(),
        );
        context.parts.insert(
            "word/diagrams/colors1.xml".to_owned(),
            br#"<dgm:colorsDef xmlns:dgm="d" xmlns:a="a"><dgm:styleLbl name="node1"><dgm:txFillClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:txFillClrLst></dgm:styleLbl><dgm:styleLbl name="revTx"><dgm:txFillClrLst meth="repeat"><a:schemeClr val="dk1"/></dgm:txFillClrLst></dgm:styleLbl></dgm:colorsDef>"#.to_vec(),
        );
        let relationship = |id: &str, target: &str| {
            (
                id.to_owned(),
                Relationship {
                    id: id.to_owned(),
                    relationship_type: "diagram".to_owned(),
                    target: target.to_owned(),
                    target_mode: None,
                },
            )
        };
        let relationships = RelationshipMap::from([
            relationship("rId1", "diagrams/data1.xml"),
            relationship("rId4", "diagrams/colors1.xml"),
        ]);
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let drawing = parse_xml(
            br#"<w:drawing xmlns:w="w" xmlns:wp="wp" xmlns:dgm="dgm" xmlns:r="r"><wp:inline><wp:extent cx="4000" cy="2000"/><dgm:relIds r:dm="rId1" r:lo="rId2" r:cs="rId4"/></wp:inline></w:drawing>"#,
            "drawing",
            &mut budget,
        )
        .unwrap()
        .root()
        .unwrap()
        .clone();
        let shape = parse_smart_art_from_drawing(
            &drawing,
            Some(&relationships),
            Some(&MediaMap::new()),
            Some(&mut context),
            &mut budget,
        )
        .unwrap()
        .unwrap();
        let scene_nodes = shape.scene.as_ref().unwrap().root.as_ref().unwrap();
        assert_eq!(scene_nodes.children.as_ref().map(Vec::len), Some(2));
        let children = shape.children.unwrap();
        let [trapezoid, frame] = children.as_slice() else {
            panic!("expected the shape and its text frame");
        };
        assert!(trapezoid.text_body.is_none());
        assert_eq!(
            frame.transform.as_ref().and_then(|value| value.rotation),
            Some(0.0)
        );
        assert_eq!(
            (frame.offset.as_ref().unwrap().x, frame.size.width),
            (1000.0, 2000.0)
        );
        let content = &frame.text_body.as_ref().unwrap().content;
        let title = &content[0];
        assert_eq!(title["formatting"]["alignment"], "center");
        assert_eq!(title["formatting"]["spaceBefore"], 0.0);
        assert_eq!(title["formatting"]["spaceAfter"], 0.35 * 2800.0 * 1.2 / 5.0);
        assert_eq!(title["formatting"]["lineSpacing"], 216.0);
        assert_eq!(
            title["content"][0]["formatting"]["color"]["themeColor"],
            "dk1"
        );
        assert_eq!(
            title["content"][0]["formatting"]["underline"]["style"],
            "single"
        );
        let bullet = &content[1];
        assert_eq!(bullet["formatting"]["spaceAfter"], 0.0);
        let item = &bullet["content"][0]["formatting"];
        assert_eq!(
            (&item["bold"], &item["italic"]),
            (&serde_json::json!(true), &serde_json::json!(false))
        );
        assert_eq!(bullet["formatting"]["indentLeft"], 360.0);
        assert_eq!(bullet["formatting"]["hangingIndent"], true);
        assert_eq!(
            bullet["bullet"],
            serde_json::json!({"text": "•", "color": {"themeColor": "dk1"}, "size": 22.0})
        );
        assert_eq!(
            bullet["content"][0]["content"],
            serde_json::json!([{"type": "text", "text": "one"}])
        );
    }

    #[test]
    fn an_unreadable_drawing_rels_leaves_the_picture_fill_unresolved_and_warns() {
        let mut context = SmartArtContext::default();
        context.parts.insert(
            "word/diagrams/drawing1.xml".to_owned(),
            br#"<dsp:drawing xmlns:dsp="dsp" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm><a:prstGeom prst="rect"/><a:blipFill><a:blip r:embed="rId1"/></a:blipFill></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>"#.to_vec(),
        );
        context.parts.insert(
            "word/diagrams/_rels/drawing1.xml.rels".to_owned(),
            b"<Relationships><not closed".to_vec(),
        );
        let relationships = RelationshipMap::from([(
            "rId1".to_owned(),
            Relationship {
                id: "rId1".to_owned(),
                relationship_type: "diagramData".to_owned(),
                target: "diagrams/data1.xml".to_owned(),
                target_mode: None,
            },
        )]);
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let shape = parse_smart_art_from_drawing(
            &drawing(),
            Some(&relationships),
            Some(&MediaMap::new()),
            Some(&mut context),
            &mut budget,
        )
        .unwrap()
        .unwrap();
        assert!(shape.children.unwrap()[0].fill_paint.is_some());
        assert_eq!(context.warnings.len(), 1);
    }

    #[test]
    fn a_drawing_rels_past_the_budget_spends_none_of_it_and_is_read_once() {
        let mut context = SmartArtContext::default();
        let rels = "word/diagrams/_rels/drawing1.xml.rels";
        context.parts.insert(
            rels.to_owned(),
            br#"<Relationships><Relationship Id="rId1" Type="image" Target="../media/a.png"/><Relationship Id="rId2" Type="image" Target="../media/b.png"/></Relationships>"#.to_vec(),
        );
        let limits = ParseLimits {
            max_relationships: 1,
            ..ParseLimits::default()
        };
        let budget = ParseBudget::new(&limits);
        let part = "word/diagrams/drawing1.xml";
        assert!(drawing_relationships(part, &mut context, &budget).is_none());
        assert!(drawing_relationships(part, &mut context, &budget).is_none());
        assert_eq!(context.warnings.len(), 1);
        let mut budget = budget;
        let own = br#"<Relationships><Relationship Id="rId1" Type="image" Target="a.png"/></Relationships>"#;
        assert_eq!(
            parse_relationships(own, "own", &mut budget).unwrap().len(),
            1
        );
    }

    #[test]
    fn an_unsafe_target_loses_only_its_own_picture() {
        let mut context = SmartArtContext::default();
        context.parts.insert(
            "word/diagrams/_rels/drawing1.xml.rels".to_owned(),
            br#"<Relationships><Relationship Id="rId1" Type="image" Target="../../../../outside.png"/><Relationship Id="rId2" Type="image" Target="../media/b.png"/></Relationships>"#.to_vec(),
        );
        let limits = ParseLimits::default();
        let budget = ParseBudget::new(&limits);
        let own =
            drawing_relationships("word/diagrams/drawing1.xml", &mut context, &budget).unwrap();
        assert!(!own.contains_key("rId1"));
        assert_eq!(own["rId2"].target, "word/media/b.png");
    }

    #[test]
    fn a_pre_rendered_picture_fill_resolves_against_its_own_part() {
        let mut context = SmartArtContext::default();
        context.parts.insert(
            "word/diagrams/drawing1.xml".to_owned(),
            br#"<dsp:drawing xmlns:dsp="dsp" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm><a:prstGeom prst="rect"/><a:blipFill><a:blip r:embed="rId1"/></a:blipFill></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>"#.to_vec(),
        );
        context.parts.insert(
            "word/diagrams/_rels/drawing1.xml.rels".to_owned(),
            br#"<Relationships><Relationship Id="rId1" Type="image" Target="../media/image4.png"/></Relationships>"#.to_vec(),
        );
        let relationships = RelationshipMap::from([(
            "rId1".to_owned(),
            Relationship {
                id: "rId1".to_owned(),
                relationship_type: "diagramData".to_owned(),
                target: "diagrams/data1.xml".to_owned(),
                target_mode: None,
            },
        )]);
        let media = MediaMap::from([(
            "word/media/image4.png".to_owned(),
            std::sync::Arc::new(crate::media::MediaFile {
                path: "word/media/image4.png".to_owned(),
                filename: None,
                mime_type: "image/png".to_owned(),
                base64: String::new(),
                data_url: "media:4".to_owned(),
            }),
        )]);
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let shape = parse_smart_art_from_drawing(
            &drawing(),
            Some(&relationships),
            Some(&media),
            Some(&mut context),
            &mut budget,
        )
        .unwrap()
        .unwrap();
        let fill = shape.children.unwrap().remove(0).fill_paint.unwrap();
        assert_eq!(fill.picture.unwrap().src.as_deref(), Some("media:4"));
    }

    #[test]
    fn basic_process_fallback_builds_two_chevrons() {
        let mut context = SmartArtContext::default();
        context.parts.insert(
            "word/diagrams/data1.xml".to_owned(),
            br#"<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a"><dgm:ptLst><dgm:pt modelId="0" type="doc"><a:t>skip</a:t></dgm:pt><dgm:pt modelId="1"><a:t>One</a:t></dgm:pt><dgm:pt modelId="2"><a:t>Two</a:t></dgm:pt></dgm:ptLst></dgm:dataModel>"#.to_vec(),
        );
        context.parts.insert(
            "word/diagrams/layout1.xml".to_owned(),
            b"<layout name=\"Basic Process\"/>".to_vec(),
        );
        let relationships = RelationshipMap::from([
            (
                "rId1".to_owned(),
                Relationship {
                    id: "rId1".to_owned(),
                    relationship_type: "diagramData".to_owned(),
                    target: "diagrams/data1.xml".to_owned(),
                    target_mode: None,
                },
            ),
            (
                "rId2".to_owned(),
                Relationship {
                    id: "rId2".to_owned(),
                    relationship_type: "diagramLayout".to_owned(),
                    target: "diagrams/layout1.xml".to_owned(),
                    target_mode: None,
                },
            ),
        ]);
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let shape = parse_smart_art_from_drawing(
            &drawing(),
            Some(&relationships),
            None,
            Some(&mut context),
            &mut budget,
        )
        .unwrap()
        .unwrap();
        let children = shape.children.unwrap();
        assert_eq!(children.len(), 2);
        assert_eq!(children[0].shape_type, "chevron");
        assert_eq!(children[0].size.width, 440.0);
        assert_eq!(children[0].offset.as_ref().unwrap().x, 40.0);
    }

    #[test]
    fn cached_connector_and_text_are_enriched_from_source() {
        let xml = br#"<dsp:spTree xmlns:dsp="dsp" xmlns:a="a"><dsp:cxnSp><dsp:nvCxnSpPr><dsp:cNvPr id="7" name="line"/></dsp:nvCxnSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></dsp:spPr><dsp:txBody><a:p><a:r><a:rPr b="1" sz="600"><a:latin typeface="Arial"/></a:rPr><a:t>Hi</a:t></a:r></a:p></dsp:txBody></dsp:cxnSp></dsp:spTree>"#;
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let parsed = parse_diagram_drawing_scene(xml, "drawing", &mut budget).unwrap();
        assert_eq!(parsed.shapes[0].shape_type, "straightConnector1");
        assert!(parsed.shapes[0].geometry_path.is_some());
        assert_eq!(
            parsed.shapes[0].text_body.as_ref().unwrap().content.len(),
            1
        );
    }

    #[test]
    fn cached_shapes_sharing_one_id_keep_their_own_text() {
        let shape = |geometry: &str, text: &str| {
            let body = if text.is_empty() {
                String::new()
            } else {
                format!("<dsp:txBody><a:p><a:r><a:t>{text}</a:t></a:r></a:p></dsp:txBody>")
            };
            format!(
                r#"<dsp:sp><dsp:nvSpPr><dsp:cNvPr id="0" name=""/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm><a:prstGeom prst="{geometry}"/></dsp:spPr>{body}</dsp:sp>"#
            )
        };
        let xml = format!(
            r#"<dsp:drawing xmlns:dsp="dsp" xmlns:a="a"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/></dsp:nvGrpSpPr>{}{}{}{}</dsp:spTree></dsp:drawing>"#,
            shape("rightArrow", ""),
            shape("roundRect", "One"),
            shape("roundRect", "Two"),
            shape("roundRect", "Three"),
        );
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let parsed = parse_diagram_drawing_scene(xml.as_bytes(), "drawing", &mut budget).unwrap();
        let text = |shape: &Shape| {
            shape
                .text_body
                .as_ref()
                .map(|body| body.content[0]["content"][0]["content"][0]["text"].clone())
        };
        let texts: Vec<_> = parsed.shapes.iter().map(text).collect();
        assert_eq!(
            texts,
            [
                None,
                Some("One".into()),
                Some("Two".into()),
                Some("Three".into())
            ]
        );
        let scene = parsed.scene.unwrap().root.unwrap().children.unwrap();
        let scene_texts: Vec<_> = scene
            .iter()
            .map(|node| text(node.shape.as_ref().unwrap()))
            .collect();
        assert_eq!(scene_texts, texts);
    }

    #[test]
    fn unsupported_and_recursive_fallbacks_are_bounded() {
        let mut context = SmartArtContext::default();
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let result = layout_basic_diagram(
            b"<data/>",
            Some(b"<layout name=\"Radial Cycle\"/>"),
            &ImageSize {
                width: 1.0,
                height: 1.0,
            },
            &mut context,
            &mut budget,
        )
        .unwrap();
        assert!(result.is_empty());
        assert_eq!(context.warnings.len(), 1);
    }
}
