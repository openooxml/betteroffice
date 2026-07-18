//! SmartArt cached-scene parsing and the incumbent basic-layout fallback.

use std::collections::HashMap;

use indexmap::IndexMap;

use crate::drawingml::{ShapeFill, ShapeOutline, preset_geometry_to_path};
use crate::image::ImageSize;
use crate::relationships::RelationshipMap;
use crate::scalars::ColorValue;
use crate::shape::{
    DrawingScene, DrawingSceneNode, Shape, ShapeOffset, ShapeTextBody, parse_drawing_scene_tree,
};
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

const MAX_DIAGRAM_DEPTH: usize = 32;
const MAX_DIAGRAM_SHAPES: usize = 4_096;
const MAX_SAFE_DRAWING_NUMBER: f64 = 1_000_000_000_000.0;

#[derive(Clone, Debug, Default)]
pub struct SmartArtContext {
    pub parts: IndexMap<String, Vec<u8>>,
    pub warnings: Vec<String>,
}

pub fn create_smart_art_context(all_xml: &IndexMap<String, Vec<u8>>) -> SmartArtContext {
    let mut parts = IndexMap::new();
    for (path, xml) in all_xml {
        if path.to_ascii_lowercase().starts_with("word/diagrams/") {
            parts.insert(normalize_package_path(path), xml.clone());
        }
    }
    SmartArtContext {
        parts,
        warnings: Vec::new(),
    }
}

pub fn is_smart_art_drawing(drawing: &XmlElement) -> bool {
    find_descendant(drawing, "relIds", 0).is_some()
}

pub fn parse_smart_art_from_drawing(
    drawing: &XmlElement,
    relationships: Option<&RelationshipMap>,
    context: Option<&mut SmartArtContext>,
    budget: &mut ParseBudget<'_>,
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

    if let Some(drawing_xml) = data_path
        .as_deref()
        .and_then(|path| find_pre_rendered_drawing(path, context))
    {
        let parsed = parse_diagram_drawing_scene(drawing_xml, "SmartArt drawing", budget)?;
        if !parsed.shapes.is_empty() {
            return Ok(Some(diagram_group(size, parsed.shapes, parsed.scene)));
        }
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
    Ok((!fallback.is_empty()).then(|| diagram_group(size, fallback, None)))
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
    let mut scene = parse_drawing_scene_tree(source);
    let mut sources = IndexMap::new();
    collect_shape_elements(source, &mut sources, 0);
    let mut shapes = Vec::new();
    collect_scene_shapes(scene.root.as_mut(), &mut shapes, &sources);
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

fn collect_shape_elements<'a>(
    parent: &'a XmlElement,
    shapes: &mut IndexMap<String, &'a XmlElement>,
    depth: usize,
) {
    if depth > MAX_DIAGRAM_DEPTH || shapes.len() >= MAX_DIAGRAM_SHAPES {
        return;
    }
    if matches!(parent.local_name(), "sp" | "cxnSp")
        && let Some(id) =
            find_descendant(parent, "cNvPr", 0).and_then(|element| element.attribute(None, "id"))
    {
        shapes.insert(id.to_owned(), parent);
    }
    for child in parent.child_elements() {
        collect_shape_elements(child, shapes, depth + 1);
        if shapes.len() >= MAX_DIAGRAM_SHAPES {
            break;
        }
    }
}

fn collect_scene_shapes(
    node: Option<&mut DrawingSceneNode>,
    shapes: &mut Vec<Shape>,
    sources: &IndexMap<String, &XmlElement>,
) {
    let Some(node) = node else { return };
    if node.kind.as_deref() == Some("shape")
        && let Some(shape) = node.shape.as_deref_mut()
    {
        if let Some(source) = shape.id.as_deref().and_then(|id| sources.get(id).copied()) {
            if source.local_name() == "cxnSp" && find_descendant(source, "prstGeom", 0).is_none() {
                shape.shape_type = "straightConnector1".to_owned();
                shape.geometry_path =
                    preset_geometry_to_path("straightConnector1", &HashMap::new());
            }
            let paragraphs = parse_drawing_text_body(source.child_by_local_name("txBody"));
            if !paragraphs.is_empty() {
                shape.text_body = Some(ShapeTextBody {
                    vertical: None,
                    rotation: None,
                    anchor: None,
                    anchor_center: None,
                    auto_fit: None,
                    margins: None,
                    content: paragraphs,
                });
            }
        }
        shapes.push(shape.clone());
    }
    if let Some(children) = &mut node.children {
        for child in children {
            collect_scene_shapes(Some(child), shapes, sources);
        }
    }
}

fn parse_drawing_text_body(text_body: Option<&XmlElement>) -> Vec<serde_json::Value> {
    let Some(text_body) = text_body else {
        return Vec::new();
    };
    text_body
        .child_elements()
        .filter(|child| child.local_name() == "p")
        .filter_map(parse_drawing_text_paragraph)
        .collect()
}

fn parse_drawing_text_paragraph(paragraph: &XmlElement) -> Option<serde_json::Value> {
    let mut runs = Vec::new();
    for child in paragraph.child_elements() {
        match child.local_name() {
            "r" | "fld" => {
                let text = collect_drawing_text(child, 0);
                if !text.is_empty() {
                    runs.push(drawing_run(&text, parse_drawing_run_formatting(child)));
                }
            }
            "br" => runs.push(drawing_run("\n", None)),
            _ => {}
        }
    }
    if runs.is_empty() {
        return None;
    }
    let alignment = match paragraph
        .child_by_local_name("pPr")
        .and_then(|properties| properties.attribute(None, "algn"))
    {
        Some("ctr") => "center",
        Some("r") => "right",
        Some("just") => "both",
        _ => "left",
    };
    Some(serde_json::json!({
        "type": "paragraph",
        "formatting": { "alignment": alignment },
        "content": runs,
    }))
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

fn parse_drawing_run_formatting(run: &XmlElement) -> Option<serde_json::Value> {
    let properties = run.child_by_local_name("rPr")?;
    let mut formatting = serde_json::Map::new();
    let mut property_count = 0;
    if properties.attribute(None, "b") == Some("1") {
        formatting.insert("bold".to_owned(), serde_json::json!(true));
        property_count += 1;
    }
    if properties.attribute(None, "i") == Some("1") {
        formatting.insert("italic".to_owned(), serde_json::json!(true));
        property_count += 1;
    }
    if let Some(size) = properties
        .parse_numeric_attribute(None, "sz", 1.0)
        .filter(safe_number)
    {
        formatting.insert("fontSize".to_owned(), serde_json::json!(size / 50.0));
        property_count += 1;
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
        property_count += 1;
    }
    if let Some(language) = properties.attribute(None, "lang") {
        formatting.insert(
            "language".to_owned(),
            serde_json::json!({"latin": language}),
        );
        property_count += 1;
    }
    if let Some(solid) = properties.child_by_full_name("a:solidFill") {
        property_count += 1;
        if let Some(color) = crate::drawingml::parse_color_element(Some(solid)) {
            formatting.insert("color".to_owned(), serde_json::to_value(color).unwrap());
        }
    }
    (property_count > 0).then_some(serde_json::Value::Object(formatting))
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
) -> Option<&'a [u8]> {
    let (directory, filename) = data_path.rsplit_once('/')?;
    let number = data_file_number(filename)?;
    let direct = format!("{directory}/drawing{number}.xml");
    if let Some(xml) = context.parts.get(&direct) {
        return Some(xml);
    }
    let suffix = format!("/drawing{number}.xml").to_ascii_lowercase();
    context
        .parts
        .iter()
        .find(|(path, _)| path.to_ascii_lowercase().ends_with(&suffix))
        .map(|(_, xml)| xml.as_slice())
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

    #[test]
    fn basic_process_fallback_matches_incumbent_layout() {
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
