use serde_json::Value;

use crate::image::{ImagePosition, ImageWrap, PositionAxis};
use crate::shape::Shape;
use crate::smart_art::{apply_frame_metadata, frame_model, other_model_sha256};
use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, XmlNode, parse_xml};

use super::context::SerializerContext;
use super::raw::validate_replayed_fragment;
use super::xml_writer::int_attr;

fn child_mut<'a>(element: &'a mut XmlElement, local: &str) -> Option<&'a mut XmlElement> {
    element.children.iter_mut().find_map(|child| match child {
        XmlNode::Element(child) if child.local_name() == local => Some(child),
        _ => None,
    })
}

fn drawing_mut(element: &mut XmlElement) -> Option<&mut XmlElement> {
    if element.local_name() == "drawing" {
        Some(element)
    } else {
        child_mut(element, "drawing")
    }
}

fn empty(name: String) -> XmlElement {
    XmlElement {
        name,
        attributes: Default::default(),
        children: Vec::new(),
        paragraph_ordinal: None,
        para_id_attribute: None,
        bookmark_xml: None,
    }
}

fn order(local: &str) -> usize {
    match local {
        "simplePos" => 0,
        "positionH" => 1,
        "positionV" => 2,
        "extent" => 3,
        "effectExtent" => 4,
        "wrapNone" | "wrapSquare" | "wrapTight" | "wrapThrough" | "wrapTopAndBottom" => 5,
        "docPr" => 6,
        "cNvGraphicFramePr" => 7,
        "graphic" => 8,
        _ => 9,
    }
}

fn ensure<'a>(parent: &'a mut XmlElement, local: &str) -> &'a mut XmlElement {
    let index = parent
        .children
        .iter()
        .position(|child| matches!(child, XmlNode::Element(child) if child.local_name() == local));
    let index = index.unwrap_or_else(|| {
        let at = parent.children.iter().position(|child| matches!(child, XmlNode::Element(child) if order(child.local_name()) > order(local)))
            .unwrap_or(parent.children.len());
        let prefix = parent.namespace_prefix().unwrap_or("wp");
        parent.children.insert(at, XmlNode::Element(empty(format!("{prefix}:{local}"))));
        at
    });
    let XmlNode::Element(element) = &mut parent.children[index] else {
        unreachable!()
    };
    element
}

fn attribute(element: &mut XmlElement, key: &str, value: Option<String>) {
    if let Some(value) = value {
        element.attributes.insert(key.to_owned(), value);
    } else {
        element.attributes.shift_remove(key);
    }
}

fn number(element: &mut XmlElement, key: &str, value: Option<f64>) {
    attribute(element, key, value.map(|value| int_attr(Some(value))));
}

fn boolean(element: &mut XmlElement, key: &str, value: Option<bool>) {
    attribute(
        element,
        key,
        value.map(|value| if value { "1" } else { "0" }.to_owned()),
    );
}

fn axis(container: &mut XmlElement, local: &str, position: &PositionAxis) {
    let element = ensure(container, local);
    attribute(element, "relativeFrom", Some(position.relative_to.clone()));
    element.children.retain(|child| !matches!(child, XmlNode::Element(child) if matches!(child.local_name(), "align" | "posOffset")));
    let prefix = element.namespace_prefix().unwrap_or("wp");
    let mut value = empty(format!(
        "{prefix}:{}",
        if position.alignment.is_some() {
            "align"
        } else {
            "posOffset"
        }
    ));
    value.children.push(XmlNode::Text(
        position
            .alignment
            .clone()
            .unwrap_or_else(|| int_attr(position.pos_offset.or(position.offset))),
    ));
    element.children.insert(0, XmlNode::Element(value));
}

fn position(container: &mut XmlElement, value: &ImagePosition) {
    boolean(
        container,
        "simplePos",
        Some(value.use_simple_pos.unwrap_or(false)),
    );
    boolean(
        container,
        "behindDoc",
        Some(value.behind_doc.unwrap_or(false)),
    );
    boolean(container, "locked", Some(value.locked.unwrap_or(false)));
    boolean(container, "hidden", value.hidden);
    number(
        container,
        "relativeHeight",
        Some(value.relative_height.unwrap_or(0.0)),
    );
    let simple = ensure(container, "simplePos");
    number(
        simple,
        "x",
        Some(
            value
                .simple_pos
                .as_ref()
                .and_then(|point| point.x)
                .unwrap_or(0.0),
        ),
    );
    number(
        simple,
        "y",
        Some(
            value
                .simple_pos
                .as_ref()
                .and_then(|point| point.y)
                .unwrap_or(0.0),
        ),
    );
    axis(container, "positionH", &value.horizontal);
    axis(container, "positionV", &value.vertical);
}

fn wrap(container: &mut XmlElement, value: &ImageWrap, context: &mut SerializerContext) {
    for (key, value) in [
        ("distT", value.dist_t),
        ("distB", value.dist_b),
        ("distL", value.dist_l),
        ("distR", value.dist_r),
    ] {
        number(container, key, value);
    }
    if container.local_name() != "anchor" {
        return;
    }
    let local = match value.wrap_type.as_str() {
        "inline" | "none" | "behind" | "inFront" => "wrapNone",
        "square" => "wrapSquare",
        "tight" => "wrapTight",
        "through" => "wrapThrough",
        "topAndBottom" => "wrapTopAndBottom",
        _ => {
            context.warn("SmartArt wrap type is unsupported; preserving the source wrap.");
            return;
        }
    };
    container.children.retain(|child| !matches!(child, XmlNode::Element(child) if child.local_name().starts_with("wrap") && child.local_name() != local));
    let element = ensure(container, local);
    for (key, value) in [
        ("distT", value.dist_t),
        ("distB", value.dist_b),
        ("distL", value.dist_l),
        ("distR", value.dist_r),
    ] {
        if element.attributes.contains_key(key) {
            number(element, key, value);
        }
    }
    if matches!(local, "wrapSquare" | "wrapTight" | "wrapThrough") {
        attribute(element, "wrapText", value.wrap_text.clone());
    }
    if let Some(points) = value
        .polygon
        .as_ref()
        .filter(|_| matches!(local, "wrapTight" | "wrapThrough"))
    {
        let polygon = ensure(element, "wrapPolygon");
        polygon.children.retain(|child| !matches!(child, XmlNode::Element(child) if matches!(child.local_name(), "start" | "lineTo")));
        for (at, point) in points.iter().enumerate() {
            let prefix = polygon.namespace_prefix().unwrap_or("wp");
            let mut element = empty(format!(
                "{prefix}:{}",
                if at == 0 { "start" } else { "lineTo" }
            ));
            number(&mut element, "x", point.x);
            number(&mut element, "y", point.y);
            polygon.children.push(XmlNode::Element(element));
        }
    } else {
        element.children.retain(|child| !matches!(child, XmlNode::Element(child) if child.local_name() == "wrapPolygon"));
    }
    if matches!(value.wrap_type.as_str(), "behind" | "inFront") {
        boolean(container, "behindDoc", Some(value.wrap_type == "behind"));
    }
}

fn clear_decorative_extensions(element: &mut XmlElement) {
    element.children.retain(|child| !matches!(child, XmlNode::Element(child) if matches!(child.name.as_str(), "adec:decorative" | "a16:decorative")));
    for child in &mut element.children {
        if let XmlNode::Element(child) = child {
            clear_decorative_extensions(child);
        }
    }
}

fn update_frame(
    drawing: &mut XmlElement,
    shape: &Shape,
    baseline: &Value,
    context: &mut SerializerContext,
) {
    let current = frame_model(shape);
    let changed = |key: &str| current.get(key) != baseline.get(key);
    let Some(container) = drawing.children.iter_mut().find_map(|node| match node {
        XmlNode::Element(node) if matches!(node.local_name(), "anchor" | "inline") => Some(node),
        _ => None,
    }) else {
        context.warn("SmartArt placement is unavailable; preserving the source frame.");
        return;
    };
    if changed("size") {
        if [shape.size.width, shape.size.height]
            .iter()
            .all(|value| value.is_finite() && *value >= 0.0)
        {
            let extent = ensure(container, "extent");
            number(extent, "cx", Some(shape.size.width));
            number(extent, "cy", Some(shape.size.height));
        } else {
            context.warn("SmartArt size is invalid; preserving the source extent.");
        }
    }
    if changed("effectExtent") {
        if let Some(value) = &shape.effect_extent {
            let extent = ensure(container, "effectExtent");
            for (key, value) in [
                ("l", value.left),
                ("t", value.top),
                ("r", value.right),
                ("b", value.bottom),
            ] {
                number(extent, key, value);
            }
        } else {
            container.children.retain(|child| !matches!(child, XmlNode::Element(child) if child.local_name() == "effectExtent"));
        }
    }
    if changed("position") || changed("wrap") {
        let floating = shape
            .wrap
            .as_ref()
            .is_some_and(|wrap| wrap.wrap_type != "inline")
            || shape.position.is_some();
        let converted = floating && container.local_name() == "inline";
        if converted {
            let prefix = container.namespace_prefix().unwrap_or("wp").to_owned();
            container.name = format!("{prefix}:anchor");
            for (key, value) in [
                ("simplePos", "0"),
                ("relativeHeight", "0"),
                ("behindDoc", "0"),
                ("locked", "0"),
                ("layoutInCell", "1"),
                ("allowOverlap", "1"),
            ] {
                container
                    .attributes
                    .insert(key.to_owned(), value.to_owned());
            }
        } else if !floating && container.local_name() == "anchor" {
            let prefix = container.namespace_prefix().unwrap_or("wp").to_owned();
            container.name = format!("{prefix}:inline");
            for key in [
                "simplePos",
                "relativeHeight",
                "behindDoc",
                "locked",
                "layoutInCell",
                "allowOverlap",
                "hidden",
            ] {
                container.attributes.shift_remove(key);
            }
            container.children.retain(|child| !matches!(child, XmlNode::Element(child) if matches!(child.local_name(), "simplePos" | "positionH" | "positionV") || child.local_name().starts_with("wrap")));
        }
        if container.local_name() == "anchor" && (changed("position") || converted) {
            if let Some(value) = &shape.position {
                position(container, value);
            } else {
                let value = ImagePosition {
                    use_simple_pos: None,
                    simple_pos: None,
                    relative_height: shape.relative_height,
                    behind_doc: None,
                    hidden: None,
                    locked: None,
                    horizontal: PositionAxis {
                        relative_to: "column".to_owned(),
                        alignment: None,
                        pos_offset: Some(0.0),
                        offset: None,
                    },
                    vertical: PositionAxis {
                        relative_to: "paragraph".to_owned(),
                        alignment: None,
                        pos_offset: Some(0.0),
                        offset: None,
                    },
                };
                position(container, &value);
            }
        }
        if let Some(value) = &shape.wrap {
            wrap(container, value, context);
        } else if container.local_name() == "anchor" {
            ensure(container, "wrapNone");
        }
    }
    if changed("relativeHeight") && container.local_name() == "anchor" {
        number(container, "relativeHeight", shape.relative_height);
    }
    for (field, key, value) in [
        ("id", "id", &shape.id),
        ("name", "name", &shape.name),
        ("title", "title", &shape.title),
        ("description", "descr", &shape.description),
    ] {
        if changed(field) {
            if field == "id" && value.as_deref().is_none_or(|id| id.parse::<u32>().is_err()) {
                context.warn("SmartArt docPr id is invalid; preserving the source id.");
            } else {
                let value = if field == "name" {
                    Some(value.clone().unwrap_or_default())
                } else {
                    value.clone()
                };
                attribute(ensure(container, "docPr"), key, value);
            }
        }
    }
    for (field, value) in [("hidden", shape.hidden), ("decorative", shape.decorative)] {
        if changed(field) {
            let properties = ensure(container, "docPr");
            boolean(properties, field, value);
            if field == "decorative" && value != Some(true) {
                clear_decorative_extensions(properties);
            }
        }
    }
}

pub(super) fn unapplied_change(baseline: &Value, current: &Value, saved: &Value) -> bool {
    if baseline == current {
        return false;
    }
    if let (Some(before), Some(after)) = (baseline.as_object(), current.as_object()) {
        return before
            .keys()
            .chain(after.keys())
            .any(|key| unapplied_change(&baseline[key], &current[key], &saved[key]));
    }
    current != saved
}

pub(super) fn replay(
    shape: &Shape,
    context: &mut SerializerContext,
    whole_run: bool,
) -> Result<Option<String>, ParseError> {
    let Some(source) = shape
        .drawing_run_xml
        .as_deref()
        .or(shape.drawing_xml.as_deref())
    else {
        return Ok(None);
    };
    validate_replayed_fragment(source)?;
    let limits = ParseLimits::default();
    let mut document = parse_xml(
        source.as_bytes(),
        "SmartArt source",
        &mut ParseBudget::new(&limits),
    )?;
    let root = document
        .roots
        .first_mut()
        .ok_or_else(|| ParseError::Canonical("SmartArt source has no root".to_owned()))?;
    if whole_run && root.local_name() != "r" {
        return Ok(None);
    }
    let Some(drawing) = drawing_mut(root) else {
        return Err(ParseError::Canonical(
            "SmartArt source has no drawing".to_owned(),
        ));
    };
    let fallback;
    let baseline = if let Some(value) = shape
        .drawing_model
        .as_ref()
        .and_then(|value| value.get("frame"))
    {
        value
    } else {
        let mut source_shape = Shape::empty("rect".to_owned(), shape.size.clone());
        apply_frame_metadata(&mut source_shape, drawing);
        fallback = frame_model(&source_shape);
        &fallback
    };
    if shape
        .drawing_model
        .as_ref()
        .and_then(|value| value.get("otherSha256"))
        .and_then(Value::as_str)
        .is_some_and(|digest| digest != other_model_sha256(shape))
    {
        context.warn("SmartArt diagram content or geometry changes cannot be represented; preserving source diagram content and saving supported frame changes.");
    }
    let changed = frame_model(shape) != *baseline;
    if changed {
        update_frame(drawing, shape, baseline, context);
        let mut saved = Shape::empty("rect".to_owned(), shape.size.clone());
        apply_frame_metadata(&mut saved, drawing);
        let saved = frame_model(&saved);
        let current = frame_model(shape);
        for &key in crate::smart_art::FRAME_FIELDS {
            if unapplied_change(&baseline[key], &current[key], &saved[key]) {
                context.warn(format!(
                    "SmartArt {key} changes cannot be represented exactly; saving supported frame changes and preserving remaining source properties."
                ));
            }
        }
    }
    if !whole_run {
        return Ok(Some(drawing.to_raw_inline_xml()));
    }
    let mut output = source.to_owned();
    if changed {
        let at = root
            .child_elements()
            .position(|child| child.local_name() == "drawing")
            .expect("source drawing");
        let span = super::s13::element_span(source, &[at as u32]).ok_or_else(|| {
            ParseError::Canonical("SmartArt drawing has no source span".to_owned())
        })?;
        let drawing = root.child_by_local_name("drawing").expect("source drawing");
        output.replace_range(span, &drawing.to_raw_inline_xml());
    }
    Ok(Some(output))
}

#[cfg(test)]
mod tests {
    use crate::image::ImagePadding;
    use crate::serializer::{CanonicalXmlEvent, SerializerDeterminism, canonical_xml_events};
    use crate::smart_art::{frame_model, other_model_sha256};
    use serde_json::json;

    use super::*;

    fn context() -> SerializerContext {
        SerializerContext::new(&SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2000-01-01T00:00:00.000Z".to_owned(),
        })
        .unwrap()
    }

    fn source(anchored: bool) -> (Shape, String) {
        let container = if anchored { "anchor" } else { "inline" };
        let placement = if anchored {
            r#"<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>100</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>200</wp:posOffset></wp:positionV>"#
        } else {
            ""
        };
        let xml = format!(
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:rPr><w:noProof/></w:rPr><w:drawing xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><wp:{container} distL="0">{placement}<wp:extent cx="5486400" cy="3200400"/><wp:effectExtent l="1" t="2" r="3" b="4"/><wp:docPr id="71" name="Keep"><a:extLst><a:ext uri="urn:keep"/></a:extLst></wp:docPr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="dm" r:lo="lo" r:qs="qs" r:cs="cs"/><a:extLst><a:ext uri="urn:payload"/></a:extLst></a:graphicData></a:graphic></wp:{container}></w:drawing></w:r>"#
        );
        capture(xml)
    }

    fn capture(xml: String) -> (Shape, String) {
        let limits = ParseLimits::default();
        let parsed = parse_xml(xml.as_bytes(), "source", &mut ParseBudget::new(&limits)).unwrap();
        let drawing = parsed
            .root()
            .unwrap()
            .child_by_local_name("drawing")
            .unwrap();
        let mut shape = Shape::empty(
            "rect".to_owned(),
            crate::image::ImageSize {
                width: 0.0,
                height: 0.0,
            },
        );
        apply_frame_metadata(&mut shape, drawing);
        shape.drawing_run_xml = Some(xml.clone());
        shape.drawing_model =
            Some(json!({"frame": frame_model(&shape), "otherSha256": other_model_sha256(&shape)}));
        (shape, xml)
    }

    #[test]
    fn smart_art_frame_edits_keep_source_run_property_bytes() {
        let (_, xml) = source(false);
        let properties = "<w:rPr keep:attribute = 'source' xmlns:keep='urn:keep'>\n<w:noProof /> <w:position w:val = '20'></w:position><!--keep--><keep:metadata />\n</w:rPr>";
        let xml = xml.replace("<w:rPr><w:noProof/></w:rPr>", properties);
        let (mut shape, _) = capture(xml.clone());
        assert_eq!(replay(&shape, &mut context(), true).unwrap().unwrap(), xml);
        shape.size.width /= 2.0;
        let saved = replay(&shape, &mut context(), true).unwrap().unwrap();
        assert!(saved.contains(properties));
        assert!(saved.contains("cx=\"2743200\""));
    }

    #[test]
    fn smart_art_wrap_distances_update_anchor_and_wrap_element() {
        let (_, xml) = source(true);
        let xml = xml.replace(
            "<wp:docPr",
            r#"<wp:wrapSquare wrapText="bothSides" distL="114300" distR="57150"/><wp:docPr"#,
        );
        for distance in [Some(228_600.0), None] {
            let (mut shape, _) = capture(xml.clone());
            let wrap = shape.wrap.as_mut().unwrap();
            wrap.dist_l = distance;
            wrap.dist_r = distance;
            let mut context = context();
            let saved = replay(&shape, &mut context, true).unwrap().unwrap();
            let (reopened, _) = capture(saved.clone());
            assert_eq!(reopened.wrap, shape.wrap);
            assert_eq!(
                saved.matches("distL=\"228600\"").count(),
                if distance.is_some() { 2 } else { 0 }
            );
            assert_eq!(
                saved.matches("distR=\"228600\"").count(),
                if distance.is_some() { 2 } else { 0 }
            );
            assert!(!saved.contains("114300"));
            assert!(!saved.contains("57150"));
            assert!(saved.contains("<dgm:relIds"));
            assert!(context.take_warnings().is_empty());
        }
    }

    #[test]
    fn smart_art_clearing_wrap_polygon_removes_source_element() {
        let (_, xml) = source(true);
        let xml = xml.replace("<wp:docPr", r#"<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="21600" y="21600"/></wp:wrapPolygon></wp:wrapTight><wp:docPr"#);
        let (mut shape, _) = capture(xml);
        assert!(shape.wrap.as_ref().unwrap().polygon.is_some());
        shape.wrap.as_mut().unwrap().polygon = None;
        shape.wrap.as_mut().unwrap().wrap_text = None;
        let mut context = context();
        let saved = replay(&shape, &mut context, true).unwrap().unwrap();
        assert!(!saved.contains("wrapPolygon"));
        assert!(!saved.contains("wrapText"));
        assert!(saved.contains("wrapTight"));
        let (reopened, _) = capture(saved);
        assert_eq!(reopened.wrap, shape.wrap);
        assert!(context.take_warnings().is_empty());
    }

    #[test]
    fn smart_art_inline_moved_to_anchor_writes_a_wrap_element() {
        let (mut shape, _) = source(false);
        assert_eq!(shape.wrap.as_ref().map(|wrap| wrap.wrap_type.as_str()), Some("inline"));
        let (anchored, _) = source(true);
        shape.position = anchored.position.clone();
        let mut context = context();
        let saved = replay(&shape, &mut context, true).unwrap().unwrap();
        assert!(saved.contains("<wp:anchor"));
        assert_eq!(saved.matches("<wp:wrapNone/>").count(), 1);
        let (reopened, _) = capture(saved);
        assert_eq!(reopened.wrap.as_ref().map(|wrap| wrap.wrap_type.as_str()), Some("inFront"));
        assert!(context.take_warnings().is_empty());
    }

    #[test]
    fn smart_art_polygon_stays_out_of_wraps_without_one() {
        let (_, xml) = source(true);
        for from in ["wrapTight", "wrapThrough"] {
            let xml = xml.replace("<wp:docPr", &format!(r#"<wp:{from} wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="21600" y="21600"/></wp:wrapPolygon></wp:{from}><wp:docPr"#));
            for (to, element, reopens) in [
                ("square", "wrapSquare", "square"),
                ("topAndBottom", "wrapTopAndBottom", "topAndBottom"),
                ("none", "wrapNone", "inFront"),
            ] {
                let (mut shape, _) = capture(xml.clone());
                assert!(shape.wrap.as_ref().unwrap().polygon.is_some());
                shape.wrap.as_mut().unwrap().wrap_type = to.to_owned();
                let mut context = context();
                let saved = replay(&shape, &mut context, true).unwrap().unwrap();
                assert!(saved.contains(&format!("<wp:{element}")), "{from} -> {to}");
                assert!(!saved.contains(from), "{from} -> {to}");
                assert!(!saved.contains("wrapPolygon"), "{from} -> {to}");
                let (reopened, _) = capture(saved);
                assert_eq!(reopened.wrap.as_ref().unwrap().wrap_type, reopens);
                assert!(reopened.wrap.as_ref().unwrap().polygon.is_none());
            }
        }
    }

    #[test]
    fn smart_art_unrepresentable_frame_edits_warn() {
        for field in ["wrap", "position", "relativeHeight"] {
            let (mut shape, _) = source(field == "position");
            match field {
                "wrap" => {
                    shape.wrap.as_mut().unwrap().polygon =
                        Some(vec![crate::image::ImageWrapPoint {
                            x: Some(0.0),
                            y: Some(0.0),
                        }])
                }
                "position" => shape.position.as_mut().unwrap().horizontal.offset = Some(123.0),
                _ => shape.relative_height = Some(42.0),
            }
            shape.size.width = 2_743_200.0;
            let mut context = context();
            let saved = replay(&shape, &mut context, true).unwrap().unwrap();
            let (reopened, _) = capture(saved);
            assert_eq!(reopened.size, shape.size);
            assert!(
                context
                    .take_warnings()
                    .iter()
                    .any(|warning| warning.contains(&format!(
                        "SmartArt {field} changes cannot be represented exactly"
                    )))
            );
        }
    }

    #[test]
    fn smart_art_frame_edits_patch_source_without_changing_diagram_content() {
        let graphic_events = |xml: &str| {
            let events = canonical_xml_events(xml, false).unwrap();
            let start = events
                .iter()
                .position(|event| {
                    matches!(
                        event,
                        CanonicalXmlEvent::Start { namespace_uri, local_name, .. }
                            if namespace_uri == crate::xml::namespaces::A && local_name == "graphic"
                    )
                })
                .unwrap();
            let end = events
                .iter()
                .rposition(|event| {
                    matches!(
                        event,
                        CanonicalXmlEvent::End { namespace_uri, local_name }
                            if namespace_uri == crate::xml::namespaces::A && local_name == "graphic"
                    )
                })
                .unwrap();
            events[start..=end].to_vec()
        };
        for anchored in [false, true] {
            let (mut shape, original) = source(anchored);
            shape.size.width = 2_743_200.0;
            shape.effect_extent = Some(ImagePadding {
                left: Some(100.0),
                top: Some(200.0),
                right: Some(300.0),
                bottom: Some(400.0),
            });
            shape.id = Some("72".to_owned());
            shape.name = Some("Resized & retained".to_owned());
            shape.title = Some("Updated title".to_owned());
            shape.description = Some("Updated description".to_owned());
            if let Some(value) = &mut shape.position {
                value.horizontal.pos_offset = Some(500.0);
                value.vertical.pos_offset = Some(600.0);
                shape.relative_height = Some(29.0);
            }
            let mut context = context();
            let saved = replay(&shape, &mut context, true).unwrap().unwrap();
            let limits = ParseLimits::default();
            let mut parsed =
                parse_xml(saved.as_bytes(), "saved", &mut ParseBudget::new(&limits)).unwrap();
            let drawing = drawing_mut(parsed.roots.first_mut().unwrap()).unwrap();
            let mut reopened = Shape::empty("rect".to_owned(), shape.size.clone());
            apply_frame_metadata(&mut reopened, drawing);
            assert_eq!(reopened.size, shape.size);
            assert_eq!(reopened.effect_extent, shape.effect_extent);
            assert_eq!(reopened.id, shape.id);
            assert_eq!(reopened.name, shape.name);
            assert_eq!(reopened.title, shape.title);
            assert_eq!(reopened.description, shape.description);
            if anchored {
                assert_eq!(
                    reopened.position.as_ref().unwrap().horizontal.pos_offset,
                    Some(500.0)
                );
                assert_eq!(
                    reopened.position.as_ref().unwrap().vertical.pos_offset,
                    Some(600.0)
                );
                assert_eq!(reopened.relative_height, Some(29.0));
            }
            assert_eq!(graphic_events(&saved), graphic_events(&original));
            assert!(saved.contains("<w:noProof/>"));
            assert!(!saved.contains("wps:wsp"));
            assert!(context.take_warnings().is_empty());
        }
    }

    #[test]
    fn smart_art_unsupported_content_edits_warn_and_keep_source_content() {
        let (mut shape, original) = source(false);
        shape.shape_type = "ellipse".to_owned();
        shape.size.width = 2_743_200.0;
        let mut context = context();
        let saved = replay(&shape, &mut context, true).unwrap().unwrap();
        assert_eq!(saved, original.replace("cx=\"5486400\"", "cx=\"2743200\""));
        assert_eq!(
            context.take_warnings(),
            [
                "SmartArt diagram content or geometry changes cannot be represented; preserving source diagram content and saving supported frame changes."
            ]
        );
    }

    #[test]
    fn smart_art_body_namespaces_are_closed_on_captured_drawings() {
        let (_, source) = source(false);
        let source = source
            .replace(
                " xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"",
                "",
            )
            .replace("<w:rPr>", "<w:rPr><meta:preserve/>");
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:meta="urn:run-metadata">{source}</w:body></w:document>"#
        );
        let limits = ParseLimits::default();
        let parsed = parse_xml(xml.as_bytes(), "source", &mut ParseBudget::new(&limits)).unwrap();
        let run = parsed
            .root()
            .unwrap()
            .child_by_local_name("body")
            .unwrap()
            .child_by_local_name("r")
            .unwrap();
        assert_eq!(
            run.attributes.get("xmlns:meta").map(String::as_str),
            Some("urn:run-metadata")
        );
        let drawing = parsed
            .root()
            .unwrap()
            .child_by_local_name("body")
            .unwrap()
            .child_by_local_name("r")
            .unwrap()
            .child_by_local_name("drawing")
            .unwrap();
        assert_eq!(
            drawing.attributes.get("xmlns:dgm").map(String::as_str),
            Some("http://schemas.openxmlformats.org/drawingml/2006/diagram")
        );
    }
}
