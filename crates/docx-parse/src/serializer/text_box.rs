use crate::block::BlockContent;
use crate::shape::Shape;
use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, XmlNode, parse_xml};

use super::context::SerializerContext;
use super::raw::validate_replayed_fragment;
use super::s13::element_span;
use super::sdt::serialize_block_content;

fn drawing_path(element: &XmlElement) -> Option<Vec<u32>> {
    if element.name == "w:drawing" {
        return Some(Vec::new());
    }
    element
        .child_elements()
        .enumerate()
        .find_map(|(index, child)| {
            let mut path = drawing_path(child)?;
            path.insert(0, index as u32);
            Some(path)
        })
}

fn element_at<'a>(element: &'a XmlElement, path: &[u32]) -> Option<&'a XmlElement> {
    let Some((&index, rest)) = path.split_first() else {
        return Some(element);
    };
    element_at(element.child_elements().nth(index as usize)?, rest)
}

fn at_path<'a>(element: &'a mut XmlElement, path: &[u32]) -> Option<&'a mut XmlElement> {
    let Some((&index, rest)) = path.split_first() else {
        return Some(element);
    };
    let child = element
        .children
        .iter_mut()
        .filter_map(|node| match node {
            XmlNode::Element(child) => Some(child),
            _ => None,
        })
        .nth(index as usize)?;
    at_path(child, rest)
}

fn content_paths(element: &XmlElement, path: &mut Vec<u32>, paths: &mut Vec<Vec<u32>>) {
    if element.name == "w:txbxContent" {
        paths.push(path.clone());
        return;
    }
    for (index, child) in element.child_elements().enumerate() {
        path.push(index as u32);
        content_paths(child, path, paths);
        path.pop();
    }
}

/// Replays a group or canvas from the part's source drawing with the same `wp:docPr` id,
/// while the model matches that source apart from its frame and text-box text.
/// Whether the model still says what its retained source says, apart from text and media. A
/// stored hash can differ without an edit when it was computed with other number parsing.
fn matches_source(drawing: &XmlElement, shape: &Shape) -> bool {
    let Some(mut baseline) = crate::shape::parse_shape_from_drawing(drawing) else {
        return false;
    };
    crate::paragraph::apply_shape_metadata(&mut baseline, drawing);
    crate::smart_art::frame_model(&baseline) == crate::smart_art::frame_model(shape)
        && crate::text_box::group_model_sha256(&baseline)
            == crate::text_box::group_model_sha256(shape)
}

pub(super) fn replay_group(
    shape: &Shape,
    context: &mut SerializerContext,
) -> Result<Option<String>, ParseError> {
    if shape.children.is_none() || shape.drawing_xml.is_some() || shape.drawing_run_xml.is_some() {
        return Ok(None);
    }
    let Some(source) = shape
        .id
        .as_deref()
        .and_then(|id| context.group_drawing(id))
        .map(str::to_owned)
    else {
        return Ok(None);
    };
    let limits = ParseLimits::default();
    let Ok(parsed) = parse_xml(
        source.as_bytes(),
        "group source",
        &mut ParseBudget::new(&limits),
    ) else {
        return Ok(None);
    };
    let Some(drawing) = parsed
        .roots
        .first()
        .filter(|root| root.local_name() == "drawing")
    else {
        return Ok(None);
    };
    let Some(mut baseline) = crate::shape::parse_shape_from_drawing(drawing) else {
        return Ok(None);
    };
    crate::paragraph::apply_shape_metadata(&mut baseline, drawing);
    let mut containers = Vec::new();
    content_paths(drawing, &mut Vec::new(), &mut containers);
    let bodies = crate::text_box::shape_text_bodies(shape);
    let baseline_bodies = crate::text_box::shape_text_bodies(&baseline);
    if bodies.len() != baseline_bodies.len() {
        return Ok(None);
    }
    let mask: Vec<bool> = baseline_bodies
        .iter()
        .map(|body| !body.content.is_empty())
        .collect();
    let mapped = mask.iter().filter(|kept| **kept).count() == containers.len();
    let mut kept = 0;
    let mut hashes = Vec::with_capacity(bodies.len());
    for (body, has_box) in bodies.iter().zip(&mask) {
        let current = crate::text_box::content_sha256(&body.content);
        let unchanged = (!*has_box && (mapped || body.content.is_empty()))
            || (mapped
                && element_at(drawing, &containers[kept])
                    .is_some_and(|container| same_text(container, &body.content, context)));
        if *has_box {
            kept += 1;
        }
        hashes.push(if unchanged { current } else { String::new() });
    }
    let mut replayed = shape.clone();
    replayed.drawing_xml = Some(source);
    replayed.drawing_run_xml = None;
    replayed.drawing_model = Some(serde_json::json!({
        "frame": crate::smart_art::frame_model(&baseline),
        "otherSha256": crate::smart_art::other_model_sha256(&baseline),
        "textBoxContentSha256": hashes,
        "textBoxBodies": mapped.then_some(mask),
        "groupSha256": crate::text_box::group_model_sha256(&baseline),
    }));
    replay(&replayed, context, false)
}

/// Whether the model text of one text box writes the same XML as its source paragraphs.
fn same_text(
    container: &XmlElement,
    content: &[serde_json::Value],
    context: &SerializerContext,
) -> bool {
    let mut scratch = context.clone();
    let mut model = String::new();
    for value in content {
        let Ok(block) = serde_json::from_value::<BlockContent>(value.clone()) else {
            return false;
        };
        let Ok(xml) = serialize_block_content(&block, &mut scratch) else {
            return false;
        };
        model.push_str(&xml);
    }
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let Ok(mut ids) = crate::paragraph::HexIdAllocator::from_sha256(&"0".repeat(64)) else {
        return false;
    };
    let mut scratch = context.clone();
    let mut source = String::new();
    for child in container.child_elements() {
        if child.local_name() != "p" {
            return false;
        }
        let Ok(paragraph) = crate::paragraph::parse_paragraph(
            child,
            None,
            None,
            None,
            None,
            None,
            "group source",
            &mut budget,
            &mut ids,
            None,
            false,
            0,
        ) else {
            return false;
        };
        let Ok(xml) = serialize_block_content(
            &BlockContent::Paragraph(std::sync::Arc::new(paragraph)),
            &mut scratch,
        ) else {
            return false;
        };
        source.push_str(&xml);
    }
    model == source
}

pub(super) fn replay(
    shape: &Shape,
    context: &mut SerializerContext,
    whole_run: bool,
) -> Result<Option<String>, ParseError> {
    let Some(model) = shape.drawing_model.as_ref() else {
        return Ok(None);
    };
    let Some(baseline) = model
        .get("textBoxContentSha256")
        .and_then(serde_json::Value::as_array)
    else {
        return Ok(None);
    };
    let group = model.get("groupSha256").and_then(serde_json::Value::as_str);
    if group.is_some_and(|before| before != crate::text_box::group_model_sha256(shape)) {
        return Ok(None);
    }
    let Some(source) = shape
        .drawing_run_xml
        .as_deref()
        .or(shape.drawing_xml.as_deref())
    else {
        return Ok(None);
    };
    validate_replayed_fragment(source)?;
    let limits = ParseLimits::default();
    let mut parsed = parse_xml(
        source.as_bytes(),
        "text box source",
        &mut ParseBudget::new(&limits),
    )?;
    let root = parsed
        .roots
        .first_mut()
        .ok_or_else(|| ParseError::Canonical("text box source has no root".to_owned()))?;
    let root_is_run = root.name == "w:r";
    if whole_run && !root_is_run {
        return Ok(None);
    }
    let drawing = drawing_path(root)
        .ok_or_else(|| ParseError::Canonical("text box source has no drawing".to_owned()))?;
    let mut paths = Vec::new();
    content_paths(
        at_path(root, &drawing).unwrap(),
        &mut drawing.clone(),
        &mut paths,
    );
    let mut mapped = false;
    let mut bodies = crate::text_box::shape_text_bodies(shape);
    let mut baseline: Vec<&serde_json::Value> = baseline.iter().collect();
    if let Some(mask) = model
        .get("textBoxBodies")
        .and_then(serde_json::Value::as_array)
        .filter(|mask| mask.len() == bodies.len() && mask.len() == baseline.len())
    {
        mapped = true;
        let kept = |index: &usize| mask[*index].as_bool() == Some(true);
        if bodies
            .iter()
            .enumerate()
            .any(|(index, body)| !kept(&index) && !body.content.is_empty())
        {
            return Ok(None);
        }
        bodies = bodies
            .into_iter()
            .enumerate()
            .filter(|(index, _)| kept(index))
            .map(|(_, body)| body)
            .collect();
        baseline = baseline
            .into_iter()
            .enumerate()
            .filter(|(index, _)| kept(index))
            .map(|(_, before)| before)
            .collect();
    }
    let unchanged = baseline.len() == bodies.len()
        && bodies.iter().zip(&baseline).all(|(body, before)| {
            before.as_str() == Some(crate::text_box::content_sha256(&body.content).as_str())
        });
    if group.is_some() && !unchanged && !mapped {
        return Ok(None);
    }
    if baseline.len() != bodies.len() || (paths.len() != bodies.len() && !unchanged) {
        if group.is_some() {
            return Ok(None);
        }
        return Err(ParseError::Canonical(
            "text box source content mismatch".to_owned(),
        ));
    }
    let mut output = source.to_owned();
    let frame_changed = model
        .get("frame")
        .is_some_and(|frame| crate::smart_art::frame_model(shape) != *frame);
    let redraw = shape.children.is_none()
        && (frame_changed
            || model
                .get("otherSha256")
                .and_then(serde_json::Value::as_str)
                .is_some_and(|before| before != crate::smart_art::other_model_sha256(shape)))
        && !matches_source(at_path(root, &drawing).unwrap(), shape);
    if redraw {
        let mut authored = shape.clone();
        authored.drawing_xml = None;
        authored.drawing_run_xml = None;
        authored.drawing_model = None;
        let xml = super::run::serialize_shape_content(&authored, context)?;
        let span = element_span(source, &drawing).ok_or_else(|| {
            ParseError::Canonical("text box drawing has no source span".to_owned())
        })?;
        output.replace_range(span, &xml);
    } else if paths.len() == bodies.len() {
        for ((path, body), before) in paths.iter().zip(bodies).zip(baseline).rev() {
            if before.as_str() == Some(crate::text_box::content_sha256(&body.content).as_str()) {
                continue;
            }
            let span = element_span(source, path).ok_or_else(|| {
                ParseError::Canonical("text box content has no source span".to_owned())
            })?;
            let fragment = &source[span.clone()];
            let tag = super::splice::root_tag(fragment)
                .ok_or_else(|| ParseError::Canonical("text box content has no root".to_owned()))?;
            let mut xml = if tag.empty {
                format!("{}>", &fragment[..tag.range.end - 2])
            } else {
                fragment[..tag.range.end].to_owned()
            };
            for value in &body.content {
                let block: BlockContent = serde_json::from_value(value.clone())
                    .map_err(|error| ParseError::Canonical(error.to_string()))?;
                xml.push_str(&serialize_block_content(&block, context)?);
            }
            xml.push_str(&format!("</{}>", tag.name));
            output.replace_range(span, &xml);
        }
    }
    if let Some(frame) = model.get("frame")
        && frame_changed
        && !redraw
    {
        parsed = parse_xml(
            output.as_bytes(),
            "text box source",
            &mut ParseBudget::new(&limits),
        )?;
        let root = parsed.roots.first_mut().unwrap();
        let drawing_element = at_path(root, &drawing).unwrap();
        super::smart_art::update_frame(drawing_element, shape, frame, context);
        let xml = drawing_element.to_bookmarked_raw_inline_xml();
        let span = element_span(&output, &drawing).ok_or_else(|| {
            ParseError::Canonical("text box drawing has no source span".to_owned())
        })?;
        output.replace_range(span, &xml);
    }
    if !whole_run && root_is_run {
        let root = parsed.roots.first().unwrap();
        let index = root
            .child_elements()
            .position(|child| matches!(child.local_name(), "drawing" | "AlternateContent"))
            .ok_or_else(|| ParseError::Canonical("text box run has no drawing".to_owned()))?;
        let span = element_span(&output, &[index as u32]).ok_or_else(|| {
            ParseError::Canonical("text box drawing has no source span".to_owned())
        })?;
        output = output[span].to_owned();
    }
    Ok(Some(output))
}
