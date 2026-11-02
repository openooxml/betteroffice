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
    let bodies = crate::text_box::shape_text_bodies(shape);
    let unchanged = baseline.len() == bodies.len()
        && bodies.iter().zip(baseline).all(|(body, before)| {
            before.as_str() == Some(crate::text_box::content_sha256(&body.content).as_str())
        });
    if baseline.len() != bodies.len() || (paths.len() != bodies.len() && !unchanged) {
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
                .is_some_and(|before| before != crate::smart_art::other_model_sha256(shape)));
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
