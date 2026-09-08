//! Speaker notes: reads the notes body placeholder's plain text out of a
//! `notesSlide` part, and mints or patches that part on write. Everything
//! else in the part (slide image, slide number, master link) is left as is.

use crate::PptxError;
use crate::relationships::{Relationship, relationship_types};
use crate::xml::{ParseBudget, XmlElement, XmlNode, parse_xml, serialize_xml};

pub(crate) const CT_NOTES_SLIDE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";

const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

pub(crate) fn slide_notes_part(relationships: &[Relationship]) -> Option<String> {
    relationships
        .iter()
        .find(|relationship| relationship.is_type(relationship_types::NOTES_SLIDE))
        .and_then(|relationship| relationship.resolved_target.clone())
}

/// The notes body placeholder's plain text, paragraphs joined by `\n`.
/// Empty when the part has no body placeholder.
pub(crate) fn parse_notes_text(
    bytes: &[u8],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<String, PptxError> {
    let root = parse_xml(bytes, part, budget)?;
    let Some(tree) = root.child("cSld").and_then(|common| common.child("spTree")) else {
        return Ok(String::new());
    };
    let Some(body) = notes_body_shape(tree).and_then(|shape| shape.child("txBody")) else {
        return Ok(String::new());
    };
    Ok(body
        .children_named("p")
        .map(|paragraph| paragraph.text_content())
        .collect::<Vec<_>>()
        .join("\n"))
}

fn notes_body_shape(tree: &XmlElement) -> Option<&XmlElement> {
    tree.children_named("sp").find(|shape| is_notes_body(shape))
}

fn is_notes_body(shape: &XmlElement) -> bool {
    shape
        .child("nvSpPr")
        .and_then(|nv| nv.child("nvPr"))
        .and_then(|nv_pr| nv_pr.child("ph"))
        .is_some_and(|ph| ph.attribute("type") == Some("body"))
}

/// Replaces the notes body placeholder's paragraphs in an existing part,
/// inserting a fresh body placeholder shape if the part does not have one
/// yet (a notes part can exist with only a slide-image placeholder, or none
/// at all).
pub(crate) fn patch_notes_xml(
    bytes: &[u8],
    part: &str,
    text: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<u8>, PptxError> {
    let mut root = parse_xml(bytes, part, budget)?;
    let tree = match root
        .child_mut("cSld")
        .and_then(|common| common.child_mut("spTree"))
    {
        Some(tree) => tree,
        None => return Ok(serialize_xml(&root)),
    };
    let existing = tree.children.iter_mut().find_map(|child| match child {
        XmlNode::Element(element) if element.local_name() == "sp" && is_notes_body(element) => {
            Some(element)
        }
        _ => None,
    });
    match existing {
        Some(shape) => {
            let body = match shape.child_mut("txBody") {
                Some(body) => body,
                None => {
                    shape
                        .children
                        .push(XmlNode::Element(XmlElement::new("p:txBody")));
                    shape.child_mut("txBody").expect("just inserted")
                }
            };
            body.children.retain(
                |child| !matches!(child, XmlNode::Element(element) if element.local_name() == "p"),
            );
            for line in text.split('\n') {
                body.children.push(XmlNode::Element(notes_paragraph(line)));
            }
        }
        None => tree
            .children
            .push(XmlNode::Element(notes_body_shape_xml(text))),
    }
    Ok(serialize_xml(&root))
}

/// Mints a minimal `notesSlide` part for a slide that has none yet: a single
/// body placeholder, no slide image or slide number placeholder.
pub(crate) fn notes_slide_xml(text: &str) -> Vec<u8> {
    let tree = XmlElement::new("p:spTree")
        .with_child(
            XmlElement::new("p:nvGrpSpPr")
                .with_child(
                    XmlElement::new("p:cNvPr")
                        .with_attribute("id", "1")
                        .with_attribute("name", ""),
                )
                .with_child(XmlElement::new("p:cNvGrpSpPr"))
                .with_child(XmlElement::new("p:nvPr")),
        )
        .with_child(XmlElement::new("p:grpSpPr"))
        .with_child(notes_body_shape_xml(text));
    let root = XmlElement::new("p:notes")
        .with_attribute("xmlns:a", NS_A)
        .with_attribute("xmlns:p", NS_P)
        .with_child(XmlElement::new("p:cSld").with_child(tree));
    serialize_xml(&root)
}

/// A `<p:sp>` body placeholder shape carrying `text` as its paragraphs.
fn notes_body_shape_xml(text: &str) -> XmlElement {
    let mut body = XmlElement::new("p:txBody")
        .with_child(XmlElement::new("a:bodyPr"))
        .with_child(XmlElement::new("a:lstStyle"));
    for line in text.split('\n') {
        body = body.with_child(notes_paragraph(line));
    }
    XmlElement::new("p:sp")
        .with_child(
            XmlElement::new("p:nvSpPr")
                .with_child(
                    XmlElement::new("p:cNvPr")
                        .with_attribute("id", "2")
                        .with_attribute("name", "Notes Placeholder"),
                )
                .with_child(
                    XmlElement::new("p:cNvSpPr")
                        .with_child(XmlElement::new("a:spLocks").with_attribute("noGrp", "1")),
                )
                .with_child(
                    XmlElement::new("p:nvPr").with_child(
                        XmlElement::new("p:ph")
                            .with_attribute("type", "body")
                            .with_attribute("idx", "1"),
                    ),
                ),
        )
        .with_child(
            XmlElement::new("p:spPr").with_child(
                XmlElement::new("a:xfrm")
                    .with_child(
                        XmlElement::new("a:off")
                            .with_attribute("x", "685800")
                            .with_attribute("y", "4351338"),
                    )
                    .with_child(
                        XmlElement::new("a:ext")
                            .with_attribute("cx", "5486400")
                            .with_attribute("cy", "3200400"),
                    ),
            ),
        )
        .with_child(body)
}

fn notes_paragraph(text: &str) -> XmlElement {
    XmlElement::new("a:p").with_child(
        XmlElement::new("a:r").with_child(XmlElement::new("a:t").with_text(text.to_owned())),
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::ParseLimits;

    fn budget(limits: &ParseLimits) -> ParseBudget<'_> {
        ParseBudget::new(limits)
    }

    const EXISTING: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="3" name="Slide Image Placeholder"/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Old line</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#;

    #[test]
    fn parses_body_placeholder_text() {
        let limits = ParseLimits::default();
        let mut budget = budget(&limits);
        let text = parse_notes_text(EXISTING, "ppt/notesSlides/notesSlide1.xml", &mut budget)
            .expect("parse");
        assert_eq!(text, "Old line");
    }

    #[test]
    fn parses_multi_paragraph_text() {
        let bytes = br#"<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Line one</a:t></a:r></a:p><a:p><a:r><a:t>Line two</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#;
        let limits = ParseLimits::default();
        let mut budget = budget(&limits);
        let text =
            parse_notes_text(bytes, "ppt/notesSlides/notesSlide1.xml", &mut budget).expect("parse");
        assert_eq!(text, "Line one\nLine two");
    }

    #[test]
    fn missing_body_placeholder_is_empty() {
        let bytes = br#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:notes>"#;
        let limits = ParseLimits::default();
        let mut budget = budget(&limits);
        let text =
            parse_notes_text(bytes, "ppt/notesSlides/notesSlide1.xml", &mut budget).expect("parse");
        assert_eq!(text, "");
    }

    #[test]
    fn patch_preserves_other_shapes_and_replaces_text() {
        let limits = ParseLimits::default();
        let mut budget = budget(&limits);
        let patched = patch_notes_xml(
            EXISTING,
            "ppt/notesSlides/notesSlide1.xml",
            "New line one\nNew line two",
            &mut budget,
        )
        .expect("patch");
        let mut budget = budget_unused(&limits);
        let text = parse_notes_text(&patched, "ppt/notesSlides/notesSlide1.xml", &mut budget)
            .expect("parse patched");
        assert_eq!(text, "New line one\nNew line two");
        let patched_str = String::from_utf8(patched).expect("utf8");
        assert!(patched_str.contains("Slide Image Placeholder"));
        assert!(!patched_str.contains("Old line"));
    }

    fn budget_unused(limits: &ParseLimits) -> ParseBudget<'_> {
        ParseBudget::new(limits)
    }

    #[test]
    fn patch_inserts_a_body_placeholder_when_the_part_has_none() {
        let bytes = br#"<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:notes>"#;
        let limits = ParseLimits::default();
        let mut budget = budget(&limits);
        let patched = patch_notes_xml(
            bytes,
            "ppt/notesSlides/notesSlide1.xml",
            "Fresh text",
            &mut budget,
        )
        .expect("patch");
        let mut budget = budget_unused(&limits);
        let text = parse_notes_text(&patched, "ppt/notesSlides/notesSlide1.xml", &mut budget)
            .expect("parse patched");
        assert_eq!(text, "Fresh text");
    }

    #[test]
    fn mints_minimal_valid_part() {
        let bytes = notes_slide_xml("Fresh notes");
        let limits = ParseLimits::default();
        let mut budget = budget(&limits);
        let text = parse_notes_text(&bytes, "ppt/notesSlides/notesSlide1.xml", &mut budget)
            .expect("parse minted");
        assert_eq!(text, "Fresh notes");
    }
}
