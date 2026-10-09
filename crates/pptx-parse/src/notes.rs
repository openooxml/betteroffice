//! Speaker notes parsing and write-back.

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

/// Reads speaker notes as plain text.
#[cfg(test)]
pub(crate) fn parse_notes_text(
    bytes: &[u8],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<String, PptxError> {
    Ok(notes_text(&parse_xml(bytes, part, budget)?))
}

/// A parsed notes page's speaker notes as plain text.
pub(crate) fn notes_text(root: &XmlElement) -> String {
    let Some(tree) = root.child("cSld").and_then(|common| common.child("spTree")) else {
        return String::new();
    };
    let Some(body) = notes_body_shape(tree).and_then(|shape| shape.child("txBody")) else {
        return String::new();
    };
    body.children_named("p")
        .map(|paragraph| {
            paragraph
                .children
                .iter()
                .filter_map(|node| match node {
                    XmlNode::Element(element) => match element.local_name() {
                        "r" | "fld" => Some(
                            element
                                .children_named("t")
                                .map(XmlElement::text_content)
                                .collect::<String>(),
                        ),
                        "br" => Some("\n".to_owned()),
                        _ => None,
                    },
                    _ => None,
                })
                .collect::<String>()
        })
        .collect::<Vec<_>>()
        .join("\n")
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

/// Replaces notes text while preserving other notes-page content.
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
        None => {
            return Err(PptxError::Write {
                part: part.to_owned(),
                message: "notes slide has no shape tree".to_owned(),
            });
        }
    };
    let shape_id = next_shape_id(tree);
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
                    shape.children.push(XmlNode::Element(notes_text_body("")));
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
            .push(XmlNode::Element(notes_body_shape_xml(text, shape_id))),
    }
    Ok(serialize_xml(&root))
}

/// Mints a minimal notes slide.
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
        .with_child(notes_body_shape_xml(text, 2));
    let root = XmlElement::new("p:notes")
        .with_attribute("xmlns:a", NS_A)
        .with_attribute("xmlns:p", NS_P)
        .with_child(XmlElement::new("p:cSld").with_child(tree));
    serialize_xml(&root)
}

/// A `<p:sp>` body placeholder shape carrying `text` as its paragraphs.
fn notes_body_shape_xml(text: &str, id: u64) -> XmlElement {
    XmlElement::new("p:sp")
        .with_attribute("xmlns:p", NS_P)
        .with_attribute("xmlns:a", NS_A)
        .with_child(
            XmlElement::new("p:nvSpPr")
                .with_child(
                    XmlElement::new("p:cNvPr")
                        .with_attribute("id", id.to_string())
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
        .with_child(notes_text_body(text))
}

fn next_shape_id(element: &XmlElement) -> u64 {
    element
        .children
        .iter()
        .filter_map(|node| match node {
            XmlNode::Element(child) => Some(next_shape_id(child)),
            _ => None,
        })
        .chain(
            (element.local_name() == "cNvPr")
                .then(|| {
                    element
                        .attribute("id")
                        .and_then(|id| id.parse::<u32>().ok())
                        .map(|id| u64::from(id) + 1)
                })
                .flatten(),
        )
        .max()
        .unwrap_or(1)
}

fn notes_text_body(text: &str) -> XmlElement {
    let mut body = XmlElement::new("p:txBody")
        .with_attribute("xmlns:p", NS_P)
        .with_attribute("xmlns:a", NS_A)
        .with_child(XmlElement::new("a:bodyPr"))
        .with_child(XmlElement::new("a:lstStyle"));
    for line in text.split('\n') {
        body = body.with_child(notes_paragraph(line));
    }
    body
}

fn notes_paragraph(text: &str) -> XmlElement {
    XmlElement::new("a:p")
        .with_attribute("xmlns:a", NS_A)
        .with_child(
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
        let mut budget = ParseBudget::new(&limits);
        let text = parse_notes_text(&patched, "ppt/notesSlides/notesSlide1.xml", &mut budget)
            .expect("parse patched");
        assert_eq!(text, "New line one\nNew line two");
        let patched_str = String::from_utf8(patched).expect("utf8");
        assert!(patched_str.contains("Slide Image Placeholder"));
        assert!(!patched_str.contains("Old line"));
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
        let mut budget = ParseBudget::new(&limits);
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
    #[test]
    fn reads_soft_breaks_and_fields_without_extension_text() {
        let xml = String::from_utf8(EXISTING.to_vec()).unwrap().replace(
            "<a:r><a:t>Old line</a:t></a:r>",
            "<a:r><a:t>First</a:t></a:r><a:br/><a:fld id=\"field\"><a:t>Second</a:t></a:fld><a:extLst><a:ext>Hidden</a:ext></a:extLst>",
        );
        let limits = ParseLimits::default();
        assert_eq!(
            parse_notes_text(xml.as_bytes(), "notes.xml", &mut budget(&limits)).unwrap(),
            "First\nSecond"
        );
    }

    #[test]
    fn inserting_notes_avoids_shape_id_collisions_and_declares_namespaces() {
        let xml = String::from_utf8(EXISTING.to_vec())
            .unwrap()
            .replace("type=\"body\"", "type=\"sldNum\"")
            .replace("xmlns:p=", "xmlns:ppt=")
            .replace("p:", "ppt:")
            .replace("xmlns:a=", "xmlns:d=")
            .replace("a:", "d:");
        let limits = ParseLimits::default();
        let patched = patch_notes_xml(
            xml.as_bytes(),
            "notes.xml",
            "New notes",
            &mut budget(&limits),
        )
        .unwrap();
        let root = parse_xml(&patched, "notes.xml", &mut budget(&limits)).unwrap();
        let tree = root.child("cSld").unwrap().child("spTree").unwrap();
        let body = notes_body_shape(tree).unwrap();
        assert_eq!(body.attribute("xmlns:p"), Some(NS_P));
        assert_eq!(body.attribute("xmlns:a"), Some(NS_A));
        assert_eq!(
            body.child("nvSpPr")
                .unwrap()
                .child("cNvPr")
                .unwrap()
                .attribute("id"),
            Some("4")
        );
        assert_eq!(
            parse_notes_text(&patched, "notes.xml", &mut budget(&limits)).unwrap(),
            "New notes"
        );
    }

    #[test]
    fn patching_a_placeholder_without_text_adds_required_body_properties() {
        let xml = String::from_utf8(EXISTING.to_vec()).unwrap();
        let start = xml.find("<p:txBody>").unwrap();
        let end = xml.find("</p:txBody>").unwrap() + "</p:txBody>".len();
        let xml = format!("{}{}", &xml[..start], &xml[end..]);
        let limits = ParseLimits::default();
        let patched = patch_notes_xml(
            xml.as_bytes(),
            "notes.xml",
            "New notes",
            &mut budget(&limits),
        )
        .unwrap();
        let root = parse_xml(&patched, "notes.xml", &mut budget(&limits)).unwrap();
        let body = notes_body_shape(root.child("cSld").unwrap().child("spTree").unwrap())
            .unwrap()
            .child("txBody")
            .unwrap();
        assert!(body.child("bodyPr").is_some());
        assert_eq!(body.children_named("p").count(), 1);
    }
}
