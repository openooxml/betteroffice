use serde_json::{Value, json};

use crate::paragraph_patch::{ParagraphPatchV1, PatchErrorCode, UnitV1};
use crate::xml::{ParseBudget, ParseLimits, parse_xml};

use super::lower::{Lowered, Node, Xml};
use super::{Failure, flatten};

const DIAGRAM: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

pub(super) fn is_smart_art(source: &str, xml: &Xml, node: &Node) -> bool {
    node.word("drawing")
        && xml.nodes.iter().any(|child| {
            node.span.start < child.span.start
                && child.span.end <= node.span.end
                && child.namespace == DRAWING
                && child.local == "graphicData"
                && child.attr(source, "uri").as_deref() == Some(DIAGRAM)
                && child.children.iter().any(|&at| {
                    xml.nodes[at].namespace == DIAGRAM && xml.nodes[at].local == "relIds"
                })
        })
}

fn canonical(xml: &str) -> Option<String> {
    let limits = ParseLimits::default();
    let document = parse_xml(xml.as_bytes(), "drawing", &mut ParseBudget::new(&limits)).ok()?;
    Some(document.root()?.to_raw_inline_xml())
}

pub(super) fn payload(source: &str, node: &Node) -> Result<super::Attrs, Failure> {
    let xml = canonical(&source[node.span.clone()])
        .ok_or_else(|| Failure::structure("drawing: invalid SmartArt XML"))?;
    Ok(super::Attrs::from([("drawingXml".to_owned(), json!(xml))]))
}

pub(super) fn bind(source: &str, lowered: &mut Lowered, units: &[UnitV1]) -> Result<(), Failure> {
    if !lowered
        .cells
        .iter()
        .any(|cell| lowered.xml.nodes[cell.node].word("drawing"))
    {
        return Ok(());
    }
    for (cell, expected) in lowered.cells.iter_mut().zip(flatten(units)) {
        if !lowered.xml.nodes[cell.node].word("drawing") {
            continue;
        }
        let UnitV1::Atom {
            atom_kind, payload, ..
        } = &expected
        else {
            return Err(Failure::new(
                PatchErrorCode::SeedMismatch,
                "drawing: missing SmartArt atom",
            ));
        };
        let shape = payload
            .get("shapeJson")
            .and_then(Value::as_str)
            .and_then(|json| serde_json::from_str::<Value>(json).ok());
        let xml = match atom_kind.as_str() {
            "shape" => shape
                .as_ref()
                .and_then(|shape| shape.get("drawingXml"))
                .and_then(Value::as_str),
            "opaqueDrawing" => payload.get("xml").and_then(Value::as_str),
            _ => None,
        };
        let UnitV1::Atom {
            payload: actual, ..
        } = &cell.unit
        else {
            unreachable!()
        };
        let xml = xml.and_then(|xml| {
            if atom_kind == "opaqueDrawing"
                && payload.get("kind").and_then(Value::as_str) == Some("drawingRun")
            {
                let limits = ParseLimits::default();
                let parsed = parse_xml(
                    xml.as_bytes(),
                    "drawing run",
                    &mut ParseBudget::new(&limits),
                )
                .ok()?;
                Some(
                    parsed
                        .root()?
                        .child_by_local_name("drawing")?
                        .to_raw_inline_xml(),
                )
            } else {
                canonical(xml)
            }
        });
        if xml.as_deref() != actual.get("drawingXml").and_then(Value::as_str) {
            return Err(Failure::new(
                PatchErrorCode::SeedMismatch,
                "drawing: SmartArt XML differs from seed",
            ));
        }
        if let Some(xml) = shape
            .as_ref()
            .and_then(|shape| shape.get("drawingRunXml"))
            .and_then(Value::as_str)
            && canonical(xml)
                != canonical(&source[lowered.xml.nodes[lowered.runs[cell.run].node].span.clone()])
        {
            return Err(Failure::new(
                PatchErrorCode::SeedMismatch,
                "drawing: SmartArt run differs from seed",
            ));
        }
        cell.unit = expected;
    }
    Ok(())
}

pub(super) fn validate_edits(lowered: &Lowered, patch: &ParagraphPatchV1) -> Result<(), Failure> {
    let drawings: Vec<_> = lowered
        .cells
        .iter()
        .filter(|cell| lowered.xml.nodes[cell.node].word("drawing"))
        .collect();
    if drawings.is_empty() {
        return Ok(());
    }
    for edit in &patch.run_edits {
        if drawings
            .iter()
            .any(|cell| edit.from < cell.to && cell.from < edit.to)
        {
            return Err(Failure::structure("drawing: formatting touches SmartArt"));
        }
    }
    for edit in &patch.text_edits {
        let mut runs = lowered
            .cells
            .iter()
            .filter(|cell| edit.from <= cell.from && cell.to <= edit.to)
            .map(|cell| cell.run);
        if let Some(first) = runs.next()
            && runs.any(|run| run != first)
        {
            return Err(Failure::structure("drawing: text edit crosses source runs"));
        }
    }
    Ok(())
}
