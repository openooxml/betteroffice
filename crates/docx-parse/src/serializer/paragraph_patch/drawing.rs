use serde_json::Value;

use crate::paragraph_patch::{ParagraphPatchV1, PatchErrorCode, UnitV1};
use crate::xml::{ParseBudget, ParseLimits, parse_xml};

use super::lower::{Cell, Lowered, Node, Xml};
use super::{Failure, flatten};

const DIAGRAM: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

pub(super) fn is_smart_art(source: &str, xml: &Xml, node: &Node) -> bool {
    if !node.word("drawing") {
        return false;
    }
    let mut pending = node.children.clone();
    while let Some(at) = pending.pop() {
        #[cfg(test)]
        super::work::add(1);
        let child = &xml.nodes[at];
        if child.namespace == DRAWING
            && child.local == "graphicData"
            && child.attr(source, "uri").as_deref() == Some(DIAGRAM)
            && child
                .children
                .iter()
                .any(|&at| xml.nodes[at].namespace == DIAGRAM && xml.nodes[at].local == "relIds")
        {
            return true;
        }
        pending.extend(&child.children);
    }
    false
}

fn canonical(xml: &str) -> Option<String> {
    let limits = ParseLimits::default();
    let document = parse_xml(xml.as_bytes(), "drawing", &mut ParseBudget::new(&limits)).ok()?;
    Some(document.root()?.to_raw_inline_xml())
}

fn equivalent(source: &str, lowered: &Lowered, node: usize, left: &str, right: &str) -> bool {
    let mut namespaces = lowered.xml.root_namespaces.clone();
    let mut ancestors = Vec::new();
    let mut next = Some(node);
    while let Some(at) = next {
        #[cfg(test)]
        super::work::add(1);
        ancestors.push(at);
        next = lowered.xml.nodes[at].parent;
    }
    for at in ancestors.into_iter().rev() {
        let ancestor = &lowered.xml.nodes[at];
        for attribute in &ancestor.attributes {
            if let Some(prefix) = attribute.name.strip_prefix("xmlns:") {
                if let Some(value) = ancestor.attr(source, &attribute.name) {
                    namespaces.insert(prefix.to_owned(), value);
                }
            } else if attribute.name == "xmlns"
                && let Some(value) = ancestor.attr(source, "xmlns")
            {
                namespaces.insert(String::new(), value);
            }
        }
    }
    let mut wrapper = String::from("<sourceContext");
    for (prefix, namespace) in namespaces {
        wrapper.push_str(&format!(
            " {}=\"{}\"",
            if prefix.is_empty() {
                "xmlns".to_owned()
            } else {
                format!("xmlns:{prefix}")
            },
            crate::serializer::xml_writer::escape_xml(&namespace)
        ));
    }
    wrapper.push('>');
    let events = |xml: &str| {
        crate::serializer::canonical_xml_events(&format!("{wrapper}{xml}</sourceContext>"), false)
    };
    match (events(left), events(right)) {
        (Ok(left), Ok(right)) => left == right,
        _ => false,
    }
}

fn seeded_smart_art(unit: &UnitV1) -> bool {
    let UnitV1::Atom {
        atom_kind, payload, ..
    } = unit
    else {
        return false;
    };
    match atom_kind.as_str() {
        "shape" => payload
            .get("shapeJson")
            .and_then(Value::as_str)
            .and_then(|json| serde_json::from_str::<Value>(json).ok())
            .is_some_and(|shape| {
                ["drawingXml", "drawingRunXml"]
                    .iter()
                    .any(|key| shape.get(key).and_then(Value::as_str).is_some())
            }),
        "opaqueDrawing" => payload
            .get("xml")
            .and_then(Value::as_str)
            .is_some_and(|xml| xml.contains(DIAGRAM)),
        _ => false,
    }
}

fn smart_art_cell(source: &str, lowered: &Lowered, cell: &Cell) -> bool {
    let node = &lowered.xml.nodes[cell.node];
    node.word("drawing")
        && (is_smart_art(source, &lowered.xml, node) || seeded_smart_art(&cell.unit))
}

pub(super) fn bind(source: &str, lowered: &Lowered, units: &[UnitV1]) -> Result<(), Failure> {
    if !lowered
        .cells
        .iter()
        .any(|cell| smart_art_cell(source, lowered, cell))
    {
        return Ok(());
    }
    for (at, expected) in flatten(units)
        .into_iter()
        .enumerate()
        .take(lowered.cells.len())
    {
        let cell = &lowered.cells[at];
        if !smart_art_cell(source, lowered, cell) {
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
                .and_then(|shape| {
                    shape
                        .get("drawingRunXml")
                        .or_else(|| shape.get("drawingXml"))
                })
                .and_then(Value::as_str),
            "opaqueDrawing" => payload.get("xml").and_then(Value::as_str),
            _ => None,
        };
        let actual = canonical(&source[lowered.xml.nodes[cell.node].span.clone()]);
        let xml = xml.and_then(|xml| {
            let whole_run = shape
                .as_ref()
                .and_then(|shape| shape.get("drawingRunXml"))
                .and_then(Value::as_str)
                .is_some()
                || (atom_kind == "opaqueDrawing"
                    && payload.get("kind").and_then(Value::as_str) == Some("drawingRun"));
            if whole_run {
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
        if !xml
            .as_deref()
            .zip(actual.as_deref())
            .is_some_and(|(expected, actual)| {
                equivalent(source, lowered, cell.node, expected, actual)
            })
        {
            return Err(Failure::new(
                PatchErrorCode::SeedMismatch,
                "drawing: SmartArt XML differs from seed",
            ));
        }
        if let Some(xml) = shape
            .as_ref()
            .and_then(|shape| shape.get("drawingRunXml"))
            .and_then(Value::as_str)
            && !equivalent(
                source,
                lowered,
                lowered.runs[cell.run].node,
                xml,
                &source[lowered.xml.nodes[lowered.runs[cell.run].node].span.clone()],
            )
        {
            return Err(Failure::new(
                PatchErrorCode::SeedMismatch,
                "drawing: SmartArt run differs from seed",
            ));
        }
    }
    Ok(())
}

pub(super) fn validate_edits(
    source: &str,
    lowered: &Lowered,
    patch: &ParagraphPatchV1,
) -> Result<(), Failure> {
    let drawings: Vec<_> = lowered
        .cells
        .iter()
        .filter(|cell| smart_art_cell(source, lowered, cell))
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
    let properties = |run: usize| {
        lowered.runs[run].properties.filter(|&at| {
            let node = &lowered.xml.nodes[at];
            !node.children.is_empty() || !node.attributes.is_empty()
        })
    };
    let same_properties = |left: usize, right: usize| match (properties(left), properties(right)) {
        (None, None) => true,
        (Some(left), Some(right)) => equivalent(
            source,
            lowered,
            left,
            &source[lowered.xml.nodes[left].span.clone()],
            &source[lowered.xml.nodes[right].span.clone()],
        ),
        _ => false,
    };
    for edit in &patch.text_edits {
        let mut runs = lowered
            .cells
            .iter()
            .filter(|cell| edit.from <= cell.from && cell.to <= edit.to)
            .map(|cell| cell.run);
        if let Some(first) = runs.next()
            && runs.any(|run| !same_properties(first, run))
        {
            return Err(Failure::structure("drawing: text edit crosses source runs"));
        }
    }
    Ok(())
}
