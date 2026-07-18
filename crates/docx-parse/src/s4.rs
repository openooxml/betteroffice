//! Complete S4 drawing/media package projection for the differential corpus gate.

use base64::Engine as _;
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::canonical::{canonical_sha256, from_serializable, to_canonical_bytes};
use crate::chart::{Chart, parse_chart_from_drawing, parse_chart_parts};
use crate::image::parse_drawing;
use crate::media::{MediaFile, build_media_map};
use crate::relationships::{RelationshipMap, parse_relationships};
use crate::shape::{is_shape_drawing, parse_shape_from_drawing, resolve_shape_fill_pictures};
use crate::smart_art::{
    SmartArtContext, create_smart_art_context, is_smart_art_drawing, parse_smart_art_from_drawing,
};
use crate::text_box::parse_text_box;
use crate::vml::{Watermark, extract_watermark, parse_vml_image_content};
use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, parse_xml};

const MAX_PROJECTION_DEPTH: usize = 64;
const MAX_DRAWING_LEAVES: usize = 100_000;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S4Projection {
    pub media_entries: Vec<(String, MediaFile)>,
    pub chart_entries: Vec<(String, Chart)>,
    pub xml_parts: Vec<S4XmlPart>,
    pub smart_art_warnings: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S4XmlPart {
    pub path: String,
    pub drawings: Vec<DrawingLeaf>,
    pub vml_images: Vec<DrawingLeaf>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub watermark: Option<Watermark>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct DrawingLeaf {
    pub element: String,
    pub kind: String,
    pub value: serde_json::Value,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S4WireEnvelope {
    pub wire_version: u8,
    pub projection: S4Projection,
    pub canonical_base64: String,
    pub canonical_sha256: String,
}

pub fn parse_docx_s4_projection(data: &[u8]) -> Result<S4Projection, ParseError> {
    let parts = ooxml_opc::unzip_parts(data).map_err(ParseError::Container)?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let media = build_media_map(&parts);
    let all_xml = parts
        .iter()
        .filter(|(path, _)| {
            let lower = path.to_ascii_lowercase();
            lower.ends_with(".xml") || lower.ends_with(".rels")
        })
        .cloned()
        .collect::<IndexMap<_, _>>();
    let charts = parse_chart_parts(&all_xml, &mut budget)?;
    let mut smart_art = create_smart_art_context(&all_xml);
    let mut relationship_parts = IndexMap::new();
    for (path, xml) in &all_xml {
        if path.to_ascii_lowercase().ends_with(".rels") {
            relationship_parts.insert(path.clone(), parse_relationships(xml, path, &mut budget)?);
        }
    }
    let document_relationships = relationship_parts.get("word/_rels/document.xml.rels");
    let mut xml_parts = Vec::new();
    for (path, xml) in &all_xml {
        if !path.to_ascii_lowercase().ends_with(".xml") || xml.contains(&0) {
            continue;
        }
        let document = parse_xml(xml, path, &mut budget)?;
        let Some(root) = document.root() else {
            continue;
        };
        let owner_path = relationship_path_for_part(path);
        let relationships = relationship_parts
            .get(&owner_path)
            .or(document_relationships);
        if let Some(part) = project_s4_xml_part(
            path,
            root,
            relationships,
            &media,
            &charts,
            &mut smart_art,
            &mut budget,
        )? {
            xml_parts.push(part);
        }
    }
    Ok(S4Projection {
        media_entries: media.into_iter().collect(),
        chart_entries: charts.into_iter().collect(),
        xml_parts,
        smart_art_warnings: smart_art.warnings,
    })
}

#[allow(clippy::too_many_arguments)]
pub fn project_s4_xml_part(
    path: &str,
    root: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: &crate::media::MediaMap,
    charts: &crate::chart::ChartPartsMap,
    smart_art: &mut SmartArtContext,
    budget: &mut ParseBudget<'_>,
) -> Result<Option<S4XmlPart>, ParseError> {
    let watermark = extract_watermark(Some(root), relationships, Some(media));
    let mut part = S4XmlPart {
        path: path.to_owned(),
        drawings: Vec::new(),
        vml_images: Vec::new(),
        watermark,
    };
    collect_projection_leaves(
        root,
        relationships,
        media,
        charts,
        smart_art,
        budget,
        &mut part,
        0,
    )?;
    Ok(
        (!part.drawings.is_empty() || !part.vml_images.is_empty() || part.watermark.is_some())
            .then_some(part),
    )
}

#[allow(clippy::too_many_arguments)]
fn collect_projection_leaves(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: &crate::media::MediaMap,
    charts: &crate::chart::ChartPartsMap,
    smart_art: &mut SmartArtContext,
    budget: &mut ParseBudget<'_>,
    output: &mut S4XmlPart,
    depth: usize,
) -> Result<(), ParseError> {
    if depth > MAX_PROJECTION_DEPTH
        || output.drawings.len() + output.vml_images.len() >= MAX_DRAWING_LEAVES
    {
        return Ok(());
    }
    match element.local_name() {
        "drawing" => {
            if let Some(leaf) =
                project_drawing(element, relationships, media, charts, smart_art, budget)?
            {
                output.drawings.push(leaf);
            }
        }
        "pict" | "object" => {
            if let Some(image) = parse_vml_image_content(element, relationships, Some(media)) {
                output.vml_images.push(DrawingLeaf {
                    element: element.local_name().to_owned(),
                    kind: "vmlImage".to_owned(),
                    value: serde_json::json!({ "type": "drawing", "image": image }),
                });
            }
        }
        _ => {}
    }
    for child in element.child_elements() {
        collect_projection_leaves(
            child,
            relationships,
            media,
            charts,
            smart_art,
            budget,
            output,
            depth + 1,
        )?;
        if output.drawings.len() + output.vml_images.len() >= MAX_DRAWING_LEAVES {
            break;
        }
    }
    Ok(())
}

fn project_drawing(
    drawing: &XmlElement,
    relationships: Option<&RelationshipMap>,
    media: &crate::media::MediaMap,
    charts: &crate::chart::ChartPartsMap,
    smart_art: &mut SmartArtContext,
    budget: &mut ParseBudget<'_>,
) -> Result<Option<DrawingLeaf>, ParseError> {
    let (kind, value) = if let Some(text_box) = parse_text_box(drawing) {
        ("textBox", serde_json::to_value(text_box))
    } else if let Some(chart) = parse_chart_from_drawing(drawing, relationships, Some(charts))? {
        ("chart", serde_json::to_value(chart))
    } else if is_smart_art_drawing(drawing) {
        let Some(shape) =
            parse_smart_art_from_drawing(drawing, relationships, Some(smart_art), budget)?
        else {
            return Ok(None);
        };
        ("smartArt", serde_json::to_value(shape))
    } else if is_shape_drawing(drawing) {
        let Some(mut shape) = parse_shape_from_drawing(drawing) else {
            return Ok(None);
        };
        resolve_shape_fill_pictures(&mut shape, relationships, Some(media));
        ("shape", serde_json::to_value(shape))
    } else {
        let Some(image) = parse_drawing(drawing, relationships, Some(media)) else {
            return Ok(None);
        };
        ("image", serde_json::to_value(image))
    };
    Ok(Some(DrawingLeaf {
        element: "drawing".to_owned(),
        kind: kind.to_owned(),
        value: value.map_err(|error| ParseError::Canonical(error.to_string()))?,
    }))
}

pub fn s4_wire_envelope(projection: S4Projection) -> Result<S4WireEnvelope, ParseError> {
    let canonical =
        from_serializable(&projection).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let bytes =
        to_canonical_bytes(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let sha =
        canonical_sha256(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    Ok(S4WireEnvelope {
        wire_version: 1,
        projection,
        canonical_base64: base64::engine::general_purpose::STANDARD.encode(bytes),
        canonical_sha256: sha,
    })
}

pub fn parse_docx_s4_wire(data: &[u8]) -> Result<S4WireEnvelope, ParseError> {
    s4_wire_envelope(parse_docx_s4_projection(data)?)
}

fn relationship_path_for_part(path: &str) -> String {
    match path.rsplit_once('/') {
        Some((directory, filename)) => format!("{directory}/_rels/{filename}.rels"),
        None => format!("_rels/{path}.rels"),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn relationship_owner_paths_and_empty_envelope_are_stable() {
        assert_eq!(
            relationship_path_for_part("word/header1.xml"),
            "word/_rels/header1.xml.rels"
        );
        assert_eq!(
            relationship_path_for_part("word/diagrams/drawing1.xml"),
            "word/diagrams/_rels/drawing1.xml.rels"
        );
        let envelope = s4_wire_envelope(S4Projection {
            media_entries: Vec::new(),
            chart_entries: Vec::new(),
            xml_parts: Vec::new(),
            smart_art_warnings: Vec::new(),
        })
        .unwrap();
        assert_eq!(envelope.wire_version, 1);
        assert_eq!(envelope.canonical_sha256.len(), 64);
    }
}
