use std::collections::{BTreeMap, HashMap, HashSet};

use crate::drawing::{common_slide_data, parse_text_styles};
use crate::model::*;
use crate::relationships::{Relationship, parse_relationships, relationship_types};
use crate::theme::parse_theme;
use crate::xml::{ParseBudget, XmlElement, parse_xml};
use crate::{ParseLimits, PptxError};

pub fn parse_pptx(data: &[u8]) -> Result<PptxPackage, PptxError> {
    parse_pptx_with_limits(data, &ParseLimits::default())
}

pub fn parse_pptx_with_limits(data: &[u8], limits: &ParseLimits) -> Result<PptxPackage, PptxError> {
    let source_parts = ooxml_opc::unzip_parts(data).map_err(PptxError::Container)?;
    let parts: HashMap<&str, &[u8]> = source_parts
        .iter()
        .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
        .collect();
    let mut budget = ParseBudget::new(limits);
    let relationships = parse_package_relationships(&source_parts, &mut budget)?;
    let presentation_path = relationships
        .get("")
        .and_then(|entries| {
            entries
                .iter()
                .find(|relationship| relationship.has_type("/officeDocument"))
        })
        .and_then(|relationship| relationship.resolved_target.clone())
        .unwrap_or_else(|| "ppt/presentation.xml".to_owned());
    let presentation_root = parse_part(&parts, &presentation_path, &mut budget)?;
    let presentation_relationships = relationships
        .get(&presentation_path)
        .map(Vec::as_slice)
        .unwrap_or_default();
    let presentation = parse_presentation(
        &presentation_root,
        &presentation_path,
        presentation_relationships,
    )?;

    let mut slides = Vec::with_capacity(presentation.slides.len());
    for reference in &presentation.slides {
        let root = parse_part(&parts, &reference.part_path, &mut budget)?;
        let slide_relationships = relationships
            .get(&reference.part_path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        let data = common_slide_data(
            &root,
            slide_relationships,
            &reference.part_path,
            &mut budget,
        )?;
        slides.push(Slide {
            part_path: reference.part_path.clone(),
            name: data.name,
            layout_part_path: relationship_by_type(
                slide_relationships,
                relationship_types::SLIDE_LAYOUT,
            ),
            show_master_shapes: bool_attribute(&root, "showMasterSp", true),
            background: data.background,
            shapes: data.shapes,
        });
    }

    let master_paths = ordered_part_paths(
        presentation.master_part_paths.clone(),
        &source_parts,
        "ppt/slideMasters/",
    );
    let mut masters = Vec::with_capacity(master_paths.len());
    for part_path in &master_paths {
        let root = parse_part(&parts, part_path, &mut budget)?;
        let master_relationships = relationships
            .get(part_path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        let data = common_slide_data(&root, master_relationships, part_path, &mut budget)?;
        masters.push(SlideMaster {
            part_path: part_path.clone(),
            name: data.name,
            theme_part_path: relationship_by_type(master_relationships, relationship_types::THEME),
            layout_part_paths: master_relationships
                .iter()
                .filter(|relationship| relationship.has_type(relationship_types::SLIDE_LAYOUT))
                .filter_map(|relationship| relationship.resolved_target.clone())
                .collect(),
            background: data.background,
            shapes: data.shapes,
            text_styles: parse_text_styles(&root),
        });
    }

    let layout_paths = ordered_part_paths(
        masters
            .iter()
            .flat_map(|master| master.layout_part_paths.iter().cloned())
            .collect(),
        &source_parts,
        "ppt/slideLayouts/",
    );
    let mut layouts = Vec::with_capacity(layout_paths.len());
    for part_path in &layout_paths {
        let root = parse_part(&parts, part_path, &mut budget)?;
        let layout_relationships = relationships
            .get(part_path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        let data = common_slide_data(&root, layout_relationships, part_path, &mut budget)?;
        layouts.push(SlideLayout {
            part_path: part_path.clone(),
            name: root
                .attribute("matchingName")
                .map(str::to_owned)
                .or(data.name),
            layout_type: root.attribute("type").map(str::to_owned),
            master_part_path: relationship_by_type(
                layout_relationships,
                relationship_types::SLIDE_MASTER,
            ),
            show_master_shapes: bool_attribute(&root, "showMasterSp", true),
            background: data.background,
            shapes: data.shapes,
        });
    }

    let theme_paths = ordered_part_paths(
        masters
            .iter()
            .filter_map(|master| master.theme_part_path.clone())
            .collect(),
        &source_parts,
        "ppt/theme/",
    );
    let mut themes = Vec::with_capacity(theme_paths.len());
    for part_path in theme_paths {
        let root = parse_part(&parts, &part_path, &mut budget)?;
        themes.push(ThemePart {
            part_path,
            theme: parse_theme(&root),
        });
    }

    let content_types = parse_content_types(&parts, &mut budget)?;
    let media = source_parts
        .iter()
        .filter(|(path, _)| path.starts_with("ppt/media/"))
        .map(|(path, bytes)| MediaPart {
            part_path: path.clone(),
            content_type: content_type_for(path, &content_types),
            bytes: bytes.clone(),
        })
        .collect();
    let parts = source_parts
        .into_iter()
        .map(|(path, bytes)| PackagePart { path, bytes })
        .collect();
    Ok(PptxPackage {
        presentation,
        slides,
        layouts,
        masters,
        themes,
        media,
        relationships,
        parts,
    })
}

pub fn write_pptx(package: &PptxPackage) -> Result<Vec<u8>, PptxError> {
    let parts = package
        .parts
        .iter()
        .map(|part| (part.path.clone(), part.bytes.clone()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).map_err(PptxError::Container)
}

fn parse_package_relationships(
    parts: &[(String, Vec<u8>)],
    budget: &mut ParseBudget<'_>,
) -> Result<BTreeMap<String, Vec<Relationship>>, PptxError> {
    let mut output = BTreeMap::new();
    for (path, bytes) in parts {
        let Some(source) = relationship_source(path) else {
            continue;
        };
        let parsed = parse_relationships(bytes, path, &source, budget)?;
        output.insert(source, parsed);
    }
    Ok(output)
}

fn relationship_source(path: &str) -> Option<String> {
    if path == "_rels/.rels" {
        return Some(String::new());
    }
    let (directory, filename) = path.rsplit_once("/_rels/")?;
    let source_filename = filename.strip_suffix(".rels")?;
    Some(format!("{directory}/{source_filename}"))
}

fn parse_part(
    parts: &HashMap<&str, &[u8]>,
    path: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlElement, PptxError> {
    let bytes = parts
        .get(path)
        .ok_or_else(|| PptxError::MissingPart(path.to_owned()))?;
    parse_xml(bytes, path, budget)
}

fn parse_presentation(
    root: &XmlElement,
    part_path: &str,
    relationships: &[Relationship],
) -> Result<Presentation, PptxError> {
    let slide_size = root.child("sldSz");
    let width_emu = positive_integer_attribute(slide_size, "cx").unwrap_or(12_192_000);
    let height_emu = positive_integer_attribute(slide_size, "cy").unwrap_or(6_858_000);
    let mut slides = Vec::new();
    if let Some(list) = root.child("sldIdLst") {
        for slide in list.children_named("sldId") {
            let relationship_id = slide
                .attribute("r:id")
                .or_else(|| slide.attribute_local("id"))
                .unwrap_or_default()
                .to_owned();
            let part_path =
                relationship_target(relationships, &relationship_id).ok_or_else(|| {
                    PptxError::MissingPart(format!("slide relationship {relationship_id}"))
                })?;
            slides.push(SlideReference {
                id: slide
                    .attribute("id")
                    .and_then(|value| value.parse().ok())
                    .unwrap_or_default(),
                relationship_id,
                part_path,
            });
        }
    }
    let master_part_paths = root
        .child("sldMasterIdLst")
        .into_iter()
        .flat_map(|list| list.children_named("sldMasterId"))
        .filter_map(|master| {
            master
                .attribute("r:id")
                .or_else(|| master.attribute_local("id"))
                .and_then(|id| relationship_target(relationships, id))
        })
        .collect();
    Ok(Presentation {
        part_path: part_path.to_owned(),
        width_emu,
        height_emu,
        slides,
        master_part_paths,
    })
}

#[derive(Default)]
struct ContentTypes {
    defaults: HashMap<String, String>,
    overrides: HashMap<String, String>,
}

fn parse_content_types(
    parts: &HashMap<&str, &[u8]>,
    budget: &mut ParseBudget<'_>,
) -> Result<ContentTypes, PptxError> {
    let root = parse_part(parts, "[Content_Types].xml", budget)?;
    let mut content_types = ContentTypes::default();
    for child in root.child_elements() {
        match child.local_name() {
            "Default" => {
                if let (Some(extension), Some(content_type)) =
                    (child.attribute("Extension"), child.attribute("ContentType"))
                {
                    content_types
                        .defaults
                        .insert(extension.to_ascii_lowercase(), content_type.to_owned());
                }
            }
            "Override" => {
                if let (Some(part_name), Some(content_type)) =
                    (child.attribute("PartName"), child.attribute("ContentType"))
                {
                    content_types.overrides.insert(
                        part_name.trim_start_matches('/').to_owned(),
                        content_type.to_owned(),
                    );
                }
            }
            _ => {}
        }
    }
    Ok(content_types)
}

fn content_type_for(path: &str, content_types: &ContentTypes) -> String {
    content_types
        .overrides
        .get(path)
        .cloned()
        .or_else(|| {
            path.rsplit_once('.').and_then(|(_, extension)| {
                content_types
                    .defaults
                    .get(&extension.to_ascii_lowercase())
                    .cloned()
            })
        })
        .unwrap_or_else(|| "application/octet-stream".to_owned())
}

fn ordered_part_paths(
    preferred: Vec<String>,
    parts: &[(String, Vec<u8>)],
    prefix: &str,
) -> Vec<String> {
    let mut seen = HashSet::new();
    preferred
        .into_iter()
        .chain(
            parts
                .iter()
                .map(|(path, _)| path.clone())
                .filter(|path| path.starts_with(prefix) && path.ends_with(".xml")),
        )
        .filter(|path| seen.insert(path.clone()))
        .collect()
}

fn relationship_by_type(relationships: &[Relationship], suffix: &str) -> Option<String> {
    relationships
        .iter()
        .find(|relationship| relationship.has_type(suffix))
        .and_then(|relationship| relationship.resolved_target.clone())
}

fn relationship_target(relationships: &[Relationship], id: &str) -> Option<String> {
    relationships
        .iter()
        .find(|relationship| relationship.id == id)
        .and_then(|relationship| relationship.resolved_target.clone())
}

fn positive_integer_attribute(element: Option<&XmlElement>, name: &str) -> Option<i64> {
    let value = element?.attribute(name)?.parse::<i64>().ok()?;
    (value > 0 && value <= 1_000_000_000_000_000).then_some(value)
}

fn bool_attribute(element: &XmlElement, name: &str, default: bool) -> bool {
    match element.attribute(name) {
        Some("1" | "true" | "on") => true,
        Some("0" | "false" | "off") => false,
        _ => default,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

    #[test]
    fn parses_betteroffice_demo_deck_surface() {
        let package = parse_pptx(FIXTURE).unwrap();
        assert_eq!(package.presentation.width_emu, 12_192_000);
        assert_eq!(package.presentation.height_emu, 6_858_000);
        assert_eq!(package.slides.len(), 3);
        assert_eq!(package.layouts.len(), 1);
        assert_eq!(package.masters.len(), 1);
        assert_eq!(package.themes.len(), 1);
        assert_eq!(package.media.len(), 1);
        assert_eq!(package.themes[0].theme.name, "BetterOffice");

        let kinds = package
            .slides
            .iter()
            .flat_map(|slide| slide.shapes.iter())
            .map(|shape| match shape {
                ShapeNode::Shape(_) => "shape",
                ShapeNode::Picture(_) => "picture",
                ShapeNode::GraphicFrame(_) => "graphicFrame",
                ShapeNode::Group(_) => "group",
            })
            .collect::<HashSet<_>>();
        assert_eq!(
            kinds,
            HashSet::from(["shape", "picture", "graphicFrame", "group"])
        );
        assert!(package.slides.iter().flat_map(|slide| &slide.shapes).any(|shape| {
            matches!(
                shape,
                ShapeNode::Shape(Shape {
                    text: Some(TextBody { paragraphs, .. }),
                    ..
                }) if paragraphs.iter().flat_map(|paragraph| &paragraph.runs).any(|run| run.text.contains("Rust"))
            )
        }));
    }

    #[test]
    fn untouched_save_preserves_every_part_byte_and_order() {
        let package = parse_pptx(FIXTURE).unwrap();
        let written = write_pptx(&package).unwrap();
        let before = ooxml_opc::unzip_parts(FIXTURE).unwrap();
        let after = ooxml_opc::unzip_parts(&written).unwrap();
        assert_eq!(after, before);
    }

    #[test]
    fn package_limits_apply_across_xml_parts() {
        let limits = ParseLimits {
            max_xml_bytes: 100,
            ..ParseLimits::default()
        };
        assert!(matches!(
            parse_pptx_with_limits(FIXTURE, &limits),
            Err(PptxError::ResourceLimit {
                kind: "xmlBytes",
                ..
            })
        ));
    }
}
