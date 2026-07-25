//! S6 body projection used by the independent corpus gate.

use base64::Engine as _;
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::block::{BlockContent, StoryParser};
use crate::canonical::{canonical_sha256, from_serializable, to_canonical_bytes};
use crate::chart::parse_chart_parts;
use crate::document::{DocumentBody, extract_all_template_variables, parse_document_body};
use crate::inline::{InlineNode, Run, RunContent, StructuredFieldContent, StructuredFieldTree};
use crate::media::build_media_map;
use crate::numbering::parse_numbering;
use crate::paragraph::{HexIdAllocator, ParagraphContent};
use crate::relationships::parse_relationships;
use crate::settings::parse_settings;
use crate::shape::{DrawingSceneNode, Shape};
use crate::smart_art::create_smart_art_context;
use crate::styles::{StyleMap, parse_style_definitions};
use crate::theme::{apply_theme_font_lang, parse_theme};
use crate::xml::{ParseBudget, ParseError, ParseLimits, parse_xml};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S6Projection {
    pub body: DocumentBody,
    pub template_variables: Vec<String>,
    pub smart_art_warnings: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S6WireEnvelope {
    pub wire_version: u8,
    pub projection: S6Projection,
    pub canonical_base64: String,
    pub canonical_sha256: String,
}

pub fn parse_docx_s6_projection(data: &[u8]) -> Result<S6Projection, ParseError> {
    let mut projection = parse_docx_story_projection(data)?;
    project_s6_body(&mut projection.body);
    projection.template_variables = extract_all_template_variables(&projection.body.content);
    Ok(projection)
}

pub(crate) fn parse_docx_story_projection(data: &[u8]) -> Result<S6Projection, ParseError> {
    let parts = ooxml_opc::unzip_parts(data).map_err(ParseError::Container)?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let settings = parse_settings(
        find_part(&parts, "word/settings.xml"),
        "word/settings.xml",
        &mut budget,
    )?;
    let mut theme = parse_theme(
        find_part(&parts, "word/theme/theme1.xml"),
        "word/theme/theme1.xml",
        &mut budget,
    )?;
    apply_theme_font_lang(&mut theme, settings.theme_font_lang.as_ref());
    let style_definitions = find_part(&parts, "word/styles.xml")
        .map(|xml| parse_style_definitions(xml, Some(&theme), "word/styles.xml", &mut budget))
        .transpose()?;
    let styles: StyleMap = style_definitions
        .as_ref()
        .map(|definitions| {
            definitions
                .styles
                .iter()
                .map(|style| (style.style_id.clone(), style.clone()))
                .collect()
        })
        .unwrap_or_default();
    let doc_defaults = style_definitions
        .as_ref()
        .and_then(|definitions| definitions.doc_defaults.as_ref());
    let numbering = parse_numbering(
        find_part(&parts, "word/numbering.xml"),
        "word/numbering.xml",
        &mut budget,
    )?;
    let relationships = find_part(&parts, "word/_rels/document.xml.rels")
        .map(|xml| parse_relationships(xml, "word/_rels/document.xml.rels", &mut budget))
        .transpose()?
        .unwrap_or_default();
    let media = build_media_map(&parts);
    let all_xml: IndexMap<_, _> = parts
        .iter()
        .filter(|(path, _)| {
            let lower = path.to_ascii_lowercase();
            lower.ends_with(".xml") || lower.ends_with(".rels")
        })
        .cloned()
        .collect();
    let charts = parse_chart_parts(&all_xml, &mut budget)?;
    let mut smart_art = create_smart_art_context(&all_xml);
    let digest = format!("{:x}", Sha256::digest(data));
    let mut ids = HexIdAllocator::from_sha256(&digest)?;
    let body = match find_part(&parts, "word/document.xml") {
        Some(xml) => {
            let document = parse_xml(xml, "word/document.xml", &mut budget)?;
            match document.root() {
                Some(root) => {
                    let mut parser = StoryParser {
                        relationships: Some(&relationships),
                        theme: Some(&theme),
                        styles: Some(&styles),
                        doc_defaults,
                        numbering: Some(&numbering),
                        media: &media,
                        charts: &charts,
                        smart_art: &mut smart_art,
                        budget: &mut budget,
                        ids: &mut ids,
                        part: "word/document.xml",
                    };
                    parse_document_body(root, &mut parser)?
                }
                None => DocumentBody::default(),
            }
        }
        None => DocumentBody::default(),
    };
    let template_variables = extract_all_template_variables(&body.content);
    Ok(S6Projection {
        body,
        template_variables,
        smart_art_warnings: smart_art.warnings,
    })
}

fn project_s6_body(body: &mut DocumentBody) {
    project_s6_blocks(&mut body.content, 0);
    if let Some(sections) = &mut body.sections {
        for section in sections {
            project_s6_blocks(&mut section.content, 0);
        }
    }
}

fn project_s6_blocks(blocks: &mut [BlockContent], depth: usize) {
    if depth > 128 {
        return;
    }
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => {
                for content in &mut paragraph.content {
                    match content {
                        ParagraphContent::Inline(node) => project_s6_inline(node, depth + 1),
                        ParagraphContent::Tracked(change) => {
                            for node in &mut change.content {
                                project_s6_inline(node, depth + 1);
                            }
                        }
                        _ => {}
                    }
                }
            }
            BlockContent::Table(table) => *table = crate::table::Table::empty(),
            BlockContent::BlockSdt(sdt) => project_s6_blocks(&mut sdt.content, depth + 1),
        }
    }
}

fn project_s6_inline(node: &mut InlineNode, depth: usize) {
    if depth > 128 {
        return;
    }
    match node {
        InlineNode::Run(run) => project_s6_run(run, depth + 1),
        InlineNode::Hyperlink(link) => {
            for child in &mut link.children {
                project_s6_inline(child, depth + 1);
            }
            if let Some(children) = &mut link.structured_children {
                for child in children {
                    project_s6_inline(child, depth + 1);
                }
            }
        }
        InlineNode::SimpleField(field) => {
            project_s6_runs(&mut field.content, depth + 1);
            if let Some(content) = &mut field.structured_result {
                project_s6_field_content(content, depth + 1);
            }
            if let Some(tree) = &mut field.field_tree {
                project_s6_field_tree(tree, depth + 1);
            }
        }
        InlineNode::ComplexField(field) => {
            project_s6_runs(&mut field.field_code, depth + 1);
            project_s6_runs(&mut field.field_result, depth + 1);
            if let Some(content) = &mut field.structured_code {
                project_s6_field_content(content, depth + 1);
            }
            if let Some(content) = &mut field.structured_result {
                project_s6_field_content(content, depth + 1);
            }
            if let Some(tree) = &mut field.field_tree {
                project_s6_field_tree(tree, depth + 1);
            }
        }
        InlineNode::InlineSdt(sdt) => {
            for child in &mut sdt.content {
                project_s6_inline(child, depth + 1);
            }
        }
        _ => {}
    }
}

fn project_s6_runs(runs: &mut [Run], depth: usize) {
    for run in runs {
        project_s6_run(run, depth + 1);
    }
}

fn project_s6_run(run: &mut Run, depth: usize) {
    if depth > 128 {
        return;
    }
    for content in &mut run.content {
        if let RunContent::Shape { shape } = content {
            project_s6_shape(shape, depth + 1);
        }
    }
}

fn project_s6_shape(shape: &mut Shape, depth: usize) {
    if depth > 128 {
        return;
    }
    if let Some(text_body) = &mut shape.text_body {
        for value in &mut text_body.content {
            project_table_boundaries(value, depth + 1);
        }
    }
    if let Some(children) = &mut shape.children {
        for child in children {
            project_s6_shape(child, depth + 1);
        }
    }
    if let Some(scene) = &mut shape.scene
        && let Some(root) = &mut scene.root
    {
        project_s6_scene_node(root, depth + 1);
    }
}

fn project_s6_scene_node(node: &mut DrawingSceneNode, depth: usize) {
    if depth > 128 {
        return;
    }
    if let Some(shape) = &mut node.shape {
        project_s6_shape(shape, depth + 1);
    }
    if let Some(children) = &mut node.children {
        for child in children {
            project_s6_scene_node(child, depth + 1);
        }
    }
}

fn project_s6_field_content(content: &mut StructuredFieldContent, depth: usize) {
    if let Some(inline) = &mut content.inline {
        for node in inline {
            project_s6_inline(node, depth + 1);
        }
    }
    if let Some(blocks) = &mut content.blocks {
        project_s6_blocks(blocks, depth + 1);
    }
}

fn project_s6_field_tree(tree: &mut StructuredFieldTree, depth: usize) {
    if depth > 128 {
        return;
    }
    if let Some(code) = &mut tree.code {
        project_s6_field_content(code, depth + 1);
    }
    if let Some(result) = &mut tree.result {
        project_s6_field_content(result, depth + 1);
    }
    if let Some(children) = &mut tree.children {
        for child in children {
            project_s6_field_tree(child, depth + 1);
        }
    }
}

fn project_table_boundaries(value: &mut serde_json::Value, depth: usize) {
    if depth > 128 {
        return;
    }
    match value {
        serde_json::Value::Array(values) => {
            for value in values {
                project_table_boundaries(value, depth.saturating_add(1));
            }
        }
        serde_json::Value::Object(object) => {
            if object.get("type").and_then(serde_json::Value::as_str) == Some("table") {
                object.clear();
                object.insert(
                    "type".to_owned(),
                    serde_json::Value::String("table".to_owned()),
                );
                object.insert("rows".to_owned(), serde_json::Value::Array(Vec::new()));
                return;
            }
            for value in object.values_mut() {
                project_table_boundaries(value, depth.saturating_add(1));
            }
        }
        _ => {}
    }
}

pub fn s6_wire_envelope(projection: S6Projection) -> Result<S6WireEnvelope, ParseError> {
    let canonical =
        from_serializable(&projection).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let bytes =
        to_canonical_bytes(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let sha =
        canonical_sha256(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    Ok(S6WireEnvelope {
        wire_version: 1,
        projection,
        canonical_base64: base64::engine::general_purpose::STANDARD.encode(bytes),
        canonical_sha256: sha,
    })
}

pub fn parse_docx_s6_wire(data: &[u8]) -> Result<S6WireEnvelope, ParseError> {
    s6_wire_envelope(parse_docx_s6_projection(data)?)
}

fn find_part<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> Option<&'a [u8]> {
    parts
        .iter()
        .find(|(candidate, _)| candidate == path)
        .map(|(_, bytes)| bytes.as_slice())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn package_projection_is_deterministic_and_uses_the_shared_wire_contract() {
        let package = ooxml_opc::rezip_parts(&[(
            "word/document.xml".to_owned(),
            br#"<w:document xmlns:w="w"><w:body><w:p w14:paraId="bad" xmlns:w14="w14"><w:r><w:t>{name}</w:t></w:r></w:p><w:tbl/></w:body></w:document>"#.to_vec(),
        )])
        .unwrap();
        let first = parse_docx_s6_wire(&package).unwrap();
        let second = parse_docx_s6_wire(&package).unwrap();
        assert_eq!(first, second);
        assert_eq!(first.projection.template_variables, ["name"]);
        assert!(matches!(
            first.projection.body.content[1],
            crate::block::BlockContent::Table(_)
        ));
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(first.canonical_base64)
            .unwrap();
        assert!(bytes.starts_with(b"docx-document-canonical-v1\n"));
    }
}
