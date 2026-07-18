//! Complete S5 runs/inline package projection used by the differential corpus gate.

use base64::Engine as _;
use serde::{Deserialize, Serialize};

use crate::canonical::{canonical_sha256, from_serializable, to_canonical_bytes};
use crate::inline::{
    BookmarkEnd, BookmarkStart, Hyperlink, InlineNode, ParsedFieldInstruction, RunProjection,
    SdtProperties, SimpleField, parse_bookmark_end, parse_bookmark_start, parse_field_instruction,
    parse_hyperlink, parse_inline_container, parse_run, parse_sdt_properties, parse_simple_field,
};
use crate::relationships::{RelationshipMap, parse_relationships};
use crate::settings::{incumbent_utf8_text_boundary, parse_settings};
use crate::styles::{DocDefaults, StyleMap, parse_style_definitions};
use crate::theme::{apply_theme_font_lang, parse_theme};
use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, parse_xml};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S5Projection {
    pub xml_parts: Vec<S5XmlPart>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S5XmlPart {
    pub path: String,
    pub runs: Vec<RunProjection>,
    pub simple_fields: Vec<SimpleField>,
    pub hyperlinks: Vec<Hyperlink>,
    pub bookmark_starts: Vec<BookmarkStart>,
    pub bookmark_ends: Vec<BookmarkEnd>,
    pub sdt_properties: Vec<SdtProperties>,
    pub paragraph_inlines: Vec<Vec<InlineNode>>,
    pub field_instructions: Vec<ParsedFieldInstruction>,
}

impl S5XmlPart {
    fn empty(path: String) -> Self {
        Self {
            path,
            runs: Vec::new(),
            simple_fields: Vec::new(),
            hyperlinks: Vec::new(),
            bookmark_starts: Vec::new(),
            bookmark_ends: Vec::new(),
            sdt_properties: Vec::new(),
            paragraph_inlines: Vec::new(),
            field_instructions: Vec::new(),
        }
    }

    fn is_empty(&self) -> bool {
        self.runs.is_empty()
            && self.simple_fields.is_empty()
            && self.hyperlinks.is_empty()
            && self.bookmark_starts.is_empty()
            && self.bookmark_ends.is_empty()
            && self.sdt_properties.is_empty()
            && self.paragraph_inlines.is_empty()
            && self.field_instructions.is_empty()
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S5WireEnvelope {
    pub wire_version: u8,
    pub projection: S5Projection,
    pub canonical_base64: String,
    pub canonical_sha256: String,
}

pub fn parse_docx_s5_projection(data: &[u8]) -> Result<S5Projection, ParseError> {
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

    let mut xml_parts = Vec::new();
    for (path, bytes) in &parts {
        if !path.to_ascii_lowercase().ends_with(".xml") || !incumbent_utf8_text_boundary(bytes) {
            continue;
        }
        let document = parse_xml(bytes, path, &mut budget)?;
        let Some(root) = document.root() else {
            continue;
        };
        if !is_story_root(root.local_name()) {
            continue;
        }
        let relationship_path = relationship_part_path(path);
        let relationships = find_part(&parts, &relationship_path)
            .map(|xml| parse_relationships(xml, &relationship_path, &mut budget))
            .transpose()?
            .unwrap_or_default();
        if let Some(projection) = project_xml_part(
            root,
            path,
            &relationships,
            Some(&theme),
            Some(&styles),
            doc_defaults,
            &mut budget,
        )? {
            xml_parts.push(projection);
        }
    }
    Ok(S5Projection { xml_parts })
}

#[allow(clippy::too_many_arguments)]
pub fn project_xml_part(
    root: &XmlElement,
    path: &str,
    relationships: &RelationshipMap,
    theme: Option<&crate::theme::Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    budget: &mut ParseBudget<'_>,
) -> Result<Option<S5XmlPart>, ParseError> {
    let mut projection = S5XmlPart::empty(path.to_owned());
    let mut stack = vec![root];
    while let Some(element) = stack.pop() {
        match element.local_name() {
            "r" => {
                budget.charge_leaf_value(path)?;
                projection
                    .runs
                    .push(parse_run(element, theme, styles, doc_defaults));
            }
            "fldSimple" => {
                budget.charge_leaf_value(path)?;
                let field = parse_simple_field(element, theme, styles, doc_defaults, path, budget)?;
                projection
                    .field_instructions
                    .push(parse_field_instruction(&field.instruction));
                projection.simple_fields.push(field);
            }
            "hyperlink" => {
                budget.charge_leaf_value(path)?;
                projection.hyperlinks.push(parse_hyperlink(
                    element,
                    Some(relationships),
                    theme,
                    styles,
                    doc_defaults,
                    path,
                    budget,
                )?);
            }
            "bookmarkStart" => {
                budget.charge_leaf_value(path)?;
                projection
                    .bookmark_starts
                    .push(parse_bookmark_start(element));
            }
            "bookmarkEnd" => {
                budget.charge_leaf_value(path)?;
                projection.bookmark_ends.push(parse_bookmark_end(element));
            }
            "sdt" => {
                budget.charge_leaf_value(path)?;
                projection.sdt_properties.push(parse_sdt_properties(
                    element.child("w", "sdtPr"),
                    element.child("w", "sdtEndPr"),
                    theme,
                ));
            }
            "p" => {
                budget.charge_leaf_value(path)?;
                let inlines = parse_inline_container(
                    element,
                    Some(relationships),
                    theme,
                    styles,
                    doc_defaults,
                    path,
                    budget,
                    0,
                )?;
                collect_complex_field_instructions(&inlines, &mut projection.field_instructions);
                projection.paragraph_inlines.push(inlines);
            }
            _ => {}
        }
        if is_drawing_seam(element.local_name()) {
            continue;
        }
        let children: Vec<_> = element.child_elements().collect();
        stack.extend(children.into_iter().rev());
    }
    Ok((!projection.is_empty()).then_some(projection))
}

fn collect_complex_field_instructions(
    nodes: &[InlineNode],
    output: &mut Vec<ParsedFieldInstruction>,
) {
    for node in nodes {
        match node {
            InlineNode::ComplexField(field) => {
                output.push(parse_field_instruction(&field.instruction));
                if let Some(tree) = field.field_tree.as_ref() {
                    collect_tree_instructions(tree, output);
                }
            }
            InlineNode::InlineSdt(sdt) => collect_complex_field_instructions(&sdt.content, output),
            _ => {}
        }
    }
}

fn collect_tree_instructions(
    tree: &crate::inline::StructuredFieldTree,
    output: &mut Vec<ParsedFieldInstruction>,
) {
    // Child trees intentionally carry no duplicate instruction string. The
    // nested ComplexField node in code/result is the authoritative source.
    for content in [tree.code.as_ref(), tree.result.as_ref()]
        .into_iter()
        .flatten()
    {
        if let Some(nodes) = content.inline.as_deref() {
            collect_complex_field_instructions(nodes, output);
        }
    }
}

pub fn s5_wire_envelope(projection: S5Projection) -> Result<S5WireEnvelope, ParseError> {
    let canonical =
        from_serializable(&projection).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let bytes =
        to_canonical_bytes(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let sha =
        canonical_sha256(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    Ok(S5WireEnvelope {
        wire_version: 1,
        projection,
        canonical_base64: base64::engine::general_purpose::STANDARD.encode(bytes),
        canonical_sha256: sha,
    })
}

pub fn parse_docx_s5_wire(data: &[u8]) -> Result<S5WireEnvelope, ParseError> {
    s5_wire_envelope(parse_docx_s5_projection(data)?)
}

fn is_story_root(local_name: &str) -> bool {
    matches!(
        local_name,
        "document" | "hdr" | "ftr" | "footnotes" | "endnotes" | "comments" | "glossaryDocument"
    )
}

fn is_drawing_seam(local_name: &str) -> bool {
    matches!(
        local_name,
        "drawing" | "pict" | "object" | "AlternateContent"
    )
}

fn relationship_part_path(path: &str) -> String {
    match path.rsplit_once('/') {
        Some((directory, name)) => format!("{directory}/_rels/{name}.rels"),
        None => format!("_rels/{path}.rels"),
    }
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
    use crate::xml::parse_xml;

    #[test]
    fn projects_complete_inline_elements_in_document_order() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(
            br##"<w:document><w:body><w:p><w:bookmarkStart w:id="1" w:name="b"/><w:r><w:t>A</w:t><w:drawing/></w:r><w:hyperlink w:anchor="b"><w:r><w:t>B</w:t></w:r></w:hyperlink><w:fldSimple w:instr="PAGE \\* MERGEFORMAT"><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DDE "cmd" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="1"/></w:p><w:sdt><w:sdtPr><w:tag w:val="x"/></w:sdtPr></w:sdt></w:body></w:document>"##,
            "word/document.xml",
            &mut budget,
        )
        .unwrap();
        let projection = project_xml_part(
            document.root().unwrap(),
            "word/document.xml",
            &RelationshipMap::new(),
            None,
            None,
            None,
            &mut budget,
        )
        .unwrap()
        .unwrap();
        assert_eq!(projection.runs.len(), 6);
        assert_eq!(projection.hyperlinks.len(), 1);
        assert_eq!(projection.simple_fields.len(), 1);
        assert_eq!(projection.bookmark_starts.len(), 1);
        assert_eq!(projection.bookmark_ends.len(), 1);
        assert_eq!(projection.sdt_properties.len(), 1);
        assert_eq!(projection.paragraph_inlines.len(), 1);
        assert_eq!(projection.field_instructions.len(), 2);
        assert!(matches!(
            projection.runs[0].run.content[1],
            crate::inline::RunContent::OpaqueDrawing { .. }
        ));
        assert_eq!(
            projection.bookmark_starts[0]
                .position
                .as_ref()
                .and_then(|position| position.offset),
            None
        );
        let InlineNode::BookmarkEnd(end) = projection.paragraph_inlines[0].last().unwrap() else {
            panic!("bookmark end")
        };
        assert_eq!(end.position.as_ref().unwrap().offset, Some(3.0));
    }

    #[test]
    fn relationship_paths_cover_body_and_story_parts() {
        assert_eq!(
            relationship_part_path("word/document.xml"),
            "word/_rels/document.xml.rels"
        );
        assert_eq!(
            relationship_part_path("word/header1.xml"),
            "word/_rels/header1.xml.rels"
        );
    }

    #[test]
    fn wire_uses_the_shared_canonical_contract() {
        let wire = s5_wire_envelope(S5Projection { xml_parts: vec![] }).unwrap();
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(wire.canonical_base64)
            .unwrap();
        assert_eq!(
            String::from_utf8(bytes).unwrap(),
            "docx-document-canonical-v1\n{\"xmlParts\":[]}\n"
        );
        assert_eq!(wire.canonical_sha256.len(), 64);
    }

    #[test]
    fn shared_leaf_budget_caps_run_projection_growth() {
        let limits = ParseLimits {
            max_leaf_values: 2,
            ..ParseLimits::default()
        };
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(
            b"<w:document><w:body><w:p><w:r/><w:r/></w:p></w:body></w:document>",
            "word/document.xml",
            &mut budget,
        )
        .unwrap();
        let error = project_xml_part(
            document.root().unwrap(),
            "word/document.xml",
            &RelationshipMap::new(),
            None,
            None,
            None,
            &mut budget,
        )
        .unwrap_err();
        assert!(matches!(
            error,
            ParseError::ResourceLimit {
                kind: "leafValues",
                ..
            }
        ));
    }
}
