//! Body-level document assembly, sections, and read-only paragraph queries.

use serde::{Deserialize, Serialize};

use crate::block::{BlockContent, StoryParser};
use crate::comments::Comment;
use crate::inline::{InlineNode, RunContent};
use crate::paragraph::{Paragraph, ParagraphContent};
use crate::section::{
    SectionProperties, apply_section_inheritance, default_section_properties,
    parse_section_properties,
};
use crate::xml::{ParseError, XmlElement};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct Section {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    pub properties: SectionProperties,
    pub content: Vec<BlockContent>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentBody {
    pub content: Vec<BlockContent>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sections: Option<Vec<Section>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub final_section_properties: Option<SectionProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comments: Option<Vec<Comment>>,
}

/// Assemble the body below an already safe-parsed `w:document` root.
pub fn parse_document_body(
    document: &XmlElement,
    parser: &mut StoryParser<'_, '_>,
) -> Result<DocumentBody, ParseError> {
    parse_document_body_impl(document, parser, true)
}

/// S9 reconstructs section content as shared host-side slices, so it does not
/// need a second cloned copy of every body block in the transport model.
pub(crate) fn parse_document_body_compact(
    document: &XmlElement,
    parser: &mut StoryParser<'_, '_>,
) -> Result<DocumentBody, ParseError> {
    parse_document_body_impl(document, parser, false)
}

fn parse_document_body_impl(
    document: &XmlElement,
    parser: &mut StoryParser<'_, '_>,
    clone_section_content: bool,
) -> Result<DocumentBody, ParseError> {
    if document.local_name() != "document" {
        return Ok(DocumentBody::default());
    }
    let Some(body) = document.child("w", "body") else {
        return Ok(DocumentBody::default());
    };
    let content = parser.parse_blocks(body, 0, false)?;
    let mut final_section_properties = body
        .child("w", "sectPr")
        .map(|element| parse_section_properties(Some(element)));
    let mut sections = build_sections(
        &content,
        final_section_properties.as_ref(),
        clone_section_content,
    );
    let mut properties: Vec<_> = sections
        .iter()
        .map(|section| section.properties.clone())
        .collect();
    apply_section_inheritance(&mut properties);
    for (section, properties) in sections.iter_mut().zip(properties) {
        section.properties = properties;
    }
    if final_section_properties.is_some()
        && let Some(last) = sections.last()
    {
        final_section_properties = Some(last.properties.clone());
    }
    Ok(DocumentBody {
        content,
        sections: Some(sections),
        final_section_properties,
        comments: None,
    })
}

fn build_sections(
    content: &[BlockContent],
    final_properties: Option<&SectionProperties>,
    clone_content: bool,
) -> Vec<Section> {
    let mut sections = Vec::new();
    let mut current = Vec::new();
    let mut current_len = 0usize;
    for block in content {
        current_len += 1;
        if clone_content {
            current.push(block.clone());
        }
        if let BlockContent::Paragraph(paragraph) = block
            && let Some(properties) = &paragraph.section_properties
        {
            sections.push(Section {
                id: None,
                properties: properties.clone(),
                content: std::mem::take(&mut current),
            });
            current_len = 0;
        }
    }
    if current_len > 0 || sections.is_empty() {
        sections.push(Section {
            id: None,
            properties: final_properties
                .cloned()
                .unwrap_or_else(default_section_properties),
            content: current,
        });
    }
    sections
}

pub fn get_paragraph_text(paragraph: &Paragraph) -> String {
    let mut text = String::new();
    for content in &paragraph.content {
        let ParagraphContent::Inline(content) = content else {
            continue;
        };
        match content {
            InlineNode::Run(run) => append_run_text(&run.content, &mut text, true),
            InlineNode::Hyperlink(link) => {
                for child in &link.children {
                    if let InlineNode::Run(run) = child {
                        append_run_text(&run.content, &mut text, false);
                    }
                }
            }
            InlineNode::SimpleField(field) => {
                for run in &field.content {
                    append_run_text(&run.content, &mut text, false);
                }
            }
            InlineNode::ComplexField(field) => {
                for run in &field.field_result {
                    append_run_text(&run.content, &mut text, false);
                }
            }
            _ => {}
        }
    }
    text
}

fn append_run_text(content: &[RunContent], output: &mut String, include_separators: bool) {
    for content in content {
        match content {
            RunContent::Text { text, .. } => output.push_str(text),
            RunContent::Tab if include_separators => output.push('\t'),
            RunContent::Break { break_type, .. } if include_separators => {
                output.push(if break_type.as_deref() == Some("page") {
                    '\u{000c}'
                } else {
                    '\n'
                });
            }
            _ => {}
        }
    }
}

pub fn is_empty_paragraph(paragraph: &Paragraph) -> bool {
    get_paragraph_text(paragraph).trim().is_empty()
        && !paragraph.content.iter().any(|content| {
            matches!(
                content,
                ParagraphContent::Inline(InlineNode::Run(run))
                    if run.content.iter().any(|content| matches!(
                        content,
                        RunContent::Drawing { .. } | RunContent::Shape { .. }
                    ))
            )
        })
}

pub fn extract_template_variables(text: &str) -> Vec<String> {
    let bytes = text.as_bytes();
    let mut variables = Vec::new();
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        let Some(relative) = bytes[cursor..].iter().position(|byte| *byte == b'{') else {
            break;
        };
        let start = cursor + relative + 1;
        let Some(relative_end) = bytes[start..].iter().position(|byte| *byte == b'}') else {
            break;
        };
        let end = start + relative_end;
        let candidate = &text[start..end];
        let valid = candidate
            .bytes()
            .enumerate()
            .all(|(index, byte)| match byte {
                b'A'..=b'Z' | b'a'..=b'z' | b'_' => true,
                b'0'..=b'9' if index > 0 => true,
                b'-' | b'.' if index > 0 => true,
                _ => false,
            });
        if valid && !candidate.is_empty() && !variables.iter().any(|value| value == candidate) {
            variables.push(candidate.to_owned());
            cursor = end.saturating_add(1);
        } else {
            // The incumbent regex can restart at the second brace in
            // `{{name}}`, producing the inner `{name}` match.
            cursor = start;
        }
    }
    variables
}

pub fn extract_all_template_variables(content: &[BlockContent]) -> Vec<String> {
    let mut variables = Vec::new();
    for block in content {
        match block {
            BlockContent::Paragraph(paragraph) => {
                for variable in extract_template_variables(&get_paragraph_text(paragraph)) {
                    if !variables.contains(&variable) {
                        variables.push(variable);
                    }
                }
            }
            BlockContent::Table(table) => {
                for variable in extract_table_template_variables(table) {
                    if !variables.contains(&variable) {
                        variables.push(variable);
                    }
                }
            }
            // Pinned incumbent boundary: block SDTs are not traversed by the
            // documentParser utility.
            BlockContent::BlockSdt(_) => {}
        }
    }
    variables
}

fn extract_table_template_variables(table: &crate::table::Table) -> Vec<String> {
    let mut variables = Vec::new();
    for row in &table.rows {
        for cell in &row.cells {
            for block in &cell.content {
                match block {
                    BlockContent::Paragraph(paragraph) => {
                        for variable in extract_template_variables(&get_paragraph_text(paragraph)) {
                            if !variables.contains(&variable) {
                                variables.push(variable);
                            }
                        }
                    }
                    BlockContent::Table(table) => {
                        for variable in extract_table_template_variables(table) {
                            if !variables.contains(&variable) {
                                variables.push(variable);
                            }
                        }
                    }
                    BlockContent::BlockSdt(_) => {}
                }
            }
        }
    }
    variables
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::ChartPartsMap;
    use crate::media::MediaMap;
    use crate::paragraph::HexIdAllocator;
    use crate::smart_art::SmartArtContext;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn parse(xml: &str) -> DocumentBody {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
        let media = MediaMap::new();
        let charts = ChartPartsMap::new();
        let mut smart_art = SmartArtContext::default();
        let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap();
        let mut parser = StoryParser {
            relationships: None,
            theme: None,
            styles: None,
            doc_defaults: None,
            numbering: None,
            media: &media,
            charts: &charts,
            smart_art: &mut smart_art,
            budget: &mut budget,
            ids: &mut ids,
            part: "word/document.xml",
        };
        parse_document_body(document.root().unwrap(), &mut parser).unwrap()
    }

    #[test]
    fn builds_sections_inherits_story_refs_and_promotes_the_effective_final_section() {
        let body = parse(
            r#"<w:document xmlns:w="w" xmlns:r="r"><w:body>
              <w:p><w:r><w:t>{first}</w:t></w:r><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rH"/><w:titlePg/></w:sectPr></w:pPr></w:p>
              <w:p><w:r><w:t>{first} {second-name}</w:t></w:r></w:p>
              <w:sectPr><w:pgMar w:left="720"/></w:sectPr>
            </w:body></w:document>"#,
        );
        let sections = body.sections.as_ref().unwrap();
        assert_eq!(sections.len(), 2);
        assert_eq!(sections[1].properties.margin_left, Some(720.0));
        assert_eq!(sections[1].properties.title_pg, Some(true));
        assert_eq!(
            sections[1].properties.header_references.as_ref().unwrap()[0].relationship_id,
            "rH"
        );
        assert_eq!(
            body.final_section_properties
                .as_ref()
                .unwrap()
                .header_references
                .as_ref()
                .unwrap()[0]
                .relationship_id,
            "rH"
        );
        assert_eq!(
            extract_all_template_variables(&body.content),
            ["first", "second-name"]
        );
    }

    #[test]
    fn normalizes_an_empty_body_to_one_default_empty_section() {
        let body = parse(r#"<w:document xmlns:w="w"><w:body/></w:document>"#);
        assert!(body.content.is_empty());
        let sections = body.sections.unwrap();
        assert_eq!(sections.len(), 1);
        assert!(sections[0].content.is_empty());
        assert_eq!(sections[0].properties.page_width, Some(12_240.0));
        assert!(body.final_section_properties.is_none());
    }

    #[test]
    fn paragraph_text_and_empty_query_pin_the_shallow_incumbent_grammar() {
        let body = parse(
            r#"<w:document xmlns:w="w"><w:body><w:p><w:r><w:t> a </w:t><w:tab/><w:br/></w:r></w:p><w:p/></w:body></w:document>"#,
        );
        let BlockContent::Paragraph(first) = &body.content[0] else {
            panic!("paragraph")
        };
        assert_eq!(get_paragraph_text(first), " a \t\n");
        let BlockContent::Paragraph(empty) = &body.content[1] else {
            panic!("paragraph")
        };
        assert!(is_empty_paragraph(empty));
    }
}
