//! Body-level document assembly, sections, and read-only paragraph queries.

use serde::{Deserialize, Serialize};

use crate::block::{BlockContent, StoryParser, transparent_children};
use crate::comments::Comment;
use crate::inline::{InlineNode, RunContent};
use crate::paragraph::RawAttribute;
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
    /// Sections with header/footer inheritance resolved forward.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sections: Option<Vec<Section>>,
    /// The body-level `w:sectPr` as authored. Resolution lives in
    /// `sections`; re-emitting a resolved copy would rehome inherited
    /// header and footer references onto the final section on save.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub final_section_properties: Option<SectionProperties>,
    /// Root namespace bindings and attributes retained outside the standard boilerplate.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_root_bindings: Vec<RawAttribute>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comments: Option<Vec<Comment>>,
}

/// Whether `document` holds what the layout may apply to the body's first
/// pages from anywhere in it, so that no cut of the body lays out like the
/// whole: a float placed from outside the text (a drawing anchored to the page
/// or margin, a table or frame not anchored to the text, VML positioned from
/// the page or margin), or a section with columns, whose settings sections
/// without their own may take.
pub(crate) fn refuses_a_body_cut(document: &XmlElement) -> bool {
    let mut pending = vec![document];
    while let Some(element) = pending.pop() {
        let refused = match element.local_name() {
            "anchor" => {
                element.attribute(None, "simplePos") == Some("1")
                    || element
                        .child_by_local_name("positionV")
                        .and_then(|position| position.attribute(None, "relativeFrom"))
                        .is_some_and(|from| !matches!(from, "paragraph" | "line"))
            }
            "tblpPr" => element.attribute(Some("w"), "vertAnchor") != Some("text"),
            "framePr" => element.attribute(Some("w"), "vAnchor") != Some("text"),
            "cols" => {
                element
                    .attribute(Some("w"), "num")
                    .is_some_and(|count| !matches!(count.trim(), "" | "0" | "1"))
                    || element.children_named("w", "col").nth(1).is_some()
            }
            _ => crate::vml::placed_off_the_text(element),
        };
        if refused {
            return true;
        }
        pending.extend(element.child_elements());
    }
    false
}

/// Assemble the body below an already safe-parsed `w:document` root.
pub fn parse_document_body(
    document: &XmlElement,
    parser: &mut StoryParser<'_, '_>,
) -> Result<DocumentBody, ParseError> {
    parse_document_body_impl(document, parser, true, None)
}

/// Parses a body without cloning blocks into section content.
pub(crate) fn parse_document_body_compact(
    document: &XmlElement,
    parser: &mut StoryParser<'_, '_>,
    body_blocks: Option<usize>,
) -> Result<DocumentBody, ParseError> {
    parse_document_body_impl(document, parser, false, body_blocks)
}

fn parse_document_body_impl(
    document: &XmlElement,
    parser: &mut StoryParser<'_, '_>,
    clone_section_content: bool,
    body_blocks: Option<usize>,
) -> Result<DocumentBody, ParseError> {
    if document.local_name() != "document" {
        return Ok(DocumentBody::default());
    }
    let Some(body) = document.child("w", "body") else {
        return Ok(DocumentBody::default());
    };
    let (content, read) = parser.parse_blocks_until(body, 0, false, body_blocks)?;
    // A body cut short ends inside the section whose properties the next
    // section-ending paragraph carries.
    let cut_section = body_blocks.and_then(|_| {
        transparent_children(body, false)
            .into_iter()
            .skip(read)
            .filter(|child| child.matches_name("w", "p"))
            .find_map(|paragraph| paragraph.child("w", "pPr")?.child("w", "sectPr"))
    });
    let final_section_properties = cut_section
        .or_else(|| body.child("w", "sectPr"))
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
    Ok(DocumentBody {
        content,
        sections: Some(sections),
        final_section_properties,
        custom_root_bindings: custom_root_bindings(document),
        comments: None,
    })
}

/// Root bindings and attributes not regenerated by the standard boilerplate.
pub(crate) fn custom_root_bindings(document: &XmlElement) -> Vec<RawAttribute> {
    document
        .attributes
        .iter()
        .filter(|(name, _)| {
            name.strip_prefix("xmlns:")
                .map_or(name.as_str() != "xmlns", |prefix| {
                    !crate::serializer::parts::is_story_root_prefix(prefix)
                })
        })
        .map(|(name, value)| RawAttribute {
            name: name.clone(),
            value: value.clone(),
        })
        .collect()
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
            // A failed match restarts at the second brace.
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
            // Block SDTs are not traversed.
            BlockContent::BlockSdt(_) => {}
            BlockContent::RawXml(_) => {}
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
                    BlockContent::RawXml(_) => {}
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
    fn builds_sections_inherits_story_refs_and_keeps_the_final_section_as_authored() {
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
        assert_eq!(sections[1].properties.title_pg, None);
        assert_eq!(
            sections[1].properties.header_references.as_ref().unwrap()[0].relationship_id,
            "rH"
        );
        // Kept as authored; header inheritance is resolved through `sections`.
        let final_properties = body.final_section_properties.as_ref().unwrap();
        assert_eq!(final_properties.margin_left, Some(720.0));
        assert!(final_properties.header_references.is_none());
        assert!(final_properties.title_pg.is_none());
        assert_eq!(
            extract_all_template_variables(&body.content),
            ["first", "second-name"]
        );
    }

    #[test]
    fn refuses_a_cut_of_a_body_with_floats_off_the_text_or_columns() {
        let refused = |body: &str| {
            let limits = ParseLimits::default();
            let mut budget = ParseBudget::new(&limits);
            let xml = format!(
                r#"<w:document xmlns:w="w" xmlns:wp="wp" xmlns:v="v"><w:body>{body}</w:body></w:document>"#
            );
            let document = parse_xml(xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
            refuses_a_body_cut(document.root().unwrap())
        };
        let anchor = |attributes: &str, vertical: &str| {
            format!(
                "<w:p><w:r><w:drawing><wp:anchor {attributes}>{vertical}</wp:anchor></w:drawing></w:r></w:p>"
            )
        };
        let table = |attributes: &str| {
            format!("<w:tbl><w:tblPr><w:tblpPr {attributes}/></w:tblPr></w:tbl>")
        };
        let frame =
            |attributes: &str| format!("<w:p><w:pPr><w:framePr {attributes}/></w:pPr></w:p>");
        let section =
            |columns: &str| format!("<w:p><w:pPr><w:sectPr>{columns}</w:sectPr></w:pPr></w:p>");
        let vml = |style: &str| {
            format!(r#"<w:p><w:r><w:pict><v:shape style="{style}"/></w:pict></w:r></w:p>"#)
        };
        for body in [
            anchor("", r#"<wp:positionV relativeFrom="margin"/>"#),
            anchor("", r#"<wp:positionV relativeFrom="page"/>"#),
            anchor("", r#"<wp:positionV relativeFrom="topMargin"/>"#),
            anchor(
                r#"simplePos="1""#,
                r#"<wp:positionV relativeFrom="paragraph"/>"#,
            ),
            table(r#"w:vertAnchor="margin""#),
            table(r#"w:tblpY="60""#),
            frame(r#"w:vAnchor="page""#),
            frame(r#"w:w="2000""#),
            vml("position:absolute;mso-position-vertical-relative:margin"),
            vml("position:absolute;mso-position-vertical-relative:page"),
            section(r#"<w:cols w:num="2"/>"#),
            section(r#"<w:cols><w:col w:w="3000"/><w:col w:w="3000"/></w:cols>"#),
        ] {
            assert!(refused(&body), "{body}");
        }
        for body in [
            anchor("", r#"<wp:positionV relativeFrom="paragraph"/>"#),
            anchor("", r#"<wp:positionV relativeFrom="line"/>"#),
            anchor("", ""),
            table(r#"w:vertAnchor="text""#),
            frame(r#"w:dropCap="drop" w:vAnchor="text""#),
            vml("position:absolute;mso-position-vertical-relative:text"),
            vml("width:300pt;height:165pt"),
            section(r#"<w:cols w:num="1" w:space="720"/>"#),
            section(r#"<w:cols w:space="720"/>"#),
        ] {
            assert!(!refused(&body), "{body}");
        }
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
    fn paragraph_text_and_empty_query_use_shallow_run_grammar() {
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
