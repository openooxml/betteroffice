//! Header/footer story ownership built on the shared block dispatcher.

use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::block::{BlockContent, StoryParser};
use crate::chart::ChartPartsMap;
use crate::media::MediaMap;
use crate::numbering::NumberingMap;
use crate::paragraph::HexIdAllocator;
use crate::relationships::{RelationshipMap, parse_relationships, relationship_types};
use crate::smart_art::SmartArtContext;
use crate::styles::{DocDefaults, StyleMap};
use crate::theme::Theme;
use crate::vml::{Watermark, extract_watermark};
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooter {
    #[serde(rename = "type")]
    pub story_type: String,
    pub hdr_ftr_type: String,
    pub content: Vec<BlockContent>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub watermark: Option<Watermark>,
}

/// Parse one `w:hdr` or `w:ftr` root through the same dispatcher used by the
/// document body, tables, SDTs, notes, comments, and text boxes.
pub fn parse_header_footer(
    root: &XmlElement,
    is_header: bool,
    hdr_ftr_type: &str,
    parser: &mut StoryParser<'_, '_>,
) -> Result<HeaderFooter, ParseError> {
    let expected = if is_header { "hdr" } else { "ftr" };
    let content = if root.local_name() == expected {
        parser.parse_blocks(root, 0, true)?
    } else {
        Vec::new()
    };
    // VML watermarks are extracted only from headers. Footer VML remains
    // ordinary story content (where supported) and never becomes the
    // dedicated package watermark field.
    let watermark = is_header
        .then(|| extract_watermark(Some(root), parser.relationships, Some(parser.media)))
        .flatten();
    Ok(HeaderFooter {
        story_type: if is_header { "header" } else { "footer" }.to_owned(),
        hdr_ftr_type: normalize_header_footer_type(Some(hdr_ftr_type)).to_owned(),
        content,
        watermark,
    })
}

pub fn normalize_header_footer_type(value: Option<&str>) -> &'static str {
    match value {
        Some("first") => "first",
        Some("even") => "even",
        _ => "default",
    }
}

/// Resolve document relationship ids to package parts and parse every owned
/// header/footer in document-relationship insertion order. Each story uses its
/// own relationship part when present; the incumbent document relationships
/// are the compatibility fallback when that part is absent.
#[allow(clippy::too_many_arguments)]
pub fn parse_related_header_footers(
    parts: &[(String, Vec<u8>)],
    document_relationships: &RelationshipMap,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    numbering: Option<&NumberingMap>,
    media: &MediaMap,
    charts: &ChartPartsMap,
    smart_art: &mut SmartArtContext,
    budget: &mut ParseBudget<'_>,
    ids: &mut HexIdAllocator,
) -> Result<
    (
        IndexMap<String, HeaderFooter>,
        IndexMap<String, HeaderFooter>,
    ),
    ParseError,
> {
    let mut headers = IndexMap::new();
    let mut footers = IndexMap::new();
    for (relationship_id, relationship) in document_relationships {
        let is_header = relationship.relationship_type == relationship_types::HEADER;
        let is_footer = relationship.relationship_type == relationship_types::FOOTER;
        if !is_header && !is_footer {
            continue;
        }
        let Some(filename) = relationship
            .target
            .rsplit(|character| character == '/' || character == '\\')
            .next()
        else {
            continue;
        };
        if filename.is_empty() {
            continue;
        }
        let expected_path = format!("word/{filename}");
        let Some((part_path, xml)) = find_part_case_insensitive(parts, &expected_path) else {
            // External and missing targets stay inert; no resolver or fetch is
            // available anywhere in this crate.
            continue;
        };
        let relationship_part_path = format!("word/_rels/{filename}.rels");
        let part_relationships = find_part_case_insensitive(parts, &relationship_part_path)
            .map(|(path, xml)| parse_relationships(xml, path, budget))
            .transpose()?;
        let relationships = part_relationships
            .as_ref()
            .unwrap_or(document_relationships);
        let document = parse_xml(xml, part_path, budget)?;
        let Some(root) = document.root() else {
            continue;
        };
        let mut parser = StoryParser {
            relationships: Some(relationships),
            theme,
            styles,
            doc_defaults,
            numbering,
            media,
            charts,
            smart_art: &mut *smart_art,
            budget,
            ids,
            part: part_path,
        };
        let story = parse_header_footer(root, is_header, "default", &mut parser)?;
        if is_header {
            headers.insert(relationship_id.clone(), story);
        } else {
            footers.insert(relationship_id.clone(), story);
        }
    }
    Ok((headers, footers))
}

fn find_part_case_insensitive<'a>(
    parts: &'a [(String, Vec<u8>)],
    expected: &str,
) -> Option<(&'a str, &'a [u8])> {
    parts
        .iter()
        .find(|(path, _)| path.eq_ignore_ascii_case(expected))
        .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
}

/// Select the effective type with the incumbent precedence: an enabled first
/// page wins over an enabled even page, then the default story is used.
pub fn select_for_page<'a>(
    stories: &'a IndexMap<String, HeaderFooter>,
    page_number: usize,
    is_first_page: bool,
    different_first_page: bool,
    different_odd_even: bool,
) -> Option<&'a HeaderFooter> {
    if is_first_page
        && different_first_page
        && let Some(story) = by_type(stories, "first")
    {
        return Some(story);
    }
    if different_odd_even
        && page_number.is_multiple_of(2)
        && let Some(story) = by_type(stories, "even")
    {
        return Some(story);
    }
    by_type(stories, "default")
}

fn by_type<'a>(
    stories: &'a IndexMap<String, HeaderFooter>,
    story_type: &str,
) -> Option<&'a HeaderFooter> {
    stories
        .values()
        .find(|story| story.hdr_ftr_type == story_type)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::ChartPartsMap;
    use crate::media::MediaMap;
    use crate::paragraph::HexIdAllocator;
    use crate::relationships::{Relationship, TargetMode};
    use crate::smart_art::SmartArtContext;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn parse(xml: &str, is_header: bool) -> HeaderFooter {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(xml.as_bytes(), "word/header1.xml", &mut budget).unwrap();
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
            part: "word/header1.xml",
        };
        parse_header_footer(document.root().unwrap(), is_header, "future", &mut parser).unwrap()
    }

    #[test]
    fn owner_uses_shared_dispatcher_for_full_block_grammar() {
        let header = parse(
            r#"<w:hdr xmlns:w="w"><w:p><w:r><w:t>before</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt></w:hdr>"#,
            true,
        );
        assert_eq!(header.story_type, "header");
        assert_eq!(header.hdr_ftr_type, "default");
        assert_eq!(
            header
                .content
                .iter()
                .map(BlockContent::node_type)
                .collect::<Vec<_>>(),
            ["paragraph", "table", "blockSdt"]
        );
    }

    #[test]
    fn first_page_type_has_precedence_over_even_then_default() {
        let mut stories = IndexMap::new();
        for kind in ["default", "even", "first"] {
            stories.insert(
                kind.to_owned(),
                HeaderFooter {
                    story_type: "header".to_owned(),
                    hdr_ftr_type: kind.to_owned(),
                    content: Vec::new(),
                    watermark: None,
                },
            );
        }
        assert_eq!(
            select_for_page(&stories, 2, true, true, true)
                .unwrap()
                .hdr_ftr_type,
            "first"
        );
        assert_eq!(
            select_for_page(&stories, 2, false, true, true)
                .unwrap()
                .hdr_ftr_type,
            "even"
        );
        assert_eq!(
            select_for_page(&stories, 3, false, true, true)
                .unwrap()
                .hdr_ftr_type,
            "default"
        );
    }

    #[test]
    fn relationship_id_resolution_uses_the_owning_part_relationships() {
        let parts = vec![
            (
                "WORD/Header1.XML".to_owned(),
                br#"<w:hdr xmlns:w="w" xmlns:r="r"><w:p><w:hyperlink r:id="rLink"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p></w:hdr>"#.to_vec(),
            ),
            (
                "word/_rels/header1.xml.rels".to_owned(),
                br#"<Relationships><Relationship Id="rLink" Type="hyperlink" Target="https://header.example/" TargetMode="External"/></Relationships>"#.to_vec(),
            ),
        ];
        let document_relationships = RelationshipMap::from([
            (
                "rHeader".to_owned(),
                Relationship {
                    id: "rHeader".to_owned(),
                    relationship_type: relationship_types::HEADER.to_owned(),
                    target: "header1.xml".to_owned(),
                    target_mode: None,
                },
            ),
            (
                "rLink".to_owned(),
                Relationship {
                    id: "rLink".to_owned(),
                    relationship_type: relationship_types::HYPERLINK.to_owned(),
                    target: "https://document.example/".to_owned(),
                    target_mode: Some(TargetMode::External),
                },
            ),
        ]);
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let media = MediaMap::new();
        let charts = ChartPartsMap::new();
        let mut smart_art = SmartArtContext::default();
        let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap();
        let (headers, footers) = parse_related_header_footers(
            &parts,
            &document_relationships,
            None,
            None,
            None,
            None,
            &media,
            &charts,
            &mut smart_art,
            &mut budget,
            &mut ids,
        )
        .unwrap();
        assert!(footers.is_empty());
        let BlockContent::Paragraph(paragraph) = &headers["rHeader"].content[0] else {
            panic!("paragraph")
        };
        let crate::paragraph::ParagraphContent::Inline(crate::inline::InlineNode::Hyperlink(link)) =
            &paragraph.content[0]
        else {
            panic!("hyperlink")
        };
        assert_eq!(link.href.as_deref(), Some("https://header.example/"));
    }
}
