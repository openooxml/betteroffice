//! Comment story ownership, modern metadata joins, reply threading, and
//! comment-range integrity.

use std::collections::HashSet;

use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::block::{BlockContent, StoryParser};
use crate::paragraph::{Paragraph, ParagraphContent};
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_javascript_integer_prefix, parse_xml};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Comment {
    pub id: f64,
    pub author: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub initials: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    pub content: Vec<Paragraph>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent_id: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub done: Option<bool>,
    #[serde(default)]
    pub status: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub durable_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub para_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date_utc: Option<String>,
    #[serde(default)]
    pub palette_index: f64,
    #[serde(default)]
    pub block_content: Vec<BlockContent>,
}

#[derive(Clone, Debug, Default)]
struct ExtendedComments {
    parent_by_para_id: IndexMap<String, String>,
    done_by_para_id: IndexMap<String, bool>,
}

/// Parse `comments.xml` and join the two modern metadata parts. Comment body
/// blocks use the same story dispatcher as every other owner.
pub fn parse_comments(
    root: &XmlElement,
    extensible_xml: Option<&[u8]>,
    extended_xml: Option<&[u8]>,
    parser: &mut StoryParser<'_, '_>,
) -> Result<Vec<Comment>, ParseError> {
    if root.local_name() != "comments" {
        return Ok(Vec::new());
    }
    let date_utc_by_para_id =
        parse_comments_extensible(extensible_xml, "word/commentsExtensible.xml", parser.budget)?;
    let extended =
        parse_comments_extended(extended_xml, "word/commentsExtended.xml", parser.budget)?;
    let mut comments = Vec::new();
    let mut palette_by_author = IndexMap::<String, usize>::new();
    let mut last_para_ids = Vec::new();

    // Incumbent quirk: the 100k cap applies to all direct children before
    // filtering for w:comment, so leading unknown nodes consume the cap.
    for child in root.child_elements().take(100_000) {
        if child.local_name() != "comment" {
            continue;
        }
        parser.budget.charge_comment(parser.part)?;
        let id = integer_attribute(child, "w", "id").unwrap_or(0.0);
        let author = child
            .attribute(Some("w"), "author")
            .unwrap_or("Unknown")
            .to_owned();
        let initials = child.attribute(Some("w"), "initials").map(str::to_owned);
        let local_date = child.attribute(Some("w"), "date").map(str::to_owned);
        let metadata_para_id = attribute_any(child, &[("w14", "paraId"), ("w", "paraId")]);
        let date_utc = metadata_para_id
            .as_deref()
            .and_then(|value| date_utc_by_para_id.get(&value.to_ascii_uppercase()))
            .cloned();
        let date = date_utc.clone().or(local_date);
        let mut done =
            matches!(child.attribute(Some("w"), "done"), Some("1" | "true")).then_some(true);
        let parent_id = attribute_any(child, &[("w16cid", "parentId"), ("w", "parentId")])
            .as_deref()
            .and_then(parse_javascript_integer_prefix);
        let block_content = parser.parse_blocks(child, 0, false)?;
        let content = block_content
            .iter()
            .filter_map(|block| match block {
                BlockContent::Paragraph(paragraph) => Some(paragraph.clone()),
                _ => None,
            })
            .collect();
        let mut last_para_id = String::new();
        for paragraph in child
            .child_elements()
            .take(100_000)
            .filter(|element| element.local_name() == "p")
        {
            if let Some(para_id) = paragraph.attribute(Some("w14"), "paraId") {
                last_para_id = para_id.to_ascii_uppercase();
            }
        }
        if done.is_none()
            && !last_para_id.is_empty()
            && extended.done_by_para_id.contains_key(&last_para_id)
        {
            done = Some(true);
        }
        let next_palette = palette_by_author.len();
        let palette_index = *palette_by_author
            .entry(author.clone())
            .or_insert(next_palette);
        let durable_id = attribute_any(child, &[("w16cid", "durableId"), ("w16cex", "durableId")])
            .map(|value| truncate_utf16_scalars(&value, 255));
        let author_id = attribute_any(child, &[("w16du", "personId"), ("w16cid", "personId")])
            .map(|value| truncate_utf16_scalars(&value, 255));
        comments.push(Comment {
            id,
            author,
            initials,
            date,
            content,
            parent_id,
            done,
            status: if done == Some(true) {
                "resolved"
            } else {
                "active"
            }
            .to_owned(),
            author_id,
            durable_id,
            para_id: (!last_para_id.is_empty()).then_some(last_para_id.clone()),
            date_utc,
            palette_index: palette_index as f64,
            block_content,
        });
        last_para_ids.push(last_para_id);
    }

    if !extended.parent_by_para_id.is_empty() {
        let mut comment_id_by_para_id = IndexMap::new();
        for (comment, para_id) in comments.iter().zip(&last_para_ids) {
            if !para_id.is_empty() {
                comment_id_by_para_id.insert(para_id.clone(), comment.id);
            }
        }
        for (comment, para_id) in comments.iter_mut().zip(&last_para_ids) {
            if comment.parent_id.is_some() || para_id.is_empty() {
                continue;
            }
            comment.parent_id = extended
                .parent_by_para_id
                .get(para_id)
                .and_then(|parent_para_id| comment_id_by_para_id.get(parent_para_id))
                .map(|value| *value);
        }
    }
    Ok(comments)
}

fn parse_comments_extensible(
    xml: Option<&[u8]>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<IndexMap<String, String>, ParseError> {
    let Some(xml) = xml else {
        return Ok(IndexMap::new());
    };
    let document = parse_xml(xml, part, budget)?;
    let Some(root) = document.root() else {
        return Ok(IndexMap::new());
    };
    let mut dates = IndexMap::new();
    for child in root.child_elements() {
        if child.local_name() != "comment" {
            continue;
        }
        let para_id = attribute_any(child, &[("w16cex", "paraId"), ("w15", "paraId")]);
        let date_utc = attribute_any(child, &[("w16cex", "dateUtc"), ("w15", "dateUtc")]);
        if let (Some(para_id), Some(date_utc)) = (para_id, date_utc) {
            dates.insert(para_id.to_ascii_uppercase(), date_utc);
        }
    }
    Ok(dates)
}

fn parse_comments_extended(
    xml: Option<&[u8]>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<ExtendedComments, ParseError> {
    let Some(xml) = xml else {
        return Ok(ExtendedComments::default());
    };
    let document = parse_xml(xml, part, budget)?;
    let Some(root) = document.root() else {
        return Ok(ExtendedComments::default());
    };
    let mut extended = ExtendedComments::default();
    for child in root.child_elements() {
        if child.local_name() != "commentEx" {
            continue;
        }
        let Some(para_id) = child.attribute(Some("w15"), "paraId") else {
            continue;
        };
        let para_id = para_id.to_ascii_uppercase();
        if let Some(parent) = child.attribute(Some("w15"), "paraIdParent") {
            extended
                .parent_by_para_id
                .insert(para_id.clone(), parent.to_ascii_uppercase());
        }
        if child.attribute(Some("w15"), "done") == Some("1") {
            extended.done_by_para_id.insert(para_id, true);
        }
    }
    Ok(extended)
}

fn attribute_any(element: &XmlElement, names: &[(&str, &str)]) -> Option<String> {
    names
        .iter()
        .find_map(|(namespace, name)| element.attribute(Some(namespace), name))
        .map(str::to_owned)
}

fn integer_attribute(element: &XmlElement, namespace: &str, name: &str) -> Option<f64> {
    element
        .attribute(Some(namespace), name)
        .and_then(parse_javascript_integer_prefix)
}

fn truncate_utf16_scalars(value: &str, max_units: usize) -> String {
    let mut units = 0usize;
    value
        .chars()
        .take_while(|character| {
            let next = units + character.len_utf16();
            if next > max_units {
                false
            } else {
                units = next;
                true
            }
        })
        .collect()
}

/// Remove body-story comment anchors whose id does not resolve to a parsed
/// comment. This deliberately follows only the incumbent block recursion
/// (paragraphs, table cells, block SDTs), not shape/textbox JSON payloads.
pub fn remove_orphan_comment_ranges(blocks: &mut [BlockContent], comment_ids: &[f64]) {
    let ids: HashSet<u64> = comment_ids.iter().map(|id| number_key(*id)).collect();
    prune_blocks(blocks, &ids);
}

fn prune_blocks(blocks: &mut [BlockContent], comment_ids: &HashSet<u64>) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => paragraph.content.retain(|content| {
                let ParagraphContent::CommentRange(marker) = content else {
                    return true;
                };
                comment_ids.contains(&number_key(marker.id))
            }),
            BlockContent::Table(table) => {
                for row in &mut table.rows {
                    for cell in &mut row.cells {
                        prune_blocks(&mut cell.content, comment_ids);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => prune_blocks(&mut sdt.content, comment_ids),
        }
    }
}

fn number_key(value: f64) -> u64 {
    if value == 0.0 { 0 } else { value.to_bits() }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::ChartPartsMap;
    use crate::media::MediaMap;
    use crate::paragraph::HexIdAllocator;
    use crate::smart_art::SmartArtContext;
    use crate::xml::{ParseLimits, parse_xml};

    fn parse(comments: &str, extensible: Option<&str>, extended: Option<&str>) -> Vec<Comment> {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(comments.as_bytes(), "word/comments.xml", &mut budget).unwrap();
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
            part: "word/comments.xml",
        };
        parse_comments(
            document.root().unwrap(),
            extensible.map(str::as_bytes),
            extended.map(str::as_bytes),
            &mut parser,
        )
        .unwrap()
    }

    #[test]
    fn joins_utc_done_and_reply_metadata_case_insensitively() {
        let comments = parse(
            r#"<w:comments xmlns:w="w" xmlns:w14="w14"><w:comment w:id="1" w:author="Ada" w14:paraId="ABCD"><w:p w14:paraId="AAAA"><w:r><w:t>parent</w:t></w:r></w:p></w:comment><w:comment w:id="2" w:author="Ada"><w:p w14:paraId="BBBB"/><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:comment></w:comments>"#,
            Some(
                r#"<w16cex:commentsExtensible xmlns:w16cex="x"><w16cex:comment w16cex:paraId="abcd" w16cex:dateUtc="2024-01-01T00:00:00Z"/></w16cex:commentsExtensible>"#,
            ),
            Some(
                r#"<w15:commentsEx xmlns:w15="x"><w15:commentEx w15:paraId="bbbb" w15:paraIdParent="aaaa" w15:done="1"/></w15:commentsEx>"#,
            ),
        );
        assert_eq!(comments[0].date.as_deref(), Some("2024-01-01T00:00:00Z"));
        assert_eq!(comments[0].palette_index, 0.0);
        assert_eq!(comments[1].palette_index, 0.0);
        assert_eq!(comments[1].parent_id, Some(1.0));
        assert_eq!(comments[1].done, Some(true));
        assert_eq!(comments[1].status, "resolved");
        assert_eq!(comments[1].content.len(), 1);
        assert_eq!(comments[1].block_content.len(), 2);
    }

    #[test]
    fn direct_parent_id_wins_over_extended_threading() {
        let comments = parse(
            r#"<w:comments xmlns:w="w" xmlns:w14="w14"><w:comment w:id="1"><w:p w14:paraId="AAAA"/></w:comment><w:comment w:id="2" w:parentId="99"><w:p w14:paraId="BBBB"/></w:comment></w:comments>"#,
            None,
            Some(
                r#"<w15:commentsEx xmlns:w15="x"><w15:commentEx w15:paraId="BBBB" w15:paraIdParent="AAAA"/></w15:commentsEx>"#,
            ),
        );
        assert_eq!(comments[1].parent_id, Some(99.0));
    }

    #[test]
    fn integrity_prunes_orphans_recursively_but_keeps_known_ranges() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(
            br#"<w:body xmlns:w="w"><w:tbl><w:tr><w:tc><w:sdt><w:sdtContent><w:p><w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="9"/><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl></w:body>"#,
            "word/document.xml",
            &mut budget,
        )
        .unwrap();
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
        let mut blocks = parser
            .parse_blocks(document.root().unwrap(), 0, false)
            .unwrap();
        remove_orphan_comment_ranges(&mut blocks, &[1.0]);
        let BlockContent::Table(table) = &blocks[0] else {
            panic!("table")
        };
        let BlockContent::BlockSdt(sdt) = &table.rows[0].cells[0].content[0] else {
            panic!("sdt")
        };
        let BlockContent::Paragraph(paragraph) = &sdt.content[0] else {
            panic!("paragraph")
        };
        assert_eq!(
            paragraph
                .content
                .iter()
                .map(ParagraphContent::node_type)
                .collect::<Vec<_>>(),
            ["commentRangeStart", "run"]
        );
    }

    #[test]
    fn comment_budget_rejects_limit_plus_one() {
        let mut limits = ParseLimits::default();
        limits.max_comments = 1;
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(
            br#"<w:comments xmlns:w="w"><w:comment w:id="1"/><w:comment w:id="2"/></w:comments>"#,
            "word/comments.xml",
            &mut budget,
        )
        .unwrap();
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
            part: "word/comments.xml",
        };
        assert_eq!(
            parse_comments(document.root().unwrap(), None, None, &mut parser),
            Err(ParseError::ResourceLimit {
                kind: "comments",
                part: "word/comments.xml".to_owned(),
            })
        );
    }
}
