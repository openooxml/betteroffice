//! S12 versioned whole-part serializer wire and parse-back gate.

use serde::{Deserialize, Serialize};

use crate::block::StoryParser;
use crate::chart::ChartPartsMap;
use crate::comments::{Comment, parse_comments};
use crate::document::{DocumentBody, parse_document_body};
use crate::header_footer::{HeaderFooter, parse_header_footer};
use crate::media::MediaMap;
use crate::notes::{Note, parse_notes};
use crate::paragraph::HexIdAllocator;
use crate::smart_art::SmartArtContext;
use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, parse_xml};

use super::context::SerializerContext;
use super::parts::{
    serialize_comments_extended_part, serialize_comments_extensible_part,
    serialize_comments_ids_part, serialize_comments_with_info, serialize_document_part,
    serialize_endnotes_part, serialize_footnotes_part, serialize_header_footer_part,
};
use super::s10::{CanonicalXmlEvent, SerializerDeterminism, canonical_xml_events};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "family", rename_all = "camelCase", deny_unknown_fields)]
pub enum S12SerializeRequest {
    Document {
        determinism: SerializerDeterminism,
        body: DocumentBody,
    },
    HeaderFooter {
        determinism: SerializerDeterminism,
        story: HeaderFooter,
    },
    Footnotes {
        determinism: SerializerDeterminism,
        notes: Vec<Note>,
    },
    Endnotes {
        determinism: SerializerDeterminism,
        notes: Vec<Note>,
    },
    Comments {
        determinism: SerializerDeterminism,
        comments: Vec<Comment>,
    },
    CommentsExtended {
        determinism: SerializerDeterminism,
        comments: Vec<Comment>,
    },
    CommentsIds {
        determinism: SerializerDeterminism,
        comments: Vec<Comment>,
    },
    CommentsExtensible {
        determinism: SerializerDeterminism,
        comments: Vec<Comment>,
    },
}

impl S12SerializeRequest {
    fn determinism(&self) -> &SerializerDeterminism {
        match self {
            Self::Document { determinism, .. }
            | Self::HeaderFooter { determinism, .. }
            | Self::Footnotes { determinism, .. }
            | Self::Endnotes { determinism, .. }
            | Self::Comments { determinism, .. }
            | Self::CommentsExtended { determinism, .. }
            | Self::CommentsIds { determinism, .. }
            | Self::CommentsExtensible { determinism, .. } => determinism,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S12SerializeResponse {
    pub wire_version: u8,
    pub family: String,
    pub xml: String,
    pub canonical_xml_events: Vec<CanonicalXmlEvent>,
    pub parse_back: serde_json::Value,
}

pub fn serialize_s12_wire(
    request: S12SerializeRequest,
) -> Result<S12SerializeResponse, ParseError> {
    let determinism = request.determinism().clone();
    let mut context = SerializerContext::new(&determinism)?;
    let (family, xml) = match request {
        S12SerializeRequest::Document { body, .. } => {
            ("document", serialize_document_part(&body, &mut context)?)
        }
        S12SerializeRequest::HeaderFooter { story, .. } => (
            "headerFooter",
            serialize_header_footer_part(&story, &mut context)?,
        ),
        S12SerializeRequest::Footnotes { notes, .. } => {
            ("footnotes", serialize_footnotes_part(&notes, &mut context)?)
        }
        S12SerializeRequest::Endnotes { notes, .. } => {
            ("endnotes", serialize_endnotes_part(&notes, &mut context)?)
        }
        S12SerializeRequest::Comments { comments, .. } => (
            "comments",
            serialize_comments_with_info(&comments, &mut context).0,
        ),
        S12SerializeRequest::CommentsExtended { comments, .. } => {
            let (_, infos) = serialize_comments_with_info(&comments, &mut context);
            ("commentsExtended", serialize_comments_extended_part(&infos))
        }
        S12SerializeRequest::CommentsIds { comments, .. } => {
            let (_, infos) = serialize_comments_with_info(&comments, &mut context);
            ("commentsIds", serialize_comments_ids_part(&infos))
        }
        S12SerializeRequest::CommentsExtensible { comments, .. } => {
            let (_, infos) = serialize_comments_with_info(&comments, &mut context);
            (
                "commentsExtensible",
                serialize_comments_extensible_part(&infos, &comments),
            )
        }
    };
    Ok(S12SerializeResponse {
        wire_version: 1,
        family: family.to_owned(),
        canonical_xml_events: canonical_xml_events(&xml, false)?,
        parse_back: parse_back(family, &xml, &determinism.seed)?,
        xml,
    })
}

fn parse_back(family: &str, xml: &str, seed: &str) -> Result<serde_json::Value, ParseError> {
    let value = match family {
        "document" => serde_json::to_value(parse_document_back(xml, seed)?),
        "headerFooter" => serde_json::to_value(parse_header_footer_back(xml, seed)?),
        "footnotes" => serde_json::to_value(parse_notes_back(xml, seed, true)?),
        "endnotes" => serde_json::to_value(parse_notes_back(xml, seed, false)?),
        "comments" => serde_json::to_value(parse_comments_back(xml, seed)?),
        "commentsExtended" | "commentsIds" | "commentsExtensible" => {
            Ok(parse_comment_companion_back(family, xml)?)
        }
        _ => unreachable!("all S12 families are matched above"),
    };
    value.map_err(|error| ParseError::Canonical(error.to_string()))
}

fn parse_document_back(xml: &str, seed: &str) -> Result<DocumentBody, ParseError> {
    let document = parse_part(xml, "word/document.xml")?;
    let root = required_root(&document, "document")?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let media = MediaMap::new();
    let charts = ChartPartsMap::new();
    let mut smart_art = SmartArtContext::default();
    let mut ids = HexIdAllocator::from_sha256(seed)?;
    parse_document_body(
        root,
        &mut StoryParser {
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
        },
    )
}

fn parse_header_footer_back(xml: &str, seed: &str) -> Result<HeaderFooter, ParseError> {
    let document = parse_part(xml, "word/header1.xml")?;
    let root = required_root(&document, "header/footer")?;
    let is_header = root.local_name() == "hdr";
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let media = MediaMap::new();
    let charts = ChartPartsMap::new();
    let mut smart_art = SmartArtContext::default();
    let mut ids = HexIdAllocator::from_sha256(seed)?;
    parse_header_footer(
        root,
        is_header,
        "default",
        &mut StoryParser {
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
        },
    )
}

fn parse_notes_back(xml: &str, seed: &str, footnotes: bool) -> Result<Vec<Note>, ParseError> {
    let part = if footnotes {
        "word/footnotes.xml"
    } else {
        "word/endnotes.xml"
    };
    let document = parse_part(xml, part)?;
    let root = required_root(&document, "notes")?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let media = MediaMap::new();
    let charts = ChartPartsMap::new();
    let mut smart_art = SmartArtContext::default();
    let mut ids = HexIdAllocator::from_sha256(seed)?;
    parse_notes(
        root,
        footnotes,
        &mut StoryParser {
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
            part,
        },
    )
}

fn parse_comments_back(xml: &str, seed: &str) -> Result<Vec<Comment>, ParseError> {
    if xml.is_empty() {
        return Ok(Vec::new());
    }
    let document = parse_part(xml, "word/comments.xml")?;
    let root = required_root(&document, "comments")?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let media = MediaMap::new();
    let charts = ChartPartsMap::new();
    let mut smart_art = SmartArtContext::default();
    let mut ids = HexIdAllocator::from_sha256(seed)?;
    parse_comments(
        root,
        None,
        None,
        &mut StoryParser {
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
        },
    )
}

fn parse_comment_companion_back(family: &str, xml: &str) -> Result<serde_json::Value, ParseError> {
    if xml.is_empty() {
        return Ok(serde_json::Value::Array(Vec::new()));
    }
    let document = parse_part(xml, "word/comment-companion.xml")?;
    let root = required_root(&document, "comment companion")?;
    let mut rows = Vec::new();
    for child in root.child_elements() {
        let row = match family {
            "commentsExtended" if child.local_name() == "commentEx" => serde_json::json!({
                "paraId": child.attribute(Some("w15"), "paraId"),
                "done": child.attribute(Some("w15"), "done"),
                "paraIdParent": child.attribute(Some("w15"), "paraIdParent"),
            }),
            "commentsIds" if child.local_name() == "commentId" => serde_json::json!({
                "paraId": child.attribute(Some("w16cid"), "paraId"),
                "durableId": child.attribute(Some("w16cid"), "durableId"),
            }),
            "commentsExtensible" if child.local_name() == "commentExtensible" => {
                serde_json::json!({
                    "durableId": child.attribute(Some("w16cex"), "durableId"),
                    "dateUtc": child.attribute(Some("w16cex"), "dateUtc"),
                })
            }
            _ => continue,
        };
        rows.push(row);
    }
    Ok(serde_json::Value::Array(rows))
}

fn parse_part(xml: &str, part: &str) -> Result<crate::xml::XmlDocument, ParseError> {
    let limits = ParseLimits::default();
    parse_xml(xml.as_bytes(), part, &mut ParseBudget::new(&limits))
}

fn required_root<'a>(
    document: &'a crate::xml::XmlDocument,
    kind: &str,
) -> Result<&'a XmlElement, ParseError> {
    document
        .root()
        .ok_or_else(|| ParseError::Canonical(format!("serialized {kind} part has no root")))
}

#[cfg(test)]
mod tests {
    use crate::inline::{InlineNode, Run, RunContent, RunType};
    use crate::paragraph::{Paragraph, ParagraphContent};

    use super::*;

    fn determinism() -> SerializerDeterminism {
        SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2000-01-01T00:00:00.000Z".to_owned(),
        }
    }

    fn body() -> DocumentBody {
        DocumentBody {
            content: vec![crate::block::BlockContent::Paragraph(Paragraph {
                node_type: "paragraph".to_owned(),
                para_id: None,
                text_id: None,
                formatting: None,
                property_changes: None,
                p_pr_ins: None,
                p_pr_del: None,
                content: vec![ParagraphContent::Inline(InlineNode::Run(Run {
                    node_type: RunType::Run,
                    formatting: None,
                    property_changes: None,
                    content: vec![RunContent::Text {
                        text: "safe <&>".to_owned(),
                        preserve_space: None,
                    }],
                }))],
                list_rendering: None,
                rendered_page_break_before: None,
                section_properties: None,
            })],
            ..DocumentBody::default()
        }
    }

    #[test]
    fn document_wire_is_repeatable_and_parse_backs_as_a_whole_part() {
        let request = S12SerializeRequest::Document {
            determinism: determinism(),
            body: body(),
        };
        let first = serialize_s12_wire(request.clone()).unwrap();
        assert_eq!(first, serialize_s12_wire(request).unwrap());
        assert_eq!(first.family, "document");
        assert_eq!(first.parse_back["content"][0]["type"], "paragraph");
        assert!(first.canonical_xml_events.len() > 6);
    }
}
