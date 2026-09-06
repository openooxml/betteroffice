use serde::{Deserialize, Serialize};

use crate::relationships::{Relationship, relationship_types};
use crate::xml::{ParseBudget, XmlElement, serialize_xml};
use crate::{PptxError, xml::parse_xml};

pub(crate) const NS_P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const NS_P188: &str = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
pub(crate) const NS_PC: &str = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
pub(crate) const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

pub(crate) const CT_COMMENT_AUTHORS: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
pub(crate) const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
pub(crate) const CT_MODERN_AUTHORS: &str = "application/vnd.ms-powerpoint.authors+xml";
pub(crate) const CT_MODERN_COMMENTS: &str = "application/vnd.ms-powerpoint.comments+xml";

pub(crate) const LEGACY_AUTHORS_PART: &str = "ppt/commentAuthors.xml";
pub(crate) const MODERN_AUTHORS_PART: &str = "ppt/authors.xml";

pub(crate) const EMU_PER_MASTER_UNIT: f64 = 1587.5;

pub(crate) fn emu_to_master(emu: i64) -> i64 {
    (emu as f64 / EMU_PER_MASTER_UNIT).round() as i64
}

pub(crate) fn master_to_emu(master: i64) -> i64 {
    (master as f64 * EMU_PER_MASTER_UNIT).round() as i64
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum CommentFlavor {
    #[default]
    Legacy,
    Modern,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CommentAuthor {
    pub id: String,
    pub name: String,
    #[serde(default)]
    pub initials: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_index: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_index: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub user_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub provider_id: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Comment {
    pub id: String,
    pub author_id: String,
    pub slide_part_path: String,
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created: Option<String>,
    pub x_emu: i64,
    pub y_emu: i64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub status: Option<String>,
    pub flavor: CommentFlavor,
}

pub(crate) struct SlideCommentParts {
    pub legacy: Option<String>,
    pub modern: Option<String>,
}

pub(crate) fn slide_comment_parts(relationships: &[Relationship]) -> SlideCommentParts {
    SlideCommentParts {
        legacy: resolved_by_exact_type(relationships, relationship_types::COMMENTS),
        modern: resolved_by_exact_type(relationships, relationship_types::MODERN_COMMENTS),
    }
}

pub(crate) fn authors_part(
    relationships: &[Relationship],
    flavor: CommentFlavor,
) -> Option<String> {
    let relationship_type = match flavor {
        CommentFlavor::Legacy => relationship_types::COMMENT_AUTHORS,
        CommentFlavor::Modern => relationship_types::MODERN_AUTHORS,
    };
    resolved_by_exact_type(relationships, relationship_type)
}

fn resolved_by_exact_type(
    relationships: &[Relationship],
    relationship_type: &str,
) -> Option<String> {
    relationships
        .iter()
        .find(|relationship| relationship.is_type(relationship_type))
        .and_then(|relationship| relationship.resolved_target.clone())
}

pub(crate) fn parse_comment_authors(
    bytes: &[u8],
    part: &str,
    flavor: CommentFlavor,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<CommentAuthor>, PptxError> {
    let root = parse_xml(bytes, part, budget)?;
    let (list, element) = match flavor {
        CommentFlavor::Legacy => ("cmAuthorLst", "cmAuthor"),
        CommentFlavor::Modern => ("authorLst", "author"),
    };
    if root.local_name() != list {
        return Ok(Vec::new());
    }
    let mut authors = Vec::new();
    for child in root.children_named(element) {
        budget.charge_comment(part)?;
        authors.push(CommentAuthor {
            id: child.attribute_local("id").unwrap_or_default().to_owned(),
            name: child.attribute_local("name").unwrap_or_default().to_owned(),
            initials: child
                .attribute_local("initials")
                .unwrap_or_default()
                .to_owned(),
            last_index: unsigned_attribute(child, "lastIdx"),
            color_index: unsigned_attribute(child, "clrIdx"),
            user_id: child.attribute_local("userId").map(str::to_owned),
            provider_id: child.attribute_local("providerId").map(str::to_owned),
        });
    }
    Ok(authors)
}

pub(crate) fn parse_comments(
    bytes: &[u8],
    part: &str,
    slide_part_path: &str,
    flavor: CommentFlavor,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<Comment>, PptxError> {
    let root = parse_xml(bytes, part, budget)?;
    if root.local_name() != "cmLst" {
        return Ok(Vec::new());
    }
    let mut comments = Vec::new();
    for child in root.children_named("cm") {
        budget.charge_comment(part)?;
        match flavor {
            CommentFlavor::Legacy => comments.push(legacy_comment(child, slide_part_path)),
            CommentFlavor::Modern => {
                let comment = modern_comment(child, slide_part_path, None);
                let parent_id = comment.id.clone();
                comments.push(comment);
                for reply in child
                    .child("replyLst")
                    .into_iter()
                    .flat_map(|list| list.children_named("reply"))
                {
                    budget.charge_comment(part)?;
                    comments.push(modern_comment(
                        reply,
                        slide_part_path,
                        Some(parent_id.clone()),
                    ));
                }
            }
        }
    }
    Ok(comments)
}

fn legacy_comment(element: &XmlElement, slide_part_path: &str) -> Comment {
    let author_id = element
        .attribute_local("authorId")
        .unwrap_or("0")
        .to_owned();
    let index = element.attribute_local("idx").unwrap_or("1").to_owned();
    let position = element.child("pos");
    Comment {
        id: format!("{author_id}:{index}"),
        author_id,
        slide_part_path: slide_part_path.to_owned(),
        text: element
            .child("text")
            .map(XmlElement::text_content)
            .unwrap_or_default(),
        created: element.attribute_local("dt").map(str::to_owned),
        x_emu: master_to_emu(coordinate(position, "x")),
        y_emu: master_to_emu(coordinate(position, "y")),
        parent_id: None,
        status: None,
        flavor: CommentFlavor::Legacy,
    }
}

fn modern_comment(
    element: &XmlElement,
    slide_part_path: &str,
    parent_id: Option<String>,
) -> Comment {
    let position = element.child("pos");
    Comment {
        id: element.attribute_local("id").unwrap_or_default().to_owned(),
        author_id: element
            .attribute_local("authorId")
            .unwrap_or_default()
            .to_owned(),
        slide_part_path: slide_part_path.to_owned(),
        text: element
            .child("txBody")
            .map(comment_text)
            .unwrap_or_default(),
        created: element.attribute_local("created").map(str::to_owned),
        x_emu: coordinate(position, "x"),
        y_emu: coordinate(position, "y"),
        parent_id,
        status: element.attribute_local("status").map(str::to_owned),
        flavor: CommentFlavor::Modern,
    }
}

fn comment_text(body: &XmlElement) -> String {
    body.children_named("p")
        .map(|paragraph| {
            paragraph
                .child_elements()
                .filter_map(|child| match child.local_name() {
                    "r" | "fld" => child.child("t").map(XmlElement::text_content),
                    "br" => Some("\n".to_owned()),
                    _ => None,
                })
                .collect::<String>()
        })
        .collect::<Vec<_>>()
        .join("\n")
}

fn coordinate(position: Option<&XmlElement>, name: &str) -> i64 {
    position
        .and_then(|element| element.attribute_local(name))
        .and_then(|value| value.parse::<i64>().ok())
        .unwrap_or(0)
}

fn unsigned_attribute(element: &XmlElement, name: &str) -> Option<u32> {
    element.attribute_local(name)?.parse::<u32>().ok()
}

pub struct CommentsWrite {
    pub flavor: CommentFlavor,
    pub authors: Vec<CommentAuthorWrite>,
    pub per_slide: Vec<(CommentSlide, Vec<CommentWrite>)>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum CommentSlide {
    Existing(String),
    Added(usize),
}

pub struct CommentAuthorWrite {
    pub id: String,
    pub name: String,
    pub initials: String,
    pub last_index: u32,
    pub color_index: u32,
    pub user_id: Option<String>,
    pub provider_id: Option<String>,
}

pub struct CommentWrite {
    pub id: String,
    pub author_id: String,
    pub index: u32,
    pub text: String,
    pub created: Option<String>,
    pub x_emu: i64,
    pub y_emu: i64,
    pub status: Option<String>,
    pub replies: Vec<CommentWrite>,
}

impl CommentsWrite {
    pub(crate) fn authors_part_path(&self) -> &'static str {
        match self.flavor {
            CommentFlavor::Legacy => LEGACY_AUTHORS_PART,
            CommentFlavor::Modern => MODERN_AUTHORS_PART,
        }
    }

    pub(crate) fn authors_content_type(&self) -> &'static str {
        match self.flavor {
            CommentFlavor::Legacy => CT_COMMENT_AUTHORS,
            CommentFlavor::Modern => CT_MODERN_AUTHORS,
        }
    }

    pub(crate) fn authors_relationship_type(&self) -> &'static str {
        match self.flavor {
            CommentFlavor::Legacy => relationship_types::COMMENT_AUTHORS,
            CommentFlavor::Modern => relationship_types::MODERN_AUTHORS,
        }
    }

    pub(crate) fn comments_content_type(&self) -> &'static str {
        match self.flavor {
            CommentFlavor::Legacy => CT_COMMENTS,
            CommentFlavor::Modern => CT_MODERN_COMMENTS,
        }
    }

    pub(crate) fn comments_relationship_type(&self) -> &'static str {
        match self.flavor {
            CommentFlavor::Legacy => relationship_types::COMMENTS,
            CommentFlavor::Modern => relationship_types::MODERN_COMMENTS,
        }
    }

    pub(crate) fn comments_part_path(&self, slide_number: usize) -> String {
        match self.flavor {
            CommentFlavor::Legacy => format!("ppt/comments/comment{slide_number}.xml"),
            CommentFlavor::Modern => format!("ppt/comments/modernComment{slide_number}.xml"),
        }
    }
}

pub(crate) fn authors_xml(write: &CommentsWrite) -> Vec<u8> {
    let root = match write.flavor {
        CommentFlavor::Legacy => {
            let mut root = XmlElement::new("p:cmAuthorLst").with_attribute("xmlns:p", NS_P);
            for author in &write.authors {
                root = root.with_child(
                    XmlElement::new("p:cmAuthor")
                        .with_attribute("id", author.id.clone())
                        .with_attribute("name", author.name.clone())
                        .with_attribute("initials", author.initials.clone())
                        .with_attribute("lastIdx", author.last_index.to_string())
                        .with_attribute("clrIdx", author.color_index.to_string()),
                );
            }
            root
        }
        CommentFlavor::Modern => {
            let mut root = XmlElement::new("p188:authorLst").with_attribute("xmlns:p188", NS_P188);
            for author in &write.authors {
                root = root.with_child(
                    XmlElement::new("p188:author")
                        .with_attribute("id", author.id.clone())
                        .with_attribute("name", author.name.clone())
                        .with_attribute("initials", author.initials.clone())
                        .with_attribute("userId", author.user_id.clone().unwrap_or_default())
                        .with_attribute(
                            "providerId",
                            author.provider_id.clone().unwrap_or_default(),
                        ),
                );
            }
            root
        }
    };
    serialize_xml(&root)
}

pub(crate) fn comments_xml(
    write: &CommentsWrite,
    comments: &[CommentWrite],
    slide_id: u32,
) -> Vec<u8> {
    let root = match write.flavor {
        CommentFlavor::Legacy => {
            let mut root = XmlElement::new("p:cmLst").with_attribute("xmlns:p", NS_P);
            for comment in comments {
                root = root.with_child(legacy_comment_element(comment));
            }
            root
        }
        CommentFlavor::Modern => {
            let mut root = XmlElement::new("p188:cmLst")
                .with_attribute("xmlns:a", NS_A)
                .with_attribute("xmlns:pc", NS_PC)
                .with_attribute("xmlns:p188", NS_P188);
            for comment in comments {
                root = root.with_child(modern_comment_element(comment, slide_id));
            }
            root
        }
    };
    serialize_xml(&root)
}

pub(crate) fn legacy_comment_element(comment: &CommentWrite) -> XmlElement {
    let mut element = XmlElement::new("p:cm")
        .with_attribute("authorId", comment.author_id.clone())
        .with_attribute("idx", comment.index.to_string());
    if let Some(created) = &comment.created {
        element.set_attribute("dt", created.clone());
    }
    element
        .with_child(
            XmlElement::new("p:pos")
                .with_attribute("x", emu_to_master(comment.x_emu).to_string())
                .with_attribute("y", emu_to_master(comment.y_emu).to_string()),
        )
        .with_child(XmlElement::new("p:text").with_text(comment.text.clone()))
}

pub(crate) fn modern_comment_element(comment: &CommentWrite, slide_id: u32) -> XmlElement {
    let mut element = modern_body_element("p188:cm", comment)
        .with_child(
            XmlElement::new("pc:sldMkLst")
                .with_child(XmlElement::new("pc:docMk"))
                .with_child(
                    XmlElement::new("pc:sldMk").with_attribute("sldId", slide_id.to_string()),
                ),
        )
        .with_child(
            XmlElement::new("p188:pos")
                .with_attribute("x", comment.x_emu.to_string())
                .with_attribute("y", comment.y_emu.to_string()),
        );
    if !comment.replies.is_empty() {
        let mut replies = XmlElement::new("p188:replyLst");
        for reply in &comment.replies {
            replies = replies.with_child(
                modern_body_element("p188:reply", reply).with_child(text_body(&reply.text)),
            );
        }
        element = element.with_child(replies);
    }
    element.with_child(text_body(&comment.text))
}

pub(crate) fn modern_body_element(name: &str, comment: &CommentWrite) -> XmlElement {
    let mut element = XmlElement::new(name)
        .with_attribute("id", comment.id.clone())
        .with_attribute("authorId", comment.author_id.clone());
    if let Some(created) = &comment.created {
        element.set_attribute("created", created.clone());
    }
    if let Some(status) = &comment.status {
        element.set_attribute("status", status.clone());
    }
    element
}

pub(crate) fn text_body(text: &str) -> XmlElement {
    let mut body = XmlElement::new("p188:txBody")
        .with_child(XmlElement::new("a:bodyPr"))
        .with_child(XmlElement::new("a:lstStyle"));
    for paragraph in text.split('\n') {
        body = body.with_child(
            XmlElement::new("a:p").with_child(
                XmlElement::new("a:r")
                    .with_child(XmlElement::new("a:t").with_text(paragraph.to_owned())),
            ),
        );
    }
    body
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::relationships::TargetMode;
    use crate::xml::ParseLimits;

    fn relationship(relationship_type: &str, target: &str) -> Relationship {
        Relationship {
            id: "rId9".to_owned(),
            relationship_type: relationship_type.to_owned(),
            target: target.to_owned(),
            target_mode: TargetMode::Internal,
            resolved_target: Some(target.to_owned()),
        }
    }

    #[test]
    fn the_modern_comment_relationship_is_not_read_as_a_legacy_one() {
        let modern = [relationship(
            relationship_types::MODERN_COMMENTS,
            "ppt/comments/modernComment1.xml",
        )];
        let parts = slide_comment_parts(&modern);
        assert_eq!(
            parts.modern.as_deref(),
            Some("ppt/comments/modernComment1.xml")
        );
        assert!(
            parts.legacy.is_none(),
            "the modern type also ends with /comments, so suffix matching would conflate them"
        );

        let legacy = [relationship(
            relationship_types::COMMENTS,
            "ppt/comments/comment1.xml",
        )];
        let parts = slide_comment_parts(&legacy);
        assert_eq!(parts.legacy.as_deref(), Some("ppt/comments/comment1.xml"));
        assert!(parts.modern.is_none());
    }

    #[test]
    fn master_units_round_trip_through_emu() {
        assert_eq!(emu_to_master(914_400), 576, "one inch");
        assert_eq!(master_to_emu(576), 914_400);
        assert_eq!(emu_to_master(0), 0);
        assert_eq!(emu_to_master(2_381), 1);
        assert_eq!(emu_to_master(-914_400), -576);
    }

    #[test]
    fn legacy_comments_and_authors_parse() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let authors = parse_comment_authors(
            br#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Mary Smith" initials="mas" lastIdx="3" clrIdx="0"/></p:cmAuthorLst>"#,
            "ppt/commentAuthors.xml",
            CommentFlavor::Legacy,
            &mut budget,
        )
        .unwrap();
        assert_eq!(authors.len(), 1);
        assert_eq!(authors[0].name, "Mary Smith");
        assert_eq!(authors[0].last_index, Some(3));
        assert_eq!(authors[0].color_index, Some(0));

        let comments = parse_comments(
            br#"<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cm authorId="0" dt="2005-11-13T17:00:22.071" idx="1"><p:pos x="576" y="288"/><p:text>Hello &amp; welcome</p:text></p:cm></p:cmLst>"#,
            "ppt/comments/comment1.xml",
            "ppt/slides/slide1.xml",
            CommentFlavor::Legacy,
            &mut budget,
        )
        .unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].id, "0:1");
        assert_eq!(comments[0].text, "Hello & welcome");
        assert_eq!(comments[0].x_emu, 914_400);
        assert_eq!(comments[0].y_emu, 457_200);
        assert_eq!(comments[0].flavor, CommentFlavor::Legacy);
    }

    #[test]
    fn modern_replies_carry_their_thread_root() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let comments = parse_comments(
            br#"<p188:cmLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:cm id="{ROOT}" authorId="{A}" created="2024-12-30T20:26:06.503" status="resolved"><p188:replyLst><p188:reply id="{REPLY}" authorId="{A}"><p188:txBody><a:p><a:r><a:t>Second</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst><p188:txBody><a:p><a:r><a:t>First</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#,
            "ppt/comments/modernComment1.xml",
            "ppt/slides/slide1.xml",
            CommentFlavor::Modern,
            &mut budget,
        )
        .unwrap();
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[0].id, "{ROOT}");
        assert_eq!(comments[0].text, "First");
        assert_eq!(comments[0].status.as_deref(), Some("resolved"));
        assert_eq!(comments[1].id, "{REPLY}");
        assert_eq!(comments[1].text, "Second");
        assert_eq!(comments[1].parent_id.as_deref(), Some("{ROOT}"));
    }

    #[test]
    fn a_comment_body_is_xml_escaped_on_the_way_out() {
        let write = CommentsWrite {
            flavor: CommentFlavor::Legacy,
            authors: vec![CommentAuthorWrite {
                id: "0".to_owned(),
                name: "A & B".to_owned(),
                initials: "AB".to_owned(),
                last_index: 1,
                color_index: 0,
                user_id: None,
                provider_id: None,
            }],
            per_slide: Vec::new(),
        };
        let comment = CommentWrite {
            id: String::new(),
            author_id: "0".to_owned(),
            index: 1,
            text: "<script> & \"quotes\"".to_owned(),
            created: None,
            x_emu: 0,
            y_emu: 0,
            status: None,
            replies: Vec::new(),
        };
        let xml = String::from_utf8(comments_xml(&write, &[comment], 256)).unwrap();
        assert!(xml.contains("&lt;script&gt; &amp; \"quotes\""));
        assert!(!xml.contains("<script>"));
        let authors = String::from_utf8(authors_xml(&write)).unwrap();
        assert!(authors.contains("name=\"A &amp; B\""));
    }
}
