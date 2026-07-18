//! Whole WordprocessingML story-part serializers.
//!
//! These functions deliberately retain the incumbent TypeScript byte layout:
//! namespace order, XML-declaration/newline quirks, and the header/footer empty
//! paragraph fallback are observable inputs to the S12 differential gate.

use serde::{Deserialize, Serialize};

use crate::comments::Comment;
use crate::document::DocumentBody;
use crate::header_footer::HeaderFooter;
use crate::inline::{InlineNode, Run, RunContent};
use crate::notes::Note;
use crate::paragraph::{Paragraph, ParagraphContent};
use crate::xml::ParseError;

use super::context::SerializerContext;
use super::raw::validate_raw_subtree;
use super::sdt::serialize_block_content;
use super::section::serialize_section_properties;
use super::watermark::serialize_watermark;
use super::xml_writer::{escape_xml, js_number};

const XML_DECLARATION: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

const DOCUMENT_NAMESPACES: &str = concat!(
    "xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" ",
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ",
    "xmlns:o=\"urn:schemas-microsoft-com:office:office\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
    "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ",
    "xmlns:v=\"urn:schemas-microsoft-com:vml\" ",
    "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" ",
    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ",
    "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" ",
    "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
    "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" ",
    "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" ",
    "xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" ",
    "xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" ",
    "xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" ",
    "xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" ",
    "xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" ",
    "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" ",
    "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" ",
    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\""
);

const HEADER_FOOTER_NAMESPACES: &str = concat!(
    "xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" ",
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ",
    "xmlns:o=\"urn:schemas-microsoft-com:office:office\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
    "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ",
    "xmlns:v=\"urn:schemas-microsoft-com:vml\" ",
    "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" ",
    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ",
    "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" ",
    "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
    "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" ",
    "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" ",
    "xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" ",
    "xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" ",
    "xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" ",
    "xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" ",
    "xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" ",
    "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" ",
    "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" ",
    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\""
);

const FULL_NAMESPACES: &str = concat!(
    "xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" ",
    "xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" ",
    "xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" ",
    "xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" ",
    "xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" ",
    "xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" ",
    "xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" ",
    "xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" ",
    "xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" ",
    "xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" ",
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ",
    "xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" ",
    "xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" ",
    "xmlns:o=\"urn:schemas-microsoft-com:office:office\" ",
    "xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
    "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ",
    "xmlns:v=\"urn:schemas-microsoft-com:vml\" ",
    "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" ",
    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ",
    "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" ",
    "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
    "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" ",
    "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" ",
    "xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" ",
    "xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" ",
    "xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" ",
    "xmlns:w16du=\"http://schemas.microsoft.com/office/word/2023/wordml/word16du\" ",
    "xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" ",
    "xmlns:w16sdtfl=\"http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock\" ",
    "xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" ",
    "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" ",
    "xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" ",
    "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" ",
    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\""
);

const MC_IGNORABLE: &str =
    "mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14\"";

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CommentParaInfo {
    pub comment_id: f64,
    pub last_para_id: String,
    pub durable_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent_id: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub done: Option<bool>,
}

pub fn serialize_document_body(
    body: &DocumentBody,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut output = String::new();
    for block in &body.content {
        output.push_str(&serialize_block_content(block, context)?);
    }
    if let Some(properties) = body.final_section_properties.as_ref() {
        output.push_str(&serialize_section_properties(Some(properties)));
    }
    Ok(output)
}

pub fn serialize_document_part(
    body: &DocumentBody,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let body_xml = serialize_document_body(body, context)?;
    Ok(format!(
        "{XML_DECLARATION}<w:document {DOCUMENT_NAMESPACES} mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\"><w:body>{body_xml}</w:body></w:document>"
    ))
}

pub fn serialize_header_footer_part(
    story: &HeaderFooter,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut content = String::new();
    if let Some(watermark) = story.watermark.as_ref() {
        content.push_str(&serialize_watermark(watermark));
    }
    for block in &story.content {
        content.push_str(&serialize_block_content(block, context)?);
    }
    if content.is_empty() {
        content.push_str("<w:p><w:pPr/></w:p>");
    }
    let root = if story.story_type == "header" {
        "w:hdr"
    } else {
        "w:ftr"
    };
    Ok(format!(
        "{XML_DECLARATION}\n<{root} {HEADER_FOOTER_NAMESPACES}>{content}</{root}>"
    ))
}

pub fn serialize_footnotes_part(
    notes: &[Note],
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    serialize_notes_part(notes, "footnote", "footnotes", context)
}

pub fn serialize_endnotes_part(
    notes: &[Note],
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    serialize_notes_part(notes, "endnote", "endnotes", context)
}

fn serialize_notes_part(
    notes: &[Note],
    element_name: &'static str,
    root_name: &'static str,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut content = String::new();
    for note in notes {
        if let Some(xml) = note.verbatim_xml.as_deref() {
            validate_raw_subtree(xml, "w", element_name)?;
            content.push_str(xml);
            continue;
        }

        content.push_str("<w:");
        content.push_str(element_name);
        if note.note_type != "normal" {
            content.push_str(" w:type=\"");
            content.push_str(&escape_xml(&note.note_type));
            content.push('"');
        }
        content.push_str(" w:id=\"");
        content.push_str(&js_number(note.id));
        content.push_str("\">");
        for block in &note.content {
            content.push_str(&serialize_block_content(block, context)?);
        }
        content.push_str("</w:");
        content.push_str(element_name);
        content.push('>');
    }

    Ok(format!(
        "{XML_DECLARATION}<w:{root_name} {FULL_NAMESPACES} {MC_IGNORABLE}>{content}</w:{root_name}>"
    ))
}

pub fn serialize_comments_part(comments: &[Comment], context: &mut SerializerContext) -> String {
    serialize_comments_with_info(comments, context).0
}

pub fn serialize_comments_with_info(
    comments: &[Comment],
    context: &mut SerializerContext,
) -> (String, Vec<CommentParaInfo>) {
    if comments.is_empty() {
        return (String::new(), Vec::new());
    }

    let mut top_level = Vec::new();
    let mut replies = Vec::new();
    for comment in comments {
        if comment.parent_id.is_none() {
            top_level.push(comment);
        } else {
            replies.push(comment);
        }
    }

    let mut para_infos = Vec::with_capacity(comments.len());
    let mut content = String::new();
    for comment in top_level.into_iter().chain(replies) {
        content.push_str(&serialize_comment(comment, &mut para_infos, context));
    }

    (
        format!(
            "{XML_DECLARATION}<w:comments {FULL_NAMESPACES} {MC_IGNORABLE}>{content}</w:comments>"
        ),
        para_infos,
    )
}

pub fn serialize_comments_extended_part(para_infos: &[CommentParaInfo]) -> String {
    if para_infos.is_empty() {
        return String::new();
    }
    let mut content = String::new();
    for info in para_infos {
        content.push_str("<w15:commentEx w15:paraId=\"");
        content.push_str(&info.last_para_id);
        content.push_str("\" w15:done=\"");
        content.push(if info.done == Some(true) { '1' } else { '0' });
        content.push('"');
        if let Some(parent_id) = info.parent_id
            && let Some(parent) = para_infos
                .iter()
                .rev()
                .find(|candidate| same_number(candidate.comment_id, parent_id))
        {
            content.push_str(" w15:paraIdParent=\"");
            content.push_str(&parent.last_para_id);
            content.push('"');
        }
        content.push_str(" />");
    }
    format!(
        "{XML_DECLARATION}<w15:commentsEx {FULL_NAMESPACES} {MC_IGNORABLE}>{content}</w15:commentsEx>"
    )
}

pub fn serialize_comments_ids_part(para_infos: &[CommentParaInfo]) -> String {
    if para_infos.is_empty() {
        return String::new();
    }
    let mut content = String::new();
    for info in para_infos {
        content.push_str("<w16cid:commentId w16cid:paraId=\"");
        content.push_str(&info.last_para_id);
        content.push_str("\" w16cid:durableId=\"");
        content.push_str(&info.durable_id);
        content.push_str("\" />");
    }
    format!(
        "{XML_DECLARATION}<w16cid:commentsIds {FULL_NAMESPACES} {MC_IGNORABLE}>{content}</w16cid:commentsIds>"
    )
}

pub fn serialize_comments_extensible_part(
    para_infos: &[CommentParaInfo],
    comments: &[Comment],
) -> String {
    if para_infos.is_empty() {
        return String::new();
    }
    let mut content = String::new();
    for info in para_infos {
        let Some(comment) = comments
            .iter()
            .rev()
            .find(|comment| same_number(comment.id, info.comment_id))
        else {
            continue;
        };
        let Some(date) = comment.date.as_deref().filter(|date| !date.is_empty()) else {
            continue;
        };
        let mut date_utc = date.to_owned();
        if !date_utc.ends_with('Z') {
            date_utc.push('Z');
        }
        strip_milliseconds(&mut date_utc);
        content.push_str("<w16cex:commentExtensible w16cex:durableId=\"");
        content.push_str(&info.durable_id);
        content.push_str("\" w16cex:dateUtc=\"");
        content.push_str(&escape_xml(&date_utc));
        content.push_str("\"/>");
    }
    format!(
        "{XML_DECLARATION}<w16cex:commentsExtensible {FULL_NAMESPACES} {MC_IGNORABLE}>{content}</w16cex:commentsExtensible>"
    )
}

fn serialize_comment(
    comment: &Comment,
    para_infos: &mut Vec<CommentParaInfo>,
    context: &mut SerializerContext,
) -> String {
    let comment_para_id = context.allocate_hex_id();
    let mut output = String::new();
    output.push_str("<w:comment w:id=\"");
    output.push_str(&js_number(comment.id));
    output.push('"');
    if !comment.author.is_empty() {
        output.push_str(" w:author=\"");
        output.push_str(&escape_xml(&comment.author));
        output.push('"');
    }
    output.push_str(" w:initials=\"");
    output.push_str(&escape_xml(
        comment
            .initials
            .as_deref()
            .filter(|initials| !initials.is_empty())
            .unwrap_or_default(),
    ));
    output.push('"');
    if let Some(date) = comment.date.as_deref().filter(|date| !date.is_empty()) {
        let mut clean_date = date.to_owned();
        strip_milliseconds(&mut clean_date);
        output.push_str(" w:date=\"");
        output.push_str(&escape_xml(&clean_date));
        output.push('"');
    }
    output.push('>');

    match comment.content.as_slice() {
        [] => {
            output.push_str("<w:p w14:paraId=\"");
            output.push_str(&comment_para_id);
            output.push_str("\"><w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:annotationRef/></w:r></w:p>");
        }
        [paragraph] => {
            output.push_str(&serialize_comment_paragraph(
                paragraph,
                Some(&comment_para_id),
                true,
            ));
        }
        paragraphs => {
            output.push_str(&serialize_comment_paragraph(&paragraphs[0], None, true));
            for paragraph in &paragraphs[1..paragraphs.len() - 1] {
                output.push_str(&serialize_comment_paragraph(paragraph, None, false));
            }
            output.push_str(&serialize_comment_paragraph(
                &paragraphs[paragraphs.len() - 1],
                Some(&comment_para_id),
                false,
            ));
        }
    }
    output.push_str("</w:comment>");
    para_infos.push(CommentParaInfo {
        comment_id: comment.id,
        last_para_id: comment_para_id,
        durable_id: context.allocate_hex_id(),
        parent_id: comment.parent_id,
        done: comment.done,
    });
    output
}

fn serialize_comment_paragraph(
    paragraph: &Paragraph,
    para_id: Option<&str>,
    annotation_ref: bool,
) -> String {
    let mut output = String::from("<w:p");
    if let Some(para_id) = para_id {
        output.push_str(" w14:paraId=\"");
        output.push_str(para_id);
        output.push('"');
    }
    output.push('>');
    if annotation_ref {
        output.push_str(
            "<w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:annotationRef/></w:r>",
        );
    }
    for item in &paragraph.content {
        if let ParagraphContent::Inline(InlineNode::Run(run)) = item {
            output.push_str(&serialize_comment_run(run));
        }
    }
    output.push_str("</w:p>");
    output
}

fn serialize_comment_run(run: &Run) -> String {
    let mut output = String::from("<w:r>");
    let bold = run
        .formatting
        .as_ref()
        .and_then(|formatting| formatting.bold)
        == Some(true);
    let italic = run
        .formatting
        .as_ref()
        .and_then(|formatting| formatting.italic)
        == Some(true);
    if bold || italic {
        output.push_str("<w:rPr>");
        if bold {
            output.push_str("<w:b/>");
        }
        if italic {
            output.push_str("<w:i/>");
        }
        output.push_str("</w:rPr>");
    }
    for content in &run.content {
        match content {
            RunContent::Text { text, .. } => {
                output.push_str("<w:t");
                if comment_text_needs_preserve(text) {
                    output.push_str(" xml:space=\"preserve\"");
                }
                output.push('>');
                output.push_str(&escape_xml(text));
                output.push_str("</w:t>");
            }
            RunContent::Break { .. } => output.push_str("<w:br/>"),
            _ => {}
        }
    }
    output.push_str("</w:r>");
    output
}

fn comment_text_needs_preserve(text: &str) -> bool {
    text.contains("  ")
        || text
            .chars()
            .next()
            .is_some_and(is_ecmascript_trim_character)
        || text
            .chars()
            .next_back()
            .is_some_and(is_ecmascript_trim_character)
}

fn is_ecmascript_trim_character(character: char) -> bool {
    matches!(
        character,
        '\u{0009}'
            | '\u{000A}'
            | '\u{000B}'
            | '\u{000C}'
            | '\u{000D}'
            | '\u{0020}'
            | '\u{00A0}'
            | '\u{1680}'
            | '\u{2000}'
            ..='\u{200A}'
                | '\u{2028}'
                | '\u{2029}'
                | '\u{202F}'
                | '\u{205F}'
                | '\u{3000}'
                | '\u{FEFF}'
    )
}

fn strip_milliseconds(date: &mut String) {
    let bytes = date.as_bytes();
    if bytes.len() >= 5
        && bytes[bytes.len() - 5] == b'.'
        && bytes[bytes.len() - 4..bytes.len() - 1]
            .iter()
            .all(u8::is_ascii_digit)
        && bytes[bytes.len() - 1] == b'Z'
    {
        date.replace_range(date.len() - 5.., "Z");
    }
}

fn same_number(left: f64, right: f64) -> bool {
    left == right || (left.is_nan() && right.is_nan())
}

#[cfg(test)]
mod tests {
    use crate::block::BlockContent;
    use crate::comments::Comment;
    use crate::formatting::TextFormatting;
    use crate::inline::{InlineNode, Run, RunContent, RunType};
    use crate::notes::Note;
    use crate::paragraph::{Paragraph, ParagraphContent};
    use crate::serializer::s10::SerializerDeterminism;

    use super::*;

    fn context() -> SerializerContext {
        SerializerContext::new(&SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2000-01-01T00:00:00.000Z".to_owned(),
        })
        .unwrap()
    }

    fn paragraph(text: &str) -> BlockContent {
        BlockContent::Paragraph(Paragraph {
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
                    text: text.to_owned(),
                    preserve_space: None,
                }],
            }))],
            list_rendering: None,
            rendered_page_break_before: None,
            section_properties: None,
        })
    }

    fn comment(id: f64, parent_id: Option<f64>, text: &str) -> Comment {
        let BlockContent::Paragraph(mut paragraph) = paragraph(text) else {
            unreachable!()
        };
        if let ParagraphContent::Inline(InlineNode::Run(run)) = &mut paragraph.content[0] {
            run.formatting = Some(TextFormatting {
                bold: Some(true),
                italic: Some(true),
                ..TextFormatting::default()
            });
        }
        Comment {
            id,
            author: "Alice <&>".to_owned(),
            initials: None,
            date: Some("2024-01-01T12:30:45.123Z".to_owned()),
            content: vec![paragraph.clone()],
            parent_id,
            done: Some(parent_id.is_none()),
            status: "active".to_owned(),
            author_id: None,
            durable_id: None,
            para_id: None,
            date_utc: None,
            palette_index: 0.0,
            block_content: vec![BlockContent::Paragraph(paragraph)],
        }
    }

    #[test]
    fn document_part_pins_namespace_order_and_has_no_declaration_newline() {
        let xml = serialize_document_part(
            &DocumentBody {
                content: vec![paragraph("safe <&>")],
                ..DocumentBody::default()
            },
            &mut context(),
        )
        .unwrap();
        assert!(xml.starts_with(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:wpc="
        ));
        assert!(xml.contains("<w:t>safe &lt;&amp;&gt;</w:t>"));
        assert!(xml.ends_with("</w:body></w:document>"));
    }

    #[test]
    fn empty_header_has_the_incumbent_newline_and_required_paragraph() {
        let xml = serialize_header_footer_part(
            &HeaderFooter {
                story_type: "header".to_owned(),
                hdr_ftr_type: "default".to_owned(),
                content: Vec::new(),
                watermark: None,
            },
            &mut context(),
        )
        .unwrap();
        assert!(xml.contains("standalone=\"yes\"?>\n<w:hdr xmlns:wpc="));
        assert!(xml.ends_with("<w:p><w:pPr/></w:p></w:hdr>"));
    }

    #[test]
    fn note_parts_pin_attribute_order_and_share_full_block_serialization() {
        let xml = serialize_footnotes_part(
            &[Note {
                story_type: "footnote".to_owned(),
                id: -1.0,
                note_type: "separator".to_owned(),
                content: vec![paragraph("safe <&>")],
                verbatim_xml: None,
            }],
            &mut context(),
        )
        .unwrap();
        assert!(xml.starts_with(concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
            "<w:footnotes xmlns:wpc="
        )));
        assert!(xml.contains(
            "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:t>safe &lt;&amp;&gt;</w:t></w:r></w:p></w:footnote>"
        ));
        assert!(xml.ends_with("</w:footnotes>"));
    }

    #[test]
    fn note_verbatim_gate_accepts_one_expected_tree_and_rejects_opaque_injection() {
        let valid = Note {
            story_type: "endnote".to_owned(),
            id: 4.0,
            note_type: "normal".to_owned(),
            content: Vec::new(),
            verbatim_xml: Some(
                "<w:endnote w:id=\"4\"><w:customXml><w:p/></w:customXml></w:endnote>".to_owned(),
            ),
        };
        let xml = serialize_endnotes_part(&[valid.clone()], &mut context()).unwrap();
        assert!(xml.contains(valid.verbatim_xml.as_deref().unwrap()));

        for hostile in [
            "<w:endnote/><w:endnote/>",
            "<w:footnote/>",
            "<!DOCTYPE x [<!ENTITY leak SYSTEM 'file:///etc/passwd'>]><w:endnote/>",
        ] {
            let mut note = valid.clone();
            note.verbatim_xml = Some(hostile.to_owned());
            assert!(serialize_endnotes_part(&[note], &mut context()).is_err());
        }
    }

    #[test]
    fn comments_and_companions_share_scoped_ids_and_preserve_incumbent_quirks() {
        let comments = vec![
            comment(2.0, Some(1.0), " reply "),
            comment(1.0, None, "top"),
        ];
        let (xml, infos) = serialize_comments_with_info(&comments, &mut context());
        assert!(xml.find("w:id=\"1\"").unwrap() < xml.find("w:id=\"2\"").unwrap());
        assert!(xml.contains("w:author=\"Alice &lt;&amp;&gt;\" w:initials=\"\""));
        assert!(xml.contains("w:date=\"2024-01-01T12:30:45Z\""));
        assert!(xml.contains("<w:rPr><w:b/><w:i/></w:rPr>"));
        assert!(xml.contains("<w:t xml:space=\"preserve\"> reply </w:t>"));
        assert!(!xml.contains("w:done"));
        assert_eq!(infos.len(), 2);
        assert_eq!(infos[0].comment_id, 1.0);
        assert_eq!(infos[1].comment_id, 2.0);
        assert_ne!(infos[0].last_para_id, infos[0].durable_id);

        let extended = serialize_comments_extended_part(&infos);
        assert!(extended.contains(&format!(
            "w15:paraId=\"{}\" w15:done=\"0\" w15:paraIdParent=\"{}\"",
            infos[1].last_para_id, infos[0].last_para_id
        )));
        let ids = serialize_comments_ids_part(&infos);
        let extensible = serialize_comments_extensible_part(&infos, &comments);
        for info in &infos {
            assert!(ids.contains(&info.last_para_id));
            assert!(ids.contains(&info.durable_id));
            assert!(extensible.contains(&info.durable_id));
        }
        assert!(!extensible.contains(".123"));

        let repeated = serialize_comments_with_info(&comments, &mut context());
        assert_eq!((xml, infos), repeated);
    }

    #[test]
    fn part_context_pins_the_clock_and_empty_comment_parts_stay_absent() {
        let mut context = context();
        assert_eq!(context.now(), "2000-01-01T00:00:00.000Z");
        assert_eq!(
            serialize_comments_with_info(&[], &mut context).0,
            String::new()
        );
        assert!(serialize_comments_extended_part(&[]).is_empty());
        assert!(serialize_comments_ids_part(&[]).is_empty());
        assert!(serialize_comments_extensible_part(&[], &[]).is_empty());
    }
}
