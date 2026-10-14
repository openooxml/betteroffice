use std::collections::{HashMap, HashSet};
use std::ops::Range;

use indexmap::IndexMap;
use serde::Serialize;
use serde_json::Value;

use crate::block::BlockContent;
use crate::comments::Comment;
use crate::paragraph::HexIdAllocator;
use crate::paragraph_identity::{attribute, parse_paragraph_id, tags, unescaped};
use crate::relationships::parse_relationships;
use crate::s8::{find_part, parse_comment_part};
use crate::styles::StyleMap;
use crate::xml::{
    ParseBudget, ParseError, ParseLimits, parse_javascript_integer_prefix, parse_xml,
    parse_xml_strict,
};

use super::context::SerializerContext;
use super::parts::{CommentParaInfo, serialize_comments_with_info};
use super::xml_writer::js_number;

pub(super) fn splice_comments(
    source: &[u8],
    parts: &[(String, Vec<u8>)],
    seed: &str,
    comments: &[Comment],
    serialized: &str,
    infos: &mut Vec<CommentParaInfo>,
    mut source_context: SerializerContext,
) -> Result<String, ParseError> {
    if !crate::xml::reads_as_written(source) {
        return Err(error_xml("source requires whole-part writing"));
    }
    let source = std::str::from_utf8(source).map_err(|error| error_xml(error.to_string()))?;
    let original = source_comments(parts, seed)?;
    let spans = comment_spans(source)?;
    let source_ids: HashSet<_> = spans.iter().map(|(id, _)| id.as_str()).collect();
    let model_ids: HashSet<_> = comments
        .iter()
        .map(|comment| js_number(comment.id))
        .collect();
    let original_ids: HashSet<_> = original
        .iter()
        .map(|comment| js_number(comment.id))
        .collect();
    if source_ids.len() != spans.len()
        || model_ids.len() != comments.len()
        || original_ids.len() != original.len()
    {
        return Err(error_xml("duplicate comment ids"));
    }
    if comments.len() != original.len()
        || comments
            .iter()
            .zip(&original)
            .any(|(comment, original)| js_number(comment.id) != js_number(original.id))
    {
        return Err(error_xml("comment ids or order changed"));
    }
    let parent_ids: HashSet<_> = comments
        .iter()
        .filter_map(|comment| comment.parent_id.map(js_number))
        .collect();
    let fragments: HashMap<_, _> = spans
        .iter()
        .map(|(id, span)| (id.as_str(), &source[span.clone()]))
        .collect();
    for (comment, original) in comments.iter().zip(&original) {
        let id = js_number(comment.id);
        let fragment = fragments
            .get(id.as_str())
            .ok_or_else(|| error_xml("missing source comment"))?;
        let same_para_id = match (&original.para_id, &comment.para_id) {
            (Some(source), Some(written)) => source.eq_ignore_ascii_case(written),
            (None, None) => true,
            _ => false,
        };
        if (original.para_id.is_none() && parent_ids.contains(&id))
            || !same_para_id
            || source_metadata_changed(original, comment)
            || !can_replay(original, comment, fragment)?
        {
            return Err(error_xml("comment changed"));
        }
    }
    let (source_xml, source_infos) = serialize_comments_with_info(&original, &mut source_context);
    if source_xml != serialized || source_infos != *infos {
        return Err(error_xml("comment serialization changed"));
    }
    let original_para_ids: HashMap<_, _> = original
        .iter()
        .map(|comment| (js_number(comment.id), comment.para_id.as_ref()))
        .collect();
    infos.retain_mut(|info| {
        if let Some(Some(para_id)) = original_para_ids.get(&js_number(info.comment_id)) {
            info.last_para_id.clone_from(para_id);
            true
        } else {
            false
        }
    });
    if infos.is_empty()
        && [
            "word/commentsExtended.xml",
            "word/commentsIds.xml",
            "word/commentsExtensible.xml",
        ]
        .iter()
        .any(|path| find_part(parts, path).is_some())
    {
        return Err(error_xml("empty comment companions"));
    }
    Ok(source.to_owned())
}

fn source_metadata_changed(source: &Comment, written: &Comment) -> bool {
    (source.done == Some(true)) != (written.done == Some(true))
        || source.parent_id != written.parent_id
}

fn can_replay(source: &Comment, written: &Comment, fragment: &str) -> Result<bool, ParseError> {
    if source
        .para_id
        .as_deref()
        .is_some_and(|id| parse_paragraph_id(id).is_none())
        || source.durable_id != written.durable_id
        || source.date_utc != written.date_utc
        || !unchanged(source, written)?
    {
        return Ok(false);
    }
    let limits = ParseLimits::default();
    let Ok(document) = parse_xml_strict(
        fragment.as_bytes(),
        "word/comments.xml",
        &mut ParseBudget::new(&limits),
    ) else {
        return Ok(false);
    };
    let root = document
        .root()
        .ok_or_else(|| error_xml("missing comment element"))?;
    if root.attributes.get("w:author") != Some(&written.author)
        || root.attributes.get("w:date") != written.date.as_ref()
        || root.attributes.get("w:initials") != written.initials.as_ref()
    {
        return Ok(false);
    }
    let fragment_tags = tags(fragment).ok_or_else(|| error_xml("invalid comment fragment"))?;
    let root_tag = fragment_tags
        .first()
        .ok_or_else(|| error_xml("missing comment element"))?;
    let id = js_number(written.id);
    if attribute(root_tag, "w:id").map(|range| &fragment[range]) != Some(id.as_str()) {
        return Ok(false);
    }
    let source_ids: Vec<_> = fragment_tags
        .iter()
        .filter(|tag| !tag.end && tag.name == "w:p")
        .map(|tag| attribute(tag, "w14:paraId").and_then(|range| unescaped(fragment, range)))
        .collect();
    let Some(model_ids) = comment_paragraph_ids(written) else {
        return Ok(false);
    };
    if source_ids.iter().all(Option::is_none) {
        if source_ids.len() != model_ids.len() || model_ids.iter().any(Option::is_some) {
            return Ok(false);
        }
    } else {
        let source_ids: Option<Vec<_>> = source_ids
            .iter()
            .map(|id| id.as_deref().and_then(parse_paragraph_id))
            .collect();
        let model_ids: Option<Vec<_>> = model_ids
            .iter()
            .map(|id| id.and_then(parse_paragraph_id))
            .collect();
        if source_ids.is_none() || source_ids != model_ids {
            return Ok(false);
        }
    }
    let last_para_id = source_ids
        .last()
        .and_then(|id| id.as_ref())
        .map(|id| id.to_ascii_uppercase());
    if last_para_id != source.para_id {
        return Ok(false);
    }
    if source_metadata_changed(source, written) {
        let root = fragment_tags
            .first()
            .ok_or_else(|| error_xml("missing comment element"))?;
        let inline_metadata = root.attributes.iter().any(|(name, range)| {
            (name.ends_with(":done") || *name == "done")
                && written.done != Some(true)
                && matches!(
                    unescaped(fragment, range.clone()).as_deref(),
                    Some("1" | "true")
                )
                || (name.ends_with(":parentId") || *name == "parentId")
                    && source.parent_id != written.parent_id
        });
        if source.para_id.is_some() && !inline_metadata {
            return Ok(true);
        }
        return Ok(false);
    }
    Ok(true)
}

fn comment_paragraph_ids(comment: &Comment) -> Option<Vec<Option<&str>>> {
    let mut ids = Vec::new();
    if comment.block_content.is_empty() {
        ids.extend(
            comment
                .content
                .iter()
                .map(|paragraph| paragraph.para_id.as_deref()),
        );
        return Some(ids);
    }
    let mut paragraphs = comment.content.iter();
    for block in &comment.block_content {
        if let BlockContent::Paragraph(block_paragraph) = block {
            let paragraph = paragraphs.next()?;
            let id = paragraph.para_id.as_deref();
            let block_id = block_paragraph.para_id.as_deref();
            match (id, block_id) {
                (Some(id), Some(block_id)) if id.eq_ignore_ascii_case(block_id) => {}
                (None, None) => {}
                _ => return None,
            }
            ids.push(id);
        } else {
            block_paragraph_ids(block, &mut ids);
        }
    }
    paragraphs.next().is_none().then_some(ids)
}

fn block_paragraph_ids<'a>(block: &'a BlockContent, ids: &mut Vec<Option<&'a str>>) {
    match block {
        BlockContent::Paragraph(paragraph) => ids.push(paragraph.para_id.as_deref()),
        BlockContent::Table(table) => {
            for row in &table.rows {
                for cell in &row.cells {
                    for block in &cell.content {
                        block_paragraph_ids(block, ids);
                    }
                }
            }
        }
        BlockContent::BlockSdt(sdt) => {
            for block in &sdt.content {
                block_paragraph_ids(block, ids);
            }
        }
        BlockContent::RawXml(_) => {}
    }
}

fn unchanged(source: &Comment, written: &Comment) -> Result<bool, ParseError> {
    Ok(source.author == written.author
        && source.initials == written.initials
        && source.date == written.date
        && source.author_id == written.author_id
        && projection(&source.content)? == projection(&written.content)?
        && (written.block_content.is_empty()
            || projection(&source.block_content)? == projection(&written.block_content)?))
}

fn projection(value: &impl Serialize) -> Result<Value, ParseError> {
    let mut value = serde_json::to_value(value).map_err(|error| error_xml(error.to_string()))?;
    normalize(&mut value);
    Ok(value)
}

fn normalize(value: &mut Value) {
    match value {
        Value::Object(object) => {
            if object.get("type").and_then(Value::as_str) == Some("paragraph") {
                for key in [
                    "paraId",
                    "repeatedParaId",
                    "paraIdAttribute",
                    "sourceOrdinal",
                    "textId",
                ] {
                    object.remove(key);
                }
            }
            if object
                .get("rId")
                .and_then(Value::as_str)
                .is_some_and(|id| !id.is_empty())
            {
                object.remove("src");
            }
            for child in object.values_mut() {
                normalize(child);
            }
        }
        Value::Array(array) => array.iter_mut().for_each(normalize),
        _ => {}
    }
}

fn source_comments(parts: &[(String, Vec<u8>)], seed: &str) -> Result<Vec<Comment>, ParseError> {
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let settings = crate::settings::parse_settings(
        find_part(parts, "word/settings.xml").map(|(_, bytes)| bytes),
        "word/settings.xml",
        &mut budget,
    )?;
    let mut theme = crate::theme::parse_theme(
        find_part(parts, "word/theme/theme1.xml").map(|(_, bytes)| bytes),
        "word/theme/theme1.xml",
        &mut budget,
    )?;
    crate::theme::apply_theme_font_lang(&mut theme, settings.theme_font_lang.as_ref());
    let styles = find_part(parts, "word/styles.xml")
        .filter(|(_, xml)| !xml.is_empty())
        .map(|(path, xml)| {
            crate::styles::parse_style_definitions(xml, Some(&theme), path, &mut budget)
        })
        .transpose()?;
    let style_map: StyleMap = styles
        .as_ref()
        .map(|definitions| {
            definitions
                .styles
                .iter()
                .map(|style| (style.style_id.clone(), style.clone()))
                .collect()
        })
        .unwrap_or_default();
    let document_path = crate::relationships::office_document_path(parts, &mut budget)?;
    let rels_path = crate::relationships::relationship_part_path(&document_path);
    let relationships = find_part(parts, &rels_path)
        .map(|(path, xml)| parse_relationships(xml, path, &mut budget))
        .transpose()?
        .unwrap_or_default();
    let media = crate::media::build_media_map(parts);
    let all_xml: IndexMap<_, _> = parts
        .iter()
        .filter(|(path, _)| {
            let lower = path.to_ascii_lowercase();
            lower.ends_with(".xml") || lower.ends_with(".rels")
        })
        .map(|(path, bytes)| (path.clone(), bytes.as_slice()))
        .collect();
    let charts = crate::chart::parse_chart_parts(&all_xml, &limits);
    let mut smart_art = crate::smart_art::create_smart_art_context(&all_xml);
    let mut ids = HexIdAllocator::from_sha256(seed)?;
    parse_comment_part(
        parts,
        &relationships,
        Some(&theme),
        Some(&style_map),
        styles
            .as_ref()
            .and_then(|definitions| definitions.doc_defaults.as_ref()),
        &media,
        &charts,
        &mut smart_art,
        &mut budget,
        &mut ids,
    )
}

fn comment_spans(xml: &str) -> Result<Vec<(String, Range<usize>)>, ParseError> {
    if xml.is_empty() {
        return Ok(Vec::new());
    }
    let limits = ParseLimits::default();
    let document = parse_xml(
        xml.as_bytes(),
        "word/comments.xml",
        &mut ParseBudget::new(&limits),
    )?;
    let root = document
        .root()
        .ok_or_else(|| error_xml("missing comments root"))?;
    if root.local_name() != "comments" {
        return Err(error_xml("invalid comments root"));
    }
    let mut ids = root.child_elements().map(|child| {
        (child.local_name() == "comment")
            .then(|| {
                child
                    .attribute(Some("w"), "id")
                    .and_then(parse_javascript_integer_prefix)
            })
            .flatten()
            .map(js_number)
    });
    let mut spans = Vec::new();
    let mut depth = 0usize;
    let mut current = None;
    let mut end = None;
    for tag in tags(xml).ok_or_else(|| error_xml("invalid comment spans"))? {
        if tag.end {
            depth = depth
                .checked_sub(1)
                .ok_or_else(|| error_xml("unbalanced comments"))?;
            if depth == 1 {
                if let Some((id, start)) = current.take() {
                    spans.push((id, start..tag.range.end));
                }
            } else if depth == 0 {
                end = Some(tag.range.start);
            }
        } else {
            if depth == 1 {
                let id = ids
                    .next()
                    .ok_or_else(|| error_xml("comment span mismatch"))?;
                if let Some(id) = id {
                    if tag.empty {
                        spans.push((id, tag.range.clone()));
                    } else {
                        current = Some((id, tag.range.start));
                    }
                }
            }
            depth += usize::from(!tag.empty);
            if depth == 0 && tag.empty {
                end = Some(tag.range.end);
            }
        }
    }
    end.ok_or_else(|| error_xml("missing comments end"))?;
    Ok(spans)
}

fn error_xml(message: impl Into<String>) -> ParseError {
    ParseError::Canonical(format!("comments splice: {}", message.into()))
}
