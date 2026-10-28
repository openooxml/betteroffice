use std::collections::{HashMap, HashSet};
use std::ops::Range;

use indexmap::IndexMap;
use serde::Serialize;
use serde_json::Value;

use crate::block::BlockContent;
use crate::comments::Comment;
use crate::paragraph::HexIdAllocator;
use crate::paragraph_identity::{Tag, attribute, parse_paragraph_id, tags, unescaped};
use crate::relationships::parse_relationships;
use crate::s8::{find_part, parse_comment_part};
use crate::styles::StyleMap;
use crate::xml::{
    ParseBudget, ParseError, ParseLimits, parse_javascript_integer_prefix, parse_xml,
    parse_xml_strict,
};

use super::context::SerializerContext;
use super::parts::{
    CommentParaInfo, serialize_comments_extended_part, serialize_comments_extensible_part,
    serialize_comments_ids_part, serialize_comments_with_info,
};
use super::xml_writer::{escape_xml, js_number};

pub(super) struct SplicedComments {
    pub(super) xml: String,
    pub(super) companions: Option<Vec<(String, String)>>,
}

pub(super) fn splice_comments(
    source: &[u8],
    parts: &[(String, Vec<u8>)],
    seed: &str,
    comments: &[Comment],
    serialized: &str,
    infos: &mut Vec<CommentParaInfo>,
    mut source_context: SerializerContext,
) -> Result<SplicedComments, ParseError> {
    if !crate::xml::reads_as_written(source) {
        return Err(error_xml("source requires whole-part writing"));
    }
    let source = std::str::from_utf8(source).map_err(|error| error_xml(error.to_string()))?;
    let original = source_comments(parts, seed)?;
    let (spans, close) = comment_spans(source)?;
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
    if comments.len() < original.len()
        || comments
            .iter()
            .zip(&original)
            .any(|(comment, original)| js_number(comment.id) != js_number(original.id))
    {
        return Err(error_xml("comment ids or order changed"));
    }
    let (kept, added) = comments.split_at(original.len());
    let thread_changed = added.iter().any(|comment| comment.parent_id.is_some())
        || kept.iter().zip(&original).any(|(comment, source)| {
            comment.done != source.done || comment.parent_id != source.parent_id
        });
    let mut parent_ids: HashSet<_> = added
        .iter()
        .chain(
            kept.iter()
                .zip(&original)
                .filter(|(comment, source)| source_metadata_changed(source, comment))
                .map(|(comment, _)| comment),
        )
        .filter_map(|comment| comment.parent_id.map(js_number))
        .collect();
    if !thread_changed {
        parent_ids.extend(
            kept.iter()
                .filter_map(|comment| comment.parent_id.map(js_number)),
        );
    }
    let fragments: HashMap<_, _> = spans
        .iter()
        .map(|(id, span)| (id.as_str(), &source[span.clone()]))
        .collect();
    let mut text_patches: HashMap<String, Vec<(Range<usize>, String)>> = HashMap::new();
    for (comment, original) in kept.iter().zip(&original) {
        let id = js_number(comment.id);
        let fragment = fragments
            .get(id.as_str())
            .ok_or_else(|| error_xml("missing source comment"))?;
        if let Some(patch) = comment_text_patch(original, comment, fragment)? {
            text_patches.insert(id.clone(), patch);
        }
        let same_para_id = match (&original.para_id, &comment.para_id) {
            (Some(source), Some(written)) => source.eq_ignore_ascii_case(written),
            (None, None) => true,
            _ => false,
        };
        if (!thread_changed && original.para_id.is_none() && parent_ids.contains(&id))
            || !same_para_id
            || !can_replay(
                original,
                comment,
                fragment,
                thread_changed,
                text_patches.contains_key(&id),
            )?
        {
            return Err(error_xml("comment changed"));
        }
        if original.para_id.is_some()
            && (source_metadata_changed(original, comment) || parent_ids.contains(&id))
        {
            let fragment_tags = tags(fragment).ok_or_else(|| error_xml("invalid comment"))?;
            if direct_paragraphs(&fragment_tags)
                .last()
                .and_then(|paragraph| attribute(paragraph, "w14:paraId"))
                .is_none()
            {
                return Err(error_xml("missing final comment paragraph id"));
            }
        }
    }
    let (source_xml, source_infos) =
        serialize_comments_with_info(&original, &mut source_context.clone());
    let (kept_xml, kept_infos) = serialize_comments_with_info(kept, &mut source_context);
    let infos_changed = if thread_changed {
        source_infos.iter().zip(&kept_infos).any(|(source, kept)| {
            source.comment_id != kept.comment_id
                || source.last_para_id != kept.last_para_id
                || source.durable_id != kept.durable_id
        })
    } else {
        source_infos != kept_infos
    };
    if infos_changed {
        return Err(error_xml("comment serialization changed"));
    }
    if source_xml != kept_xml {
        let source_fragments: HashMap<_, _> = comment_spans(&source_xml)?.0.into_iter().collect();
        let kept_fragments: HashMap<_, _> = comment_spans(&kept_xml)?.0.into_iter().collect();
        for comment in kept.iter() {
            let id = js_number(comment.id);
            let differs = match (source_fragments.get(&id), kept_fragments.get(&id)) {
                (Some(source_span), Some(kept_span)) => {
                    source_xml[source_span.clone()] != kept_xml[kept_span.clone()]
                }
                _ => true,
            };
            if differs && !text_patches.contains_key(&id) {
                return Err(error_xml("comment serialization changed"));
            }
        }
    }
    let original_para_ids: HashMap<_, _> = original
        .iter()
        .map(|comment| (js_number(comment.id), comment.para_id.as_ref()))
        .collect();
    let added_ids: HashSet<_> = added.iter().map(|comment| js_number(comment.id)).collect();
    let needed_ids: HashSet<_> = kept
        .iter()
        .zip(&original)
        .filter(|(comment, source)| source_metadata_changed(source, comment))
        .map(|(comment, _)| js_number(comment.id))
        .chain(parent_ids)
        .collect();
    infos.retain_mut(|info| {
        let id = js_number(info.comment_id);
        if let Some(Some(para_id)) = original_para_ids.get(&id) {
            info.last_para_id.clone_from(para_id);
            true
        } else {
            added_ids.contains(&id) || needed_ids.contains(&id)
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
    let mut patches = Vec::new();
    for (id, span) in &spans {
        if let Some(text_patch) = text_patches.get(id) {
            for (range, replacement) in text_patch {
                patches.push((
                    span.start + range.start..span.start + range.end,
                    replacement.clone(),
                ));
            }
        }
        if original_para_ids.get(id).is_some_and(|id| id.is_none()) && needed_ids.contains(id) {
            let fragment = &source[span.clone()];
            let fragment_tags = tags(fragment).ok_or_else(|| error_xml("invalid parent"))?;
            let paragraph = direct_paragraphs(&fragment_tags)
                .last()
                .ok_or_else(|| error_xml("missing parent paragraph"))?;
            if paragraph
                .attributes
                .iter()
                .any(|(key, _)| key.ends_with(":paraId"))
            {
                return Err(error_xml("unsupported parent paragraph id"));
            }
            let info = infos
                .iter()
                .find(|info| js_number(info.comment_id) == *id)
                .ok_or_else(|| error_xml("missing parent info"))?;
            let opening = &fragment[paragraph.range.clone()];
            let mut patched = set_attribute(opening, "w14:paraId", Some(&info.last_para_id))?;
            patched = with_namespaces(&patched, serialized, source)?;
            patches.push((
                span.start + paragraph.range.start..span.start + paragraph.range.end,
                patched,
            ));
        }
    }
    if !added.is_empty() {
        let close = close.ok_or_else(|| error_xml("self-closing comments root"))?;
        let written: HashMap<_, _> = comment_spans(serialized)?.0.into_iter().collect();
        let mut appended = String::new();
        for comment in added {
            let span = written
                .get(&js_number(comment.id))
                .ok_or_else(|| error_xml("missing written comment"))?;
            appended.push_str(&serialized[span.clone()]);
        }
        patches.push((close..close, appended));
    }
    let spliced = apply_patches(source, patches);
    if !crate::xml::reads_as_written(spliced.as_bytes()) {
        return Err(error_xml("added comments require whole-part writing"));
    }
    let companions = if thread_changed || (!text_patches.is_empty() && !added.is_empty()) {
        Some(splice_companions(
            parts, &original, comments, infos, &added_ids,
        )?)
    } else if !text_patches.is_empty() {
        Some(Vec::new())
    } else {
        None
    };
    Ok(SplicedComments {
        xml: spliced,
        companions,
    })
}

fn direct_paragraphs<'a, 'b>(tags: &'a [Tag<'b>]) -> impl Iterator<Item = &'a Tag<'b>> {
    let mut depth = 0usize;
    tags.iter().filter(move |tag| {
        if tag.end {
            depth = depth.saturating_sub(1);
            false
        } else {
            let paragraph = depth == 1 && tag.name == "w:p";
            depth += usize::from(!tag.empty);
            paragraph
        }
    })
}

fn apply_patches(source: &str, mut patches: Vec<(Range<usize>, String)>) -> String {
    patches.sort_by_key(|(span, _)| span.start);
    let mut result = source.to_owned();
    for (span, replacement) in patches.into_iter().rev() {
        result.replace_range(span, &replacement);
    }
    result
}

fn set_attribute(xml: &str, key: &str, value: Option<&str>) -> Result<String, ParseError> {
    let parsed = tags(xml).ok_or_else(|| error_xml("invalid attribute target"))?;
    let tag = parsed
        .first()
        .ok_or_else(|| error_xml("missing attribute target"))?;
    let mut result = xml.to_owned();
    if let Some(span) = attribute(tag, key) {
        if let Some(value) = value {
            result.replace_range(span, value);
        } else {
            let start = xml[..span.start]
                .rfind(key)
                .ok_or_else(|| error_xml("missing attribute name"))?;
            let mut start = start;
            while start > 0 && xml.as_bytes()[start - 1].is_ascii_whitespace() {
                start -= 1;
            }
            result.replace_range(start..span.end + 1, "");
        }
    } else if let Some(value) = value {
        let end = tag.range.end - if tag.empty { 2 } else { 1 };
        result.insert_str(end, &format!(" {key}=\"{value}\""));
    }
    Ok(result)
}

fn with_namespaces(fragment: &str, serialized: &str, source: &str) -> Result<String, ParseError> {
    let written = tags(serialized).ok_or_else(|| error_xml("invalid written namespaces"))?;
    let source_tags = tags(source).ok_or_else(|| error_xml("invalid source namespaces"))?;
    let fragment_tags = tags(fragment).ok_or_else(|| error_xml("invalid fragment namespaces"))?;
    let written_root = written
        .first()
        .ok_or_else(|| error_xml("missing written root"))?;
    let source_root = source_tags
        .first()
        .ok_or_else(|| error_xml("missing source root"))?;
    let fragment_root = fragment_tags
        .first()
        .ok_or_else(|| error_xml("missing fragment root"))?;
    let prefixes: HashSet<_> = fragment_tags
        .iter()
        .flat_map(|tag| std::iter::once(tag.name).chain(tag.attributes.iter().map(|(key, _)| *key)))
        .filter_map(|name| name.split_once(':').map(|(prefix, _)| prefix))
        .filter(|prefix| *prefix != "xml" && *prefix != "xmlns")
        .collect();
    let mut prefixes: Vec<_> = prefixes.into_iter().collect();
    prefixes.sort_unstable();
    let mut result = fragment.to_owned();
    for prefix in prefixes {
        let key = format!("xmlns:{prefix}");
        if attribute(fragment_root, &key).is_some() {
            continue;
        }
        let Some(value) =
            attribute(written_root, &key).and_then(|range| unescaped(serialized, range))
        else {
            return Err(error_xml("missing written namespace"));
        };
        if attribute(source_root, &key)
            .and_then(|range| unescaped(source, range))
            .as_ref()
            != Some(&value)
        {
            result = set_attribute(&result, &key, Some(&value))?;
        }
    }
    Ok(result)
}

fn splice_companions(
    parts: &[(String, Vec<u8>)],
    original: &[Comment],
    comments: &[Comment],
    infos: &[CommentParaInfo],
    added_ids: &HashSet<String>,
) -> Result<Vec<(String, String)>, ParseError> {
    let changed: HashSet<_> = comments
        .iter()
        .zip(original)
        .filter(|(comment, source)| source_metadata_changed(source, comment))
        .map(|(comment, _)| js_number(comment.id))
        .collect();
    let added_infos: Vec<_> = infos
        .iter()
        .filter(|info| added_ids.contains(&js_number(info.comment_id)))
        .cloned()
        .collect();
    let identity_infos = if changed.is_empty() && added_ids.is_empty() {
        infos
    } else {
        &added_infos
    };
    let written = [
        (
            "word/commentsExtended.xml",
            "commentEx",
            serialize_comments_extended_part(infos),
        ),
        (
            "word/commentsIds.xml",
            "commentId",
            serialize_comments_ids_part(identity_infos),
        ),
        (
            "word/commentsExtensible.xml",
            "commentExtensible",
            serialize_comments_extensible_part(identity_infos, comments),
        ),
    ];
    let mut result = Vec::new();
    for (path, row_name, serialized) in written {
        let Some((_, source)) = find_part(parts, path) else {
            if !serialized.is_empty() {
                result.push((path.to_owned(), serialized));
            }
            continue;
        };
        if added_ids.is_empty() && (row_name != "commentEx" || changed.is_empty()) {
            continue;
        }
        if !crate::xml::reads_as_written(source) {
            return Err(error_xml("companion requires whole-part writing"));
        }
        let source = std::str::from_utf8(source).map_err(|error| error_xml(error.to_string()))?;
        let (rows, close) = companion_rows(source, row_name)?;
        let (written_rows, _) = companion_rows(&serialized, row_name)?;
        let (prefix, key) = match row_name {
            "commentEx" => ("w15", "paraId"),
            "commentId" => ("w16cid", "paraId"),
            _ => ("w16cex", "durableId"),
        };
        let qualified_key = format!("{prefix}:{key}");
        let row_id = |xml: &str, span: &Range<usize>| -> Result<String, ParseError> {
            let parsed =
                tags(&xml[span.clone()]).ok_or_else(|| error_xml("invalid companion row"))?;
            let tag = parsed
                .first()
                .ok_or_else(|| error_xml("missing companion row"))?;
            let range = attribute(tag, &qualified_key)
                .or_else(|| attribute(tag, key))
                .ok_or_else(|| error_xml("missing companion key"))?;
            unescaped(&xml[span.clone()], range)
                .map(|value| value.to_ascii_uppercase())
                .ok_or_else(|| error_xml("invalid companion key"))
        };
        let mut by_id = HashMap::new();
        for span in &rows {
            if by_id.insert(row_id(source, span)?, span).is_some() {
                return Err(error_xml("duplicate companion rows"));
            }
        }
        let mut patches = Vec::new();
        let mut appended = String::new();
        for span in written_rows {
            let id = row_id(&serialized, &span)?;
            let info = infos
                .iter()
                .find(|info| {
                    if key == "durableId" {
                        info.durable_id.eq_ignore_ascii_case(&id)
                    } else {
                        info.last_para_id.eq_ignore_ascii_case(&id)
                    }
                })
                .ok_or_else(|| error_xml("missing companion info"))?;
            let comment_id = js_number(info.comment_id);
            let added = added_ids.contains(&comment_id);
            let metadata_changed = row_name == "commentEx" && changed.contains(&comment_id);
            if !added && !metadata_changed {
                continue;
            }
            if let Some(source_span) = by_id.get(&id) {
                if added {
                    return Err(error_xml("added companion identity collision"));
                }
                let fragment = &source[(*source_span).clone()];
                let parsed = tags(fragment).ok_or_else(|| error_xml("invalid extended row"))?;
                let tag = parsed
                    .first()
                    .ok_or_else(|| error_xml("missing extended row"))?;
                let prefix = if attribute(tag, &qualified_key).is_some() {
                    "w15:"
                } else {
                    ""
                };
                let original = original
                    .iter()
                    .find(|comment| comment.id == info.comment_id)
                    .ok_or_else(|| error_xml("missing original comment"))?;
                let mut patched = fragment.to_owned();
                if (info.done == Some(true)) != (original.done == Some(true)) {
                    let done = if info.done == Some(true) { "1" } else { "0" };
                    patched = set_attribute(&patched, &format!("{prefix}done"), Some(done))?;
                }
                if info.parent_id != original.parent_id {
                    let parent = info
                        .parent_id
                        .and_then(|id| infos.iter().find(|candidate| candidate.comment_id == id));
                    patched = set_attribute(
                        &patched,
                        &format!("{prefix}paraIdParent"),
                        parent.map(|parent| parent.last_para_id.as_str()),
                    )?;
                }
                patches.push(((*source_span).clone(), patched));
            } else {
                appended.push_str(&with_namespaces(&serialized[span], &serialized, source)?);
            }
        }
        if !appended.is_empty() {
            let close = close.ok_or_else(|| error_xml("self-closing companion root"))?;
            patches.push((close..close, appended));
        }
        result.push((path.to_owned(), apply_patches(source, patches)));
    }
    Ok(result)
}

fn companion_rows(
    xml: &str,
    row_name: &str,
) -> Result<(Vec<Range<usize>>, Option<usize>), ParseError> {
    let mut rows = Vec::new();
    let mut depth = 0usize;
    let mut start = None;
    let mut close = None;
    for tag in tags(xml).ok_or_else(|| error_xml("invalid companion spans"))? {
        if tag.end {
            depth = depth
                .checked_sub(1)
                .ok_or_else(|| error_xml("unbalanced companion"))?;
            if depth == 1 {
                if let Some(start) = start.take() {
                    rows.push(start..tag.range.end);
                }
            } else if depth == 0 {
                close = Some(tag.range.start);
            }
        } else {
            if depth == 1 && tag.name.rsplit(':').next() == Some(row_name) {
                if tag.empty {
                    rows.push(tag.range.clone());
                } else {
                    start = Some(tag.range.start);
                }
            }
            depth += usize::from(!tag.empty);
        }
    }
    Ok((rows, close))
}

fn source_metadata_changed(source: &Comment, written: &Comment) -> bool {
    (source.done == Some(true)) != (written.done == Some(true))
        || source.parent_id != written.parent_id
}

fn can_replay(
    source: &Comment,
    written: &Comment,
    fragment: &str,
    thread_changed: bool,
    text_only: bool,
) -> Result<bool, ParseError> {
    if source
        .para_id
        .as_deref()
        .is_some_and(|id| parse_paragraph_id(id).is_none())
        || source.durable_id != written.durable_id
        || source.date_utc != written.date_utc
        || !(unchanged(source, written)? || text_only)
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
    if !thread_changed
        && source_ids.iter().any(Option::is_some)
        && source_ids.iter().any(Option::is_none)
    {
        return Ok(false);
    }
    let source_values: Option<Vec<_>> = source_ids
        .iter()
        .map(|id| match id {
            Some(id) => parse_paragraph_id(id).map(Some),
            None => Some(None),
        })
        .collect();
    let model_values: Option<Vec<_>> = model_ids
        .iter()
        .map(|id| match id {
            Some(id) => parse_paragraph_id(id).map(Some),
            None => Some(None),
        })
        .collect();
    if source_values.is_none() || source_values != model_values {
        return Ok(false);
    }
    let last_para_id = if thread_changed {
        direct_paragraphs(&fragment_tags)
            .filter_map(|tag| {
                attribute(tag, "w14:paraId").and_then(|range| unescaped(fragment, range))
            })
            .last()
            .map(|id| id.to_ascii_uppercase())
    } else {
        source_ids
            .last()
            .and_then(|id| id.as_ref())
            .map(|id| id.to_ascii_uppercase())
    };
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
                && unescaped(fragment, range.clone())
                    .as_deref()
                    .and_then(crate::xml::parse_on_off)
                    == Some(true)
                || (name.ends_with(":parentId") || *name == "parentId")
                    && source.parent_id != written.parent_id
        });
        return Ok(!inline_metadata);
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

fn blanked(value: &impl Serialize) -> Result<Value, ParseError> {
    let mut value = projection(value)?;
    blank_text(&mut value);
    Ok(value)
}

fn blank_text(value: &mut Value) {
    match value {
        Value::Object(object) => {
            if object.get("type").and_then(Value::as_str) == Some("paragraph") {
                for key in [
                    "extraAttributes",
                    "listRendering",
                    "renderedPageBreakBefore",
                ] {
                    object.remove(key);
                }
                if object.get("formatting").is_some_and(|formatting| {
                    formatting.is_null()
                        || formatting
                            .as_object()
                            .is_some_and(|values| values.is_empty())
                }) {
                    object.remove("formatting");
                }
            }
            if object.get("type").and_then(Value::as_str) == Some("text") {
                object.insert("text".to_owned(), Value::String(String::new()));
            }
            for child in object.values_mut() {
                blank_text(child);
            }
        }
        Value::Array(array) => {
            for child in array.iter_mut() {
                blank_text(child);
            }
        }
        _ => {}
    }
}

fn value_text_parts(value: &Value, out: &mut Vec<String>) {
    match value {
        Value::Array(array) => array.iter().for_each(|child| value_text_parts(child, out)),
        Value::Object(object) => {
            if object.get("type").and_then(Value::as_str) == Some("text") {
                if let Some(text) = object.get("text").and_then(Value::as_str) {
                    out.push(text.to_owned());
                }
                return;
            }
            let structured_result = object.contains_key("structuredResult");
            for key in [
                "structuredChildren",
                "structuredResult",
                "content",
                "children",
                "fieldResult",
                "inline",
            ] {
                if key == "children" && object.contains_key("structuredChildren") {
                    continue;
                }
                if key == "content" && structured_result {
                    continue;
                }
                if key == "fieldResult" && structured_result {
                    continue;
                }
                if let Some(child) = object.get(key) {
                    value_text_parts(child, out);
                }
            }
        }
        _ => {}
    }
}

fn model_paragraph_text_parts(
    paragraph: &crate::paragraph::Paragraph,
) -> Result<Vec<String>, ParseError> {
    let value = projection(paragraph)?;
    let mut parts = Vec::new();
    value_text_parts(&value, &mut parts);
    Ok(parts)
}

fn direct_paragraph_spans(tags: &[Tag]) -> Vec<Range<usize>> {
    let mut spans = Vec::new();
    let mut depth = 0usize;
    let mut start: Option<usize> = None;
    for tag in tags {
        if tag.end {
            depth = depth.saturating_sub(1);
            if depth == 1 {
                if let Some(start) = start.take() {
                    spans.push(start..tag.range.end);
                }
            }
        } else {
            if depth == 1 && tag.name == "w:p" {
                if tag.empty {
                    spans.push(tag.range.clone());
                } else {
                    start = Some(tag.range.start);
                }
            }
            depth += usize::from(!tag.empty);
        }
    }
    spans
}

struct CommentTextNode {
    range: Range<usize>,
    tag: Range<usize>,
    text: String,
    empty: bool,
    preserve: bool,
}

fn comment_text_nodes(
    fragment: &str,
    tags: &[Tag],
    span: &Range<usize>,
) -> Option<Vec<CommentTextNode>> {
    let mut nodes = Vec::new();
    let mut open: Option<&Tag> = None;
    for tag in tags {
        if tag.range.start < span.start || tag.range.start >= span.end || tag.name != "w:t" {
            continue;
        }
        if tag.end {
            let begin = open.take()?;
            let range = begin.range.end..tag.range.start;
            let preserve = attribute(begin, "xml:space")
                .and_then(|value| unescaped(fragment, value))
                .as_deref()
                == Some("preserve");
            nodes.push(CommentTextNode {
                text: unescaped(fragment, range.clone())?,
                range,
                tag: begin.range.clone(),
                empty: false,
                preserve,
            });
        } else if tag.empty {
            nodes.push(CommentTextNode {
                range: tag.range.end..tag.range.end,
                tag: tag.range.clone(),
                text: String::new(),
                empty: true,
                preserve: false,
            });
        } else {
            open = Some(tag);
        }
    }
    if open.is_some() {
        return None;
    }
    Some(nodes)
}

fn comment_text_patch(
    source: &Comment,
    written: &Comment,
    fragment: &str,
) -> Result<Option<Vec<(Range<usize>, String)>>, ParseError> {
    if source.content.len() != written.content.len()
        || blanked(&source.content)? != blanked(&written.content)?
        || (!written.block_content.is_empty()
            && blanked(&source.block_content)? != blanked(&written.block_content)?)
    {
        return Ok(None);
    }
    let Some(fragment_tags) = tags(fragment) else {
        return Ok(None);
    };
    let spans = direct_paragraph_spans(&fragment_tags);
    if spans.len() != source.content.len() {
        return Ok(None);
    }
    let mut patches = Vec::new();
    let mut changed = false;
    for (index, paragraph) in source.content.iter().enumerate() {
        let source_parts = model_paragraph_text_parts(paragraph)?;
        let written_parts = model_paragraph_text_parts(&written.content[index])?;
        if source_parts == written_parts {
            continue;
        }
        changed = true;
        let Some(nodes) = comment_text_nodes(fragment, &fragment_tags, &spans[index]) else {
            return Ok(None);
        };
        if nodes.len() != source_parts.len()
            || nodes.len() != written_parts.len()
            || nodes
                .iter()
                .zip(&source_parts)
                .any(|(node, part)| node.text != *part)
        {
            return Ok(None);
        }
        for (node, replacement) in nodes.iter().zip(&written_parts) {
            if *replacement != node.text {
                if node.empty {
                    return Ok(None);
                }
                patches.push((node.range.clone(), escape_xml(replacement)));
                let edge_whitespace = matches!(replacement.chars().next(), Some(' ' | '\t'))
                    || matches!(replacement.chars().next_back(), Some(' ' | '\t'));
                if edge_whitespace && !node.preserve {
                    patches.push((
                        node.tag.clone(),
                        set_attribute(&fragment[node.tag.clone()], "xml:space", Some("preserve"))?,
                    ));
                }
            }
        }
    }
    Ok(changed.then_some(patches))
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

/// Each comment's span and where the root's end tag starts (`None` for a self-closing root).
fn comment_spans(xml: &str) -> Result<(Vec<(String, Range<usize>)>, Option<usize>), ParseError> {
    if xml.is_empty() {
        return Ok((Vec::new(), None));
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
                end = Some(Some(tag.range.start));
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
                end = Some(None);
            }
        }
    }
    let close = end.ok_or_else(|| error_xml("missing comments end"))?;
    Ok((spans, close))
}

fn error_xml(message: impl Into<String>) -> ParseError {
    ParseError::Canonical(format!("comments splice: {}", message.into()))
}

#[cfg(test)]
mod tests {
    use crate::formatting::{ParagraphFormatting, TextFormatting};
    use crate::inline::{
        InlineNode, PropertyChangeInfo, Run, RunContent, RunPropertyChange, RunPropertyChangeType,
        RunType,
    };
    use crate::paragraph::{Paragraph, ParagraphContent};

    use super::*;

    const FRAGMENT: &str = "<w:comment w:id=\"0\" w:author=\"Reviewer\" w:date=\"2026-01-05T09:00:00Z\" w:initials=\"R\"><w:p w14:paraId=\"1A2B3C4D\" w14:textId=\"77777777\" w:rsidR=\"00AB\" w:rsidRDefault=\"00CD\"><w:pPr><w:pStyle w:val=\"CommentText\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>First identity proofed here. </w:t></w:r><w:proofErr w:type=\"spellStart\"/><w:r><w:rPr><w:b/></w:rPr><w:t>Second identity </w:t></w:r><w:r><w:rPr><w:color w:val=\"0066AA\"/></w:rPr><w:t>proofed there.</w:t></w:r><w:proofErr w:type=\"spellEnd\"/></w:p></w:comment>";

    #[test]
    fn source_guard_requires_preserve_for_edge_spaces_in_any_comment() {
        let guarded = "<w:comment w:id=\"1\"><w:p><w:fldSimple w:instr=\" PAGE \"><w:r><w:t>7</w:t></w:r></w:fldSimple><w:hyperlink w:anchor=\"target\"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p></w:comment>";
        let source = format!(
            "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">{FRAGMENT}{guarded}</w:comments>"
        );
        assert!(!crate::xml::reads_as_written(source.as_bytes()));
        let preserved = source
            .replace(
                "<w:t>First identity proofed here. </w:t>",
                "<w:t xml:space=\"preserve\">First identity proofed here. </w:t>",
            )
            .replace(
                "<w:t>Second identity </w:t>",
                "<w:t xml:space=\"preserve\">Second identity </w:t>",
            );
        assert!(crate::xml::reads_as_written(preserved.as_bytes()));
    }

    fn run(text: &str, formatting: Option<TextFormatting>) -> ParagraphContent {
        ParagraphContent::Inline(InlineNode::Run(Run {
            node_type: RunType::Run,
            formatting,
            property_changes: None,
            content: vec![RunContent::Text {
                text: text.to_owned(),
                preserve_space: None,
            }],
        }))
    }

    fn comment(first: &str, second: &str, third: &str) -> Comment {
        let reference = ParagraphContent::Inline(InlineNode::Run(Run {
            node_type: RunType::Run,
            formatting: None,
            property_changes: None,
            content: vec![RunContent::CommentReference { id: None }],
        }));
        Comment {
            id: 0.0,
            author: "Reviewer".to_owned(),
            initials: Some("R".to_owned()),
            date: Some("2026-01-05T09:00:00Z".to_owned()),
            content: vec![Paragraph {
                node_type: "paragraph".to_owned(),
                para_id: Some("1A2B3C4D".to_owned()),
                repeated_para_id: None,
                para_id_attribute: None,
                source_ordinal: None,
                text_id: Some("77777777".to_owned()),
                extra_attributes: Vec::new(),
                formatting: None,
                property_changes: None,
                p_pr_ins: None,
                p_pr_del: None,
                content: vec![
                    reference,
                    run(first, None),
                    run(
                        second,
                        Some(TextFormatting {
                            bold: Some(true),
                            ..TextFormatting::default()
                        }),
                    ),
                    run(third, None),
                ],
                list_rendering: None,
                rendered_page_break_before: None,
                section_properties: None,
            }],
            parent_id: None,
            done: Some(false),
            status: "active".to_owned(),
            author_id: None,
            durable_id: Some("2ECC71AA".to_owned()),
            para_id: Some("1A2B3C4D".to_owned()),
            date_utc: None,
            palette_index: 0.0,
            block_content: Vec::new(),
        }
    }

    #[test]
    fn text_patch_changes_only_wt_text_and_keeps_ids_and_proofing() {
        let source = comment(
            "First identity proofed here. ",
            "Second identity ",
            "proofed there.",
        );
        let written = comment(
            "First identity-proofed here. ",
            "Second identity-proofed",
            " there.",
        );
        let patch = comment_text_patch(&source, &written, FRAGMENT)
            .unwrap()
            .expect("a text-only change patches");
        let expected = FRAGMENT
            .replace(
                "<w:t>First identity proofed here. </w:t>",
                "<w:t xml:space=\"preserve\">First identity-proofed here. </w:t>",
            )
            .replace(
                "<w:t>Second identity </w:t>",
                "<w:t>Second identity-proofed</w:t>",
            )
            .replace(
                "<w:t>proofed there.</w:t>",
                "<w:t xml:space=\"preserve\"> there.</w:t>",
            );
        assert_eq!(apply_patches(FRAGMENT, patch), expected);
    }

    #[test]
    fn text_patch_adds_preserve_to_nodes_gaining_edge_whitespace() {
        const FRAGMENT: &str = "<w:comment w:id=\"0\" w:author=\"Reviewer\" w:date=\"2026-01-05T09:00:00Z\" w:initials=\"R\"><w:p w14:paraId=\"1A2B3C4D\" w14:textId=\"77777777\" w:rsidR=\"00AB\"><w:pPr><w:pStyle w:val=\"CommentText\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>alpha be</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>ta gamma</w:t></w:r><w:r><w:t> rest</w:t></w:r></w:p></w:comment>";
        let source = comment("alpha be", "ta gamma", " rest");
        let written = comment("alpha x ", " gamma", " rest");
        let patch = comment_text_patch(&source, &written, FRAGMENT)
            .unwrap()
            .expect("a text-only change patches");
        let expected = FRAGMENT
            .replace(
                "<w:t>alpha be</w:t>",
                "<w:t xml:space=\"preserve\">alpha x </w:t>",
            )
            .replace(
                "<w:t>ta gamma</w:t>",
                "<w:t xml:space=\"preserve\"> gamma</w:t>",
            );
        assert_eq!(apply_patches(FRAGMENT, patch), expected);
    }

    #[test]
    fn text_patch_replaces_default_space_with_preserve() {
        let fragment = FRAGMENT.replace(
            "<w:t>First identity proofed here. </w:t>",
            "<w:t xml:space=\"default\">alpha</w:t>",
        );
        let source = comment("alpha", "Second identity ", "proofed there.");
        let written = comment(" alpha ", "Second identity ", "proofed there.");
        let patch = comment_text_patch(&source, &written, &fragment)
            .unwrap()
            .expect("a text-only change patches");
        let expected = fragment.replace(
            "<w:t xml:space=\"default\">alpha</w:t>",
            "<w:t xml:space=\"preserve\"> alpha </w:t>",
        );
        assert_eq!(apply_patches(&fragment, patch), expected);
    }

    #[test]
    fn text_patch_keeps_a_fully_covered_run_as_an_empty_wt() {
        const FRAGMENT: &str = "<w:comment w:id=\"0\" w:author=\"Reviewer\" w:date=\"2026-01-05T09:00:00Z\" w:initials=\"R\"><w:p w14:paraId=\"1A2B3C4D\" w14:textId=\"77777777\" w:rsidR=\"00AB\"><w:pPr><w:pStyle w:val=\"CommentText\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>alpha b</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>et</w:t></w:r><w:r><w:t>a gamma</w:t></w:r></w:p></w:comment>";
        let source = comment("alpha b", "et", "a gamma");
        let written = comment("alpha x", "", "a gamma");
        let patch = comment_text_patch(&source, &written, FRAGMENT)
            .unwrap()
            .expect("a text-only change patches");
        let expected = FRAGMENT
            .replace("<w:t>alpha b</w:t>", "<w:t>alpha x</w:t>")
            .replace("<w:t>et</w:t>", "<w:t></w:t>");
        assert_eq!(apply_patches(FRAGMENT, patch), expected);
    }

    #[test]
    fn text_patch_rejects_a_paragraph_formatting_change_with_the_text() {
        let mut source = comment(
            "First identity proofed here. ",
            "Second identity ",
            "proofed there.",
        );
        source.content[0].formatting = Some(ParagraphFormatting {
            alignment: Some("left".to_owned()),
            ..ParagraphFormatting::default()
        });
        let mut written = comment(
            "First identity-proofed here. ",
            "Second identity-proofed",
            " there.",
        );
        written.content[0].formatting = Some(ParagraphFormatting {
            alignment: Some("center".to_owned()),
            ..ParagraphFormatting::default()
        });
        assert_eq!(
            comment_text_patch(&source, &written, FRAGMENT).unwrap(),
            None
        );
        written.content[0].formatting = source.content[0].formatting.clone();
        assert!(
            comment_text_patch(&source, &written, FRAGMENT)
                .unwrap()
                .is_some()
        );
    }

    #[test]
    fn text_patch_rejects_removing_an_explicit_false_paragraph_override() {
        let mut source = comment(
            "First identity proofed here. ",
            "Second identity ",
            "proofed there.",
        );
        source.content[0].formatting = Some(ParagraphFormatting {
            style_id: Some("KeepWithNext".to_owned()),
            keep_next: Some(false),
            ..ParagraphFormatting::default()
        });
        let mut written = comment(
            "First identity-proofed here. ",
            "Second identity-proofed",
            " there.",
        );
        written.content[0].formatting = source.content[0].formatting.clone();
        let fragment = FRAGMENT.replace(
            "<w:pStyle w:val=\"CommentText\"/>",
            "<w:pStyle w:val=\"KeepWithNext\"/><w:keepNext w:val=\"0\"/>",
        );
        assert!(
            comment_text_patch(&source, &written, &fragment)
                .unwrap()
                .is_some()
        );
        written.content[0].formatting.as_mut().unwrap().keep_next = None;
        assert_eq!(
            comment_text_patch(&source, &written, &fragment).unwrap(),
            None
        );
    }

    #[test]
    fn text_patch_rejects_run_property_changes_and_reordering() {
        let source = comment(
            "First identity proofed here. ",
            "Second identity ",
            "proofed there.",
        );
        let mut written = comment(
            "First identity-proofed here. ",
            "Second identity-proofed",
            " there.",
        );
        let mut revised = written.clone();
        if let ParagraphContent::Inline(InlineNode::Run(run)) = &mut revised.content[0].content[2] {
            run.property_changes = Some(vec![RunPropertyChange {
                node_type: RunPropertyChangeType::RunPropertyChange,
                info: PropertyChangeInfo {
                    id: 1.0,
                    author: "Editor".to_owned(),
                    date: None,
                    rsid: None,
                },
                previous_formatting: None,
                current_formatting: run.formatting.clone(),
            }]);
        }
        assert_eq!(
            comment_text_patch(&source, &revised, FRAGMENT).unwrap(),
            None
        );
        written.content[0].content.swap(1, 2);
        assert_eq!(
            comment_text_patch(&source, &written, FRAGMENT).unwrap(),
            None
        );
    }

    #[test]
    fn top_level_additions_without_text_patches_keep_default_companion_writing() {
        let source_comment = "<w:comment w:id=\"0\" w:author=\"Reviewer\"><w:p w14:paraId=\"1A2B3C4D\"><w:r><w:t>old</w:t></w:r></w:p></w:comment>";
        let source = format!(
            "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">{source_comment}</w:comments>"
        );
        for entry in [
            None,
            Some("<w15:commentEx w15:done='0' w15:paraId='1A2B3C4D'/>"),
            Some("<w15:commentEx w15:paraId=\"1A2B3C4D\" w15:done=\"0\" extra=\"keep\"/>"),
        ] {
            let mut parts = vec![("word/comments.xml".to_owned(), source.as_bytes().to_vec())];
            if let Some(entry) = entry {
                parts.push((
                    "word/commentsExtended.xml".to_owned(),
                    format!(
                        "<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\">{entry}</w15:commentsEx>"
                    )
                    .into_bytes(),
                ));
            }
            let seed = "0".repeat(64);
            let original = source_comments(&parts, &seed).unwrap();
            let mut comments = original.clone();
            let mut added = original[0].clone();
            added.id = 1.0;
            added.para_id = Some("1A2B3C4E".to_owned());
            added.content[0].para_id = added.para_id.clone();
            added.content[0].content = vec![run("added", None)];
            added.block_content.clear();
            added.done = Some(true);
            comments.push(added);
            let context = SerializerContext::new(&super::super::s10::SerializerDeterminism {
                seed: seed.clone(),
                now: "2026-01-05T09:00:00.000Z".to_owned(),
            })
            .unwrap();
            let (serialized, mut infos) =
                serialize_comments_with_info(&comments, &mut context.clone());
            let spliced = splice_comments(
                source.as_bytes(),
                &parts,
                &seed,
                &comments,
                &serialized,
                &mut infos,
                context,
            )
            .unwrap();
            assert!(spliced.xml.contains(source_comment));
            assert!(spliced.companions.is_none());
            assert_eq!(infos.len(), 2);
            assert_eq!(infos[1].done, Some(true));
        }
    }

    #[test]
    fn text_patch_with_an_added_resolved_comment_keeps_its_companion_entry() {
        let source = "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w:comment w:id=\"0\" w:author=\"Reviewer\"><w:p w14:paraId=\"1A2B3C4D\"><w:r><w:t>old</w:t></w:r></w:p></w:comment></w:comments>";
        let source_entry = "<w15:commentEx w15:paraId=\"1A2B3C4D\" w15:done=\"1\"/>";
        let mut parts = vec![
            ("word/comments.xml".to_owned(), source.as_bytes().to_vec()),
            (
                "word/commentsExtended.xml".to_owned(),
                format!(
                    "<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\">{source_entry}</w15:commentsEx>"
                )
                .into_bytes(),
            ),
        ];
        let seed = "0".repeat(64);
        let original = source_comments(&parts, &seed).unwrap();
        assert_eq!(original[0].done, Some(true));
        let mut comments = original.clone();
        comments[0].content[0].content = vec![run("new", None)];
        comments[0].block_content.clear();
        let mut added = original[0].clone();
        added.id = 1.0;
        added.para_id = Some("1A2B3C4E".to_owned());
        added.content[0].para_id = added.para_id.clone();
        added.content[0].content = vec![run("added", None)];
        added.block_content.clear();
        added.done = Some(true);
        comments.push(added);
        let context = SerializerContext::new(&super::super::s10::SerializerDeterminism {
            seed: seed.clone(),
            now: "2026-01-05T09:00:00.000Z".to_owned(),
        })
        .unwrap();
        let (serialized, mut infos) = serialize_comments_with_info(&comments, &mut context.clone());
        let spliced = splice_comments(
            source.as_bytes(),
            &parts,
            &seed,
            &comments,
            &serialized,
            &mut infos,
            context,
        )
        .unwrap();
        parts[0].1 = spliced.xml.into_bytes();
        for (path, xml) in spliced.companions.unwrap() {
            if let Some((_, bytes)) = parts.iter_mut().find(|(name, _)| name == &path) {
                *bytes = xml.into_bytes();
            } else {
                parts.push((path, xml.into_bytes()));
            }
        }
        let (_, extended) = find_part(&parts, "word/commentsExtended.xml").unwrap();
        assert!(
            std::str::from_utf8(extended)
                .unwrap()
                .contains(source_entry)
        );
        let reopened = source_comments(&parts, &seed).unwrap();
        assert_eq!(reopened.len(), 2);
        assert_eq!(
            model_paragraph_text_parts(&reopened[0].content[0]).unwrap(),
            ["new"]
        );
        assert_eq!(reopened[0].done, Some(true));
        assert_eq!(reopened[1].done, Some(true));
    }

    #[test]
    fn text_patch_ignores_unchanged_and_rejects_formatting_changes() {
        let source = comment(
            "First identity proofed here. ",
            "Second identity ",
            "proofed there.",
        );
        assert_eq!(
            comment_text_patch(&source, &source.clone(), FRAGMENT).unwrap(),
            None
        );
        let mut written = source.clone();
        if let ParagraphContent::Inline(InlineNode::Run(run)) = &mut written.content[0].content[2] {
            run.formatting = Some(TextFormatting {
                italic: Some(true),
                ..TextFormatting::default()
            });
            run.content = vec![RunContent::Text {
                text: "Second identity-proofed".to_owned(),
                preserve_space: None,
            }];
        }
        assert_eq!(
            comment_text_patch(&source, &written, FRAGMENT).unwrap(),
            None
        );
    }
}
