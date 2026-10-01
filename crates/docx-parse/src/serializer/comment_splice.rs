use std::collections::{HashMap, HashSet};
use std::ops::Range;

use indexmap::IndexMap;
use serde::Serialize;
use serde_json::Value;

use crate::comments::Comment;
use crate::paragraph::HexIdAllocator;
use crate::paragraph_identity::{attribute, tags, unescaped};
use crate::relationships::parse_relationships;
use crate::s8::{find_part, parse_comment_part};
use crate::styles::StyleMap;
use crate::xml::{ParseBudget, ParseError, ParseLimits, parse_javascript_integer_prefix, parse_xml};

use super::parts::CommentParaInfo;
use super::xml_writer::js_number;

pub(super) fn splice_comments(
    source: &[u8],
    parts: &[(String, Vec<u8>)],
    seed: &str,
    comments: &[Comment],
    serialized: &str,
    infos: &mut Vec<CommentParaInfo>,
) -> Result<String, ParseError> {
    let source = std::str::from_utf8(source).map_err(|error| error_xml(error.to_string()))?;
    let source = expand_empty_root(source)?;
    let original = source_comments(parts, seed)?;
    let (spans, end) = comment_spans(&source)?;
    let (written, _) = comment_spans(serialized)?;
    let fragments: HashMap<_, _> = written
        .iter()
        .map(|(id, span)| (id.as_str(), &serialized[span.clone()]))
        .collect();
    let models: HashMap<_, _> = comments
        .iter()
        .map(|comment| (js_number(comment.id), comment))
        .collect();
    let originals: HashMap<_, _> = original
        .iter()
        .map(|comment| (js_number(comment.id), comment))
        .collect();
    if models.len() != comments.len() || originals.len() != original.len() {
        return Err(error_xml("duplicate comment ids"));
    }
    let mut kept = HashSet::new();
    let mut without_ids = HashSet::new();
    let mut output = String::new();
    let mut cursor = 0;
    let mut rewritten = false;
    for (id, span) in spans {
        output.push_str(&source[cursor..span.start]);
        if let Some(comment) = models.get(&id) {
            kept.insert(id.clone());
            if let Some(original) = originals.get(&id)
                && can_replay(original, comment, &source[span.clone()])?
            {
                output.push_str(&source[span.clone()]);
                if let Some(para_id) = &original.para_id {
                    if let Some(info) = infos
                        .iter_mut()
                        .find(|info| js_number(info.comment_id) == id)
                    {
                        info.last_para_id.clone_from(para_id);
                    }
                } else {
                    without_ids.insert(id.clone());
                }
            } else {
                output.push_str(
                    fragments
                        .get(id.as_str())
                        .ok_or_else(|| error_xml("missing written comment"))?,
                );
                rewritten = true;
            }
        } else if !originals.contains_key(&id) {
            output.push_str(&source[span.clone()]);
        }
        cursor = span.end;
    }
    output.push_str(&source[cursor..end]);
    for (id, span) in written {
        if !kept.contains(&id) {
            output.push_str(&serialized[span]);
            rewritten = true;
        }
    }
    output.push_str(&source[end..]);
    infos.retain(|info| !without_ids.contains(&js_number(info.comment_id)));
    if rewritten {
        output = with_writer_namespaces(output)?;
    }
    Ok(output)
}

fn source_metadata_changed(source: &Comment, written: &Comment) -> bool {
    (source.done == Some(true)) != (written.done == Some(true))
        || source.parent_id != written.parent_id
}

fn can_replay(source: &Comment, written: &Comment, fragment: &str) -> Result<bool, ParseError> {
    if !unchanged(source, written)? {
        return Ok(false);
    }
    if source_metadata_changed(source, written) {
        let fragment_tags = tags(fragment).ok_or_else(|| error_xml("invalid comment fragment"))?;
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

fn comment_spans(xml: &str) -> Result<(Vec<(String, Range<usize>)>, usize), ParseError> {
    if xml.is_empty() {
        return Ok((Vec::new(), 0));
    }
    let limits = ParseLimits::default();
    let document = parse_xml(xml.as_bytes(), "word/comments.xml", &mut ParseBudget::new(&limits))?;
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
                let id = ids.next().ok_or_else(|| error_xml("comment span mismatch"))?;
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
    Ok((spans, end.ok_or_else(|| error_xml("missing comments end"))?))
}

fn expand_empty_root(xml: &str) -> Result<String, ParseError> {
    let all_tags = tags(xml).ok_or_else(|| error_xml("invalid comments root"))?;
    let root = all_tags
        .first()
        .ok_or_else(|| error_xml("missing comments root"))?;
    if !root.empty {
        return Ok(xml.to_owned());
    }
    let mut output = xml.to_owned();
    let close = xml[root.range.clone()]
        .rfind('/')
        .ok_or_else(|| error_xml("invalid empty comments root"))?;
    output.replace_range(
        root.range.start + close..root.range.end,
        &format!("></{}>", root.name),
    );
    Ok(output)
}

fn with_writer_namespaces(mut xml: String) -> Result<String, ParseError> {
    let all_tags = tags(&xml).ok_or_else(|| error_xml("invalid comments root"))?;
    let root = all_tags.first().ok_or_else(|| error_xml("missing comments root"))?;
    let mut added = String::new();
    for (prefix, uri) in [
        ("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        ("w14", "http://schemas.microsoft.com/office/word/2010/wordml"),
    ] {
        let name = format!("xmlns:{prefix}");
        match attribute(root, &name).and_then(|range| unescaped(&xml, range)) {
            Some(bound) if bound != uri => return Err(error_xml("conflicting comments namespace")),
            Some(_) => {}
            None => added.push_str(&format!(" {name}=\"{uri}\"")),
        }
    }
    let at = root.range.end - 1;
    xml.insert_str(at, &added);
    Ok(xml)
}

fn error_xml(message: impl Into<String>) -> ParseError {
    ParseError::Canonical(format!("comments splice: {}", message.into()))
}
