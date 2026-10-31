use serde_json::{Value, json};

use super::{EngineError, EngineResult};
use crate::{
    ParagraphAnchor, ParagraphIdDiagnostic, ParagraphIdOrigin, ParagraphOrigin, ParagraphRef,
    PersistedParagraphIds, SourceParagraphRef, SourceStory, SourceStoryKind,
};

/// Parses the render bridge's host context from JSON:
/// `{ "themeColors": {name: hex}, "defaultTabStopTwips": number|null,
/// "pageContentHeight": number|null, "numericIds": {yrsId: number},
/// "showHiddenText": bool, "showTableGridlines": bool,
/// "defaultParagraphStyleId": string,
/// "paragraphStyleFontSizes": {styleId: pt}, "docDefaultFontSize": number,
/// "revisionPreview": {revisionId: "accepted"|"rejected"} }`.
pub(crate) fn parse_render_env(env_json: &str) -> EngineResult<crate::bridge::RenderEnv> {
    let value: Value = serde_json::from_str(env_json).map_err(EngineError::message)?;
    let mut env = crate::bridge::RenderEnv::default();
    env.font_theme = value
        .get("fontTheme")
        .filter(|theme| !theme.is_null())
        .cloned()
        .map(crate::bridge::parse_font_theme)
        .transpose()
        .map_err(EngineError::message)?;
    if let Some(Value::Array(ids)) = value.get("tocStyleIds") {
        env.toc_style_ids
            .extend(ids.iter().filter_map(Value::as_str).map(str::to_owned));
    }
    if let Some(Value::Object(colors)) = value.get("themeColors") {
        for (key, entry) in colors {
            if let Some(hex) = entry.as_str() {
                env.theme_colors.insert(key.clone(), hex.to_owned());
            }
        }
    }
    env.default_tab_stop_twips = value.get("defaultTabStopTwips").and_then(Value::as_f64);
    env.page_content_height = value.get("pageContentHeight").and_then(Value::as_f64);
    env.default_paragraph_style_id = value
        .get("defaultParagraphStyleId")
        .and_then(Value::as_str)
        .filter(|style| !style.is_empty())
        .map(str::to_owned);
    if let Some(Value::Object(sizes)) = value.get("paragraphStyleFontSizes") {
        for (key, entry) in sizes {
            if let Some(size) = entry.as_f64() {
                env.paragraph_style_font_sizes.insert(key.clone(), size);
            }
        }
    }
    if let Some(Value::Object(sizes)) = value.get("hyperlinkStyleFontSizes") {
        for (key, entry) in sizes {
            if let Some(size) = entry.as_f64() {
                env.hyperlink_style_font_sizes.insert(key.clone(), size);
            }
        }
    }
    env.doc_default_font_size = value.get("docDefaultFontSize").and_then(Value::as_f64);
    env.show_hidden_text = value
        .get("showHiddenText")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    env.show_table_gridlines = value
        .get("showTableGridlines")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    env.media_tokens = value
        .get("mediaTokens")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    if let Some(Value::Object(ids)) = value.get("numericIds") {
        for (key, entry) in ids {
            if let Some(id) = entry.as_f64() {
                env.numeric_ids.insert(key.clone(), id);
            }
        }
    }
    if let Some(preview) = value.get("revisionPreview") {
        env.revision_preview = crate::bridge::RenderEnv::parse_revision_preview(preview);
    }
    Ok(env)
}

pub(crate) fn source_story_kind(kind: SourceStoryKind) -> &'static str {
    match kind {
        SourceStoryKind::Body => "body",
        SourceStoryKind::Header => "header",
        SourceStoryKind::Footer => "footer",
        SourceStoryKind::Footnote => "footnote",
        SourceStoryKind::Endnote => "endnote",
        SourceStoryKind::Comment => "comment",
    }
}

pub(crate) fn source_story_json(story: &SourceStory) -> Value {
    let mut value = json!({
        "partUri": story.part_uri,
        "kind": source_story_kind(story.kind),
    });
    if let Some(item_id) = &story.item_id {
        value["itemId"] = json!(item_id);
    }
    value
}

pub(crate) fn source_anchor_json(source: &SourceParagraphRef) -> Value {
    json!({
        "kind": "source",
        "packageSha256": source.package_sha256,
        "partUri": source.part_uri,
        "paragraphOrdinal": source.paragraph_ordinal,
    })
}

pub(crate) fn paragraph_anchor_json(session_id: &str, paragraph: &ParagraphRef) -> Value {
    match paragraph {
        ParagraphRef::Session { story, para_id } => {
            json!({ "kind": "session", "sessionId": session_id, "story": story, "paraId": para_id })
        }
        ParagraphRef::Source(source) => source_anchor_json(source),
    }
}

pub(crate) fn persisted_anchor_json(story: Option<&SourceStory>, para_id: Option<&str>) -> Value {
    match (story, para_id) {
        (Some(story), Some(para_id)) => json!({
            "kind": "persisted",
            "story": source_story_json(story),
            "paraId": para_id,
        }),
        _ => Value::Null,
    }
}

pub(crate) fn required_str<'a>(value: &'a Value, key: &str) -> EngineResult<&'a str> {
    value
        .get(key)
        .and_then(Value::as_str)
        .ok_or_else(|| EngineError::message(format!("anchor requires a string {key:?}")))
}

pub(crate) fn parse_source_story(value: Option<&Value>) -> EngineResult<SourceStory> {
    let value = value.ok_or_else(|| EngineError::message("anchor requires a \"story\" object"))?;
    let kind = match required_str(value, "kind")? {
        "body" => SourceStoryKind::Body,
        "header" => SourceStoryKind::Header,
        "footer" => SourceStoryKind::Footer,
        "footnote" => SourceStoryKind::Footnote,
        "endnote" => SourceStoryKind::Endnote,
        "comment" => SourceStoryKind::Comment,
        other => {
            return Err(EngineError::message(format!(
                "unknown source story kind {other:?}"
            )));
        }
    };
    let item_id = match kind {
        SourceStoryKind::Footnote | SourceStoryKind::Endnote | SourceStoryKind::Comment => {
            Some(required_str(value, "itemId")?.to_owned())
        }
        _ => None,
    };
    Ok(SourceStory {
        part_uri: required_str(value, "partUri")?.to_owned(),
        kind,
        item_id,
    })
}

pub(crate) fn parse_paragraph_anchor(anchor_json: &str) -> EngineResult<ParagraphAnchor> {
    let value: Value = serde_json::from_str(anchor_json).map_err(EngineError::message)?;
    match required_str(&value, "kind")? {
        "session" => Ok(ParagraphAnchor::Session {
            session_id: required_str(&value, "sessionId")?.to_owned(),
            story: required_str(&value, "story")?.to_owned(),
            para_id: required_str(&value, "paraId")?.to_owned(),
        }),
        "source" => Ok(ParagraphAnchor::Source(SourceParagraphRef {
            package_sha256: required_str(&value, "packageSha256")?.to_owned(),
            part_uri: required_str(&value, "partUri")?.to_owned(),
            paragraph_ordinal: value
                .get("paragraphOrdinal")
                .and_then(Value::as_u64)
                .and_then(|ordinal| u32::try_from(ordinal).ok())
                .ok_or_else(|| {
                    EngineError::message("anchor requires a non-negative \"paragraphOrdinal\"")
                })?,
        })),
        "persisted" => Ok(ParagraphAnchor::Persisted {
            story: parse_source_story(value.get("story"))?,
            para_id: required_str(&value, "paraId")?.to_owned(),
        }),
        other => Err(EngineError::message(format!(
            "unknown anchor kind {other:?}"
        ))),
    }
}

pub(crate) fn origin_json(origin: Option<ParagraphIdOrigin>) -> Value {
    origin.map_or(Value::Null, |origin| json!(origin.as_str()))
}

pub(crate) fn paragraph_origin_json(origin: ParagraphOrigin) -> &'static str {
    match origin {
        ParagraphOrigin::Source => "source",
        ParagraphOrigin::Authored => "authored",
        ParagraphOrigin::Synthetic => "synthetic",
    }
}

pub(crate) fn persisted_receipt_json(session_id: &str, persisted: &PersistedParagraphIds) -> Value {
    let assignments: Vec<Value> = persisted
        .assignments
        .iter()
        .map(|assignment| {
            json!({
                "paragraph": paragraph_anchor_json(session_id, &assignment.paragraph),
                "replacedParaId": assignment.replaced_para_id,
                "previousOoxmlParaId": assignment.previous_ooxml_para_id,
                "ooxmlParaId": assignment.ooxml_para_id,
                "idOrigin": assignment.origin.as_str(),
                "persisted": persisted_anchor_json(
                    assignment.source_story.as_ref(),
                    Some(&assignment.ooxml_para_id),
                ),
            })
        })
        .collect();
    let diagnostics: Vec<Value> = persisted
        .diagnostics
        .iter()
        .map(|diagnostic| match diagnostic {
            ParagraphIdDiagnostic::ConflictingSavedIds {
                ooxml_para_id,
                paragraphs,
            } => json!({
                "kind": "conflicting-saved-ids",
                "ooxmlParaId": ooxml_para_id,
                "paragraphs": paragraphs
                    .iter()
                    .map(|paragraph| paragraph_anchor_json(session_id, paragraph))
                    .collect::<Vec<_>>(),
            }),
            ParagraphIdDiagnostic::NoSourcePackage => json!({ "kind": "no-source-package" }),
        })
        .collect();
    json!({ "status": "applied", "assignments": assignments, "diagnostics": diagnostics })
}
