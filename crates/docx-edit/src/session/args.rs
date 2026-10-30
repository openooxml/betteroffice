use std::collections::{BTreeMap, HashMap};

use serde_json::{Value, json};
use yrs::{Any, Transact};

use super::{EngineError, EngineResult};
use crate::segments::SegKind;
use crate::{
    ChangeTarget, ColorPatch, EditCtx, EditingDoc, FontFamilyPatch, InlineFormatDelta,
    ParaAttrDelta, Patch, RawOp, SimpleFormat, StoryRange, TabStop, TriState,
};

/// `{"revisionId", "range"}` for a text edit's receipt.
pub(crate) fn text_receipt_json(receipt: crate::Receipt) -> String {
    let range = receipt.range.map(|range| {
        json!({
            "story": range.start.story,
            "start": { "paraId": range.start.para, "offset": range.start.offset },
            "end": { "paraId": range.end.para, "offset": range.end.offset },
        })
    });
    json!({
        "revisionId": receipt.revision_ids.into_iter().next(),
        "range": range,
    })
    .to_string()
}

/// Builds the op [`EditCtx`]. Suggesting mode crosses the boundary as an
/// optional `(name, date)` pair — both-or-neither; a plain local edit uses an
/// empty author (author/date are stamped only in suggesting mode).
pub(crate) fn edit_ctx(name: Option<String>, date: Option<String>) -> EngineResult<EditCtx> {
    match (name, date) {
        (None, None) => Ok(EditCtx::local(String::new(), String::new())),
        (Some(name), Some(date)) => Ok(EditCtx::local(name, date).suggesting()),
        _ => Err(EngineError::message(
            "suggesting requires both an author name and an ISO date",
        )),
    }
}

pub(crate) fn parse_any_object(json: &str, label: &str) -> EngineResult<HashMap<String, Any>> {
    match Any::from_json(json).map_err(EngineError::message)? {
        Any::Map(value) => Ok(value.as_ref().clone()),
        _ => Err(EngineError::message(format!(
            "{label} must be a JSON object"
        ))),
    }
}

pub(crate) struct ParaSpan {
    /// Story index of the paragraph's first unit (after the previous pilcrow).
    pub(crate) start: u32,
    /// Story index of the paragraph's own pilcrow embed.
    pub(crate) pilcrow: u32,
}

pub(crate) struct IndexedLoc {
    pub(crate) para_id: String,
    pub(crate) offset: u32,
    pub(crate) node_offset: u32,
}

#[derive(Clone, Copy)]
pub(crate) enum DeleteDirection {
    Backward,
    Forward,
}

#[derive(Clone, Copy)]
pub(crate) enum AdjacentStoryUnit {
    Content(u32),
    Pilcrow,
    Table,
}

/// Resolves a paragraph to its story span via the committed paragraph index.
/// Story-scoped: a `para_id` that lives in another story is "not found".
pub(crate) fn find_para_span(
    doc: &EditingDoc,
    story: &str,
    para_id: &str,
) -> EngineResult<ParaSpan> {
    doc.paragraph_index(story)
        .map_err(EngineError::message)?
        .para_span(para_id)
        .map(|(start, pilcrow)| ParaSpan { start, pilcrow })
        .ok_or_else(|| {
            EngineError::message(format!(
                "paragraph {para_id:?} was not found in story {story:?}"
            ))
        })
}

/// `Loc { story, paraId, offset }` -> transient story-global index.
pub(crate) fn loc_index(
    doc: &EditingDoc,
    story: &str,
    para_id: &str,
    offset: u32,
) -> EngineResult<u32> {
    let span = find_para_span(doc, story, para_id)?;
    let para_len = span.pilcrow - span.start;
    if offset > para_len {
        return Err(EngineError::message(format!(
            "offset {offset} exceeds the length {para_len} of paragraph {para_id:?}"
        )));
    }
    Ok(span.start + offset)
}

pub(crate) fn comment_ranges(doc: &EditingDoc, ranges_json: &str) -> EngineResult<Vec<StoryRange>> {
    let value: Value = serde_json::from_str(ranges_json).map_err(EngineError::message)?;
    let entries = value
        .as_array()
        .ok_or_else(|| EngineError::message("expected an array of comment ranges"))?;
    let mut ranges = Vec::with_capacity(entries.len());
    for entry in entries {
        let get_str = |key: &str| {
            entry.get(key).and_then(Value::as_str).ok_or_else(|| {
                EngineError::message(format!("a comment range requires a string {key:?}"))
            })
        };
        let get_offset = |key: &str| {
            entry
                .get(key)
                .and_then(Value::as_u64)
                .ok_or_else(|| {
                    EngineError::message(format!("a comment range requires a non-negative {key:?}"))
                })
                .and_then(|offset| {
                    u32::try_from(offset)
                        .map_err(|_| EngineError::message("comment offset exceeds u32"))
                })
        };
        let story = get_str("story")?;
        let start = loc_index(
            doc,
            story,
            get_str("startPara")?,
            get_offset("startOffset")?,
        )?;
        let end = loc_index(doc, story, get_str("endPara")?, get_offset("endOffset")?)?;
        ranges.push(StoryRange::new(story, start, end));
    }
    Ok(ranges)
}

/// Transient story-global index -> public paragraph-keyed location. Sticky
/// awareness positions resolve to story indices; the JS facade never exposes
/// that internal coordinate system.
pub(crate) fn index_loc(doc: &EditingDoc, story: &str, index: u32) -> EngineResult<IndexedLoc> {
    let paragraphs = doc.paragraph_index(story).map_err(EngineError::message)?;
    let para = paragraphs.para_at(index).ok_or_else(|| {
        EngineError::message(format!(
            "selection index {index} does not resolve in story {story:?}"
        ))
    })?;
    Ok(IndexedLoc {
        para_id: para.para_id.to_string(),
        offset: index.saturating_sub(para.start),
        node_offset: index.saturating_sub(para.node_start),
    })
}

pub(crate) fn adjacent_story_unit(
    doc: &EditingDoc,
    story: &str,
    index: u32,
    direction: DeleteDirection,
) -> EngineResult<Option<AdjacentStoryUnit>> {
    let segments = doc.segment_index(story).map_err(EngineError::message)?;
    let position = match direction {
        DeleteDirection::Backward => index.checked_sub(1),
        DeleteDirection::Forward => Some(index),
    };
    let Some(segment) = position.and_then(|pos| segments.segment_at(pos)) else {
        return Ok(None);
    };
    Ok(Some(match &segment.kind {
        SegKind::Text(text) => {
            let units: Vec<u16> = text.encode_utf16().collect();
            let relative = (index - segment.start) as usize;
            let width = match direction {
                DeleteDirection::Backward
                    if relative > 1
                        && (0xdc00..=0xdfff).contains(&units[relative - 1])
                        && (0xd800..=0xdbff).contains(&units[relative - 2]) =>
                {
                    2
                }
                DeleteDirection::Forward
                    if relative + 1 < units.len()
                        && (0xd800..=0xdbff).contains(&units[relative])
                        && (0xdc00..=0xdfff).contains(&units[relative + 1]) =>
                {
                    2
                }
                _ => 1,
            };
            AdjacentStoryUnit::Content(width)
        }
        SegKind::Pilcrow => AdjacentStoryUnit::Pilcrow,
        SegKind::Embed if table_embed_at(doc, story, segment.start)? => AdjacentStoryUnit::Table,
        SegKind::Embed => AdjacentStoryUnit::Content(1),
    }))
}

pub(crate) fn table_embed_at(doc: &EditingDoc, story: &str, index: u32) -> EngineResult<bool> {
    let txn = doc.yrs_doc().transact();
    let text = crate::story_ref(&txn, story).map_err(EngineError::message)?;
    let chunks = crate::ops::snapshot_range(&text, &txn, index, index.saturating_add(1));
    Ok(chunks.first().is_some_and(|chunk| {
        matches!(
            &chunk.kind,
            crate::ops::ChunkKind::Embed(Some(map))
                if crate::map_string(map, &txn, crate::KIND_KEY).as_deref() == Some("table")
        )
    }))
}

pub(crate) fn embed_at(doc: &EditingDoc, story: &str, index: u32) -> EngineResult<bool> {
    let segments = doc.segment_index(story).map_err(EngineError::message)?;
    Ok(segments
        .segment_at(index)
        .is_some_and(|segment| matches!(segment.kind, SegKind::Embed)))
}

/// Applies one boundary mark to `range`. `mark_json`:
/// `{"type":"bold"|"italic"|"underline"|"strike"|"superscript"|"subscript"} |
/// {"type":"fontFamily"|"color","value":string} |
/// {"type":"fontSize","value":number}`. The six simple types toggle; font,
/// size and color set.
pub(crate) fn apply_mark(doc: &EditingDoc, range: StoryRange, mark_json: &str) -> EngineResult<()> {
    let value: Value = serde_json::from_str(mark_json).map_err(EngineError::message)?;
    let kind = value
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| EngineError::message("mark JSON requires a string \"type\""))?;
    // Formatting is never itself a tracked change, so the caret's suggest mode
    // is irrelevant here.
    let ctx = EditCtx::local(String::new(), String::new());
    let simple = match kind {
        "bold" => Some(SimpleFormat::Bold),
        "italic" => Some(SimpleFormat::Italic),
        "underline" => Some(SimpleFormat::Underline),
        "strike" => Some(SimpleFormat::Strike),
        "superscript" => Some(SimpleFormat::Superscript),
        "subscript" => Some(SimpleFormat::Subscript),
        _ => None,
    };
    if let Some(format) = simple {
        doc.toggle_format(&ctx, range, format)
            .map_err(EngineError::message)?;
        return Ok(());
    }
    let delta = match kind {
        "fontFamily" => {
            let family = value.get("value").and_then(Value::as_str).ok_or_else(|| {
                EngineError::message("mark type \"fontFamily\" requires a string \"value\"")
            })?;
            InlineFormatDelta {
                font_family: Patch::Set(FontFamilyPatch {
                    ascii: family.to_owned(),
                    h_ansi: None,
                }),
                ..Default::default()
            }
        }
        "fontSize" => {
            let size = value.get("value").and_then(Value::as_f64).ok_or_else(|| {
                EngineError::message("mark type \"fontSize\" requires a numeric \"value\"")
            })?;
            InlineFormatDelta {
                font_size: Patch::Set(size),
                ..Default::default()
            }
        }
        "color" => {
            let color = value.get("value").and_then(Value::as_str).ok_or_else(|| {
                EngineError::message("mark type \"color\" requires a string \"value\"")
            })?;
            InlineFormatDelta {
                color: Patch::Set(ColorPatch::Rgb(color.to_owned())),
                ..Default::default()
            }
        }
        other => return Err(EngineError::message(format!("unknown mark type {other:?}"))),
    };
    doc.format_range(&ctx, range, &delta)
        .map_err(EngineError::message)?;
    Ok(())
}

/// Decodes a tri-state inline-formatting delta. An omitted key is
/// [`Patch::Keep`], `null` is [`Patch::Clear`], any typed value is
/// [`Patch::Set`]. Boolean `false` also clears.
pub(crate) fn parse_inline_format_delta(delta_json: &str) -> EngineResult<InlineFormatDelta> {
    let value: Value = serde_json::from_str(delta_json).map_err(EngineError::message)?;
    let object = value
        .as_object()
        .ok_or_else(|| EngineError::message("format_range expects a JSON object delta"))?;
    crate::format::parse_inline_format_delta(object).map_err(EngineError::message)
}

/// Decodes the public facade's tri-state paragraph-property delta. Typed
/// fields lower to [`ParaAttrDelta`]; list/render metadata and other passive
/// pPr fields use its `other` bag. Omitted fields are kept and `null` clears.
pub(crate) fn parse_para_attr_delta(attrs_json: &str) -> EngineResult<ParaAttrDelta> {
    let value: Value = serde_json::from_str(attrs_json).map_err(EngineError::message)?;
    let object = value
        .as_object()
        .ok_or_else(|| EngineError::message("set_paragraph_attrs expects a JSON object"))?;

    let string_patch = |key: &str| -> EngineResult<Patch<String>> {
        match object.get(key) {
            None => Ok(Patch::Keep),
            Some(Value::Null) => Ok(Patch::Clear),
            Some(Value::String(value)) => Ok(Patch::Set(value.clone())),
            Some(_) => Err(EngineError::message(format!(
                "paragraph attribute {key:?} must be a string or null"
            ))),
        }
    };
    let number_patch = |key: &str| -> EngineResult<Patch<f64>> {
        match object.get(key) {
            None => Ok(Patch::Keep),
            Some(Value::Null) => Ok(Patch::Clear),
            Some(Value::Number(value)) => Ok(Patch::Set(value.as_f64().ok_or_else(|| {
                EngineError::message(format!("paragraph attribute {key:?} must be finite"))
            })?)),
            Some(_) => Err(EngineError::message(format!(
                "paragraph attribute {key:?} must be a number or null"
            ))),
        }
    };
    let bool_patch = |key: &str| -> EngineResult<Patch<bool>> {
        match object.get(key) {
            None => Ok(Patch::Keep),
            Some(Value::Null) => Ok(Patch::Clear),
            Some(Value::Bool(value)) => Ok(Patch::Set(*value)),
            Some(_) => Err(EngineError::message(format!(
                "paragraph attribute {key:?} must be a boolean or null"
            ))),
        }
    };

    let tabs = match object.get("tabs") {
        None => Patch::Keep,
        Some(Value::Null) => Patch::Clear,
        Some(Value::Array(values)) => {
            let mut stops = Vec::with_capacity(values.len());
            for value in values {
                let stop = value
                    .as_object()
                    .ok_or_else(|| EngineError::message("each paragraph tab must be an object"))?;
                let pos = stop
                    .get("position")
                    .and_then(Value::as_f64)
                    .ok_or_else(|| {
                        EngineError::message("a paragraph tab requires numeric \"position\"")
                    })?;
                let alignment = stop
                    .get("alignment")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        EngineError::message("a paragraph tab requires string \"alignment\"")
                    })?;
                let leader = stop
                    .get("leader")
                    .map(|entry| {
                        entry.as_str().map(str::to_owned).ok_or_else(|| {
                            EngineError::message("a paragraph tab \"leader\" must be a string")
                        })
                    })
                    .transpose()?;
                stops.push(TabStop {
                    pos,
                    alignment: alignment.to_owned(),
                    leader,
                });
            }
            Patch::Set(stops)
        }
        Some(_) => {
            return Err(EngineError::message(
                "paragraph attribute \"tabs\" must be an array or null",
            ));
        }
    };

    let default_text_formatting = match object.get("defaultTextFormatting") {
        None => Patch::Keep,
        Some(Value::Null) => Patch::Clear,
        Some(Value::Object(values)) => {
            let mut formatting = BTreeMap::new();
            for (key, value) in values {
                formatting.insert(key.clone(), json_to_any(value)?);
            }
            Patch::Set(formatting)
        }
        Some(_) => {
            return Err(EngineError::message(
                "paragraph attribute \"defaultTextFormatting\" must be an object or null",
            ));
        }
    };

    const TYPED_KEYS: [&str; 13] = [
        "alignment",
        "lineSpacing",
        "lineSpacingRule",
        "spaceBefore",
        "spaceAfter",
        "indentLeft",
        "indentRight",
        "indentFirstLine",
        "hangingIndent",
        "bidi",
        "tabs",
        "defaultTextFormatting",
        "other",
    ];
    let mut other = BTreeMap::new();
    for (key, value) in object {
        if TYPED_KEYS.contains(&key.as_str()) {
            continue;
        }
        other.insert(
            key.clone(),
            if value.is_null() {
                None
            } else {
                Some(json_to_any(value)?)
            },
        );
    }
    if let Some(value) = object.get("other") {
        let entries = value.as_object().ok_or_else(|| {
            EngineError::message("paragraph attribute \"other\" must be an object")
        })?;
        for (key, value) in entries {
            other.insert(
                key.clone(),
                if value.is_null() {
                    None
                } else {
                    Some(json_to_any(value)?)
                },
            );
        }
    }

    Ok(ParaAttrDelta {
        alignment: string_patch("alignment")?,
        line_spacing: number_patch("lineSpacing")?,
        line_spacing_rule: string_patch("lineSpacingRule")?,
        space_before: number_patch("spaceBefore")?,
        space_after: number_patch("spaceAfter")?,
        indent_left: number_patch("indentLeft")?,
        indent_right: number_patch("indentRight")?,
        indent_first_line: number_patch("indentFirstLine")?,
        hanging_indent: number_patch("hangingIndent")?,
        bidi: bool_patch("bidi")?,
        tabs,
        default_text_formatting,
        other,
    })
}

pub(crate) fn attrs_value(attrs: &std::collections::BTreeMap<String, Any>) -> EngineResult<Value> {
    serde_json::to_string(attrs)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .ok_or_else(|| EngineError::message("serialize attrs"))
}

pub(crate) fn tri_state_value(state: TriState) -> Value {
    match state {
        TriState::On => Value::Bool(true),
        TriState::Off => Value::Bool(false),
        TriState::Mixed => Value::String("mixed".to_owned()),
    }
}

pub(crate) fn json_to_any(value: &Value) -> EngineResult<Any> {
    Any::from_json(&value.to_string()).map_err(EngineError::message)
}

/// Image geometry JSON → payload entries, with a nested `"other"` object
/// flattened alongside the named fields.
pub(crate) fn image_geometry_entries(geometry_json: &str) -> EngineResult<Vec<(String, Any)>> {
    let value: Value = serde_json::from_str(geometry_json).map_err(EngineError::message)?;
    let object = value
        .as_object()
        .ok_or_else(|| EngineError::message("image geometry must be a JSON object"))?;
    let mut entries = Vec::new();
    for (key, value) in object {
        if key == "other" {
            let other = value.as_object().ok_or_else(|| {
                EngineError::message("image geometry \"other\" must be an object")
            })?;
            for (other_key, other_value) in other {
                entries.push((other_key.clone(), json_to_any(other_value)?));
            }
        } else {
            entries.push((key.clone(), json_to_any(value)?));
        }
    }
    Ok(entries)
}

/// A JSON object → yrs text/format attributes (`Arc<str>` keys, `Any` values).
pub(crate) fn parse_attrs(value: Option<&Value>) -> EngineResult<yrs::types::Attrs> {
    let mut attrs = yrs::types::Attrs::new();
    if let Some(Value::Object(map)) = value {
        for (key, entry) in map {
            attrs.insert(std::sync::Arc::from(key.as_str()), json_to_any(entry)?);
        }
    }
    Ok(attrs)
}

/// A JSON object → ordered `(key, Any)` embed payload entries.
pub(crate) fn parse_payload(value: Option<&Value>) -> EngineResult<Vec<(String, Any)>> {
    let mut out = Vec::new();
    if let Some(Value::Object(map)) = value {
        for (key, entry) in map {
            out.push((key.clone(), json_to_any(entry)?));
        }
    }
    Ok(out)
}

/// Parses one `{ "op", "index", … }` raw operation.
pub(crate) fn parse_raw_op(value: &Value) -> EngineResult<RawOp> {
    let op = value
        .get("op")
        .and_then(Value::as_str)
        .ok_or_else(|| EngineError::message("a raw op requires a string \"op\""))?;
    let index = value
        .get("index")
        .and_then(Value::as_u64)
        .map(|i| i as u32)
        .ok_or_else(|| EngineError::message("a raw op requires a non-negative \"index\""));
    let u32_field = |key: &str| {
        value
            .get(key)
            .and_then(Value::as_u64)
            .map(|v| v as u32)
            .ok_or_else(|| {
                EngineError::message(format!("op {op:?} requires a non-negative {key:?}"))
            })
    };
    match op {
        "insert" => Ok(RawOp::Insert {
            index: index?,
            text: value
                .get("text")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .to_owned(),
            attrs: parse_attrs(value.get("attrs"))?,
        }),
        "delete" => Ok(RawOp::Delete {
            index: index?,
            len: u32_field("len")?,
        }),
        "format" => Ok(RawOp::Format {
            index: index?,
            len: u32_field("len")?,
            attrs: parse_attrs(value.get("attrs"))?,
        }),
        "insertEmbed" => Ok(RawOp::InsertEmbed {
            index: index?,
            kind: value
                .get("kind")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .to_owned(),
            payload: parse_payload(value.get("payload"))?,
            attrs: parse_attrs(value.get("attrs"))?,
        }),
        "setEmbedAttr" => Ok(RawOp::SetEmbedAttr {
            index: index?,
            key: value
                .get("key")
                .and_then(Value::as_str)
                .ok_or_else(|| EngineError::message("setEmbedAttr requires a string \"key\""))?
                .to_owned(),
            value: json_to_any(value.get("value").unwrap_or(&Value::Null))?,
        }),
        "setComment" => {
            let ranges_value = value
                .get("ranges")
                .and_then(Value::as_array)
                .ok_or_else(|| EngineError::message("setComment requires an array \"ranges\""))?;
            let mut ranges = Vec::with_capacity(ranges_value.len());
            for entry in ranges_value {
                let pair = entry
                    .as_array()
                    .filter(|pair| pair.len() == 2)
                    .ok_or_else(|| {
                        EngineError::message("a setComment range must be a [start, end] pair")
                    })?;
                let bound = |index: usize| {
                    pair[index]
                        .as_u64()
                        .map(|bound| bound as u32)
                        .ok_or_else(|| {
                            EngineError::message(
                                "a setComment range bound must be a non-negative integer",
                            )
                        })
                };
                ranges.push((bound(0)?, bound(1)?));
            }
            Ok(RawOp::SetComment {
                id: value
                    .get("id")
                    .and_then(Value::as_str)
                    .ok_or_else(|| EngineError::message("setComment requires a string \"id\""))?
                    .to_owned(),
                ranges,
                author: value
                    .get("author")
                    .and_then(Value::as_str)
                    .unwrap_or_default()
                    .to_owned(),
                date: value
                    .get("date")
                    .and_then(Value::as_str)
                    .unwrap_or_default()
                    .to_owned(),
                body: json_to_any(value.get("body").unwrap_or(&Value::Null))?,
            })
        }
        "removeComment" => Ok(RawOp::RemoveComment {
            id: value
                .get("id")
                .and_then(Value::as_str)
                .ok_or_else(|| EngineError::message("removeComment requires a string \"id\""))?
                .to_owned(),
        }),
        other => Err(EngineError::message(format!("unknown raw op {other:?}"))),
    }
}

pub(crate) fn parse_raw_ops(ops_json: &str) -> EngineResult<Vec<RawOp>> {
    let value: Value = serde_json::from_str(ops_json).map_err(EngineError::message)?;
    let entries = value
        .as_array()
        .ok_or_else(|| EngineError::message("apply_raw_ops expects an array of ops"))?;
    entries.iter().map(parse_raw_op).collect()
}

/// Parses an accept/reject target from JSON: `{"revisionId": string}` (one
/// coalesced revision, any story) or a Loc range
/// `{"story","startPara","startOffset","endPara","endOffset"}`.
pub(crate) fn parse_change_target(
    doc: &EditingDoc,
    target_json: &str,
) -> EngineResult<ChangeTarget> {
    let value: Value = serde_json::from_str(target_json).map_err(EngineError::message)?;
    if let Some(id) = value.get("revisionId").and_then(Value::as_str) {
        return Ok(ChangeTarget::Revision(id.to_owned()));
    }
    let get_str = |key: &str| {
        value.get(key).and_then(Value::as_str).ok_or_else(|| {
            EngineError::message(format!("a change range requires a string {key:?}"))
        })
    };
    let get_offset = |key: &str| {
        value
            .get(key)
            .and_then(Value::as_u64)
            .map(|offset| offset as u32)
            .ok_or_else(|| {
                EngineError::message(format!("a change range requires a non-negative {key:?}"))
            })
    };
    let story = get_str("story")?;
    let start = loc_index(
        doc,
        story,
        get_str("startPara")?,
        get_offset("startOffset")?,
    )?;
    let end = loc_index(doc, story, get_str("endPara")?, get_offset("endOffset")?)?;
    Ok(ChangeTarget::Range(StoryRange::new(story, start, end)))
}
