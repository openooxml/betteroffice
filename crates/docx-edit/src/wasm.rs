//! The wasm-bindgen session boundary over [`EditingDoc`].
//!
//! This module is the ONLY JS-visible surface of the crate (compiled behind
//! `--features wasm`; see `scripts/embed-edit-wasm.mjs`). Its JS twin is the
//! `packages/core/src/yrs/` facade — the sole JS entry to this crate, the
//! `docx/zipContainer.ts` precedent.
//!
//! Boundary conventions (mirrors `crates/docx-layout/src/lib.rs`):
//! - values cross as JSON strings (receipts, queries) or raw bytes (yrs
//!   updates); errors cross as `Err(JsValue)` carrying a display string;
//! - op addressing is the op-contract public vocabulary
//!   `Loc { story, paraId, offset }` — offsets are UTF-16 units within one
//!   paragraph (`offset ∈ [0, para_len]`, the paragraph's own pilcrow
//!   excluded). Story-global u32 indices stay transient inside each call;
//! - suggesting mode is an optional `(author_name, author_date)` pair — both
//!   or neither; it maps to an [`EditCtx`] (a plain local edit uses an empty
//!   author, which is stamped only in suggesting mode).
//!
//! Everything here is composition of the crate's PUBLIC ops — no op internals
//! live in this file (the ops track owns those).

use std::cell::RefCell;
use std::collections::{BTreeMap, HashMap};

use js_sys::{Function, Uint8Array};
use serde::Serialize;
use serde_json::{Value, json};
use wasm_bindgen::prelude::*;
use yrs::{Any, Assoc, IndexedSequence, Map, ReadTxn, StickyIndex, Subscription, Transact};

use crate::{
    CellLoc, ChangeKind, ChangeTarget, ColorPatch, DocUndoManager, EditCtx, EditingDoc,
    EngineSession, FontFamilyPatch, FormatPolicy, InlineFormatDelta, MergeDirection, ParaAttrDelta,
    ParaSelector, Patch, Position, RawOp, SegmentContent, SimpleFormat, StoryRange, TabStop,
    TableLocator, TableRange, TriState, story_ref,
};

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = performance, js_name = now)]
    fn performance_now() -> f64;
}

const STORIES: &str = "stories";

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ApplyInputProfile {
    selection_ms: f64,
    edit_ms: f64,
    lower_ms: f64,
    measure_ms: f64,
    paginate_ms: f64,
    display_input_ms: f64,
    display_build_ms: f64,
    display_finalize_ms: f64,
    display_ms: f64,
    encode_ms: f64,
}

fn js_err(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}
/// Builds the op [`EditCtx`]. Suggesting mode crosses the boundary as an
/// optional `(name, date)` pair — both-or-neither; a plain local edit uses an
/// empty author (author/date are stamped only in suggesting mode).
fn edit_ctx(name: Option<String>, date: Option<String>) -> Result<EditCtx, JsValue> {
    match (name, date) {
        (None, None) => Ok(EditCtx::local(String::new(), String::new())),
        (Some(name), Some(date)) => Ok(EditCtx::local(name, date).suggesting()),
        _ => Err(js_err(
            "suggesting requires both an author name and an ISO date",
        )),
    }
}

fn parse_any_object(json: &str, label: &str) -> Result<HashMap<String, Any>, JsValue> {
    match Any::from_json(json).map_err(js_err)? {
        Any::Map(value) => Ok(value.as_ref().clone()),
        _ => Err(js_err(format!("{label} must be a JSON object"))),
    }
}

struct ParaSpan {
    /// Story index of the paragraph's first unit (after the previous pilcrow).
    start: u32,
    /// Story index of the paragraph's own pilcrow embed.
    pilcrow: u32,
}

/// Resolves a paragraph to its story span by walking the public segment view.
/// Story-scoped: a `para_id` that lives in another story is "not found".
fn find_para_span(doc: &EditingDoc, story: &str, para_id: &str) -> Result<ParaSpan, JsValue> {
    let mut offset: u32 = 0;
    let mut para_start: u32 = 0;
    for segment in doc.story_segments(story).map_err(js_err)? {
        match segment.content {
            SegmentContent::Text(text) => offset += text.encode_utf16().count() as u32,
            SegmentContent::Pilcrow(properties) => {
                if properties.para_id == para_id {
                    return Ok(ParaSpan {
                        start: para_start,
                        pilcrow: offset,
                    });
                }
                offset += 1;
                para_start = offset;
            }
            SegmentContent::OtherEmbed { .. } => offset += 1,
        }
    }
    Err(js_err(format!(
        "paragraph {para_id:?} was not found in story {story:?}"
    )))
}

/// `Loc { story, paraId, offset }` -> transient story-global index.
fn loc_index(doc: &EditingDoc, story: &str, para_id: &str, offset: u32) -> Result<u32, JsValue> {
    let span = find_para_span(doc, story, para_id)?;
    let para_len = span.pilcrow - span.start;
    if offset > para_len {
        return Err(js_err(format!(
            "offset {offset} exceeds the length {para_len} of paragraph {para_id:?}"
        )));
    }
    Ok(span.start + offset)
}

/// Transient story-global index -> public paragraph-keyed location. Sticky
/// awareness positions resolve to story indices; the JS facade never exposes
/// that internal coordinate system.
fn index_loc(doc: &EditingDoc, story: &str, index: u32) -> Result<(String, u32), JsValue> {
    let mut cursor = 0_u32;
    let mut para_start = 0_u32;
    for segment in doc.story_segments(story).map_err(js_err)? {
        match segment.content {
            SegmentContent::Text(text) => cursor += text.encode_utf16().count() as u32,
            SegmentContent::Pilcrow(properties) => {
                if index <= cursor {
                    return Ok((properties.para_id, index.saturating_sub(para_start)));
                }
                cursor += 1;
                para_start = cursor;
            }
            SegmentContent::OtherEmbed { .. } => cursor += 1,
        }
    }
    Err(js_err(format!(
        "selection index {index} does not resolve in story {story:?}"
    )))
}

/// Per-peer selection state. These sticky positions are deliberately held
/// outside the yrs document: an awareness transport may publish them, but they
/// are never serialized as document content or included in save updates.
struct LocalSelection {
    story: String,
    anchor: StickyIndex,
    head: StickyIndex,
}

/// One endpoint of a local table selection. The cell story is the stable
/// identity; row/column are a fallback if that cell is removed locally.
struct LocalCellPoint {
    cell_story: String,
    row: u32,
    column: u32,
}

/// Per-peer cell selection, held outside the yrs document. The sticky table
/// position survives unrelated edits in the parent story; cell-story identity
/// lets each endpoint follow inserted/deleted rows and columns.
struct LocalCellSelection {
    parent_story: String,
    table: StickyIndex,
    anchor: LocalCellPoint,
    head: LocalCellPoint,
}

/// Applies one boundary "mark" to `range`, routing the six simple toggles
/// through [`EditingDoc::toggle_format`] and font/size/color through the
/// tri-state [`EditingDoc::format_range`] (a set, not a toggle — the
/// op-contract inline-formatting split). `mark_json`:
/// `{"type":"bold"|"italic"|"underline"|"strike"|"superscript"|"subscript"} |
/// {"type":"fontFamily"|"color","value":string} |
/// {"type":"fontSize","value":number}`.
fn apply_mark(doc: &EditingDoc, range: StoryRange, mark_json: &str) -> Result<(), JsValue> {
    let value: Value = serde_json::from_str(mark_json).map_err(js_err)?;
    let kind = value
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| js_err("mark JSON requires a string \"type\""))?;
    // Formatting is not itself tracked at S1 (op-contract deviation 3), so a
    // plain local context is correct regardless of the caret's suggest mode.
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
        doc.toggle_format(&ctx, range, format).map_err(js_err)?;
        return Ok(());
    }
    let delta = match kind {
        "fontFamily" => {
            let family = value
                .get("value")
                .and_then(Value::as_str)
                .ok_or_else(|| js_err("mark type \"fontFamily\" requires a string \"value\""))?;
            InlineFormatDelta {
                font_family: Patch::Set(FontFamilyPatch {
                    ascii: family.to_owned(),
                    h_ansi: None,
                }),
                ..Default::default()
            }
        }
        "fontSize" => {
            let size = value
                .get("value")
                .and_then(Value::as_f64)
                .ok_or_else(|| js_err("mark type \"fontSize\" requires a numeric \"value\""))?;
            InlineFormatDelta {
                font_size: Patch::Set(size),
                ..Default::default()
            }
        }
        "color" => {
            let color = value
                .get("value")
                .and_then(Value::as_str)
                .ok_or_else(|| js_err("mark type \"color\" requires a string \"value\""))?;
            InlineFormatDelta {
                color: Patch::Set(ColorPatch::Rgb(color.to_owned())),
                ..Default::default()
            }
        }
        other => return Err(js_err(format!("unknown mark type {other:?}"))),
    };
    doc.format_range(&ctx, range, &delta).map_err(js_err)?;
    Ok(())
}

/// Decodes the public facade's tri-state inline-formatting delta. An omitted
/// key is [`Patch::Keep`], `null` is [`Patch::Clear`], and any typed value is
/// [`Patch::Set`]. Boolean `false` values also clear in
/// [`InlineFormatDelta::to_attrs`], matching the PM formatting commands.
fn parse_inline_format_delta(delta_json: &str) -> Result<InlineFormatDelta, JsValue> {
    let value: Value = serde_json::from_str(delta_json).map_err(js_err)?;
    let object = value
        .as_object()
        .ok_or_else(|| js_err("format_range expects a JSON object delta"))?;

    let bool_patch = |key: &str| -> Result<Patch<bool>, JsValue> {
        match object.get(key) {
            None => Ok(Patch::Keep),
            Some(Value::Null) => Ok(Patch::Clear),
            Some(Value::Bool(value)) => Ok(Patch::Set(*value)),
            Some(_) => Err(js_err(format!(
                "format delta {key:?} must be a boolean or null"
            ))),
        }
    };

    let underline = match object.get("underline") {
        None => Patch::Keep,
        Some(Value::Null | Value::Bool(false)) => Patch::Clear,
        Some(Value::Bool(true)) => Patch::Set(Default::default()),
        Some(Value::Object(value)) => {
            let string = |key: &str| -> Result<Option<String>, JsValue> {
                value
                    .get(key)
                    .map(|entry| {
                        entry
                            .as_str()
                            .map(str::to_owned)
                            .ok_or_else(|| js_err(format!("underline {key:?} must be a string")))
                    })
                    .transpose()
            };
            Patch::Set(crate::UnderlinePatch {
                style: string("style")?,
                color: string("color")?,
            })
        }
        Some(_) => {
            return Err(js_err(
                "format delta \"underline\" must be a boolean, object, or null",
            ));
        }
    };

    let strike = match object.get("strike") {
        None => Patch::Keep,
        Some(Value::Null | Value::Bool(false)) => Patch::Clear,
        Some(Value::Bool(true)) => Patch::Set(crate::StrikePatch { double: false }),
        Some(Value::Object(value)) => {
            let double = value
                .get("double")
                .map(|entry| {
                    entry
                        .as_bool()
                        .ok_or_else(|| js_err("strike \"double\" must be a boolean"))
                })
                .transpose()?
                .unwrap_or(false);
            Patch::Set(crate::StrikePatch { double })
        }
        Some(_) => {
            return Err(js_err(
                "format delta \"strike\" must be a boolean, object, or null",
            ));
        }
    };

    let color = match object.get("color") {
        None => Patch::Keep,
        Some(Value::Null) => Patch::Clear,
        Some(Value::Object(value)) => match (
            value.get("rgb").and_then(Value::as_str),
            value.get("themeColor").and_then(Value::as_str),
        ) {
            (Some(rgb), None) => Patch::Set(ColorPatch::Rgb(rgb.to_owned())),
            (None, Some(theme)) => Patch::Set(ColorPatch::Theme(theme.to_owned())),
            _ => {
                return Err(js_err(
                    "format delta \"color\" requires exactly one of \"rgb\" or \"themeColor\"",
                ));
            }
        },
        Some(_) => {
            return Err(js_err("format delta \"color\" must be an object or null"));
        }
    };

    let highlight = match object.get("highlight") {
        None => Patch::Keep,
        Some(Value::Null) => Patch::Clear,
        Some(Value::String(value)) => Patch::Set(value.clone()),
        Some(_) => {
            return Err(js_err(
                "format delta \"highlight\" must be a string or null",
            ));
        }
    };

    let font_size = match object.get("fontSize") {
        None => Patch::Keep,
        Some(Value::Null) => Patch::Clear,
        Some(Value::Number(value)) => Patch::Set(
            value
                .as_f64()
                .ok_or_else(|| js_err("format delta \"fontSize\" must be finite"))?,
        ),
        Some(_) => {
            return Err(js_err("format delta \"fontSize\" must be a number or null"));
        }
    };

    let font_family = match object.get("fontFamily") {
        None => Patch::Keep,
        Some(Value::Null) => Patch::Clear,
        Some(Value::Object(value)) => {
            let ascii = value
                .get("ascii")
                .and_then(Value::as_str)
                .ok_or_else(|| js_err("format delta \"fontFamily\" requires string \"ascii\""))?;
            let h_ansi = value
                .get("hAnsi")
                .map(|entry| {
                    entry
                        .as_str()
                        .map(str::to_owned)
                        .ok_or_else(|| js_err("fontFamily \"hAnsi\" must be a string"))
                })
                .transpose()?;
            Patch::Set(FontFamilyPatch {
                ascii: ascii.to_owned(),
                h_ansi,
            })
        }
        Some(_) => {
            return Err(js_err(
                "format delta \"fontFamily\" must be an object or null",
            ));
        }
    };

    let mut other = BTreeMap::new();
    if let Some(value) = object.get("other") {
        let entries = value
            .as_object()
            .ok_or_else(|| js_err("format delta \"other\" must be an object"))?;
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

    Ok(InlineFormatDelta {
        bold: bool_patch("bold")?,
        italic: bool_patch("italic")?,
        underline,
        strike,
        color,
        highlight,
        font_size,
        font_family,
        other,
    })
}

/// Decodes the public facade's tri-state paragraph-property delta. Typed
/// fields lower to [`ParaAttrDelta`]; list/render metadata and other passive
/// pPr fields use its `other` bag. Omitted fields are kept and `null` clears.
fn parse_para_attr_delta(attrs_json: &str) -> Result<ParaAttrDelta, JsValue> {
    let value: Value = serde_json::from_str(attrs_json).map_err(js_err)?;
    let object = value
        .as_object()
        .ok_or_else(|| js_err("set_paragraph_attrs expects a JSON object"))?;

    let string_patch = |key: &str| -> Result<Patch<String>, JsValue> {
        match object.get(key) {
            None => Ok(Patch::Keep),
            Some(Value::Null) => Ok(Patch::Clear),
            Some(Value::String(value)) => Ok(Patch::Set(value.clone())),
            Some(_) => Err(js_err(format!(
                "paragraph attribute {key:?} must be a string or null"
            ))),
        }
    };
    let number_patch =
        |key: &str| -> Result<Patch<f64>, JsValue> {
            match object.get(key) {
                None => Ok(Patch::Keep),
                Some(Value::Null) => Ok(Patch::Clear),
                Some(Value::Number(value)) => Ok(Patch::Set(value.as_f64().ok_or_else(|| {
                    js_err(format!("paragraph attribute {key:?} must be finite"))
                })?)),
                Some(_) => Err(js_err(format!(
                    "paragraph attribute {key:?} must be a number or null"
                ))),
            }
        };
    let bool_patch = |key: &str| -> Result<Patch<bool>, JsValue> {
        match object.get(key) {
            None => Ok(Patch::Keep),
            Some(Value::Null) => Ok(Patch::Clear),
            Some(Value::Bool(value)) => Ok(Patch::Set(*value)),
            Some(_) => Err(js_err(format!(
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
                    .ok_or_else(|| js_err("each paragraph tab must be an object"))?;
                let pos = stop
                    .get("position")
                    .and_then(Value::as_f64)
                    .ok_or_else(|| js_err("a paragraph tab requires numeric \"position\""))?;
                let alignment = stop
                    .get("alignment")
                    .and_then(Value::as_str)
                    .ok_or_else(|| js_err("a paragraph tab requires string \"alignment\""))?;
                let leader =
                    stop.get("leader")
                        .map(|entry| {
                            entry.as_str().map(str::to_owned).ok_or_else(|| {
                                js_err("a paragraph tab \"leader\" must be a string")
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
            return Err(js_err(
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
            return Err(js_err(
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
        let entries = value
            .as_object()
            .ok_or_else(|| js_err("paragraph attribute \"other\" must be an object"))?;
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

fn attrs_value(attrs: &std::collections::BTreeMap<String, Any>) -> Result<Value, JsValue> {
    serde_json::to_value(attrs).map_err(js_err)
}

fn tri_state_value(state: TriState) -> Value {
    match state {
        TriState::On => Value::Bool(true),
        TriState::Off => Value::Bool(false),
        TriState::Mixed => Value::String("mixed".to_owned()),
    }
}

fn json_to_any(value: &Value) -> Result<Any, JsValue> {
    Any::from_json(&value.to_string()).map_err(js_err)
}

/// A JSON object → yrs text/format attributes (`Arc<str>` keys, `Any` values).
fn parse_attrs(value: Option<&Value>) -> Result<yrs::types::Attrs, JsValue> {
    let mut attrs = yrs::types::Attrs::new();
    if let Some(Value::Object(map)) = value {
        for (key, entry) in map {
            attrs.insert(std::sync::Arc::from(key.as_str()), json_to_any(entry)?);
        }
    }
    Ok(attrs)
}

/// A JSON object → ordered `(key, Any)` embed payload entries.
fn parse_payload(value: Option<&Value>) -> Result<Vec<(String, Any)>, JsValue> {
    let mut out = Vec::new();
    if let Some(Value::Object(map)) = value {
        for (key, entry) in map {
            out.push((key.clone(), json_to_any(entry)?));
        }
    }
    Ok(out)
}

/// Parses one `{ "op", "index", … }` coexistence mirror op.
fn parse_raw_op(value: &Value) -> Result<RawOp, JsValue> {
    let op = value
        .get("op")
        .and_then(Value::as_str)
        .ok_or_else(|| js_err("a raw op requires a string \"op\""))?;
    let index = value
        .get("index")
        .and_then(Value::as_u64)
        .map(|i| i as u32)
        .ok_or_else(|| js_err("a raw op requires a non-negative \"index\""));
    let u32_field = |key: &str| {
        value
            .get(key)
            .and_then(Value::as_u64)
            .map(|v| v as u32)
            .ok_or_else(|| js_err(format!("op {op:?} requires a non-negative {key:?}")))
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
                .ok_or_else(|| js_err("setEmbedAttr requires a string \"key\""))?
                .to_owned(),
            value: json_to_any(value.get("value").unwrap_or(&Value::Null))?,
        }),
        "setComment" => {
            let ranges_value = value
                .get("ranges")
                .and_then(Value::as_array)
                .ok_or_else(|| js_err("setComment requires an array \"ranges\""))?;
            let mut ranges = Vec::with_capacity(ranges_value.len());
            for entry in ranges_value {
                let pair = entry
                    .as_array()
                    .filter(|pair| pair.len() == 2)
                    .ok_or_else(|| js_err("a setComment range must be a [start, end] pair"))?;
                let bound = |index: usize| {
                    pair[index]
                        .as_u64()
                        .map(|bound| bound as u32)
                        .ok_or_else(|| {
                            js_err("a setComment range bound must be a non-negative integer")
                        })
                };
                ranges.push((bound(0)?, bound(1)?));
            }
            Ok(RawOp::SetComment {
                id: value
                    .get("id")
                    .and_then(Value::as_str)
                    .ok_or_else(|| js_err("setComment requires a string \"id\""))?
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
                .ok_or_else(|| js_err("removeComment requires a string \"id\""))?
                .to_owned(),
        }),
        other => Err(js_err(format!("unknown raw op {other:?}"))),
    }
}

/// Parses an accept/reject target from JSON: `{"revisionId": string}` (one
/// coalesced revision, any story) or a Loc range
/// `{"story","startPara","startOffset","endPara","endOffset"}`.
fn parse_change_target(doc: &EditingDoc, target_json: &str) -> Result<ChangeTarget, JsValue> {
    let value: Value = serde_json::from_str(target_json).map_err(js_err)?;
    if let Some(id) = value.get("revisionId").and_then(Value::as_str) {
        return Ok(ChangeTarget::Revision(id.to_owned()));
    }
    let get_str = |key: &str| {
        value
            .get(key)
            .and_then(Value::as_str)
            .ok_or_else(|| js_err(format!("a change range requires a string {key:?}")))
    };
    let get_offset = |key: &str| {
        value
            .get(key)
            .and_then(Value::as_u64)
            .map(|offset| offset as u32)
            .ok_or_else(|| js_err(format!("a change range requires a non-negative {key:?}")))
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

/// Parses the render bridge's host context from JSON:
/// `{ "themeColors": {name: hex}, "defaultTabStopTwips": number|null,
/// "pageContentHeight": number|null, "numericIds": {yrsId: number} }`.
fn parse_render_env(env_json: &str) -> Result<crate::bridge::RenderEnv, JsValue> {
    let value: Value = serde_json::from_str(env_json).map_err(js_err)?;
    let mut env = crate::bridge::RenderEnv::default();
    if let Some(Value::Object(colors)) = value.get("themeColors") {
        for (key, entry) in colors {
            if let Some(hex) = entry.as_str() {
                env.theme_colors.insert(key.clone(), hex.to_owned());
            }
        }
    }
    env.default_tab_stop_twips = value.get("defaultTabStopTwips").and_then(Value::as_f64);
    env.page_content_height = value.get("pageContentHeight").and_then(Value::as_f64);
    if let Some(Value::Object(ids)) = value.get("numericIds") {
        for (key, entry) in ids {
            if let Some(id) = entry.as_f64() {
                env.numeric_ids.insert(key.clone(), id);
            }
        }
    }
    Ok(env)
}

fn seed_paragraph(value: &Value) -> (String, String, String) {
    let text = value
        .get("text")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_owned();
    let p_style = value
        .get("pStyle")
        .and_then(Value::as_str)
        .unwrap_or("Normal")
        .to_owned();
    let alignment = value
        .get("alignment")
        .and_then(Value::as_str)
        .unwrap_or("left")
        .to_owned();
    (text, p_style, alignment)
}

/// One yrs replica of the DOCX editing model, held for a JS host.
///
/// Owns the [`EditingDoc`] plus the (single) JS update observer. The JS facade
/// multiplexes its own listener set over that one callback.
#[wasm_bindgen]
pub struct EditSession {
    engine: EngineSession,
    update_observer: Option<Subscription>,
    undo: RefCell<Option<DocUndoManager>>,
    undo_story: RefCell<Option<String>>,
    selection: RefCell<Option<LocalSelection>>,
    cell_selection: RefCell<Option<LocalCellSelection>>,
    last_apply_profile_json: RefCell<String>,
}

impl EditSession {
    fn collapsed_resident_input_selection(&self) -> Result<(String, String, u32, u32), JsValue> {
        let selection = self.selection.borrow();
        let selection = selection
            .as_ref()
            .ok_or_else(|| js_err("resident input requires a selection"))?;
        let story = selection.story.clone();
        let txn = self.engine.doc().yrs_doc().transact();
        let anchor = selection
            .anchor
            .get_offset(&txn)
            .ok_or_else(|| js_err("selection anchor no longer resolves"))?
            .index;
        let head = selection
            .head
            .get_offset(&txn)
            .ok_or_else(|| js_err("selection head no longer resolves"))?
            .index;
        drop(txn);
        if anchor != head {
            return Err(js_err(
                "resident input currently requires a collapsed selection",
            ));
        }
        let (para_id, offset) = index_loc(self.engine.doc(), &story, head)?;
        if !self.engine.can_apply_input(&story, &para_id) {
            return Err(js_err(
                "resident input state is not ready for this paragraph",
            ));
        }
        Ok((story, para_id, offset, head))
    }

    fn delete_resident_input(
        &self,
        direction: &str,
        selection: (String, String, u32, u32),
    ) -> Result<String, JsValue> {
        let (story, para_id, offset, head) = selection;
        let paragraphs = self.engine.doc().paragraphs(&story).map_err(js_err)?;
        let paragraph_index = paragraphs
            .iter()
            .position(|paragraph| paragraph.para_id == para_id)
            .ok_or_else(|| js_err("resident input paragraph no longer resolves"))?;
        let paragraph = &paragraphs[paragraph_index];
        let units: Vec<u16> = paragraph.text.encode_utf16().collect();
        let ctx = EditCtx::local("", "");

        match direction {
            "backward" if offset > 0 => {
                let previous = if offset > 1
                    && (0xdc00..=0xdfff).contains(&units[offset as usize - 1])
                    && (0xd800..=0xdbff).contains(&units[offset as usize - 2])
                {
                    offset - 2
                } else {
                    offset - 1
                };
                self.engine
                    .doc()
                    .delete_range(
                        &ctx,
                        StoryRange::new(&story, head - (offset - previous), head),
                    )
                    .map_err(js_err)?;
            }
            "backward" if paragraph_index > 0 => {
                self.engine
                    .doc()
                    .merge_paragraphs(&ctx, &para_id, MergeDirection::Backward)
                    .map_err(js_err)?;
            }
            "forward" if (offset as usize) < units.len() => {
                let next = if offset as usize + 1 < units.len()
                    && (0xd800..=0xdbff).contains(&units[offset as usize])
                    && (0xdc00..=0xdfff).contains(&units[offset as usize + 1])
                {
                    offset + 2
                } else {
                    offset + 1
                };
                self.engine
                    .doc()
                    .delete_range(&ctx, StoryRange::new(&story, head, head + (next - offset)))
                    .map_err(js_err)?;
            }
            "forward" if paragraph_index + 1 < paragraphs.len() => {
                self.engine
                    .doc()
                    .merge_paragraphs(&ctx, &para_id, MergeDirection::Forward)
                    .map_err(js_err)?;
            }
            "backward" | "forward" => {
                return Err(js_err("resident input has no character in that direction"));
            }
            _ => return Err(js_err("delete direction must be backward or forward")),
        }
        Ok(story)
    }
}

#[wasm_bindgen]
impl EditSession {
    /// Creates a replica. `client_id` must be a non-negative safe integer —
    /// the host allocates it (yjs-style random 32-bit ids are fine).
    #[wasm_bindgen(constructor)]
    pub fn new(client_id: f64) -> Result<EditSession, JsValue> {
        const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
        if !(client_id.is_finite()
            && client_id >= 0.0
            && client_id.fract() == 0.0
            && client_id <= MAX_SAFE_INTEGER)
        {
            return Err(js_err("client_id must be a non-negative safe integer"));
        }
        Ok(Self {
            engine: EngineSession::new(client_id as u64),
            update_observer: None,
            undo: RefCell::new(None),
            undo_story: RefCell::new(None),
            selection: RefCell::new(None),
            cell_selection: RefCell::new(None),
            last_apply_profile_json: RefCell::new("{}".to_owned()),
        })
    }

    pub fn client_id(&self) -> f64 {
        self.engine.doc().client_id() as f64
    }

    /// Register font bytes in this editing wasm's resident measurement store.
    /// Returned ids are valid for measurement and display work in this module.
    pub fn register_measure_font(&self, bytes: &[u8]) -> Result<u32, JsValue> {
        docx_layout::register_measure_font(bytes)
    }

    /// Clear this editing wasm's resident measurement fonts.
    pub fn clear_measure_fonts(&self) {
        docx_layout::clear_measure_fonts();
        self.engine.clear_measurement_templates();
    }

    /// Paragraph-measure compatibility export on the resident engine module.
    pub fn measure_paragraph_json(&self, input: &str) -> Result<String, JsValue> {
        self.engine.measure_paragraph_json(input).map_err(js_err)
    }

    /// Paginate and retain the measured input and Layout. The full JSON return
    /// remains the migration parity bridge until binary frames consume it.
    pub fn layout_document_json(&self, input: &str) -> Result<String, JsValue> {
        self.engine
            .layout_document_json(input)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Paginate and compose section/page regions inside the resident engine.
    pub fn layout_document_with_regions_json(&self, input: &str) -> Result<String, JsValue> {
        self.engine
            .layout_document_with_regions_json(input)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Build display primitives against the same resident font store used by
    /// this session's measurement path.
    pub fn build_display_list_json(&self, input: &str) -> Result<String, JsValue> {
        self.engine
            .build_display_list_json(input)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Binary FrameDelta v1 display output. The returned `Vec<u8>` is exposed
    /// by wasm-bindgen as a transferable-friendly `Uint8Array`.
    pub fn build_display_list_frame(
        &self,
        input: &str,
        expected_frame_epoch: f64,
    ) -> Result<Vec<u8>, JsValue> {
        const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
        if !(expected_frame_epoch.is_finite()
            && expected_frame_epoch >= 0.0
            && expected_frame_epoch.fract() == 0.0
            && expected_frame_epoch <= MAX_SAFE_INTEGER)
        {
            return Err(js_err(
                "expected_frame_epoch must be a non-negative safe integer",
            ));
        }
        self.engine
            .build_display_list_frame(input, expected_frame_epoch as u64)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Apply one ordinary collapsed body-text insertion and return the
    /// resulting FrameDelta. Selection, measurement inputs, pagination
    /// checkpoints, and display state all remain resident in this session.
    pub fn apply_input(&self, text: &str, expected_frame_epoch: f64) -> Result<Vec<u8>, JsValue> {
        const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
        if text.is_empty() || text.contains(['\r', '\n']) {
            return Err(js_err(
                "apply_input requires non-empty paragraph-break-free text",
            ));
        }
        if !(expected_frame_epoch.is_finite()
            && expected_frame_epoch >= 0.0
            && expected_frame_epoch.fract() == 0.0
            && expected_frame_epoch <= MAX_SAFE_INTEGER)
        {
            return Err(js_err(
                "expected_frame_epoch must be a non-negative safe integer",
            ));
        }

        let selection = self.selection.borrow();
        let selection = selection
            .as_ref()
            .ok_or_else(|| js_err("apply_input requires a resident selection"))?;
        let story = selection.story.clone();
        let txn = self.engine.doc().yrs_doc().transact();
        let anchor = selection
            .anchor
            .get_offset(&txn)
            .ok_or_else(|| js_err("selection anchor no longer resolves"))?
            .index;
        let head = selection
            .head
            .get_offset(&txn)
            .ok_or_else(|| js_err("selection head no longer resolves"))?
            .index;
        drop(txn);
        if anchor != head {
            return Err(js_err(
                "apply_input currently requires a collapsed selection",
            ));
        }
        let (para_id, _) = index_loc(self.engine.doc(), &story, head)?;
        if !self.engine.can_apply_input(&story, &para_id) {
            return Err(js_err(
                "resident input state is not ready for this paragraph",
            ));
        }

        self.engine
            .doc()
            .insert_text(
                &EditCtx::local("", ""),
                Position::new(&story, head),
                text,
                FormatPolicy::Inherit,
            )
            .map_err(js_err)?;
        self.engine
            .apply_and_layout(&story, expected_frame_epoch as u64)
            .map_err(js_err)
    }

    /// Instrumented twin of `apply_input`, used only by opt-in browser perf
    /// traces. Keeping this separate leaves the production hot path timer-free.
    pub fn apply_input_profiled(
        &self,
        text: &str,
        expected_frame_epoch: f64,
    ) -> Result<Vec<u8>, JsValue> {
        const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
        if text.is_empty() || text.contains(['\r', '\n']) {
            return Err(js_err(
                "apply_input requires non-empty paragraph-break-free text",
            ));
        }
        if !(expected_frame_epoch.is_finite()
            && expected_frame_epoch >= 0.0
            && expected_frame_epoch.fract() == 0.0
            && expected_frame_epoch <= MAX_SAFE_INTEGER)
        {
            return Err(js_err(
                "expected_frame_epoch must be a non-negative safe integer",
            ));
        }

        let started = performance_now();
        let selection = self.selection.borrow();
        let selection = selection
            .as_ref()
            .ok_or_else(|| js_err("apply_input requires a resident selection"))?;
        let story = selection.story.clone();
        let txn = self.engine.doc().yrs_doc().transact();
        let anchor = selection
            .anchor
            .get_offset(&txn)
            .ok_or_else(|| js_err("selection anchor no longer resolves"))?
            .index;
        let head = selection
            .head
            .get_offset(&txn)
            .ok_or_else(|| js_err("selection head no longer resolves"))?
            .index;
        drop(txn);
        if anchor != head {
            return Err(js_err(
                "apply_input currently requires a collapsed selection",
            ));
        }
        let (para_id, _) = index_loc(self.engine.doc(), &story, head)?;
        if !self.engine.can_apply_input(&story, &para_id) {
            return Err(js_err(
                "resident input state is not ready for this paragraph",
            ));
        }
        let selection_ms = performance_now() - started;

        let started = performance_now();
        self.engine
            .doc()
            .insert_text(
                &EditCtx::local("", ""),
                Position::new(&story, head),
                text,
                FormatPolicy::Inherit,
            )
            .map_err(js_err)?;
        let edit_ms = performance_now() - started;
        let (frame, engine_profile) = self
            .engine
            .apply_and_layout_profiled(&story, expected_frame_epoch as u64, &mut performance_now)
            .map_err(js_err)?;
        let profile = ApplyInputProfile {
            selection_ms,
            edit_ms,
            lower_ms: engine_profile.lower_ms,
            measure_ms: engine_profile.measure_ms,
            paginate_ms: engine_profile.paginate_ms,
            display_input_ms: engine_profile.display_input_ms,
            display_build_ms: engine_profile.display_build_ms,
            display_finalize_ms: engine_profile.display_finalize_ms,
            display_ms: engine_profile.display_ms,
            encode_ms: engine_profile.encode_ms,
        };
        *self.last_apply_profile_json.borrow_mut() =
            serde_json::to_string(&profile).map_err(js_err)?;
        Ok(frame)
    }

    /// Apply one ordinary collapsed character deletion (or adjacent paragraph
    /// merge at a boundary) and return the resulting resident FrameDelta.
    pub fn apply_delete(
        &self,
        direction: &str,
        expected_frame_epoch: f64,
    ) -> Result<Vec<u8>, JsValue> {
        const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
        if !(expected_frame_epoch.is_finite()
            && expected_frame_epoch >= 0.0
            && expected_frame_epoch.fract() == 0.0
            && expected_frame_epoch <= MAX_SAFE_INTEGER)
        {
            return Err(js_err(
                "expected_frame_epoch must be a non-negative safe integer",
            ));
        }
        let selection = self.collapsed_resident_input_selection()?;
        let story = self.delete_resident_input(direction, selection)?;
        self.engine
            .apply_and_layout(&story, expected_frame_epoch as u64)
            .map_err(js_err)
    }

    /// Instrumented twin of `apply_delete`, used only by opt-in browser perf
    /// traces. Keeping this separate leaves the production hot path timer-free.
    pub fn apply_delete_profiled(
        &self,
        direction: &str,
        expected_frame_epoch: f64,
    ) -> Result<Vec<u8>, JsValue> {
        const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
        if !(expected_frame_epoch.is_finite()
            && expected_frame_epoch >= 0.0
            && expected_frame_epoch.fract() == 0.0
            && expected_frame_epoch <= MAX_SAFE_INTEGER)
        {
            return Err(js_err(
                "expected_frame_epoch must be a non-negative safe integer",
            ));
        }

        let started = performance_now();
        let selection = self.collapsed_resident_input_selection()?;
        let selection_ms = performance_now() - started;

        let started = performance_now();
        let story = self.delete_resident_input(direction, selection)?;
        let edit_ms = performance_now() - started;
        let (frame, engine_profile) = self
            .engine
            .apply_and_layout_profiled(&story, expected_frame_epoch as u64, &mut performance_now)
            .map_err(js_err)?;
        let profile = ApplyInputProfile {
            selection_ms,
            edit_ms,
            lower_ms: engine_profile.lower_ms,
            measure_ms: engine_profile.measure_ms,
            paginate_ms: engine_profile.paginate_ms,
            display_input_ms: engine_profile.display_input_ms,
            display_build_ms: engine_profile.display_build_ms,
            display_finalize_ms: engine_profile.display_finalize_ms,
            display_ms: engine_profile.display_ms,
            encode_ms: engine_profile.encode_ms,
        };
        *self.last_apply_profile_json.borrow_mut() =
            serde_json::to_string(&profile).map_err(js_err)?;
        Ok(frame)
    }

    /// Last opt-in `apply_input_profiled` stage timings as a compact JSON object.
    pub fn apply_input_profile_json(&self) -> String {
        self.last_apply_profile_json.borrow().clone()
    }

    /// Hit-test without re-serializing the resident display list through JS.
    pub fn display_hit_test_regions_json(
        &self,
        page_index: u32,
        x: f64,
        y: f64,
    ) -> Result<String, JsValue> {
        self.engine
            .display_hit_test_regions_json(page_index as usize, x, y)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Body range geometry without a display-list JSON round trip.
    pub fn display_range_rects_json(&self, from: f64, to: f64) -> Result<String, JsValue> {
        self.engine
            .display_range_rects_json(from as i64, to as i64)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Region-scoped range geometry without a display-list JSON round trip.
    pub fn display_range_rects_region_json(
        &self,
        region: &str,
        r_id: &str,
        from: f64,
        to: f64,
    ) -> Result<String, JsValue> {
        self.engine
            .display_range_rects_region_json(region, r_id, from as i64, to as i64)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Resolve one glyph outline from the session's resident font store.
    pub fn outline_glyph_json(&self, font_id: u32, glyph_id: u32) -> Result<String, JsValue> {
        docx_layout::outline_glyph_json(font_id, glyph_id)
    }

    // -- lifecycle (op-contract §1.6: load / encode_state / apply_update / subscribe) --

    /// Hydrates this replica from an encoded yrs update (the bytes form of
    /// `load` — typically another replica's `encode_state()` output).
    pub fn load(&self, update: &[u8]) -> Result<(), JsValue> {
        self.engine.doc().apply_update_v1(update).map_err(js_err)
    }

    /// Seeds stories from JSON (the json form of `load`):
    /// `[{"storyId","paragraphs":[{"text","pStyle"?,"alignment"?}, …]}, …]`.
    /// Paragraph text must not contain paragraph breaks. Returns
    /// `{storyId: [paraId, …]}` in document order. This is an S1 seeding
    /// scaffold composed from public ops; the real `load(ParsedDocument)`
    /// belongs to the ops track.
    pub fn load_json(&self, stories_json: &str) -> Result<String, JsValue> {
        let value: Value = serde_json::from_str(stories_json).map_err(js_err)?;
        let entries = value
            .as_array()
            .ok_or_else(|| js_err("load_json expects an array of stories"))?;
        let mut receipt = serde_json::Map::new();
        for entry in entries {
            let story_id = entry
                .get("storyId")
                .and_then(Value::as_str)
                .ok_or_else(|| js_err("a story entry requires a string \"storyId\""))?;
            let paragraphs = entry
                .get("paragraphs")
                .and_then(Value::as_array)
                .ok_or_else(|| js_err("a story entry requires a \"paragraphs\" array"))?;
            if paragraphs.is_empty() {
                return Err(js_err(format!(
                    "story {story_id:?} requires at least one paragraph"
                )));
            }

            let (text, p_style, alignment) = seed_paragraph(&paragraphs[0]);
            self.engine
                .doc()
                .create_story(story_id, &text, &p_style, &alignment)
                .map_err(js_err)?;
            let seed_ctx = EditCtx::local(String::new(), String::new());
            for paragraph in &paragraphs[1..] {
                let (text, p_style, alignment) = seed_paragraph(paragraph);
                // Boundary = index of the final pilcrow, before which the new
                // paragraph's text lands and at which the split inserts a new
                // pilcrow. Under the S1 split the FIRST half keeps the original
                // id and the SECOND half is re-minted, so the just-appended
                // paragraph is `second_para_id` — restamp ITS properties.
                let boundary = self.engine.doc().story_len(story_id).map_err(js_err)? - 1;
                if !text.is_empty() {
                    self.engine
                        .doc()
                        .insert_text(
                            &seed_ctx,
                            Position::new(story_id, boundary),
                            &text,
                            FormatPolicy::Inherit,
                        )
                        .map_err(js_err)?;
                }
                let split = self
                    .engine
                    .doc()
                    .split_paragraph(&seed_ctx, Position::new(story_id, boundary), None)
                    .map_err(js_err)?;
                self.engine
                    .doc()
                    .set_paragraph_attr(
                        &split.second_para_id,
                        "pStyle",
                        Any::from(p_style.as_str()),
                    )
                    .map_err(js_err)?;
                self.engine
                    .doc()
                    .set_paragraph_attr(
                        &split.second_para_id,
                        "alignment",
                        Any::from(alignment.as_str()),
                    )
                    .map_err(js_err)?;
            }

            let para_ids: Vec<Value> = self
                .engine
                .doc()
                .paragraphs(story_id)
                .map_err(js_err)?
                .into_iter()
                .map(|paragraph| Value::String(paragraph.para_id))
                .collect();
            receipt.insert(story_id.to_owned(), Value::Array(para_ids));
        }
        serde_json::to_string(&Value::Object(receipt)).map_err(js_err)
    }

    /// Full document state as one yrs v1 update (Yjs wire format).
    pub fn encode_state(&self) -> Vec<u8> {
        self.engine.doc().encode_state_as_update_v1()
    }

    /// Applies a remote/incremental yrs v1 update.
    pub fn apply_update(&self, update: &[u8]) -> Result<(), JsValue> {
        self.engine.doc().apply_update_v1(update).map_err(js_err)
    }

    /// Applies an update produced by this document's dedicated local worker.
    /// The local origin lets the main replica's UndoManager retain ownership of
    /// the edit; remote/collaboration updates must use `apply_update` instead.
    pub fn apply_local_update(&self, update: &[u8]) -> Result<(), JsValue> {
        self.engine
            .doc()
            .apply_local_update_v1(update)
            .map_err(js_err)
    }

    /// Subscribes `callback(update: Uint8Array)` to every committed
    /// transaction (v1 encoding — feed it straight to `apply_update` on a
    /// peer). One observer per session; a second call replaces the first.
    /// The facade fans out to multiple JS listeners over this single hook.
    pub fn set_update_observer(&mut self, callback: &Function) -> Result<(), JsValue> {
        let callback = callback.clone();
        let subscription = self
            .engine
            .doc()
            .yrs_doc()
            .observe_update_v1(move |_txn, event| {
                // Uint8Array::from copies out of wasm memory — the JS side
                // owns the bytes and may hold them across further edits.
                let bytes = Uint8Array::from(event.update.as_slice());
                let _ = callback.call1(&JsValue::NULL, &bytes.into());
            })
            .map_err(js_err)?;
        self.update_observer = Some(subscription);
        Ok(())
    }

    /// Drops the update observer registered by [`EditSession::set_update_observer`].
    pub fn clear_update_observer(&mut self) {
        self.update_observer = None;
    }

    // -- local input state (undo + awareness selection) --

    /// Starts local-origin undo tracking for one story. Hosts call this lazily
    /// after import/seeding but before the first direct input operation, so the
    /// initial document is not an undo step.
    pub fn track_undo(&self, story: &str) -> Result<(), JsValue> {
        if self.undo_story.borrow().as_deref() == Some(story) {
            return Ok(());
        }
        let manager = self.engine.doc().undo_scope(&[story]).map_err(js_err)?;
        *self.undo.borrow_mut() = Some(manager);
        *self.undo_story.borrow_mut() = Some(story.to_owned());
        Ok(())
    }

    /// Starts local undo tracking for a structural table transaction. Besides
    /// the parent story (which owns the table embed), the stories root must be
    /// in scope so undo/redo also removes/restores cell-story map entries.
    pub fn track_table_undo(&self, story: &str) -> Result<(), JsValue> {
        let scope_key = format!("table:{story}");
        if self.undo_story.borrow().as_deref() == Some(scope_key.as_str()) {
            return Ok(());
        }
        let mut manager = self.engine.doc().undo_scope(&[story]).map_err(js_err)?;
        let txn = self.engine.doc().yrs_doc().transact();
        let stories = txn
            .get_map(STORIES)
            .expect("stories root is declared by EditingDoc::new");
        drop(txn);
        manager
            .raw()
            .expand_scope(self.engine.doc().yrs_doc(), &stories);
        *self.undo.borrow_mut() = Some(manager);
        *self.undo_story.borrow_mut() = Some(scope_key);
        Ok(())
    }

    /// Reverts the latest local-origin transaction. Remote/system mirror
    /// transactions are excluded by `DocUndoManager`'s tracked-origin policy.
    pub fn undo(&self) -> bool {
        self.undo
            .borrow_mut()
            .as_mut()
            .is_some_and(DocUndoManager::undo)
    }

    /// Reapplies the latest locally undone transaction.
    pub fn redo(&self) -> bool {
        self.undo
            .borrow_mut()
            .as_mut()
            .is_some_and(DocUndoManager::redo)
    }

    pub fn can_undo(&self) -> bool {
        self.undo
            .borrow()
            .as_ref()
            .is_some_and(DocUndoManager::can_undo)
    }

    pub fn can_redo(&self) -> bool {
        self.undo
            .borrow()
            .as_ref()
            .is_some_and(DocUndoManager::can_redo)
    }

    /// Current local undo stack size. Zero before a story starts tracking.
    pub fn undo_depth(&self) -> u32 {
        self.undo
            .borrow()
            .as_ref()
            .map_or(0, |undo| undo.undo_depth() as u32)
    }

    /// Current local redo stack size. Zero before a story starts tracking.
    pub fn redo_depth(&self) -> u32 {
        self.undo
            .borrow()
            .as_ref()
            .map_or(0, |undo| undo.redo_depth() as u32)
    }

    /// Stores this peer's anchor/head as sticky positions. `Assoc::After`
    /// makes a collapsed caret advance with text inserted at the caret.
    #[allow(clippy::too_many_arguments)]
    pub fn set_selection(
        &self,
        story: &str,
        anchor_para: &str,
        anchor_offset: u32,
        head_para: &str,
        head_offset: u32,
    ) -> Result<(), JsValue> {
        let anchor_index = loc_index(self.engine.doc(), story, anchor_para, anchor_offset)?;
        let head_index = loc_index(self.engine.doc(), story, head_para, head_offset)?;
        let txn = self.engine.doc().yrs_doc().transact();
        let text = story_ref(&txn, story).map_err(js_err)?;
        let anchor = text
            .sticky_index(&txn, anchor_index, Assoc::After)
            .ok_or_else(|| js_err("selection anchor could not be made sticky"))?;
        let head = text
            .sticky_index(&txn, head_index, Assoc::After)
            .ok_or_else(|| js_err("selection head could not be made sticky"))?;
        drop(txn);
        *self.selection.borrow_mut() = Some(LocalSelection {
            story: story.to_owned(),
            anchor,
            head,
        });
        Ok(())
    }

    /// Resolves this peer's current sticky selection as two public Locs, or
    /// `null` before the host establishes an initial selection.
    pub fn selection(&self) -> Result<String, JsValue> {
        let selection = self.selection.borrow();
        let Some(selection) = selection.as_ref() else {
            return Ok("null".to_owned());
        };
        let txn = self.engine.doc().yrs_doc().transact();
        let anchor_index = selection
            .anchor
            .get_offset(&txn)
            .ok_or_else(|| js_err("selection anchor no longer resolves"))?
            .index;
        let head_index = selection
            .head
            .get_offset(&txn)
            .ok_or_else(|| js_err("selection head no longer resolves"))?
            .index;
        drop(txn);
        let (anchor_para, anchor_offset) =
            index_loc(self.engine.doc(), &selection.story, anchor_index)?;
        let (head_para, head_offset) = index_loc(self.engine.doc(), &selection.story, head_index)?;
        Ok(json!({
            "anchor": {
                "story": selection.story,
                "paraId": anchor_para,
                "offset": anchor_offset,
            },
            "head": {
                "story": selection.story,
                "paraId": head_para,
                "offset": head_offset,
            }
        })
        .to_string())
    }

    /// Stores a rectangular anchor-cell → head-cell selection outside the yrs
    /// document. `range_json` is a [`TableRange`]. The table embed is held by a
    /// sticky index and the endpoints by stable cell-story identity.
    pub fn set_cell_selection(&self, range_json: &str) -> Result<(), JsValue> {
        let range: TableRange = serde_json::from_str(range_json).map_err(js_err)?;
        if range.anchor.story != range.head.story
            || range.anchor.table_index != range.head.table_index
        {
            return Err(js_err("a cell selection must stay inside one table"));
        }
        let (anchor, anchor_story) = self
            .engine
            .doc()
            .resolve_cell_identity(&range.anchor)
            .map_err(js_err)?;
        let (head, head_story) = self
            .engine
            .doc()
            .resolve_cell_identity(&range.head)
            .map_err(js_err)?;
        let locator = anchor.table();
        let table_index = self
            .engine
            .doc()
            .table_embed_index(&locator)
            .map_err(js_err)?;
        let txn = self.engine.doc().yrs_doc().transact();
        let story = story_ref(&txn, &locator.story).map_err(js_err)?;
        let table = story
            .sticky_index(&txn, table_index, Assoc::Before)
            .ok_or_else(|| js_err("table selection could not be made sticky"))?;
        drop(txn);
        *self.cell_selection.borrow_mut() = Some(LocalCellSelection {
            parent_story: locator.story,
            table,
            anchor: LocalCellPoint {
                cell_story: anchor_story,
                row: anchor.row,
                column: anchor.column,
            },
            head: LocalCellPoint {
                cell_story: head_story,
                row: head.row,
                column: head.column,
            },
        });
        Ok(())
    }

    /// Resolves the current local cell selection, or `null` before the host
    /// establishes one. Deleted endpoints clamp to a surviving nearby cell.
    pub fn cell_selection(&self) -> Result<String, JsValue> {
        let selection = self.cell_selection.borrow();
        let Some(selection) = selection.as_ref() else {
            return Ok("null".to_owned());
        };
        let txn = self.engine.doc().yrs_doc().transact();
        let table_index = selection
            .table
            .get_offset(&txn)
            .ok_or_else(|| js_err("cell selection table no longer resolves"))?
            .index;
        drop(txn);
        let locator = self
            .engine
            .doc()
            .table_locator_at_index(&selection.parent_story, table_index)
            .map_err(js_err)?;
        let anchor = self
            .engine
            .doc()
            .cell_loc_for_story(
                &locator,
                &selection.anchor.cell_story,
                selection.anchor.row,
                selection.anchor.column,
            )
            .map_err(js_err)?;
        let head = self
            .engine
            .doc()
            .cell_loc_for_story(
                &locator,
                &selection.head.cell_story,
                selection.head.row,
                selection.head.column,
            )
            .map_err(js_err)?;
        serde_json::to_string(&TableRange { anchor, head }).map_err(js_err)
    }

    // -- S1 ops (Loc addressing; JSON receipts) --

    /// Adds a story with one paragraph. Receipt: `{"paraId"}` (the final
    /// pilcrow's paragraph).
    pub fn create_story(
        &self,
        story_id: &str,
        initial_text: &str,
        p_style: &str,
        alignment: &str,
    ) -> Result<String, JsValue> {
        let para_id = self
            .engine
            .doc()
            .create_story(story_id, initial_text, p_style, alignment)
            .map_err(js_err)?;
        Ok(json!({ "paraId": para_id }).to_string())
    }

    /// Removes one complete story (used for unreachable table-cell stories).
    pub fn delete_story(&self, story_id: &str) -> Result<(), JsValue> {
        self.engine.doc().delete_story(story_id).map_err(js_err)
    }

    // -- native table ops (cell-grid addressing; JSON receipts) --

    /// Inserts a rectangular structural table at a paragraph-keyed location.
    #[allow(clippy::too_many_arguments)]
    pub fn insert_table(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        rows: u32,
        columns: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .insert_table(&ctx, Position::new(story, index), rows, columns)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Inserts a row above (`after = false`) or below (`after = true`) a cell.
    pub fn insert_row(
        &self,
        at_json: &str,
        after: bool,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        let at: CellLoc = serde_json::from_str(at_json).map_err(js_err)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .insert_row(&ctx, &at, after)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Inserts a column left (`after = false`) or right (`after = true`) of a cell.
    pub fn insert_column(&self, at_json: &str, after: bool) -> Result<String, JsValue> {
        let at: CellLoc = serde_json::from_str(at_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .insert_column(&EditCtx::local("", ""), &at, after)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Deletes every row covered by an explicit cell range.
    pub fn delete_row(
        &self,
        range_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        let range: TableRange = serde_json::from_str(range_json).map_err(js_err)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self.engine.doc().delete_row(&ctx, &range).map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Deletes every column covered by an explicit cell range.
    pub fn delete_column(&self, range_json: &str) -> Result<String, JsValue> {
        let range: TableRange = serde_json::from_str(range_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .delete_column(&EditCtx::local("", ""), &range)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Removes one complete table plus all of its reachable cell stories.
    pub fn delete_table(&self, table_json: &str) -> Result<String, JsValue> {
        let table: TableLocator = serde_json::from_str(table_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .delete_table(&EditCtx::local("", ""), &table)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Merges a rectangular cell range into its top-left cell.
    pub fn merge_cells(&self, range_json: &str) -> Result<String, JsValue> {
        let range: TableRange = serde_json::from_str(range_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .merge_cells(&EditCtx::local("", ""), &range)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Splits the cell covering `at` into the requested grid. Omitted
    /// dimensions unmerge the cell into its existing covered slots.
    pub fn split_cell(
        &self,
        at_json: &str,
        rows: Option<u32>,
        columns: Option<u32>,
    ) -> Result<String, JsValue> {
        let at: CellLoc = serde_json::from_str(at_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .split_cell_grid(&EditCtx::local("", ""), &at, rows, columns)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Sets/clears selected cells' background color (hex without or with `#`).
    pub fn set_cell_shading(
        &self,
        range_json: &str,
        color: Option<String>,
    ) -> Result<String, JsValue> {
        let range: TableRange = serde_json::from_str(range_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .set_cell_shading(&EditCtx::local("", ""), &range, color.as_deref())
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Merges a JSON cell-format patch into every selected cell's `tcPr`.
    pub fn set_cell_text_format(
        &self,
        range_json: &str,
        patch_json: &str,
    ) -> Result<String, JsValue> {
        let range: TableRange = serde_json::from_str(range_json).map_err(js_err)?;
        let patch = parse_any_object(patch_json, "cell format patch")?;
        let receipt = self
            .engine
            .doc()
            .set_cell_text_format(&EditCtx::local("", ""), &range, &patch)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Replaces the selected cells' complete border property object.
    pub fn set_cell_borders(
        &self,
        range_json: &str,
        borders_json: &str,
    ) -> Result<String, JsValue> {
        let range: TableRange = serde_json::from_str(range_json).map_err(js_err)?;
        let borders = parse_any_object(borders_json, "cell borders")?;
        let receipt = self
            .engine
            .doc()
            .set_cell_borders(&EditCtx::local("", ""), &range, &borders)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Sets one grid-column width in twips.
    pub fn set_column_width(&self, at_json: &str, width_twips: f64) -> Result<String, JsValue> {
        let at: CellLoc = serde_json::from_str(at_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .set_column_width(&EditCtx::local("", ""), &at, width_twips)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Sets the table-wide preferred width in twips.
    pub fn set_table_width(&self, table_json: &str, width_twips: f64) -> Result<String, JsValue> {
        let table: TableLocator = serde_json::from_str(table_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .set_table_width(&EditCtx::local("", ""), &table, width_twips)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Inserts paragraph-break-free text at `(story, para_id, offset)`.
    /// Receipt: `{"revisionId": string|null}` (non-null in suggesting mode).
    pub fn insert_text(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        text: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        let at = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .insert_text(&ctx, Position::new(story, at), text, FormatPolicy::Inherit)
            .map_err(js_err)?;
        Ok(json!({ "revisionId": receipt.revision_ids.into_iter().next() }).to_string())
    }

    /// Deletes `[start, end)` given as two Locs in one story. A range whose
    /// ends sit in different paragraphs spans the boundary pilcrows, so the
    /// plain delete also merges (the pilcrow-as-character dividend).
    /// Suggesting mode retains the content with a `del` revision instead.
    /// Receipt: `{"revisionId": string|null}`.
    #[allow(clippy::too_many_arguments)]
    pub fn delete_range(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .delete_range(&ctx, StoryRange::new(story, start, end))
            .map_err(js_err)?;
        Ok(json!({ "revisionId": receipt.revision_ids.into_iter().next() }).to_string())
    }

    /// Replaces `[start, end)` with text in one transaction. The inserted text
    /// adopts the first replaced unit's formatting; in suggesting mode the
    /// deletion and insertion share one revision id.
    #[allow(clippy::too_many_arguments)]
    pub fn replace_range(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        text: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .replace_range(&ctx, StoryRange::new(story, start, end), text)
            .map_err(js_err)?;
        Ok(json!({ "revisionId": receipt.revision_ids.into_iter().next() }).to_string())
    }

    /// Splits a paragraph at `(story, para_id, offset)` by inserting one
    /// pilcrow. Under the S1 split the FIRST half keeps the original paraId and
    /// the SECOND half is re-minted. Receipt:
    /// `{"firstParaId","secondParaId","revisionId": string|null}`.
    pub fn split_paragraph(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        let at = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .split_paragraph(&ctx, Position::new(story, at), None)
            .map_err(js_err)?;
        Ok(json!({
            "firstParaId": receipt.first_para_id,
            "secondParaId": receipt.second_para_id,
            "revisionId": receipt.revision_ids.into_iter().next(),
        })
        .to_string())
    }

    /// Merges `para_id` with the FOLLOWING paragraph by deleting (plain) or
    /// `del`-marking (suggesting) its pilcrow. Errors on the story's final
    /// paragraph. Receipt: `{"revisionId": string|null}`.
    pub fn merge_paragraphs(
        &self,
        story: &str,
        para_id: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        // Validate story membership (story-scoped "not found") before merging.
        find_para_span(self.engine.doc(), story, para_id)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .merge_paragraphs(&ctx, para_id, MergeDirection::Forward)
            .map_err(js_err)?;
        Ok(json!({ "revisionId": receipt.revision_ids.into_iter().next() }).to_string())
    }

    /// Applies one run mark over `[start, end)` (two Locs in one story). Simple
    /// marks toggle; font/size/color set (see [`apply_mark`]). `mark_json`:
    /// `{"type":"bold"|"italic"|"underline"|"strike"|"superscript"|"subscript"} |
    /// {"type":"fontFamily"|"color","value":string} |
    /// {"type":"fontSize","value":number}`.
    #[allow(clippy::too_many_arguments)]
    pub fn toggle_mark(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        mark_json: &str,
    ) -> Result<(), JsValue> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        apply_mark(
            self.engine.doc(),
            StoryRange::new(story, start, end),
            mark_json,
        )
    }

    /// Applies a set-valued, tri-state inline formatting delta over
    /// `[start, end)` in one transaction. Omitted fields are kept and `null`
    /// fields are cleared.
    #[allow(clippy::too_many_arguments)]
    pub fn format_range(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        delta_json: &str,
    ) -> Result<(), JsValue> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let delta = parse_inline_format_delta(delta_json)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .format_range(&ctx, StoryRange::new(story, start, end), &delta)
            .map(|_| ())
            .map_err(js_err)
    }

    /// Sets or clears the protected hyperlink attribute over `[start, end)`.
    /// `hyperlink_json` is an object (`{href, tooltip?, rId?}`) or `null`.
    #[allow(clippy::too_many_arguments)]
    pub fn set_hyperlink(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        hyperlink_json: &str,
    ) -> Result<(), JsValue> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let value: Value = serde_json::from_str(hyperlink_json).map_err(js_err)?;
        let hyperlink = if value.is_null() {
            None
        } else if value.is_object() {
            Some(json_to_any(&value)?)
        } else {
            return Err(js_err("set_hyperlink expects an object or null"));
        };
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_hyperlink(&ctx, StoryRange::new(story, start, end), hyperlink)
            .map(|_| ())
            .map_err(js_err)
    }

    /// Clears every direct formatting attribute over `[start, end)`, while
    /// retaining hyperlinks and tracked-change stamps.
    #[allow(clippy::too_many_arguments)]
    pub fn clear_formatting(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
    ) -> Result<(), JsValue> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .clear_formatting(&ctx, StoryRange::new(story, start, end))
            .map(|_| ())
            .map_err(js_err)
    }

    /// Applies a paragraph style id to every paragraph intersecting
    /// `[start, end)`. With no host style resolver at this boundary, this is
    /// the PM fallback path: write `pStyle` without fabricating a resolved
    /// paragraph/run formatting projection.
    #[allow(clippy::too_many_arguments)]
    pub fn apply_paragraph_style(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        style_id: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<(), JsValue> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let selector = ParaSelector::Range(StoryRange::new(story, start, end));
        let mut delta = ParaAttrDelta::default();
        delta
            .other
            .insert("pStyle".to_owned(), Some(Any::from(style_id)));
        let ctx = edit_ctx(author_name, author_date)?;
        self.engine
            .doc()
            .set_paragraph_attrs(&ctx, &selector, &delta)
            .map(|_| ())
            .map_err(js_err)
    }

    /// Applies a tri-state paragraph-property delta to every paragraph
    /// intersecting `[start, end)` in one transaction.
    #[allow(clippy::too_many_arguments)]
    pub fn set_paragraph_attrs(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        attrs_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<(), JsValue> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let selector = ParaSelector::Range(StoryRange::new(story, start, end));
        let delta = parse_para_attr_delta(attrs_json)?;
        let ctx = edit_ctx(author_name, author_date)?;
        self.engine
            .doc()
            .set_paragraph_attrs(&ctx, &selector, &delta)
            .map(|_| ())
            .map_err(js_err)
    }

    /// Inserts one native inline image embed at a paragraph-keyed location.
    #[allow(clippy::too_many_arguments)]
    pub fn insert_image(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        payload_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let Any::Map(payload) = Any::from_json(payload_json).map_err(js_err)? else {
            return Err(js_err("insert_image expects a JSON object"));
        };
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .insert_embed(
                &ctx,
                Position::new(story, index),
                "image",
                payload
                    .iter()
                    .map(|(key, value)| (key.clone(), value.clone()))
                    .collect(),
            )
            .map_err(js_err)?;
        Ok(json!({ "revisionId": receipt.revision_ids.into_iter().next() }).to_string())
    }

    /// Sets the authored `value` on a stable-id content-control embed.
    pub fn set_content_control_value(
        &self,
        embed_id: &str,
        value_json: &str,
    ) -> Result<(), JsValue> {
        let value = Any::from_json(value_json).map_err(js_err)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_embed_attrs_by_id(&ctx, embed_id, vec![("value".to_owned(), value)])
            .map(|_| ())
            .map_err(js_err)
    }

    /// Sets the authored `value` on a content-control embed at a paragraph-keyed
    /// position. This is the fallback for valid controls that have no authored
    /// `w:id`/tag and therefore cannot be addressed by stable payload identity.
    pub fn set_content_control_value_at(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        value_json: &str,
    ) -> Result<(), JsValue> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let value = Any::from_json(value_json).map_err(js_err)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_embed_attrs(
                &ctx,
                Position::new(story, index),
                vec![("value".to_owned(), value)],
            )
            .map(|_| ())
            .map_err(js_err)
    }

    /// Clears the authored `value` from a stable-id content-control embed.
    pub fn clear_content_control_value(&self, embed_id: &str) -> Result<(), JsValue> {
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_embed_attrs_by_id(&ctx, embed_id, vec![("value".to_owned(), Any::Null)])
            .map(|_| ())
            .map_err(js_err)
    }

    /// Commits image geometry fields to a stable-id image embed in one
    /// transaction. `null` fields clear; `other` is flattened into the payload.
    pub fn set_image_geometry(&self, embed_id: &str, geometry_json: &str) -> Result<(), JsValue> {
        let value: Value = serde_json::from_str(geometry_json).map_err(js_err)?;
        let object = value
            .as_object()
            .ok_or_else(|| js_err("set_image_geometry expects a JSON object"))?;
        let mut entries = Vec::new();
        for (key, value) in object {
            if key == "other" {
                let other = value
                    .as_object()
                    .ok_or_else(|| js_err("image geometry \"other\" must be an object"))?;
                for (other_key, other_value) in other {
                    entries.push((other_key.clone(), json_to_any(other_value)?));
                }
            } else {
                entries.push((key.clone(), json_to_any(value)?));
            }
        }
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_embed_attrs_by_id(&ctx, embed_id, entries)
            .map(|_| ())
            .map_err(js_err)
    }

    /// Inserts a native page-break embed at a Loc.
    pub fn insert_page_break(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
    ) -> Result<(), JsValue> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .insert_embed(&ctx, Position::new(story, index), "pageBreak", vec![])
            .map(|_| ())
            .map_err(js_err)
    }

    /// Inserts a native section-break embed at a Loc.
    pub fn insert_section_break(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        break_type: &str,
    ) -> Result<(), JsValue> {
        if !matches!(
            break_type,
            "nextPage" | "continuous" | "oddPage" | "evenPage"
        ) {
            return Err(js_err(format!("unknown section break type {break_type:?}")));
        }
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .insert_embed(
                &ctx,
                Position::new(story, index),
                "sectionBreak",
                vec![("type".to_owned(), Any::from(break_type))],
            )
            .map(|_| ())
            .map_err(js_err)
    }

    /// Inserts a typed watermark embed at a Loc.
    pub fn insert_watermark(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        watermark_json: &str,
    ) -> Result<(), JsValue> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let value: Value = serde_json::from_str(watermark_json).map_err(js_err)?;
        if !value.is_object() {
            return Err(js_err("insert_watermark expects a JSON object"));
        }
        let payload = parse_payload(Some(&value))?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .insert_embed(&ctx, Position::new(story, index), "watermark", payload)
            .map(|_| ())
            .map_err(js_err)
    }

    /// Sets one paragraph property (any JSON value) on `para_id`'s pilcrow.
    /// `paraId` / the embed discriminator are reserved.
    pub fn set_paragraph_attr(
        &self,
        para_id: &str,
        key: &str,
        value_json: &str,
    ) -> Result<(), JsValue> {
        let value = Any::from_json(value_json).map_err(js_err)?;
        self.engine
            .doc()
            .set_paragraph_attr(para_id, key, value)
            .map_err(js_err)
    }

    /// Adds a sticky-anchored comment. `ranges_json`:
    /// `[{"story","startPara","startOffset","endPara","endOffset"}, …]`;
    /// `body_json` is any JSON value. Receipt: `{"commentId"}`.
    pub fn add_comment(
        &self,
        ranges_json: &str,
        author: &str,
        date: &str,
        body_json: &str,
    ) -> Result<String, JsValue> {
        let value: Value = serde_json::from_str(ranges_json).map_err(js_err)?;
        let entries = value
            .as_array()
            .ok_or_else(|| js_err("add_comment expects an array of ranges"))?;
        let mut ranges = Vec::with_capacity(entries.len());
        for entry in entries {
            let get_str = |key: &str| {
                entry
                    .get(key)
                    .and_then(Value::as_str)
                    .ok_or_else(|| js_err(format!("a comment range requires a string {key:?}")))
            };
            let get_offset = |key: &str| {
                entry
                    .get(key)
                    .and_then(Value::as_u64)
                    .ok_or_else(|| {
                        js_err(format!("a comment range requires a non-negative {key:?}"))
                    })
                    .map(|offset| offset as u32)
            };
            let story = get_str("story")?;
            let start = loc_index(
                self.engine.doc(),
                story,
                get_str("startPara")?,
                get_offset("startOffset")?,
            )?;
            let end = loc_index(
                self.engine.doc(),
                story,
                get_str("endPara")?,
                get_offset("endOffset")?,
            )?;
            ranges.push(StoryRange::new(story, start, end));
        }
        let body = Any::from_json(body_json).map_err(js_err)?;
        let comment_id = self
            .engine
            .doc()
            .add_comment(&ranges, author, date, body)
            .map_err(js_err)?;
        Ok(json!({ "commentId": comment_id }).to_string())
    }

    /// Accepts tracked changes (S4b): pending insertions become plain content,
    /// pending deletions are carried out; `pPrIns` marks clear (the split
    /// stays), `pPrDel` marks join with the following paragraph (its pPr
    /// survives). `target_json`: `{"revisionId": string}` for one coalesced
    /// revision (any story) or
    /// `{"story","startPara","startOffset","endPara","endOffset"}` for a Loc
    /// range. Receipt: `{"revisionIds": [string, …]}` — the revision ids
    /// resolved. Resolving never stamps a new revision.
    pub fn accept_change(&self, target_json: &str) -> Result<String, JsValue> {
        let target = parse_change_target(self.engine.doc(), target_json)?;
        let ctx = EditCtx::local(String::new(), String::new());
        let receipt = self
            .engine
            .doc()
            .accept_change(&ctx, &target)
            .map_err(js_err)?;
        Ok(json!({ "revisionIds": receipt.revision_ids }).to_string())
    }

    /// Rejects tracked changes — the inverse of [`EditSession::accept_change`]:
    /// pending insertions roll back, pending deletions restore their text;
    /// `pPrIns` marks join back with the following paragraph, `pPrDel` marks
    /// clear (the split stays). Same target and receipt shapes.
    pub fn reject_change(&self, target_json: &str) -> Result<String, JsValue> {
        let target = parse_change_target(self.engine.doc(), target_json)?;
        let ctx = EditCtx::local(String::new(), String::new());
        let receipt = self
            .engine
            .doc()
            .reject_change(&ctx, &target)
            .map_err(js_err)?;
        Ok(json!({ "revisionIds": receipt.revision_ids }).to_string())
    }

    /// Applies a batch of raw story mutations in ONE transaction — the
    /// coexistence bridge's mirror-into-yrs path (not a user-intent op). `ops_json`
    /// is `[{ "op":"insert"|"delete"|"format"|"insertEmbed"|"setEmbedAttr"
    /// |"setComment"|"removeComment", "index", … }, …]`; each op's index (and each
    /// `setComment` `[start, end)` range) is read against the story state after all
    /// prior ops in the batch. Attributes/payloads are faithful mirrors of the
    /// bridge's lowered PM state (tracked-change stamps arrive inside `attrs`;
    /// comments are keyed by the PM comment id and anchored sticky, side-map only).
    pub fn apply_raw_ops(&self, story: &str, ops_json: &str) -> Result<(), JsValue> {
        let value: Value = serde_json::from_str(ops_json).map_err(js_err)?;
        let entries = value
            .as_array()
            .ok_or_else(|| js_err("apply_raw_ops expects an array of ops"))?;
        let ops = entries
            .iter()
            .map(parse_raw_op)
            .collect::<Result<Vec<_>, _>>()?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .apply_raw_ops(story, ops, &ctx)
            .map_err(js_err)
    }

    // -- read queries (pure snapshots; JSON out) --

    /// Aggregated toolbar/a11y state over one paragraph-addressed story
    /// range. Toggle marks are `true`, `false`, or `"mixed"`; value marks
    /// are their uniform value or `null` when absent/mixed.
    #[allow(clippy::too_many_arguments)]
    pub fn selection_context(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
    ) -> Result<String, JsValue> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let context = self
            .engine
            .doc()
            .selection_context(&StoryRange::new(story, start, end))
            .map_err(js_err)?;
        let is_single_embed = context.embed_kind.is_some();
        let is_image = context.embed_kind.as_deref() == Some("image");
        Ok(json!({
            "bold": tri_state_value(context.bold),
            "italic": tri_state_value(context.italic),
            "underline": tri_state_value(context.underline),
            "strike": tri_state_value(context.strike),
            "fontFamily": context.font_family,
            "fontSize": context.font_size,
            "color": context.color,
            "paraId": context.para_id,
            "styleId": context.style_id,
            "alignment": context.alignment,
            "paragraphProperties": attrs_value(&context.paragraph_properties)?,
            "hasSelection": context.has_selection,
            "isMultiParagraph": context.is_multi_paragraph,
            "inTable": context.in_table,
            "isSingleEmbed": is_single_embed,
            "embedKind": context.embed_kind,
            "isImage": is_image,
            "inInsertion": context.in_insertion,
            "inDeletion": context.in_deletion,
        })
        .to_string())
    }

    /// Every tracked-change run/paragraph-mark revision across all stories,
    /// in deterministic story/position order.
    pub fn list_revisions(&self) -> Result<String, JsValue> {
        let revisions = self.engine.doc().list_revisions().map_err(js_err)?;
        let items: Vec<Value> = revisions
            .into_iter()
            .map(|revision| {
                let kind = match revision.change.kind {
                    ChangeKind::Insertion => "insertion",
                    ChangeKind::Deletion => "deletion",
                    ChangeKind::ParagraphMarkInsertion => "pPrIns",
                    ChangeKind::ParagraphMarkDeletion => "pPrDel",
                    ChangeKind::ParagraphPropertiesChanged => "pPrChange",
                    ChangeKind::TableRowInsertion => "trIns",
                    ChangeKind::TableRowDeletion => "trDel",
                    ChangeKind::TableInsertion => "tableIns",
                    ChangeKind::TableDeletion => "tableDel",
                };
                json!({
                    "revisionId": revision.change.revision_id,
                    "author": revision.change.author,
                    "date": revision.change.date,
                    "kind": kind,
                    "story": revision.story,
                    "preview": revision.preview,
                    "range": {
                        "story": revision.change.range.start.story,
                        "start": {
                            "paraId": revision.change.range.start.para,
                            "offset": revision.change.range.start.offset,
                        },
                        "end": {
                            "paraId": revision.change.range.end.para,
                            "offset": revision.change.range.end.offset,
                        },
                    },
                })
            })
            .collect();
        serde_json::to_string(&items).map_err(js_err)
    }

    /// Story ids currently in the document, sorted for determinism.
    pub fn story_ids(&self) -> Vec<String> {
        let txn = self.engine.doc().yrs_doc().transact();
        let Some(stories) = txn.get_map(STORIES) else {
            return Vec::new();
        };
        let mut ids: Vec<String> = stories.iter(&txn).map(|(id, _)| id.to_string()).collect();
        ids.sort();
        ids
    }

    /// Story length in UTF-16 units (every embed, pilcrows included, = 1).
    pub fn story_len(&self, story: &str) -> Result<u32, JsValue> {
        self.engine.doc().story_len(story).map_err(js_err)
    }

    /// The story's `canonical-stream-v1` FNV-1a checksum as a decimal string
    /// (u64 exceeds JS safe-integer range). The coexistence watchdog compares
    /// this against the PM projector's checksum after every mirrored edit.
    pub fn story_checksum(&self, story: &str) -> Result<String, JsValue> {
        crate::story_checksum(self.engine.doc(), story)
            .map(|checksum| checksum.to_string())
            .map_err(js_err)
    }

    /// Lowers a story straight to the renderer's `LayoutBlock[]` (JSON) — the
    /// yrs-authoritative render path (the eventual replacement for the TS
    /// `toLayoutBlocks(pmDoc)`). Errors with an unsupported-embed message on any
    /// non-native content (e.g. an opaque page-break blob) until that class is
    /// promoted to native; the host falls back to the PM render path there.
    /// `env_json` carries theme colors, the default tab stop, and list numeric
    /// ids (see [`parse_render_env`]).
    pub fn yrs_blocks_for_story(&self, story: &str, env_json: &str) -> Result<String, JsValue> {
        let env = parse_render_env(env_json)?;
        self.engine.lower_story_json(story, &env).map_err(js_err)
    }

    /// `[{"paraId","text","properties"}]` in document order. `properties`
    /// carries pStyle/alignment plus any op-set extras.
    pub fn paragraphs(&self, story: &str) -> Result<String, JsValue> {
        let paragraphs = self.engine.doc().paragraphs(story).map_err(js_err)?;
        let items = paragraphs
            .into_iter()
            .map(|paragraph| {
                Ok(json!({
                    "paraId": paragraph.para_id,
                    "text": paragraph.text,
                    "properties": attrs_value(&paragraph.properties)?,
                }))
            })
            .collect::<Result<Vec<Value>, JsValue>>()?;
        serde_json::to_string(&items).map_err(js_err)
    }

    /// Compact paragraph-position projection in one story traversal:
    /// `[{"paraId","length"}]`. Length counts UTF-16 text and inline embed
    /// units before each paragraph's pilcrow. The JS input shim uses this
    /// instead of crossing the wasm boundary once per paragraph.
    pub fn paragraph_spans(&self, story: &str) -> Result<String, JsValue> {
        let mut items = Vec::new();
        let mut cursor = 0_u32;
        let mut paragraph_start = 0_u32;
        for segment in self.engine.doc().story_segments(story).map_err(js_err)? {
            match segment.content {
                SegmentContent::Text(text) => {
                    cursor += text.encode_utf16().count() as u32;
                }
                SegmentContent::Pilcrow(properties) => {
                    items.push(json!({
                        "paraId": properties.para_id,
                        "length": cursor - paragraph_start,
                    }));
                    cursor += 1;
                    paragraph_start = cursor;
                }
                SegmentContent::OtherEmbed { .. } => cursor += 1,
            }
        }
        serde_json::to_string(&items).map_err(js_err)
    }

    /// The raw formatted-segment view (the render bridge's input):
    /// `[{"kind":"text","text",…} | {"kind":"pilcrow","paraId","properties",…}
    /// | {"kind":"embed",…}]`, each with `"attributes"` (run marks plus
    /// `ins`/`del` revision values).
    pub fn story_segments(&self, story: &str) -> Result<String, JsValue> {
        let segments = self.engine.doc().story_segments(story).map_err(js_err)?;
        let items = segments
            .into_iter()
            .map(|segment| {
                let attributes = attrs_value(&segment.attributes)?;
                Ok(match segment.content {
                    SegmentContent::Text(text) => {
                        json!({ "kind": "text", "text": text, "attributes": attributes })
                    }
                    SegmentContent::Pilcrow(properties) => json!({
                        "kind": "pilcrow",
                        "paraId": properties.para_id,
                        "properties": attrs_value(&properties.values)?,
                        "attributes": attributes,
                    }),
                    SegmentContent::OtherEmbed { kind, payload } => {
                        json!({
                            "kind": "embed",
                            "embedKind": kind,
                            "payload": attrs_value(&payload)?,
                            "attributes": attributes,
                        })
                    }
                })
            })
            .collect::<Result<Vec<Value>, JsValue>>()?;
        serde_json::to_string(&items).map_err(js_err)
    }

    /// `{"start","end"}` — the paragraph's story span; `end` is its pilcrow's
    /// index, so `end - start` is the paragraph length (`offset` domain).
    pub fn locate_paragraph(&self, story: &str, para_id: &str) -> Result<String, JsValue> {
        let span = find_para_span(self.engine.doc(), story, para_id)?;
        Ok(json!({ "start": span.start, "end": span.pilcrow }).to_string())
    }

    /// Current offsets of a comment's sticky anchors:
    /// `[{"story","start","end"}]`. Errors when an anchor no longer resolves.
    pub fn resolve_comment(&self, comment_id: &str) -> Result<String, JsValue> {
        let anchors = self
            .engine
            .doc()
            .resolve_comment(comment_id)
            .map_err(js_err)?;
        let items: Vec<Value> = anchors
            .into_iter()
            .map(
                |anchor| json!({ "story": anchor.story, "start": anchor.start, "end": anchor.end }),
            )
            .collect();
        serde_json::to_string(&items).map_err(js_err)
    }
}
