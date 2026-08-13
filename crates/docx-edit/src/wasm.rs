//! wasm-bindgen session boundary over [`EditingDoc`] — the only JS-visible
//! surface of this crate, compiled behind `--features wasm`.
//!
//! Conventions every entry point below obeys, so its own docs need only state
//! what is specific to it:
//!
//! - **Values cross as JSON strings or raw bytes.** Arguments named `*_json`
//!   are JSON text; a method returning `Result<String, _>` returns JSON text.
//!   Yrs updates and binary display frames cross as byte slices / `Vec<u8>`
//!   (wasm-bindgen exposes the latter as a transferable `Uint8Array`).
//! - **Errors cross as `Err(JsValue)` carrying a display string** and nothing
//!   else — there is no error code or structured payload. Every entry point
//!   taking JSON fails that way on malformed JSON or a wrong value type, and
//!   every entry point addressing content fails that way on an unknown story,
//!   an unknown paragraph, or an out-of-range offset. A failed call leaves the
//!   document untouched: arguments are fully validated before any mutation,
//!   and each op commits in a single yrs transaction.
//! - **Content is addressed as `Loc { story, paraId, offset }`.** `offset`
//!   counts UTF-16 units within ONE paragraph and lives in `[0, para_len]`,
//!   where `para_len` excludes the paragraph's own pilcrow; `offset ==
//!   para_len` addresses the paragraph mark. A paragraph id is resolved
//!   story-scoped, so an id living in another story reads as "not found".
//!   Story-global indices exist only transiently inside a call and are never
//!   returned.
//! - **A range is two Locs in one story** and is half-open, `[start, end)`.
//!   Because pilcrows occupy a unit, a range whose ends sit in different
//!   paragraphs spans the boundary marks.
//! - **Suggesting mode is the optional `(author_name, author_date)` pair.**
//!   Pass both to record the edit as a tracked change stamped with that author
//!   and ISO date, or neither for a plain local edit; passing one alone is an
//!   error. Ops without the pair are always plain and local.
//! - **Receipts are JSON objects.** An op that can stamp a tracked change
//!   reports `{"revisionId": string|null}` — null outside suggesting mode.
//!   Ops with no receipt content return `()`.
//!
//! Story lengths, selection indices and every other unit count in this module
//! are UTF-16 units in which each embed, pilcrows included, counts as one.

use std::cell::RefCell;
use std::collections::{BTreeMap, HashMap, VecDeque};
use std::rc::Rc;

use js_sys::{Function, Uint8Array};
use serde::Serialize;
use serde_json::{Value, json};
use wasm_bindgen::prelude::*;
use yrs::{Any, Assoc, IndexedSequence, Map, ReadTxn, StickyIndex, Subscription, Transact};

use crate::presence::{
    apply_update_with_typing_inference, encode_sticky, resolve_sticky_selection,
};
use crate::{
    CellLoc, ChangeKind, ChangeTarget, ColorPatch, EditCtx, EditingDoc, EngineSession,
    FontFamilyPatch, FormatPolicy, InlineFormatDelta, MergeDirection, ParaAttrDelta, ParaSelector,
    Patch, Position, RawOp, SegmentContent, SimpleFormat, StoryRange, TabStop, TableLocator,
    TableRange, TriState, UndoSession, story_ref,
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

struct IndexedLoc {
    para_id: String,
    offset: u32,
    node_offset: u32,
}

#[derive(Clone, Copy)]
enum DeleteDirection {
    Backward,
    Forward,
}

#[derive(Clone, Copy)]
enum AdjacentStoryUnit {
    Content(u32),
    Pilcrow,
}

/// Whether the layout gives an embed its own block.
fn is_block_embed(kind: &str) -> bool {
    matches!(kind, "table" | "blockSdt" | "pageBreak" | "columnBreak")
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
fn index_loc(doc: &EditingDoc, story: &str, index: u32) -> Result<IndexedLoc, JsValue> {
    let mut cursor = 0_u32;
    let mut para_start = 0_u32;
    let mut node_start = 0_u32;
    for segment in doc.story_segments(story).map_err(js_err)? {
        match segment.content {
            SegmentContent::Text(text) => cursor += text.encode_utf16().count() as u32,
            SegmentContent::Pilcrow(properties) => {
                if index <= cursor {
                    return Ok(IndexedLoc {
                        para_id: properties.para_id,
                        offset: index.saturating_sub(para_start),
                        node_offset: index.saturating_sub(node_start),
                    });
                }
                cursor += 1;
                para_start = cursor;
                node_start = cursor;
            }
            SegmentContent::OtherEmbed { ref kind, .. } => {
                if cursor == node_start && is_block_embed(kind) {
                    node_start = cursor + 1;
                }
                cursor += 1;
            }
        }
    }
    Err(js_err(format!(
        "selection index {index} does not resolve in story {story:?}"
    )))
}

fn adjacent_story_unit(
    doc: &EditingDoc,
    story: &str,
    index: u32,
    direction: DeleteDirection,
) -> Result<Option<AdjacentStoryUnit>, JsValue> {
    let mut cursor = 0_u32;
    for segment in doc.story_segments(story).map_err(js_err)? {
        match segment.content {
            SegmentContent::Text(text) => {
                let units: Vec<u16> = text.encode_utf16().collect();
                let end = cursor + units.len() as u32;
                let relative = match direction {
                    DeleteDirection::Backward if index > cursor && index <= end => {
                        Some((index - cursor) as usize)
                    }
                    DeleteDirection::Forward if index >= cursor && index < end => {
                        Some((index - cursor) as usize)
                    }
                    _ => None,
                };
                if let Some(relative) = relative {
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
                    return Ok(Some(AdjacentStoryUnit::Content(width)));
                }
                cursor = end;
            }
            SegmentContent::Pilcrow(_) => {
                let adjacent = match direction {
                    DeleteDirection::Backward => index == cursor + 1,
                    DeleteDirection::Forward => index == cursor,
                };
                if adjacent {
                    return Ok(Some(AdjacentStoryUnit::Pilcrow));
                }
                cursor += 1;
            }
            SegmentContent::OtherEmbed { .. } => {
                let adjacent = match direction {
                    DeleteDirection::Backward => index == cursor + 1,
                    DeleteDirection::Forward => index == cursor,
                };
                if adjacent {
                    return Ok(Some(AdjacentStoryUnit::Content(1)));
                }
                cursor += 1;
            }
        }
    }
    Ok(None)
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

/// Applies one boundary mark to `range`. `mark_json`:
/// `{"type":"bold"|"italic"|"underline"|"strike"|"superscript"|"subscript"} |
/// {"type":"fontFamily"|"color","value":string} |
/// {"type":"fontSize","value":number}`. The six simple types toggle; font,
/// size and color set.
fn apply_mark(doc: &EditingDoc, range: StoryRange, mark_json: &str) -> Result<(), JsValue> {
    let value: Value = serde_json::from_str(mark_json).map_err(js_err)?;
    let kind = value
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| js_err("mark JSON requires a string \"type\""))?;
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

/// Decodes a tri-state inline-formatting delta. An omitted key is
/// [`Patch::Keep`], `null` is [`Patch::Clear`], any typed value is
/// [`Patch::Set`]. Boolean `false` also clears.
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

/// Parses one `{ "op", "index", … }` raw operation.
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
    docx_source: RefCell<Option<Vec<u8>>>,
    update_observer: Option<Subscription>,
    update_event_observer: Option<UpdateEventObserver>,
    undo: UndoSession,
    selection: RefCell<Option<LocalSelection>>,
    cell_selection: RefCell<Option<LocalCellSelection>>,
    last_apply_profile_json: RefCell<String>,
}

struct UpdateEventObserver {
    pending: Rc<RefCell<VecDeque<Vec<u8>>>>,
    _subscription: Subscription,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct DocxHostWire {
    envelope: docx_parse::S9WireEnvelope,
    referenced_fonts: Vec<String>,
}

fn thin_header_footer(
    entries: &Option<Vec<(String, docx_parse::HeaderFooter)>>,
) -> Option<Vec<(String, docx_parse::HeaderFooter)>> {
    entries.as_ref().map(|entries| {
        entries
            .iter()
            .map(|(relationship_id, part)| {
                (
                    relationship_id.clone(),
                    docx_parse::HeaderFooter {
                        story_type: part.story_type.clone(),
                        hdr_ftr_type: part.hdr_ftr_type.clone(),
                        content: Vec::new(),
                        watermark: part.watermark.clone(),
                    },
                )
            })
            .collect()
    })
}

fn thin_notes(notes: &Option<Vec<docx_parse::Note>>) -> Option<Vec<docx_parse::Note>> {
    notes.as_ref().map(|notes| {
        notes
            .iter()
            .map(|note| docx_parse::Note {
                story_type: note.story_type.clone(),
                id: note.id,
                note_type: note.note_type.clone(),
                content: Vec::new(),
                verbatim_xml: None,
            })
            .collect()
    })
}

fn thin_docx_envelope(envelope: &docx_parse::S9WireEnvelope) -> docx_parse::S9WireEnvelope {
    let package = &envelope.document.package;
    let sections = package.document.sections.as_ref().map(|sections| {
        sections
            .iter()
            .map(|section| docx_parse::S9SectionWire {
                id: section.id.clone(),
                properties: section.properties.clone(),
                content_start: 0,
                content_end: 0,
            })
            .collect()
    });
    docx_parse::S9WireEnvelope {
        wire_version: envelope.wire_version,
        document: docx_parse::S9DocumentWire {
            package: docx_parse::S9PackageWire {
                document: docx_parse::S9DocumentBodyWire {
                    content: Vec::new(),
                    sections,
                    final_section_properties: package.document.final_section_properties.clone(),
                    comments: package.document.comments.clone(),
                },
                styles: package.styles.clone(),
                theme: package.theme.clone(),
                numbering: docx_parse::NumberingDefinitions::default(),
                settings: package.settings.clone(),
                font_table: package.font_table.clone(),
                header_entries: thin_header_footer(&package.header_entries),
                footer_entries: thin_header_footer(&package.footer_entries),
                footnotes: thin_notes(&package.footnotes),
                endnotes: thin_notes(&package.endnotes),
                footnote_separators: None,
                endnote_separators: None,
                relationship_entries: package.relationship_entries.clone(),
                media_entries: Vec::new(),
                chart_entries: Vec::new(),
            },
            template_variables: None,
            warnings: envelope.document.warnings.clone(),
        },
        embedded_font_parts: envelope.embedded_font_parts.clone(),
        font_table_relationships_xml: envelope.font_table_relationships_xml.clone(),
        canonical_base64: None,
        canonical_sha256: None,
    }
}

impl EditSession {
    fn open_docx_inner(&self, bytes: &[u8], seed_stories: bool) -> Result<String, JsValue> {
        let envelope = crate::seed::parse_docx_for_edit(bytes).map_err(js_err)?;
        let host_envelope = thin_docx_envelope(&envelope);
        let referenced_fonts = if seed_stories {
            crate::seed::seed_parsed_docx(self.engine.doc(), envelope).map_err(js_err)?
        } else {
            let fonts = crate::seed::referenced_fonts(&envelope).map_err(js_err)?;
            drop(envelope);
            fonts
        };
        let host = DocxHostWire {
            envelope: host_envelope,
            referenced_fonts,
        };
        let json = serde_json::to_string(&host).map_err(js_err)?;
        self.docx_source.replace(Some(bytes.to_vec()));
        Ok(json)
    }

    fn collapsed_resident_input_selection(&self) -> Result<(String, String, u32), JsValue> {
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
        let loc = index_loc(self.engine.doc(), &story, head)?;
        if !self.engine.can_apply_input(&story, &loc.para_id) {
            return Err(js_err(
                "resident input state is not ready for this paragraph",
            ));
        }
        Ok((story, loc.para_id, head))
    }

    fn delete_resident_input(
        &self,
        direction: &str,
        selection: (String, String, u32),
    ) -> Result<String, JsValue> {
        let (story, para_id, head) = selection;
        let direction = match direction {
            "backward" => DeleteDirection::Backward,
            "forward" => DeleteDirection::Forward,
            _ => return Err(js_err("delete direction must be backward or forward")),
        };
        let adjacent = adjacent_story_unit(self.engine.doc(), &story, head, direction)?;
        let ctx = EditCtx::local("", "");

        match (direction, adjacent) {
            (DeleteDirection::Backward, Some(AdjacentStoryUnit::Content(width))) => {
                self.engine
                    .doc()
                    .delete_range(&ctx, StoryRange::new(&story, head - width, head))
                    .map_err(js_err)?;
            }
            (DeleteDirection::Forward, Some(AdjacentStoryUnit::Content(width))) => {
                self.engine
                    .doc()
                    .delete_range(&ctx, StoryRange::new(&story, head, head + width))
                    .map_err(js_err)?;
            }
            (direction, Some(AdjacentStoryUnit::Pilcrow)) => {
                let paragraphs = self.engine.doc().paragraphs(&story).map_err(js_err)?;
                let paragraph_index = paragraphs
                    .iter()
                    .position(|paragraph| paragraph.para_id == para_id)
                    .ok_or_else(|| js_err("resident input paragraph no longer resolves"))?;
                let merge = match direction {
                    DeleteDirection::Backward if paragraph_index > 0 => {
                        Some(MergeDirection::Backward)
                    }
                    DeleteDirection::Forward if paragraph_index + 1 < paragraphs.len() => {
                        Some(MergeDirection::Forward)
                    }
                    _ => None,
                };
                let Some(merge) = merge else {
                    return Err(js_err("resident input has no character in that direction"));
                };
                self.engine
                    .doc()
                    .merge_paragraphs(&ctx, &para_id, merge)
                    .map_err(js_err)?;
            }
            (_, None) => {
                return Err(js_err("resident input has no character in that direction"));
            }
        }
        Ok(story)
    }
}

#[wasm_bindgen]
impl EditSession {
    /// Creates a replica. The host allocates `client_id` (a random 32-bit id
    /// is fine) and must keep it unique across the replicas that will merge.
    /// Errors unless it is a non-negative safe integer.
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
            docx_source: RefCell::new(None),
            update_observer: None,
            update_event_observer: None,
            undo: UndoSession::new(),
            selection: RefCell::new(None),
            cell_selection: RefCell::new(None),
            last_apply_profile_json: RefCell::new("{}".to_owned()),
        })
    }

    /// This replica's client id, as passed to the constructor.
    pub fn client_id(&self) -> f64 {
        self.engine.doc().client_id() as f64
    }

    /// Registers raw sfnt bytes in this session's resident measurement store
    /// and returns the font id that measurement and display inputs reference.
    /// Errors on bytes the font parser rejects.
    pub fn register_measure_font(&self, bytes: &[u8]) -> Result<u32, JsValue> {
        docx_layout::register_measure_font(bytes)
    }

    /// Drops every registered measurement font (ids restart at zero) and
    /// invalidates the retained paragraph measurement templates, so the next
    /// edit must pass back through the full layout path.
    pub fn clear_measure_fonts(&self) {
        docx_layout::clear_measure_fonts();
        self.engine.clear_measurement_templates();
    }

    /// Measurement input JSON in, `ParagraphExtent` JSON out. Also records the
    /// paragraph's immutable width/font envelope under its stable block id, so
    /// a later resident edit re-measures only the changed block. Errors with
    /// the engine's message for input it cannot measure.
    pub fn measure_paragraph_json(&self, input: &str) -> Result<String, JsValue> {
        self.engine.measure_paragraph_json(input).map_err(js_err)
    }

    /// `{ measured, options }` JSON in, `Layout` JSON out. Both the measured
    /// input and the resulting layout are retained for the resident edit path.
    /// Errors on unparseable input or on layout failure.
    pub fn layout_document_json(&self, input: &str) -> Result<String, JsValue> {
        self.engine
            .layout_document_json(input)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Region-layout input JSON in, the font families and sizes that input
    /// needs as JSON out, so the host can register fonts before laying out.
    pub fn layout_font_requirements_json(&self, input: &str) -> Result<String, JsValue> {
        self.engine
            .layout_font_requirements_json(input)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Region-layout input JSON in, a paginated envelope with section and page
    /// regions already composed as JSON out — ready for the display builder
    /// with no host-side layout mutation. Retains the pass for resident edits.
    pub fn layout_document_with_regions_json(&self, input: &str) -> Result<String, JsValue> {
        self.engine
            .layout_document_with_regions_json(input)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// `{ measured, options, layout }` JSON in, `DisplayList` JSON out, built
    /// against the same resident font store this session measures with.
    pub fn build_display_list_json(&self, input: &str) -> Result<String, JsValue> {
        self.engine
            .build_display_list_json(input)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Display-only input JSON in, one binary `FrameDelta` v1 out (exposed as
    /// a transferable `Uint8Array`). `expected_frame_epoch` is the epoch of the
    /// frame the caller currently holds; pass `0` for the first frame. A
    /// mismatch makes the engine emit a full frame instead of a delta. Errors
    /// unless the epoch is a non-negative safe integer, and on build failure.
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

    /// `{"frameEpoch", "caretRect": {…}|null}` for the session's own collapsed
    /// body selection. `caretRect` is null whenever there is no selection, the
    /// selection is not a collapsed body caret, or the retained layout has no
    /// geometry for it. `frameEpoch` identifies the frame the rect belongs to.
    pub fn resident_caret_snapshot_json(&self) -> Result<String, JsValue> {
        let paragraph = {
            let selection = self.selection.borrow();
            let Some(selection) = selection.as_ref() else {
                return serde_json::to_string(
                    &self.engine.resident_caret_snapshot(None).map_err(js_err)?,
                )
                .map_err(js_err);
            };
            if selection.story != "body" {
                None
            } else {
                let txn = self.engine.doc().yrs_doc().transact();
                let anchor = selection.anchor.get_offset(&txn).map(|offset| offset.index);
                let head = selection.head.get_offset(&txn).map(|offset| offset.index);
                drop(txn);
                match (anchor, head) {
                    (Some(anchor), Some(head)) if anchor == head => {
                        let loc = index_loc(self.engine.doc(), &selection.story, head)?;
                        Some((loc.para_id, loc.node_offset))
                    }
                    _ => None,
                }
            }
        };
        let snapshot = self
            .engine
            .resident_caret_snapshot(
                paragraph
                    .as_ref()
                    .map(|(id, offset)| (id.as_str(), *offset)),
            )
            .map_err(js_err)?;
        serde_json::to_string(&snapshot).map_err(js_err)
    }

    /// Applies one ordinary insertion at this session's collapsed selection
    /// and returns the resulting binary `FrameDelta`. The inserted text
    /// inherits the formatting at the caret; selection, measurement inputs,
    /// pagination checkpoints and display state all stay resident, so nothing
    /// but the frame crosses the boundary.
    ///
    /// Errors when `text` is empty or contains `\r`/`\n`, when
    /// `expected_frame_epoch` is not a non-negative safe integer, when there is
    /// no resident selection or it no longer resolves, when the selection is
    /// not collapsed, or when the resident layout state cannot absorb an edit
    /// in that paragraph. The caller must then run the full layout path.
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
        let loc = index_loc(self.engine.doc(), &story, head)?;
        if !self.engine.can_apply_input(&story, &loc.para_id) {
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

    /// Instrumented twin of [`EditSession::apply_input`]: identical arguments,
    /// result and error contract, but it also records stage timings for
    /// [`EditSession::apply_input_profile_json`]. Separate so the ordinary
    /// input path pays no timer calls.
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
        let loc = index_loc(self.engine.doc(), &story, head)?;
        if !self.engine.can_apply_input(&story, &loc.para_id) {
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

    /// Deletes one character at this session's collapsed selection and returns
    /// the resulting binary `FrameDelta`. `direction` is `"backward"` or
    /// `"forward"`; a surrogate pair is removed whole. At a paragraph boundary
    /// this merges with the neighbouring paragraph instead.
    ///
    /// Errors on an unknown `direction`, when `expected_frame_epoch` is not a
    /// non-negative safe integer, under the same selection and readiness
    /// conditions as [`EditSession::apply_input`], and when there is no
    /// character to delete in that direction (document start or end).
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

    /// Instrumented twin of [`EditSession::apply_delete`]: identical arguments,
    /// result and error contract, but it also records stage timings for
    /// [`EditSession::apply_input_profile_json`].
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

    /// Stage timings of the last profiled apply, in milliseconds:
    /// `{"selectionMs","editMs","lowerMs","measureMs","paginateMs",
    /// "displayInputMs","displayBuildMs","displayFinalizeMs","displayMs",
    /// "encodeMs"}`. `{}` before the first profiled call.
    pub fn apply_input_profile_json(&self) -> String {
        self.last_apply_profile_json.borrow().clone()
    }

    /// Region-aware hit test against the resident display list, so no
    /// display-list JSON crosses the boundary. `x`/`y` are page-local px.
    /// Returns
    /// `{"region":"body"|"header"|"footer","rId"?,"pos":n|null,"target":"text"|"image"|"none"}`,
    /// or `"null"` for a page index outside the frame. A header/footer `pos`
    /// refers to that header/footer's own document, not the body; `target` is
    /// what sits under the point, which is what a pointer cursor needs.
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

    /// Nearest caret position on the adjacent visual line. `direction` is
    /// `"up"` or `"down"`; `goal_x` is the page-local x the caret is trying to
    /// hold across successive moves (a non-finite value falls back to the
    /// caret's own x). Returns `{"position","goalX"}`, or `"null"` when there
    /// is no line in that direction. Errors on an unknown `direction` and when
    /// no display list is resident.
    pub fn display_vertical_move_json(
        &self,
        position: f64,
        direction: &str,
        goal_x: f64,
    ) -> Result<String, JsValue> {
        self.engine
            .display_vertical_move_json(position as i64, direction, goal_x)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Highlight rectangles for a body document range against the resident
    /// display list: a JSON array of `{pageIndex,x,y,width,height}` in
    /// page-local px, one entry per page the range touches. Body positions
    /// only. Errors when no display list is resident.
    pub fn display_range_rects_json(&self, from: f64, to: f64) -> Result<String, JsValue> {
        self.engine
            .display_range_rects_json(from as i64, to as i64)
            .map_err(|error| JsValue::from_str(&error))
    }

    /// Same rectangles as [`EditSession::display_range_rects_json`], scoped to
    /// a region. `region` is `"body"`, `"header"` or `"footer"`; `r_id` names
    /// one header/footer part, and an empty string matches any. `from`/`to`
    /// are positions in THAT region's document, and a header/footer part
    /// paints on every page carrying it, so the result holds one rect set per
    /// such page, each tagged with its `pageIndex`. Errors on an unknown
    /// `region` and when no display list is resident.
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

    /// One glyph outline from this session's resident font store:
    /// `{"upem":n,"cmds":[{"t":"M"|"L"|"Q"|"C"|"Z", …}]}` — commands in font
    /// units, y-up. `font_id` comes from
    /// [`EditSession::register_measure_font`] and `glyph_id` from shaping.
    /// Errors when the glyph cannot be extracted.
    pub fn outline_glyph_json(&self, font_id: u32, glyph_id: u32) -> Result<String, JsValue> {
        docx_layout::outline_glyph_json(font_id, glyph_id)
    }

    /// Hydrates this replica from an encoded yrs v1 update, typically another
    /// replica's [`EditSession::encode_state`] output. Identical to
    /// [`EditSession::apply_update`]; the separate name marks the initial-load
    /// call site. Errors on a malformed update.
    pub fn load(&self, update: &[u8]) -> Result<(), JsValue> {
        self.engine.doc().apply_update_v1(update).map_err(js_err)
    }

    /// [`EditSession::open_docx`] with seeding always on.
    pub fn seed_from_docx(&self, bytes: &[u8]) -> Result<String, JsValue> {
        self.open_docx_inner(bytes, true)
    }

    /// Parses a DOCX package, optionally seeds its editable stories into this
    /// replica, and retains the source bytes for
    /// [`EditSession::materialize_docx`].
    ///
    /// Returns `{"envelope","referencedFonts":[string, …]}`. The envelope is
    /// the parsed package with the parts the host does not need stripped —
    /// body content, header/footer and note content, numbering, media and
    /// charts are emptied, section entries keep only their properties — so
    /// styles, theme, settings, fonts and relationships still cross while the
    /// bulk of the document stays in Rust. Errors on bytes that are not a
    /// readable DOCX.
    pub fn open_docx(&self, bytes: &[u8], seed_stories: bool) -> Result<String, JsValue> {
        self.open_docx_inner(bytes, seed_stories)
    }

    /// Re-parses the DOCX bytes retained by the last
    /// [`EditSession::open_docx`] and returns the COMPLETE package envelope as
    /// JSON, or `None` when no DOCX has been opened. Unlike `open_docx` this
    /// keeps every part; it reflects the source file, not later edits.
    pub fn materialize_docx(&self) -> Result<Option<String>, JsValue> {
        let source = self.docx_source.borrow();
        let Some(source) = source.as_ref() else {
            return Ok(None);
        };
        let envelope = crate::seed::parse_docx_for_edit(source).map_err(js_err)?;
        let json = serde_json::to_string(&envelope).map_err(js_err)?;
        Ok(Some(json))
    }

    /// Seeds stories from JSON:
    /// `[{"storyId","paragraphs":[{"text","pStyle"?,"alignment"?}, …]}, …]`.
    /// `pStyle` defaults to `"Normal"` and `alignment` to `"left"`; paragraph
    /// text must not contain paragraph breaks. Receipt:
    /// `{storyId: [paraId, …]}` with each story's paragraphs in document
    /// order. Errors when a story entry has no `storyId`, no `paragraphs`
    /// array, or an empty one, and when a story id already exists.
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
                // Splitting at the final pilcrow appends: the first half keeps
                // the original paraId, so the appended paragraph — whose
                // properties this seeds — is the SECOND half.
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

    /// The full document state as one yrs v1 update. Hand it to
    /// [`EditSession::load`] on a fresh replica to reproduce this document.
    pub fn encode_state(&self) -> Vec<u8> {
        self.engine.doc().encode_state_as_update_v1()
    }

    /// This replica's yrs v1 state vector — what a peer sends so
    /// [`EditSession::encode_diff`] can compute the update it is missing.
    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.engine.doc().encode_state_vector_v1()
    }

    /// The yrs v1 update carrying everything this replica has that the peer
    /// described by `remote_state_vector` does not. Errors on a malformed
    /// state vector.
    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>, JsValue> {
        self.engine
            .doc()
            .encode_diff_v1(remote_state_vector)
            .map_err(js_err)
    }

    /// Applies a remote or incremental yrs v1 update. It commits without a
    /// local origin, so this replica's undo manager never claims it. Errors on
    /// a malformed update.
    pub fn apply_update(&self, update: &[u8]) -> Result<(), JsValue> {
        self.engine.doc().apply_update_v1(update).map_err(js_err)
    }

    /// [`EditSession::apply_update`] plus an inference about where the peer
    /// that authored it is typing, for caret presence. Returns
    /// `{"clientId","story","paraId","endOffset"}` for the end of that peer's
    /// last text insertion, or `"null"` when the update carries work from more
    /// than one client, ends in a non-text insertion, or inserts nothing.
    /// Errors on a malformed update.
    pub fn apply_update_with_inference(&self, update: &[u8]) -> Result<String, JsValue> {
        let inference =
            apply_update_with_typing_inference(self.engine.doc(), update).map_err(js_err)?;
        serde_json::to_string(&inference.map(|inference| {
            json!({
                "clientId": inference.client_id,
                "story": inference.story,
                "paraId": inference.para_id,
                "endOffset": inference.end_offset,
            })
        }))
        .map_err(js_err)
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

    /// Subscribes `callback(update: Uint8Array, isRemote: 0|1)` to every
    /// committed transaction. `update` is v1-encoded — feed it straight to
    /// [`EditSession::apply_update`] on a peer — and is copied out of wasm
    /// memory, so JS may hold it across later edits. The second argument is
    /// `1` when the transaction had no local origin. One observer per session;
    /// a second call replaces the first, and a throwing callback is ignored.
    pub fn set_update_observer(&mut self, callback: &Function) -> Result<(), JsValue> {
        let callback = callback.clone();
        let subscription = self
            .engine
            .doc()
            .yrs_doc()
            .observe_update_v1(move |txn, event| {
                // Uint8Array::from copies out of wasm memory — the JS side
                // owns the bytes and may hold them across further edits.
                let bytes = Uint8Array::from(event.update.as_slice());
                let origin = if txn.origin().is_none() { 1.0 } else { 0.0 };
                let _ = callback.call2(&JsValue::NULL, &bytes.into(), &JsValue::from_f64(origin));
            })
            .map_err(js_err)?;
        self.update_observer = Some(subscription);
        Ok(())
    }

    /// Drops the update observer registered by [`EditSession::set_update_observer`].
    pub fn clear_update_observer(&mut self) {
        self.update_observer = None;
    }

    /// Starts queueing committed transactions for
    /// [`EditSession::drain_update_event`] instead of pushing them through a
    /// callback. Idempotent while observation is running.
    pub fn start_update_event_observation(&mut self) -> Result<(), JsValue> {
        if self.update_event_observer.is_some() {
            return Ok(());
        }
        let pending = Rc::new(RefCell::new(VecDeque::new()));
        let observed = Rc::clone(&pending);
        let subscription = self
            .engine
            .doc()
            .yrs_doc()
            .observe_update_v1(move |txn, event| {
                let mut encoded = Vec::with_capacity(event.update.len() + 1);
                encoded.push(if txn.origin().is_none() { 1 } else { 0 });
                encoded.extend_from_slice(&event.update);
                observed.borrow_mut().push_back(encoded);
            })
            .map_err(js_err)?;
        self.update_event_observer = Some(UpdateEventObserver {
            pending,
            _subscription: subscription,
        });
        Ok(())
    }

    /// Pops the oldest queued transaction, in arrival order. Byte 0 is `1`
    /// when the transaction had no local origin and `0` otherwise; the rest is
    /// the v1 update. Empty when the queue is drained or observation never
    /// started.
    pub fn drain_update_event(&self) -> Vec<u8> {
        self.update_event_observer
            .as_ref()
            .and_then(|observer| observer.pending.borrow_mut().pop_front())
            .unwrap_or_default()
    }

    /// Stops queueing and discards anything not yet drained.
    pub fn clear_update_event_observation(&mut self) {
        self.update_event_observer = None;
    }

    // -- local input state (undo + awareness selection) --

    /// Starts local-origin undo tracking for one story, replacing any scope
    /// already tracked (and its history). Call this after import or seeding but
    /// before the first edit, so the initial document is not an undo step.
    /// Re-tracking the same story is a no-op that preserves the history.
    /// Errors on an unknown story.
    pub fn track_undo(&self, story: &str) -> Result<(), JsValue> {
        self.undo.track(self.engine.doc(), story).map_err(js_err)
    }

    /// Starts local undo tracking for a structural table edit in `story`.
    /// Besides the parent story, which owns the table embed, this widens the
    /// scope to the stories root so undo and redo also remove and restore the
    /// cell stories the edit created or destroyed. Tracked separately from
    /// [`EditSession::track_undo`] on the same story, so switching between
    /// them starts a fresh history. Errors on an unknown story.
    pub fn track_table_undo(&self, story: &str) -> Result<(), JsValue> {
        self.undo
            .track_table(self.engine.doc(), story)
            .map_err(js_err)
    }

    /// Reverts the latest local-origin transaction and reports whether
    /// anything was reverted. Remote and system transactions are excluded by
    /// the manager's tracked-origin policy; `false` before a story is tracked.
    pub fn undo(&self) -> bool {
        self.undo.undo()
    }

    /// Reapplies the latest locally undone transaction and reports whether
    /// anything was reapplied.
    pub fn redo(&self) -> bool {
        self.undo.redo()
    }

    /// Whether [`EditSession::undo`] would revert something.
    pub fn can_undo(&self) -> bool {
        self.undo.can_undo()
    }

    /// Whether [`EditSession::redo`] would reapply something.
    pub fn can_redo(&self) -> bool {
        self.undo.can_redo()
    }

    /// Current local undo stack size. Zero before a story starts tracking.
    pub fn undo_depth(&self) -> u32 {
        self.undo.undo_depth() as u32
    }

    /// Current local redo stack size. Zero before a story starts tracking.
    pub fn redo_depth(&self) -> u32 {
        self.undo.redo_depth() as u32
    }

    /// Stores this peer's anchor and head as sticky positions, replacing any
    /// previous selection. Both endpoints must lie in `story`. The positions
    /// live outside the yrs document, so they are never serialized as content
    /// or carried in an update. `Assoc::After` makes a collapsed caret advance
    /// with text inserted at it.
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

    /// Encodes one Loc as opaque sticky-position bytes that keep pointing at
    /// the same content as the story is edited. Publish them over an awareness
    /// transport and resolve them with
    /// [`EditSession::resolve_sticky_position`].
    pub fn encode_sticky_position(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
    ) -> Result<Vec<u8>, JsValue> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let txn = self.engine.doc().yrs_doc().transact();
        let text = story_ref(&txn, story).map_err(js_err)?;
        let position = text
            .sticky_index(&txn, index, Assoc::After)
            .ok_or_else(|| js_err("position could not be made sticky"))?;
        Ok(encode_sticky(&position))
    }

    /// Resolves sticky bytes from [`EditSession::encode_sticky_position`] back
    /// to `{"story","paraId","offset"}`. Errors when the bytes are malformed
    /// or the position no longer resolves in `story`.
    pub fn resolve_sticky_position(&self, story: &str, position: &[u8]) -> Result<String, JsValue> {
        let (index, _) = resolve_sticky_selection(self.engine.doc(), story, position, position)
            .map_err(js_err)?;
        let loc = index_loc(self.engine.doc(), story, index)?;
        Ok(json!({
            "story": story,
            "paraId": loc.para_id,
            "offset": loc.offset,
        })
        .to_string())
    }

    /// This peer's current selection as `{"anchor":{"story","paraId","offset"},
    /// "head":{…}}`, or `"null"` before [`EditSession::set_selection`] is
    /// called. Errors when an endpoint no longer resolves.
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
        let anchor = index_loc(self.engine.doc(), &selection.story, anchor_index)?;
        let head = index_loc(self.engine.doc(), &selection.story, head_index)?;
        Ok(json!({
            "anchor": {
                "story": selection.story,
                "paraId": anchor.para_id,
                "offset": anchor.offset,
            },
            "head": {
                "story": selection.story,
                "paraId": head.para_id,
                "offset": head.offset,
            }
        })
        .to_string())
    }

    /// This peer's selection in transportable form:
    /// `{"story","anchor":[byte, …],"head":[byte, …]}` with sticky-encoded
    /// endpoints, or `"null"` before a selection is set. A peer resolves it
    /// with [`EditSession::resolve_encoded_selection`].
    pub fn encoded_selection(&self) -> Result<String, JsValue> {
        let selection = self.selection.borrow();
        let Some(selection) = selection.as_ref() else {
            return Ok("null".to_owned());
        };
        Ok(json!({
            "story": selection.story,
            "anchor": encode_sticky(&selection.anchor),
            "head": encode_sticky(&selection.head),
        })
        .to_string())
    }

    /// Resolves another peer's encoded selection against this replica's
    /// current state: `{"anchor":{"story","paraId","offset"},"head":{…}}`.
    /// Errors on malformed bytes or an endpoint that no longer resolves.
    pub fn resolve_encoded_selection(
        &self,
        story: &str,
        anchor: &[u8],
        head: &[u8],
    ) -> Result<String, JsValue> {
        let (anchor_index, head_index) =
            resolve_sticky_selection(self.engine.doc(), story, anchor, head).map_err(js_err)?;
        let anchor = index_loc(self.engine.doc(), story, anchor_index)?;
        let head = index_loc(self.engine.doc(), story, head_index)?;
        Ok(json!({
            "anchor": {
                "story": story,
                "paraId": anchor.para_id,
                "offset": anchor.offset,
            },
            "head": {
                "story": story,
                "paraId": head.para_id,
                "offset": head.offset,
            }
        })
        .to_string())
    }

    /// Stores a rectangular anchor-cell to head-cell selection outside the yrs
    /// document. `range_json` is a [`TableRange`]:
    /// `{"anchor":{"story","tableIndex","row","column"},"head":{…}}`. The table
    /// embed is held by a sticky index and each endpoint by its cell story's
    /// stable identity, so the selection survives unrelated edits and follows
    /// inserted or deleted rows and columns. Errors when the two endpoints are
    /// not in the same table, or when either cell does not resolve.
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

    /// The current local cell selection as a [`TableRange`], or `"null"`
    /// before one is set. An endpoint whose cell was deleted clamps to a
    /// surviving nearby cell. Errors when the table itself no longer resolves.
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

    /// Adds a story holding one paragraph with `initial_text` (which must not
    /// contain paragraph breaks), `p_style` and `alignment`. Receipt:
    /// `{"paraId"}` — the paragraph ending at the story's pilcrow. Errors when
    /// the story id already exists.
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

    /// Removes one complete story and its content. Errors on an unknown story.
    pub fn delete_story(&self, story_id: &str) -> Result<(), JsValue> {
        self.engine.doc().delete_story(story_id).map_err(js_err)
    }

    // -- native table ops --
    //
    // Tables are addressed in the resolved rectangular grid, not by OOXML row
    // and cell order: `TableLocator` is `{"story","tableIndex"}` with the
    // zero-based ordinal of the table embed in its story, `CellLoc` adds
    // `{"row","column"}`, and `TableRange` is `{"anchor":CellLoc,"head":
    // CellLoc}` covering the rectangle they span. Each cell's content lives in
    // its own story, so structural edits create and destroy stories.
    //
    // Every op below returns the same receipt:
    // `{"table":TableLocator,"rows","columns","createdStoryIds":[string, …],
    // "deletedStoryIds":[string, …],"newParaIds":[string, …],
    // "deletedTable":bool,"revisionIds":[string, …]}` — `rows`/`columns` are
    // the grid AFTER the op, the story lists let the host follow cell content
    // in and out of existence, and `revisionIds` is non-empty only for the ops
    // that accept a suggesting author. All of them additionally error on an
    // unknown table or a cell outside the grid.

    /// Inserts a `rows` x `columns` table at `(story, para_id, offset)`,
    /// creating one story per cell. Errors when either dimension is zero.
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

    /// Inserts a row above (`after = false`) or below (`after = true`) the
    /// cell `at_json` names. `at_json` is a [`CellLoc`].
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

    /// Inserts a column left (`after = false`) or right (`after = true`) of
    /// the cell `at_json` ([`CellLoc`]) names. Always a plain local edit.
    pub fn insert_column(&self, at_json: &str, after: bool) -> Result<String, JsValue> {
        let at: CellLoc = serde_json::from_str(at_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .insert_column(&EditCtx::local("", ""), &at, after)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Deletes every row the [`TableRange`] `range_json` covers. In suggesting
    /// mode the rows are marked `trDel` instead of removed.
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

    /// Deletes every column the [`TableRange`] `range_json` covers, along with
    /// the stories of the cells removed. Always a plain local edit.
    pub fn delete_column(&self, range_json: &str) -> Result<String, JsValue> {
        let range: TableRange = serde_json::from_str(range_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .delete_column(&EditCtx::local("", ""), &range)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Removes the table `table_json` ([`TableLocator`]) names plus every one
    /// of its reachable cell stories. Always a plain local edit; the receipt
    /// has `deletedTable: true`.
    pub fn delete_table(&self, table_json: &str) -> Result<String, JsValue> {
        let table: TableLocator = serde_json::from_str(table_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .delete_table(&EditCtx::local("", ""), &table)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Merges the rectangle the [`TableRange`] `range_json` covers into its
    /// top-left cell, whose story survives; the other cells' stories are
    /// deleted. Always a plain local edit. Errors when the range covers fewer
    /// than two cells or has no top-left anchor cell.
    pub fn merge_cells(&self, range_json: &str) -> Result<String, JsValue> {
        let range: TableRange = serde_json::from_str(range_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .merge_cells(&EditCtx::local("", ""), &range)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Splits the cell `at_json` ([`CellLoc`]) covers into a `rows` x
    /// `columns` grid; the original story stays in the top-left slot and every
    /// other slot gets a fresh one-paragraph cell story. Omitting both
    /// dimensions unmerges the cell into the slots it already covers, which
    /// then requires a merged cell. Always a plain local edit. Errors on a
    /// zero dimension or a grid smaller than the cell already covers.
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

    /// Sets every selected cell's background to `color` (hex, with or without
    /// a leading `#`), or clears it when `color` is absent. `range_json` is a
    /// [`TableRange`]. Always a plain local edit.
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

    /// Merges the JSON object `patch_json` into every selected cell's `tcPr`;
    /// a `null` value removes that key. `range_json` is a [`TableRange`].
    /// Always a plain local edit. Errors when the patch touches `rowspan`,
    /// `colspan`, `gridSpan` or `vMerge`, which only merge and split may set.
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

    /// Merges the sides of the JSON object `borders_json` into every selected
    /// cell's `tcPr.borders`; `insideH`/`insideV` resolve to the physical edges
    /// interior to the selection. `range_json` is a [`TableRange`]. Always a
    /// plain local edit.
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

    /// Sets the grid width, in twips, of the column holding the cell `at_json`
    /// ([`CellLoc`]) names. Always a plain local edit. Errors unless
    /// `width_twips` is finite and positive.
    pub fn set_column_width(&self, at_json: &str, width_twips: f64) -> Result<String, JsValue> {
        let at: CellLoc = serde_json::from_str(at_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .set_column_width(&EditCtx::local("", ""), &at, width_twips)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Sets the preferred width, in twips, of the table `table_json`
    /// ([`TableLocator`]) names. Always a plain local edit. Errors unless
    /// `width_twips` is finite and positive.
    pub fn set_table_width(&self, table_json: &str, width_twips: f64) -> Result<String, JsValue> {
        let table: TableLocator = serde_json::from_str(table_json).map_err(js_err)?;
        let receipt = self
            .engine
            .doc()
            .set_table_width(&EditCtx::local("", ""), &table, width_twips)
            .map_err(js_err)?;
        serde_json::to_string(&receipt).map_err(js_err)
    }

    /// Inserts `text` at `(story, para_id, offset)`. It must contain no
    /// paragraph or line breaks, and it inherits the formatting at the
    /// insertion point. Receipt: `{"revisionId": string|null}` — non-null in
    /// suggesting mode, where the text is stamped `ins` and coalesces into an
    /// adjacent insertion by the same author rather than opening a second
    /// revision.
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

    /// Deletes `[start, end)`. Because a range crossing a paragraph boundary
    /// includes the boundary pilcrow, a plain delete also merges those
    /// paragraphs. Suggesting mode removes nothing and stamps the content
    /// `del` instead. Receipt: `{"revisionId": string|null}`.
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

    /// Replaces `[start, end)` with `text` in one transaction. The inserted
    /// text adopts the first replaced unit's formatting; in suggesting mode
    /// the deletion and the insertion share one revision id. Receipt:
    /// `{"revisionId": string|null}`.
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
    /// pilcrow. The FIRST half keeps the original paraId and the second is
    /// re-minted. A split at the paragraph end leaves the empty second half
    /// with only the inherited property subset; a mid-paragraph split keeps
    /// its properties. Paragraph borders are cleared either way. Suggesting
    /// mode stamps the new pilcrow `ins` and `pPrIns`. Receipt:
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
    /// `del`- and `pPrDel`-marking (suggesting) its pilcrow. On a plain merge
    /// the survivor adopts the deleted mark's properties and paraId, so the
    /// earlier paragraph's identity wins. Receipt:
    /// `{"revisionId": string|null}`. Errors on the story's final paragraph,
    /// which has no following paragraph to merge with.
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

    /// Applies one run mark over `[start, end)`. `mark_json`:
    /// `{"type":"bold"|"italic"|"underline"|"strike"|"superscript"|"subscript"} |
    /// {"type":"fontFamily"|"color","value":string} |
    /// {"type":"fontSize","value":number}`. The six boolean types TOGGLE —
    /// they turn on unless the whole range already carries the mark — while
    /// font family, size and color SET. Formatting is always a plain local
    /// edit, never a tracked change, so there is no receipt. Errors on an
    /// unknown `"type"` or a missing/mistyped `"value"`.
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
    /// `[start, end)` in one transaction. An omitted key keeps the current
    /// value, `null` clears it, and any other value sets it. `delta_json`:
    ///
    /// ```json
    /// {
    ///   "bold": true, "italic": true,
    ///   "underline": true | {"style"?: string, "color"?: string},
    ///   "strike": true | {"double"?: boolean},
    ///   "color": {"rgb": string} | {"themeColor": string},
    ///   "highlight": string,
    ///   "fontSize": 12,
    ///   "fontFamily": {"ascii": string, "hAnsi"?: string},
    ///   "other": {"anyAttr": value}
    /// }
    /// ```
    ///
    /// `false` clears `bold`, `italic`, `underline` and `strike`. `color`
    /// takes exactly one of `rgb` or `themeColor`. `other` writes arbitrary
    /// run attributes, with `null` removing one. Always a plain local edit, so
    /// there is no receipt. Errors when a key carries a type not listed here.
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

    /// Sets or clears the hyperlink attribute over `[start, end)`.
    /// `hyperlink_json` is `{"href", "tooltip"?, "rId"?}` or `null` to unlink.
    /// The attribute is protected: ordinary formatting ops cannot write or
    /// erase it, only this one. Errors when the JSON is neither an object nor
    /// `null`.
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

    /// Clears every direct run-formatting attribute over `[start, end)`.
    /// Protected attributes — hyperlinks and tracked-change stamps — survive,
    /// as do paragraph properties.
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

    /// Writes `style_id` as the `pStyle` of every paragraph intersecting
    /// `[start, end)`. Only that key changes: this boundary has no style
    /// resolver, so it never fabricates the paragraph attributes or run marks
    /// the style definition would imply. In suggesting mode the property
    /// change is recorded as a `pPrChange` revision.
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
    /// intersecting `[start, end)` in one transaction. An omitted key keeps
    /// the current value and `null` clears it. `attrs_json` recognises
    /// `alignment`, `lineSpacing`, `lineSpacingRule`, `spaceBefore`,
    /// `spaceAfter`, `indentLeft`, `indentRight`, `indentFirstLine`,
    /// `hangingIndent`, `bidi`, `tabs`
    /// (`[{"position":number,"alignment":string,"leader"?:string}, …]`) and
    /// `defaultTextFormatting` (an object of run defaults). Any other key —
    /// whether written at the top level or nested under `"other"` — is stored
    /// as an opaque paragraph property. Spacing and indents are authored OOXML
    /// units (twips, line-spacing units), never pixels. In suggesting mode a
    /// change is recorded as a `pPrChange` revision. Errors when a recognised
    /// key carries the wrong type, and when a key names schema-managed
    /// identity such as `paraId`.
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

    /// Inserts one inline image embed at `(story, para_id, offset)`.
    /// `payload_json` is the image's authored payload object, stored as given.
    /// The embed occupies one story unit. Receipt:
    /// `{"revisionId": string|null}`. Errors when the payload is not an
    /// object.
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

    /// Sets the authored `value` (any JSON) on the content-control embed
    /// carrying `embed_id`, searching every story. Errors when no embed has
    /// that id.
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

    /// Sets the authored `value` on the content-control embed at
    /// `(story, para_id, offset)` — the way to reach a control with no
    /// authored `w:id` or tag, which
    /// [`EditSession::set_content_control_value`] cannot address. Errors when
    /// that position holds no embed.
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

    /// Removes the authored `value` from the content-control embed carrying
    /// `embed_id`, leaving the control itself in place. Errors when no embed
    /// has that id.
    pub fn clear_content_control_value(&self, embed_id: &str) -> Result<(), JsValue> {
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_embed_attrs_by_id(&ctx, embed_id, vec![("value".to_owned(), Any::Null)])
            .map(|_| ())
            .map_err(js_err)
    }

    /// Writes geometry fields onto the image embed carrying `embed_id` in one
    /// transaction. Every key of the `geometry_json` object becomes a payload
    /// entry, with `null` clearing it; the entries of a nested `"other"`
    /// object are flattened into the payload alongside them. Errors when
    /// `geometry_json` or its `"other"` is not an object, and when no embed
    /// has that id.
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

    /// Inserts a page-break embed at `(story, para_id, offset)`, occupying one
    /// story unit. Always a plain local edit.
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

    /// Inserts a section-break embed at `(story, para_id, offset)`.
    /// `break_type` must be `"nextPage"`, `"continuous"`, `"oddPage"` or
    /// `"evenPage"`; anything else errors. Always a plain local edit.
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

    /// Inserts a watermark embed at `(story, para_id, offset)`, its payload
    /// taken verbatim from the `watermark_json` object. Always a plain local
    /// edit. Errors when the JSON is not an object.
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

    /// Sets one paragraph property to any JSON value on `para_id`'s pilcrow,
    /// searching every story. Unlike
    /// [`EditSession::set_paragraph_attrs`] this writes a single key and never
    /// records a revision. Errors when the paragraph is unknown and when `key`
    /// names schema-managed identity (`paraId` or the embed discriminator).
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

    /// Adds a comment anchored to one or more ranges. `ranges_json`:
    /// `[{"story","startPara","startOffset","endPara","endOffset"}, …]`;
    /// `body_json` is any JSON value, stored as given. The anchors are sticky,
    /// so they follow the text they cover, and the comment lives outside the
    /// story rather than as an attribute on it. Receipt: `{"commentId"}` — a
    /// freshly minted id. Errors when `ranges_json` is not an array of
    /// well-formed ranges, or is empty.
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

    /// Accepts tracked changes: pending insertions become plain content,
    /// pending deletions are carried out; `pPrIns` marks clear so the split
    /// stays, `pPrDel` marks join with the following paragraph, whose own
    /// properties survive.
    ///
    /// `target_json` is either `{"revisionId": string}`, resolving that one
    /// coalesced revision wherever it appears in any story, or
    /// `{"story","startPara","startOffset","endPara","endOffset"}`, resolving
    /// every tracked change overlapping that range regardless of revision id.
    /// Receipt: `{"revisionIds": [string, …]}` — the ids actually resolved, in
    /// resolution order and deduplicated. Resolving applies a revision rather
    /// than authoring one, so it never stamps a new revision and ignores
    /// suggesting mode. Errors on an empty range and on a `revisionId` that
    /// matches nothing.
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

    /// Rejects tracked changes — the inverse of
    /// [`EditSession::accept_change`]: pending insertions roll back, pending
    /// deletions restore their text; `pPrIns` marks join back with the
    /// following paragraph, `pPrDel` marks clear so the split stays, and a
    /// `pPrChange` restores the paragraph's previous properties. Same target,
    /// receipt and error contract.
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

    /// Applies a batch of raw story mutations in ONE transaction. Unlike every
    /// other op here these carry no user intent: indices are story-global
    /// UTF-16 units, not Locs, and nothing is inferred or stamped on the
    /// caller's behalf. `ops_json` is an array of:
    ///
    /// - `{"op":"insert","index","text"?,"attrs"?}`
    /// - `{"op":"delete","index","len"}`
    /// - `{"op":"format","index","len","attrs"?}`
    /// - `{"op":"insertEmbed","index","kind"?,"payload"?,"attrs"?}`
    /// - `{"op":"setEmbedAttr","index","key","value"?}`
    /// - `{"op":"setComment","id","ranges":[[start,end], …],"author"?,"date"?,"body"?}`
    /// - `{"op":"removeComment","id"}`
    ///
    /// Each op's `index`, and each `setComment` range, is read against the
    /// story state AFTER every preceding op in the batch. `attrs` and
    /// `payload` are JSON objects written verbatim, so tracked-change stamps
    /// travel inside `attrs`; comments are keyed by the given id and anchored
    /// sticky. Errors on an unknown `"op"`, a missing or negative `"index"`,
    /// or a missing required field, and leaves the story untouched — parsing
    /// completes before the transaction opens.
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

    // -- read queries --
    //
    // Every query below is a pure snapshot of the document as it stands: it
    // mutates nothing, commits no transaction, and its JSON is fully
    // materialized before it returns.

    /// Aggregated toolbar and accessibility state over `[start, end)`:
    ///
    /// ```json
    /// {
    ///   "bold": true | false | "mixed", "italic": …, "underline": …, "strike": …,
    ///   "fontFamily": string|null, "fontSize": number|null, "color": string|null,
    ///   "paraId": string, "styleId": string|null, "alignment": string|null,
    ///   "paragraphProperties": {…},
    ///   "hasSelection": bool, "isMultiParagraph": bool, "inTable": bool,
    ///   "isSingleEmbed": bool, "embedKind": string|null, "isImage": bool,
    ///   "inInsertion": bool, "inDeletion": bool
    /// }
    /// ```
    ///
    /// A toggle mark is `"mixed"` when the range disagrees about it; a value
    /// mark is `null` when it is absent OR not uniform, so the two cases are
    /// not distinguished. `isImage` is `embedKind == "image"`, and
    /// `inInsertion`/`inDeletion` report whether the range sits inside a
    /// pending tracked change.
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

    /// Every pending tracked change across all stories, in deterministic
    /// story-then-position order:
    /// `[{"revisionId","author","date","kind","story","preview",
    /// "range":{"story","start":{"paraId","offset"},"end":{…}}}, …]`. `kind`
    /// is one of `"insertion"`, `"deletion"`, `"pPrIns"`, `"pPrDel"`,
    /// `"pPrChange"`, `"trIns"`, `"trDel"`, `"tableIns"` or `"tableDel"`.
    /// `preview` holds the first few characters of the affected text, and is
    /// empty for every structural kind, which covers no text.
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

    /// Every story id in the document, sorted so the order is stable across
    /// replicas.
    pub fn story_ids(&self) -> Vec<String> {
        let txn = self.engine.doc().yrs_doc().transact();
        let Some(stories) = txn.get_map(STORIES) else {
            return Vec::new();
        };
        let mut ids: Vec<String> = stories.iter(&txn).map(|(id, _)| id.to_string()).collect();
        ids.sort();
        ids
    }

    /// Story length in UTF-16 units, every embed (pilcrows included) counting
    /// as one. Errors on an unknown story.
    pub fn story_len(&self, story: &str) -> Result<u32, JsValue> {
        self.engine.doc().story_len(story).map_err(js_err)
    }

    /// The story's `canonical-stream-v1` FNV-1a checksum (see
    /// [`crate::canonical`]) as a DECIMAL STRING, because a u64 exceeds the
    /// JavaScript safe-integer range. Two stories with the same authored
    /// content share a checksum even when their paragraph and comment ids
    /// differ. Errors on an unknown story.
    pub fn story_checksum(&self, story: &str) -> Result<String, JsValue> {
        crate::story_checksum(self.engine.doc(), story)
            .map(|checksum| checksum.to_string())
            .map_err(js_err)
    }

    /// Lowers one story to a `LayoutBlock[]` JSON array — the block, run and
    /// table vocabulary the layout engine consumes. `env_json` supplies the
    /// document-level values lowering cannot read off the story:
    /// `{"themeColors":{slot: hex},"defaultTabStopTwips":number|null,
    /// "pageContentHeight":number|null,"numericIds":{yrsId: number}}`, all
    /// optional. Errors when the story does not end in a pilcrow, holds a
    /// malformed table, references itself through a cell story, or contains an
    /// embed lowering does not support.
    pub fn yrs_blocks_for_story(&self, story: &str, env_json: &str) -> Result<String, JsValue> {
        let env = parse_render_env(env_json)?;
        self.engine.lower_story_json(story, &env).map_err(js_err)
    }

    /// `[{"paraId","text","properties"}, …]` in document order. `text` is the
    /// paragraph's plain text without its pilcrow; `properties` is the
    /// pilcrow's authored property map (`pStyle`, `alignment` and whatever
    /// else has been set on it). Errors on an unknown story.
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

    /// Compact paragraph-position projection built in one story traversal:
    /// `[{"paraId","length"}, …]` in document order, where `length` counts the
    /// UTF-16 text and inline embed units before that paragraph's pilcrow —
    /// exactly the `offset` domain of a Loc in it. One call replaces crossing
    /// the boundary once per paragraph. Errors on an unknown story.
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

    /// The story as an ordered run of formatted segments — the same view
    /// lowering reads:
    ///
    /// - `{"kind":"text","text","attributes"}`
    /// - `{"kind":"pilcrow","paraId","properties","attributes"}`
    /// - `{"kind":"embed","embedKind","payload","attributes"}`
    ///
    /// A text segment covers one maximal run of identically formatted
    /// characters; each pilcrow and embed is its own segment worth one unit.
    /// `attributes` holds the segment's run marks together with any `ins`/`del`
    /// tracked-change stamps. Errors on an unknown story.
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

    /// `{"start","end"}` — the paragraph's span in story-global UTF-16 units.
    /// `end` is the index of its own pilcrow, so `end - start` is the
    /// paragraph length and the upper bound of a Loc `offset` in it. Errors
    /// when the paragraph is not in `story`.
    pub fn locate_paragraph(&self, story: &str, para_id: &str) -> Result<String, JsValue> {
        let span = find_para_span(self.engine.doc(), story, para_id)?;
        Ok(json!({ "start": span.start, "end": span.pilcrow }).to_string())
    }

    /// Where a comment's sticky anchors currently sit:
    /// `[{"story","start","end"}, …]`, one entry per anchored range, in
    /// story-global UTF-16 units. Errors on an unknown comment id and when an
    /// anchor no longer resolves.
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{EditCtx, RawOp};

    fn seed_paragraph_after_embeds(doc: &EditingDoc, embeds: &[&str], text: &str) {
        doc.create_story_with_paragraph_id("body", "p0", "Alpha", "Normal", "left")
            .unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        let mut index = 6;
        let mut ops = Vec::new();
        for kind in embeds {
            ops.push(RawOp::InsertEmbed {
                index,
                kind: (*kind).into(),
                payload: Default::default(),
                attrs: Default::default(),
            });
            index += 1;
        }
        ops.push(RawOp::Insert {
            index,
            text: text.into(),
            attrs: Default::default(),
        });
        index += text.encode_utf16().count() as u32;
        ops.push(RawOp::InsertEmbed {
            index,
            kind: "pilcrow".into(),
            payload: vec![
                ("paraId".into(), Any::from("p1")),
                ("pStyle".into(), Any::from("Normal")),
                ("alignment".into(), Any::from("left")),
            ],
            attrs: Default::default(),
        });
        doc.apply_raw_ops("body", ops, &ctx).unwrap();
    }

    fn caret_layout_input() -> String {
        json!({
            "measured": [{
                "block": {
                    "kind": "paragraph",
                    "id": "p1",
                    "paraId": "p1",
                    "runs": [{
                        "kind": "text",
                        "text": "ABCDE",
                        "pmStart": 1,
                        "pmEnd": 6
                    }],
                    "attrs": {},
                    "pmStart": 0,
                    "pmEnd": 7
                },
                "measure": {
                    "kind": "paragraph",
                    "lines": [{
                        "headRun": 0,
                        "headChar": 0,
                        "tailRun": 0,
                        "tailChar": 5,
                        "width": 50,
                        "ascent": 8,
                        "descent": 2,
                        "lineHeight": 20
                    }],
                    "totalHeight": 20
                }
            }],
            "options": {
                "pageSize": { "w": 200, "h": 120 },
                "margins": { "top": 10, "right": 10, "bottom": 10, "left": 10 }
            }
        })
        .to_string()
    }

    #[test]
    fn index_loc_returns_the_node_offset_for_cumulative_block_embeds() {
        let kinds = ["table", "pageBreak", "columnBreak", "blockSdt"];
        for count in 0..=kinds.len() {
            let doc = EditingDoc::new(7 + count as u64);
            seed_paragraph_after_embeds(&doc, &kinds[..count], "ABCDE");
            let index = 6 + count as u32 + 3;
            let loc = index_loc(&doc, "body", index).unwrap();

            assert_eq!(loc.para_id, "p1");
            assert_eq!(loc.offset, count as u32 + 3);
            assert_eq!(loc.node_offset, 3);
        }
    }

    #[test]
    fn resident_backspace_after_a_leading_embed_deletes_the_previous_character() {
        let session = EditSession::new(20.0).unwrap();
        seed_paragraph_after_embeds(session.engine.doc(), &["table"], "ABCDE");
        let head = loc_index(session.engine.doc(), "body", "p1", 6).unwrap();

        session
            .delete_resident_input("backward", ("body".into(), "p1".into(), head))
            .unwrap();

        let paragraphs = session.engine.doc().paragraphs("body").unwrap();
        assert_eq!(paragraphs[1].text, "ABCD");
    }

    #[test]
    fn resident_caret_snapshot_rebases_cumulative_block_embeds() {
        let session = EditSession::new(21.0).unwrap();
        seed_paragraph_after_embeds(
            session.engine.doc(),
            &["table", "pageBreak", "columnBreak", "blockSdt"],
            "ABCDE",
        );
        session
            .engine
            .layout_document_json(&caret_layout_input())
            .unwrap();
        session.engine.build_display_list_frame("{}", 0).unwrap();
        session.set_selection("body", "p1", 7, "p1", 7).unwrap();

        let snapshot: Value =
            serde_json::from_str(&session.resident_caret_snapshot_json().unwrap()).unwrap();
        let caret = &snapshot["caretRect"];

        assert_eq!(snapshot["frameEpoch"], 1);
        assert_eq!(caret["pageIndex"], 0);
        assert!((caret["x"].as_f64().unwrap() - 40.0).abs() < 0.001);
    }
}
