//! Inline formatting operations.

use std::collections::{BTreeMap, HashMap};
use std::sync::Arc;

use serde_json::Value;
use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{Any, Out, ReadTxn, Text, TextRef};

use crate::op::{OpError, OpResult, Receipt};
use crate::{DEL, EditCtx, EditingDoc, INS, StoryRange, out_len, story_ref};

/// The hyperlink text attribute; retained by [`EditingDoc::clear_formatting`].
pub const HYPERLINK: &str = "hyperlink";

/// Attributes protected from formatting sweeps.
pub(crate) const PROTECTED_ATTRS: [&str; 3] = [INS, DEL, HYPERLINK];

/// Word's closed highlight palette keyed by exact uppercase hex.
const HIGHLIGHT_HEX_TO_NAME: [(&str, &str); 16] = [
    ("FFFF00", "yellow"),
    ("00FF00", "green"),
    ("00FFFF", "cyan"),
    ("FF00FF", "magenta"),
    ("0000FF", "blue"),
    ("FF0000", "red"),
    ("00008B", "darkBlue"),
    ("008080", "darkCyan"),
    ("008000", "darkGreen"),
    ("800080", "darkMagenta"),
    ("8B0000", "darkRed"),
    ("808000", "darkYellow"),
    ("808080", "darkGray"),
    ("C0C0C0", "lightGray"),
    ("000000", "black"),
    ("FFFFFF", "white"),
];

/// Maps exact palette colors and preserves other values.
pub fn highlight_color_name(input: &str) -> String {
    let normalized = input.trim_start_matches('#').to_ascii_uppercase();
    HIGHLIGHT_HEX_TO_NAME
        .iter()
        .find(|(hex, _)| *hex == normalized)
        .map(|(_, name)| (*name).to_owned())
        .unwrap_or_else(|| input.to_owned())
}

/// Tri-state field patch.
#[derive(Clone, Debug, Default, PartialEq)]
pub enum Patch<T> {
    /// Leave the attribute as it is.
    #[default]
    Keep,
    /// Remove the attribute.
    Clear,
    /// Write the attribute.
    Set(T),
}

/// Six simple formatting toggles.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SimpleFormat {
    Bold,
    Italic,
    Underline,
    Strike,
    Superscript,
    Subscript,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct UnderlinePatch {
    /// `w:u` val; `None` writes the default `"single"`.
    pub style: Option<String>,
    pub color: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct StrikePatch {
    pub double: bool,
}

/// Text color is rgb XOR theme by construction.
#[derive(Clone, Debug, PartialEq)]
pub enum ColorPatch {
    Rgb(String),
    Theme(String),
}

#[derive(Clone, Debug, PartialEq)]
pub struct FontFamilyPatch {
    pub ascii: String,
    /// Defaults to `ascii` when omitted (Word writes both `w:ascii` and `w:hAnsi`).
    pub h_ansi: Option<String>,
}

/// Tri-state inline formatting patch.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct InlineFormatDelta {
    pub bold: Patch<bool>,
    pub italic: Patch<bool>,
    pub underline: Patch<UnderlinePatch>,
    pub strike: Patch<StrikePatch>,
    pub color: Patch<ColorPatch>,
    /// Named highlight or hex (mapped through Word's palette; unmapped hex kept raw).
    pub highlight: Patch<String>,
    /// Points; written as half-points into both `size` and `sizeCs` (`w:sz` + `w:szCs`).
    pub font_size: Patch<f64>,
    pub font_family: Patch<FontFamilyPatch>,
    pub other: BTreeMap<String, Option<Any>>,
}

impl InlineFormatDelta {
    pub fn is_empty(&self) -> bool {
        self.bold == Patch::Keep
            && self.italic == Patch::Keep
            && self.underline == Patch::Keep
            && self.strike == Patch::Keep
            && self.color == Patch::Keep
            && self.highlight == Patch::Keep
            && self.font_size == Patch::Keep
            && self.font_family == Patch::Keep
            && self.other.is_empty()
    }

    /// Lowers the delta to yrs formatting attributes (`Any::Null` = remove).
    pub(crate) fn to_attrs(&self) -> OpResult<Attrs> {
        let mut attrs = Attrs::new();
        let mut put = |key: &str, value: Any| {
            attrs.insert(Arc::from(key), value);
        };
        match &self.bold {
            Patch::Keep => {}
            Patch::Clear | Patch::Set(false) => put("bold", Any::Null),
            Patch::Set(true) => put("bold", Any::Bool(true)),
        }
        match &self.italic {
            Patch::Keep => {}
            Patch::Clear | Patch::Set(false) => put("italic", Any::Null),
            Patch::Set(true) => put("italic", Any::Bool(true)),
        }
        match &self.underline {
            Patch::Keep => {}
            Patch::Clear => put("underline", Any::Null),
            Patch::Set(patch) => {
                let mut value = HashMap::from([(
                    "style".into(),
                    Any::from(patch.style.clone().unwrap_or_else(|| "single".into())),
                )]);
                if let Some(color) = &patch.color {
                    value.insert("color".into(), Any::from(color.as_str()));
                }
                put("underline", Any::Map(Arc::new(value)));
            }
        }
        match &self.strike {
            Patch::Keep => {}
            Patch::Clear => put("strike", Any::Null),
            Patch::Set(patch) => put(
                "strike",
                Any::Map(Arc::new(HashMap::from([(
                    "double".into(),
                    Any::Bool(patch.double),
                )]))),
            ),
        }
        match &self.color {
            Patch::Keep => {}
            Patch::Clear => put("textColor", Any::Null),
            Patch::Set(patch) => {
                let (rgb, theme) = match patch {
                    ColorPatch::Rgb(rgb) => (Any::from(rgb.as_str()), Any::Null),
                    ColorPatch::Theme(theme) => (Any::Null, Any::from(theme.as_str())),
                };
                put(
                    "textColor",
                    Any::Map(Arc::new(HashMap::from([
                        ("rgb".into(), rgb),
                        ("themeColor".into(), theme),
                    ]))),
                );
            }
        }
        match &self.highlight {
            Patch::Keep => {}
            Patch::Clear => put("highlight", Any::Null),
            Patch::Set(color) => put(
                "highlight",
                Any::Map(Arc::new(HashMap::from([(
                    "color".into(),
                    Any::from(highlight_color_name(color)),
                )]))),
            ),
        }
        match &self.font_size {
            Patch::Keep => {}
            Patch::Clear => put("fontSize", Any::Null),
            Patch::Set(points) => {
                if !points.is_finite() || *points <= 0.0 {
                    return Err(OpError::InvalidFormatValue(format!(
                        "font size must be a positive number of points, got {points}"
                    )));
                }
                let half_points = (points * 2.0).round();
                put(
                    "fontSize",
                    Any::Map(Arc::new(HashMap::from([
                        ("size".into(), Any::Number(half_points)),
                        ("sizeCs".into(), Any::Number(half_points)),
                    ]))),
                );
            }
        }
        match &self.font_family {
            Patch::Keep => {}
            Patch::Clear => put("fontFamily", Any::Null),
            Patch::Set(patch) => {
                let h_ansi = patch.h_ansi.clone().unwrap_or_else(|| patch.ascii.clone());
                put(
                    "fontFamily",
                    Any::Map(Arc::new(HashMap::from([
                        ("ascii".into(), Any::from(patch.ascii.as_str())),
                        ("hAnsi".into(), Any::from(h_ansi)),
                    ]))),
                );
            }
        }
        for (key, value) in &self.other {
            if PROTECTED_ATTRS.contains(&key.as_str()) {
                return Err(OpError::InvalidFormatValue(format!(
                    "attribute {key:?} is not a formatting attribute"
                )));
            }
            put(key, value.clone().unwrap_or(Any::Null));
        }
        Ok(attrs)
    }
}

/// Parses the `format_range` delta JSON object. Shared by the wasm entry
/// and table-fragment paste, which lowers the result with [`InlineFormatDelta::to_attrs`].
pub(crate) fn parse_inline_format_delta(
    object: &serde_json::Map<String, Value>,
) -> Result<InlineFormatDelta, String> {
    let bool_patch = |key: &str| -> Result<Patch<bool>, String> {
        match object.get(key) {
            None => Ok(Patch::Keep),
            Some(Value::Null) => Ok(Patch::Clear),
            Some(Value::Bool(value)) => Ok(Patch::Set(*value)),
            Some(_) => Err(format!("format delta {key:?} must be a boolean or null")),
        }
    };

    let underline = match object.get("underline") {
        None => Patch::Keep,
        Some(Value::Null | Value::Bool(false)) => Patch::Clear,
        Some(Value::Bool(true)) => Patch::Set(Default::default()),
        Some(Value::Object(value)) => {
            let string = |key: &str| -> Result<Option<String>, String> {
                value
                    .get(key)
                    .map(|entry| {
                        entry
                            .as_str()
                            .map(str::to_owned)
                            .ok_or_else(|| format!("underline {key:?} must be a string"))
                    })
                    .transpose()
            };
            Patch::Set(UnderlinePatch {
                style: string("style")?,
                color: string("color")?,
            })
        }
        Some(_) => {
            return Err("format delta \"underline\" must be a boolean, object, or null".to_owned());
        }
    };

    let strike = match object.get("strike") {
        None => Patch::Keep,
        Some(Value::Null | Value::Bool(false)) => Patch::Clear,
        Some(Value::Bool(true)) => Patch::Set(StrikePatch { double: false }),
        Some(Value::Object(value)) => {
            let double = value
                .get("double")
                .map(|entry| {
                    entry
                        .as_bool()
                        .ok_or_else(|| "strike \"double\" must be a boolean".to_owned())
                })
                .transpose()?
                .unwrap_or(false);
            Patch::Set(StrikePatch { double })
        }
        Some(_) => {
            return Err("format delta \"strike\" must be a boolean, object, or null".to_owned());
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
                return Err(
                    "format delta \"color\" requires exactly one of \"rgb\" or \"themeColor\""
                        .to_owned(),
                );
            }
        },
        Some(_) => {
            return Err("format delta \"color\" must be an object or null".to_owned());
        }
    };

    let highlight = match object.get("highlight") {
        None => Patch::Keep,
        Some(Value::Null) => Patch::Clear,
        Some(Value::String(value)) => Patch::Set(value.clone()),
        Some(_) => {
            return Err("format delta \"highlight\" must be a string or null".to_owned());
        }
    };

    let font_size = match object.get("fontSize") {
        None => Patch::Keep,
        Some(Value::Null) => Patch::Clear,
        Some(Value::Number(value)) => Patch::Set(
            value
                .as_f64()
                .ok_or_else(|| "format delta \"fontSize\" must be finite".to_owned())?,
        ),
        Some(_) => {
            return Err("format delta \"fontSize\" must be a number or null".to_owned());
        }
    };

    let font_family = match object.get("fontFamily") {
        None => Patch::Keep,
        Some(Value::Null) => Patch::Clear,
        Some(Value::Object(value)) => {
            let ascii = value.get("ascii").and_then(Value::as_str).ok_or_else(|| {
                "format delta \"fontFamily\" requires string \"ascii\"".to_owned()
            })?;
            let h_ansi = value
                .get("hAnsi")
                .map(|entry| {
                    entry
                        .as_str()
                        .map(str::to_owned)
                        .ok_or_else(|| "fontFamily \"hAnsi\" must be a string".to_owned())
                })
                .transpose()?;
            Patch::Set(FontFamilyPatch {
                ascii: ascii.to_owned(),
                h_ansi,
            })
        }
        Some(_) => {
            return Err("format delta \"fontFamily\" must be an object or null".to_owned());
        }
    };

    let mut other = BTreeMap::new();
    if let Some(value) = object.get("other") {
        let entries = value
            .as_object()
            .ok_or_else(|| "format delta \"other\" must be an object".to_owned())?;
        for (key, value) in entries {
            other.insert(
                key.clone(),
                if value.is_null() {
                    None
                } else {
                    Some(Any::from_json(&value.to_string()).map_err(|error| {
                        format!("format delta \"other\" entry {key:?} is not JSON: {error}")
                    })?)
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

/// Formatting policy for [`EditingDoc::insert_text`].
#[derive(Clone, Debug, Default, PartialEq)]
pub enum FormatPolicy {
    /// Match typing: copy the formatting attributes of the character before the insertion point
    /// (or after it at a paragraph start). Tracked-change stamps are never inherited; a hyperlink
    /// is inherited only when the insertion point is strictly inside it.
    #[default]
    Inherit,
    /// Insert with exactly these formatting attributes.
    Explicit(BTreeMap<String, Any>),
    /// Insert with no formatting attributes.
    Plain,
}

pub(crate) fn is_active(attrs: Option<&Attrs>, key: &str) -> bool {
    matches!(attrs.and_then(|attrs| attrs.get(key)), Some(value) if *value != Any::Null)
}

/// Tests whether every text unit carries a non-null attribute.
fn all_text_has<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32, key: &str) -> bool {
    let mut offset = 0;
    let mut saw_text = false;
    for diff in story.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        let chunk_end = offset + len;
        let overlaps = chunk_end.min(end) > offset.max(start);
        if overlaps && matches!(diff.insert, Out::Any(Any::String(_))) {
            saw_text = true;
            if !is_active(diff.attributes.as_deref(), key) {
                return false;
            }
        }
        offset = chunk_end;
        if offset >= end {
            break;
        }
    }
    saw_text
}

impl EditingDoc {
    /// Toggles one simple format across a range.
    pub fn toggle_format(
        &self,
        ctx: &EditCtx,
        mut range: StoryRange,
        format: SimpleFormat,
    ) -> OpResult<Receipt> {
        let len = range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        crate::check_range(&story, &txn, range.start, len)?;
        (range.start, range.end) =
            crate::ops::code_point_range(&story, &txn, range.start, range.end);
        let len = range.end - range.start;
        let (key, on_value, counterpart) = match format {
            SimpleFormat::Bold => ("bold", Any::Bool(true), None),
            SimpleFormat::Italic => ("italic", Any::Bool(true), None),
            SimpleFormat::Underline => (
                "underline",
                Any::Map(Arc::new(HashMap::from([(
                    "style".into(),
                    Any::from("single"),
                )]))),
                None,
            ),
            SimpleFormat::Strike => (
                "strike",
                Any::Map(Arc::new(HashMap::from([(
                    "double".into(),
                    Any::Bool(false),
                )]))),
                None,
            ),
            SimpleFormat::Superscript => ("superscript", Any::Bool(true), Some("subscript")),
            SimpleFormat::Subscript => ("subscript", Any::Bool(true), Some("superscript")),
        };
        let turn_off = all_text_has(&story, &txn, range.start, range.end, key);
        let mut attrs =
            Attrs::from([(Arc::from(key), if turn_off { Any::Null } else { on_value })]);
        if let Some(counterpart) = counterpart
            && !turn_off
        {
            attrs.insert(Arc::from(counterpart), Any::Null);
        }
        story.format(&mut txn, range.start, len, attrs);
        let loc_range =
            crate::op::loc_range_in_txn(&range.story, &story, &txn, range.start, range.end)?;
        Ok(Receipt {
            range: Some(loc_range),
            ..Receipt::default()
        })
    }

    /// Applies a formatting delta across a range in one transaction.
    pub fn format_range(
        &self,
        ctx: &EditCtx,
        mut range: StoryRange,
        delta: &InlineFormatDelta,
    ) -> OpResult<Receipt> {
        let len = range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let attrs = delta.to_attrs()?;
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        crate::check_range(&story, &txn, range.start, len)?;
        (range.start, range.end) =
            crate::ops::code_point_range(&story, &txn, range.start, range.end);
        let len = range.end - range.start;
        if !attrs.is_empty() {
            story.format(&mut txn, range.start, len, attrs);
        }
        let loc_range =
            crate::op::loc_range_in_txn(&range.story, &story, &txn, range.start, range.end)?;
        Ok(Receipt {
            range: Some(loc_range),
            ..Receipt::default()
        })
    }

    /// Sets or clears the protected hyperlink attribute over one non-empty
    /// range. Generic formatting deliberately cannot touch hyperlinks;
    /// hyperlink editing uses this explicit operation instead.
    pub fn set_hyperlink(
        &self,
        ctx: &EditCtx,
        mut range: StoryRange,
        hyperlink: Option<Any>,
    ) -> OpResult<Receipt> {
        let len = range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        crate::check_range(&story, &txn, range.start, len)?;
        (range.start, range.end) =
            crate::ops::code_point_range(&story, &txn, range.start, range.end);
        let len = range.end - range.start;
        story.format(
            &mut txn,
            range.start,
            len,
            Attrs::from([(Arc::from(HYPERLINK), hyperlink.unwrap_or(Any::Null))]),
        );
        let loc_range =
            crate::op::loc_range_in_txn(&range.story, &story, &txn, range.start, range.end)?;
        Ok(Receipt {
            range: Some(loc_range),
            ..Receipt::default()
        })
    }

    /// Clears formatting while retaining tracked-change stamps and hyperlinks.
    pub fn clear_formatting(&self, ctx: &EditCtx, mut range: StoryRange) -> OpResult<Receipt> {
        let len = range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        crate::check_range(&story, &txn, range.start, len)?;
        (range.start, range.end) =
            crate::ops::code_point_range(&story, &txn, range.start, range.end);
        let len = range.end - range.start;
        let mut keys: BTreeMap<String, ()> = BTreeMap::new();
        let mut offset = 0;
        for diff in story.diff(&txn, YChange::identity) {
            let chunk_len = out_len(&diff.insert);
            let chunk_end = offset + chunk_len;
            if chunk_end.min(range.end) > offset.max(range.start)
                && let Some(attrs) = diff.attributes.as_deref()
            {
                for (key, value) in attrs.iter() {
                    if *value != Any::Null && !PROTECTED_ATTRS.contains(&key.as_ref()) {
                        keys.insert(key.to_string(), ());
                    }
                }
            }
            offset = chunk_end;
            if offset >= range.end {
                break;
            }
        }
        if !keys.is_empty() {
            let attrs: Attrs = keys
                .into_keys()
                .map(|key| (Arc::from(key.as_str()), Any::Null))
                .collect();
            story.format(&mut txn, range.start, len, attrs);
        }
        let loc_range =
            crate::op::loc_range_in_txn(&range.story, &story, &txn, range.start, range.end)?;
        Ok(Receipt {
            range: Some(loc_range),
            ..Receipt::default()
        })
    }
}

pub(crate) fn range_len(range: &StoryRange) -> OpResult<u32> {
    range
        .end
        .checked_sub(range.start)
        .ok_or(OpError::InvalidRange {
            start: range.start,
            end: range.end,
        })
}
