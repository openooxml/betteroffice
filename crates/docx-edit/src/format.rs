//! Inline formatting operations.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use serde_json::Value;
use yrs::types::text::YChange;
use yrs::types::{Attrs, Delta};
use yrs::{Any, In, Map, MapRef, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::op::{OpError, OpResult, Receipt};
use crate::{DEL, EditCtx, EditingDoc, INS, KIND_KEY, StoryRange, map_string, out_len, story_ref};

/// The hyperlink text attribute; retained by [`EditingDoc::clear_formatting`].
pub const HYPERLINK: &str = "hyperlink";

/// Attributes protected from formatting sweeps.
pub(crate) const RPR_CHANGE: &str = "rPrChange";

pub(crate) const PROTECTED_ATTRS: [&str; 4] = [INS, DEL, HYPERLINK, RPR_CHANGE];

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

#[derive(Clone, Debug, Default, PartialEq)]
pub struct LanguagePatch {
    pub latin: Patch<String>,
    pub east_asia: Patch<String>,
    pub bidi: Patch<String>,
}

impl LanguagePatch {
    pub(crate) fn parse(value: &Value) -> Result<Self, String> {
        let object = value.as_object().ok_or("language must be an object")?;
        if object
            .keys()
            .any(|key| !matches!(key.as_str(), "lang" | "eastAsia" | "bidi"))
        {
            return Err("unknown language slot".to_owned());
        }
        let slot = |key: &str| match object.get(key) {
            None => Ok(Patch::Keep),
            Some(Value::Null) => Ok(Patch::Clear),
            Some(Value::String(tag)) if valid_language_tag(tag) => Ok(Patch::Set(tag.clone())),
            _ => Err(format!("invalid language tag for {key}")),
        };
        Ok(Self {
            latin: slot("lang")?,
            east_asia: slot("eastAsia")?,
            bidi: slot("bidi")?,
        })
    }

    pub(crate) fn attrs(&self) -> Any {
        let mut slots = HashMap::new();
        for (key, patch) in [
            ("latin", &self.latin),
            ("eastAsia", &self.east_asia),
            ("bidi", &self.bidi),
        ] {
            match patch {
                Patch::Keep => {}
                Patch::Clear => {
                    slots.insert(key.to_owned(), Any::Null);
                }
                Patch::Set(tag) => {
                    slots.insert(key.to_owned(), Any::from(tag.as_str()));
                }
            }
        }
        Any::Map(Arc::new(slots))
    }
}

fn valid_language_tag(tag: &str) -> bool {
    let mut parts = tag.split('-');
    let first = parts.next().unwrap_or_default();
    tag.len() <= 35
        && (2..=8).contains(&first.len())
        && first.bytes().all(|c| c.is_ascii_alphabetic())
        && parts.all(|part| {
            (1..=8).contains(&part.len()) && part.bytes().all(|c| c.is_ascii_alphanumeric())
        })
}

pub(crate) fn merge_language(previous: Option<&Any>, patch: &Any) -> Any {
    let mut slots = match previous {
        Some(Any::Map(slots)) => slots.as_ref().clone(),
        _ => HashMap::new(),
    };
    if let Some(value) = slots.remove("val") {
        slots.entry("latin".to_owned()).or_insert(value);
    }
    if let Any::Map(patch) = patch {
        for (key, value) in patch.iter() {
            slots.insert(
                if key == "val" {
                    "latin".to_owned()
                } else {
                    key.clone()
                },
                value.clone(),
            );
        }
    }
    Any::Map(Arc::new(slots))
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
    pub language: Patch<LanguagePatch>,
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
            && self.language == Patch::Keep
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
        match &self.language {
            Patch::Keep => {}
            Patch::Clear => put(
                "language",
                LanguagePatch {
                    latin: Patch::Clear,
                    east_asia: Patch::Clear,
                    bidi: Patch::Clear,
                }
                .attrs(),
            ),
            Patch::Set(patch) => {
                for slot in [&patch.latin, &patch.east_asia, &patch.bidi] {
                    if let Patch::Set(tag) = slot
                        && !valid_language_tag(tag)
                    {
                        return Err(OpError::InvalidFormatValue(
                            "invalid language tag".to_owned(),
                        ));
                    }
                }
                put("language", patch.attrs());
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
        language: match object.get("language") {
            None => Patch::Keep,
            Some(Value::Null) => Patch::Clear,
            Some(value) => Patch::Set(LanguagePatch::parse(value)?),
        },
        other,
    })
}

const UNDERLINE_STYLES: &[&str] = &[
    "single",
    "words",
    "double",
    "thick",
    "dotted",
    "dottedHeavy",
    "dash",
    "dashedHeavy",
    "dashLong",
    "dashLongHeavy",
    "dotDash",
    "dashDotHeavy",
    "dotDotDash",
    "dashDotDotHeavy",
    "wave",
    "wavyHeavy",
    "wavyDouble",
];

const HIGHLIGHT_NAMES: &[&str] = &[
    "black",
    "blue",
    "cyan",
    "green",
    "magenta",
    "red",
    "yellow",
    "white",
    "darkBlue",
    "darkCyan",
    "darkGreen",
    "darkMagenta",
    "darkRed",
    "darkYellow",
    "darkGray",
    "lightGray",
];

const THEME_COLORS: &[&str] = &[
    "dark1",
    "light1",
    "dark2",
    "light2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hyperlink",
    "followedHyperlink",
    "background1",
    "text1",
    "background2",
    "text2",
];

/// Lowers a host edit run's `marks` to the attributes the formatting commands write, in key
/// order; `Any::Null` removes a key. Refuses unknown marks and unsupported values.
pub(crate) fn text_marks_overlay(
    marks: &serde_json::Map<String, Value>,
) -> Result<Vec<(String, Any)>, String> {
    let flag = |key: &str, value: &Value| match value {
        Value::Bool(true) => Ok(Patch::Set(true)),
        Value::Bool(false) => Ok(Patch::Clear),
        _ => Err(format!("mark {key:?} must be a boolean")),
    };
    let mut delta = InlineFormatDelta::default();
    let mut vertical = None;
    for (key, value) in marks {
        match key.as_str() {
            "bold" => delta.bold = flag(key, value)?,
            "italic" => delta.italic = flag(key, value)?,
            "underline" => {
                delta.underline = match value {
                    Value::Bool(true) => Patch::Set(UnderlinePatch::default()),
                    Value::Bool(false) => Patch::Clear,
                    Value::String(style) if style == "none" => Patch::Clear,
                    Value::String(style) if UNDERLINE_STYLES.contains(&style.as_str()) => {
                        Patch::Set(UnderlinePatch {
                            style: Some(style.clone()),
                            color: None,
                        })
                    }
                    _ => {
                        return Err(
                            "mark \"underline\" must be a boolean or a Word underline style"
                                .to_owned(),
                        );
                    }
                }
            }
            "strike" => {
                delta.strike = match value
                    .as_bool()
                    .map(|on| if on { "sngStrike" } else { "noStrike" })
                    .or(value.as_str())
                {
                    Some("sngStrike") => Patch::Set(StrikePatch { double: false }),
                    Some("dblStrike") => Patch::Set(StrikePatch { double: true }),
                    Some("noStrike") => Patch::Clear,
                    _ => {
                        return Err("mark \"strike\" must be a boolean, \"sngStrike\", \"dblStrike\" or \"noStrike\"".to_owned());
                    }
                }
            }
            "verticalAlign" => {
                vertical = match value.as_str() {
                    Some(align @ ("superscript" | "subscript")) => Some(align),
                    Some("baseline") => {
                        return Err(
                            "mark \"verticalAlign\" \"baseline\" is not supported yet".to_owned()
                        );
                    }
                    _ => {
                        return Err(
                            "mark \"verticalAlign\" must be \"superscript\" or \"subscript\""
                                .to_owned(),
                        );
                    }
                }
            }
            "color" => {
                delta.color = match value {
                    Value::String(rgb)
                        if rgb.len() == 6 && rgb.bytes().all(|byte| byte.is_ascii_hexdigit()) =>
                    {
                        Patch::Set(ColorPatch::Rgb(rgb.clone()))
                    }
                    Value::Object(object) if object.len() == 1 => match object.get("themeColor") {
                        Some(Value::String(theme)) if THEME_COLORS.contains(&theme.as_str()) => {
                            Patch::Set(ColorPatch::Theme(theme.clone()))
                        }
                        _ => return Err("mark \"color\" has an unknown theme colour".to_owned()),
                    },
                    _ => {
                        return Err(
                            "mark \"color\" must be six hex digits or {\"themeColor\"}".to_owned()
                        );
                    }
                }
            }
            "highlight" => {
                delta.highlight = match value.as_str() {
                    Some("none") => {
                        return Err("mark \"highlight\" \"none\" is not supported yet".to_owned());
                    }
                    Some(name) if HIGHLIGHT_NAMES.contains(&name) => Patch::Set(name.to_owned()),
                    _ => {
                        return Err(
                            "mark \"highlight\" must be a Word highlight colour name".to_owned()
                        );
                    }
                }
            }
            "fontSizePt" => {
                delta.font_size = match value.as_f64() {
                    Some(points)
                        if (1.0..=1638.0).contains(&points) && (points * 2.0).fract() == 0.0 =>
                    {
                        Patch::Set(points)
                    }
                    _ => {
                        return Err(
                            "mark \"fontSizePt\" must be 1 to 1638 points in half-point steps"
                                .to_owned(),
                        );
                    }
                }
            }
            "fontFamily" => {
                delta.font_family = match value.as_str() {
                    Some(name)
                        if !name.trim().is_empty()
                            && name.chars().count() <= 255
                            && !name.chars().any(char::is_control) =>
                    {
                        Patch::Set(FontFamilyPatch {
                            ascii: name.to_owned(),
                            h_ansi: Some(name.to_owned()),
                        })
                    }
                    _ => return Err("mark \"fontFamily\" must be a font name".to_owned()),
                }
            }
            other => return Err(format!("unknown mark {other:?}")),
        }
    }
    let mut overlay: Vec<(String, Any)> = delta
        .to_attrs()
        .map_err(|error| error.to_string())?
        .into_iter()
        .map(|(key, value)| (key.to_string(), value))
        .collect();
    if let Some(align) = vertical {
        let on = |mark: &str| {
            if align == mark {
                Any::Bool(true)
            } else {
                Any::Null
            }
        };
        overlay.push(("superscript".to_owned(), on("superscript")));
        overlay.push(("subscript".to_owned(), on("subscript")));
    }
    overlay.sort_by(|(left, _), (right, _)| left.cmp(right));
    Ok(overlay)
}

/// The run attributes a host mark name stands for.
fn mark_attr_keys(mark: &str) -> Option<&'static [&'static str]> {
    Some(match mark {
        "bold" => &["bold"],
        "italic" => &["italic"],
        "underline" => &["underline"],
        "strike" => &["strike", "doubleStrike"],
        "verticalAlign" => &["superscript", "subscript", "vertAlign"],
        "color" => &["textColor", "color"],
        "highlight" => &["highlight"],
        "fontSizePt" => &["fontSize", "fontSizeCs", "sz"],
        "fontFamily" => &["fontFamily", "rFonts"],
        _ => return None,
    })
}

/// What a host `clear` returns to the styles: every character formatting attribute, or the
/// attributes of the named marks.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum ClearMarks {
    All,
    Keys(Vec<&'static str>),
}

/// Parses a host `clear`: `"all"` or a non-empty list of mark names, returned with the names.
pub(crate) fn parse_clear_marks(value: &Value) -> Result<(ClearMarks, Vec<String>), String> {
    match value {
        Value::String(all) if all == "all" => Ok((ClearMarks::All, Vec::new())),
        Value::Array(names) if !names.is_empty() => {
            let mut keys = Vec::new();
            let mut marks = Vec::new();
            for name in names {
                let name = name
                    .as_str()
                    .ok_or_else(|| "clear lists mark names".to_owned())?;
                let attrs = mark_attr_keys(name)
                    .ok_or_else(|| format!("unknown mark {name:?} in clear"))?;
                keys.extend_from_slice(attrs);
                marks.push(name.to_owned());
            }
            Ok((ClearMarks::Keys(keys), marks))
        }
        _ => Err("clear must be \"all\" or a non-empty list of mark names".to_owned()),
    }
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

/// Whether suggested formatting over `chunks` would cross a field result, a content control or
/// a field, which it refuses.
fn holds_suggestion_container<T: ReadTxn>(chunks: &[crate::ops::Chunk], txn: &T) -> bool {
    chunks.iter().any(|chunk| {
        chunk.attrs.contains_key("fieldResult")
            || matches!(&chunk.kind, crate::ops::ChunkKind::Embed(Some(map))
                if crate::map_string(map, txn, crate::KIND_KEY).as_deref() == Some("sdt")
                    || (crate::map_string(map, txn, crate::KIND_KEY).as_deref() == Some("field")
                        && crate::map_string(map, txn, "modelKind").as_deref() == Some("field")))
    })
}

/// Whether suggested formatting over `start..end` of a story, and over the text its block embeds
/// hold, would be refused anywhere, so a refusal leaves every part of the selection unchanged.
fn refuses_suggestion<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32) -> bool {
    let (spans, children) = block_free_spans(story, txn, start, end, false);
    spans.into_iter().any(|(at, len)| {
        holds_suggestion_container(&crate::ops::snapshot_range(story, txn, at, at + len), txn)
    }) || children.iter().any(|(child, _)| {
        story_ref(txn, child).is_ok_and(|text| {
            let len = text.len(txn);
            len > 0 && refuses_suggestion(&text, txn, 0, len)
        })
    })
}

/// The run attributes that are character formatting, the only ones clearing formatting removes:
/// weight, slant, underline, strike (single and double), raised or lowered text, colour,
/// highlight and shading, size and font (complex-script variants included), spacing, position,
/// scaling and kerning, caps, emboss, imprint, outline, shadow, hidden text, emphasis marks, text
/// effects, run style and proofing language. Script direction (`rtl`, `cs`) is kept, as clearing it
/// would reorder right-to-left text, and fields, hyperlinks, comment and bookmark anchors,
/// revisions and source bookkeeping always stay.
pub(crate) const FORMATTING_ATTRS: &[&str] = &[
    "bold",
    "italic",
    "boldCs",
    "italicCs",
    "underline",
    "strike",
    "doubleStrike",
    "superscript",
    "subscript",
    "vertAlign",
    "textColor",
    "color",
    "highlight",
    "shading",
    "fontSize",
    "fontSizeCs",
    "fontFamily",
    "rFonts",
    "sz",
    "characterSpacing",
    "allCaps",
    "smallCaps",
    "emboss",
    "imprint",
    "hidden",
    "textShadow",
    "textOutline",
    "textEffect",
    "modernTextEffects",
    "emphasisMark",
    "runStyle",
    "language",
    "lang",
    "snapToGrid",
];

/// The stories whose text a block embed holds: a table's cells or a block control's content.
fn block_child_stories<T: ReadTxn>(embed: &MapRef, txn: &T) -> Vec<String> {
    let mut stories = match map_string(embed, txn, KIND_KEY).as_deref() {
        Some("table") => {
            let mut cells = HashSet::new();
            if let Some(Out::Any(rows)) = embed.get(txn, "rows") {
                crate::undo::collect_row_stories(&rows, &mut cells);
            }
            cells.into_iter().collect()
        }
        Some("blockSdt") => map_string(embed, txn, "story").into_iter().collect(),
        _ => Vec::new(),
    };
    stories.sort_unstable();
    stories
}

/// The parts of `start..end` outside block embeds, and the stories inside the embeds it covers,
/// each with whether a block content control holds it.
fn block_free_spans<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    start: u32,
    end: u32,
    text_only: bool,
) -> (Vec<(u32, u32)>, Vec<(String, bool)>) {
    let mut spans = Vec::new();
    let mut children = Vec::new();
    let mut from = start;
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let chunk_end = offset + out_len(&diff.insert);
        if chunk_end.min(end) > offset.max(start) {
            let block = matches!(&diff.insert, Out::YMap(embed) if map_string(embed, txn, KIND_KEY)
                .is_some_and(|kind| crate::segments::is_block_embed(&kind)));
            let skipped = block || (text_only && !matches!(diff.insert, Out::Any(Any::String(_))));
            if skipped {
                let cut = offset.max(start);
                if cut > from {
                    spans.push((from, cut - from));
                }
                from = chunk_end.min(end);
                if block && let Out::YMap(embed) = &diff.insert {
                    let control = map_string(embed, txn, KIND_KEY).as_deref() == Some("blockSdt");
                    children.extend(
                        block_child_stories(embed, txn)
                            .into_iter()
                            .map(|child| (child, control)),
                    );
                }
            }
        }
        offset = chunk_end;
        if offset >= end {
            break;
        }
    }
    if end > from {
        spans.push((from, end - from));
    }
    (spans, children)
}

/// Tests whether every text unit carries a non-null attribute, in block embeds too; `saw_text`
/// reports whether there was any.
fn text_has<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    (start, end): (u32, u32),
    key: &str,
    saw_text: &mut bool,
) -> bool {
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        let chunk_end = offset + len;
        if chunk_end.min(end) > offset.max(start) {
            match &diff.insert {
                Out::Any(Any::String(_)) => {
                    *saw_text = true;
                    if !is_active(diff.attributes.as_deref(), key) {
                        return false;
                    }
                }
                Out::YMap(embed) => {
                    for child in block_child_stories(embed, txn) {
                        if let Ok(child) = story_ref(txn, &child)
                            && !text_has(&child, txn, (0, child.len(txn)), key, saw_text)
                        {
                            return false;
                        }
                    }
                }
                _ => {}
            }
        }
        offset = chunk_end;
        if offset >= end {
            break;
        }
    }
    true
}

fn all_text_has<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32, key: &str) -> bool {
    let mut saw_text = false;
    text_has(story, txn, (start, end), key, &mut saw_text) && saw_text
}

/// Why formatting through a table or a block content control kept one of their paragraphs as it
/// was: saving the paragraph rewritten would drop content its source holds.
#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub enum KeptReason {
    /// Alternate content, such as a text box or drawing with its fallback.
    AlternateContent,
    /// A page or column break.
    Break,
    /// A content control inside the paragraph.
    ContentControl,
    /// A field nested in another field's result, or field characters inside a hyperlink.
    Field,
    /// A bookmark or permission anchor inside a hyperlink.
    HyperlinkMarker,
    /// Source XML that could not be read.
    Unreadable,
}

impl KeptReason {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::AlternateContent => "alternateContent",
            Self::Break => "break",
            Self::ContentControl => "contentControl",
            Self::Field => "field",
            Self::HyperlinkMarker => "hyperlinkMarker",
            Self::Unreadable => "unreadable",
        }
    }
}

/// The paragraphs a formatting edit kept unchanged, counted by reason.
pub type KeptUnchanged = BTreeMap<KeptReason, u32>;

struct Keeper<'d> {
    doc: &'d EditingDoc,
    sources: Option<Option<crate::identity::SourceParagraphs>>,
    kept: KeptUnchanged,
}

impl<'d> Keeper<'d> {
    fn new(doc: &'d EditingDoc) -> Self {
        Self {
            doc,
            sources: None,
            kept: KeptUnchanged::new(),
        }
    }

    fn rewrite_loss(&mut self, key: &str) -> Option<KeptReason> {
        let doc = self.doc;
        self.sources
            .get_or_insert_with(|| {
                doc.source_index()
                    .map(crate::identity::SourceParagraphs::new)
            })
            .as_mut()?
            .rewrite_loss(key)
    }

    fn keep(&mut self, reason: KeptReason) {
        *self.kept.entry(reason).or_default() += 1;
    }
}

/// A piece of a story as one reading of it saw it. Formatting moves no offset, so the pieces stay
/// right for every paragraph a loop has yet to format.
struct StoryChunk {
    start: u32,
    end: u32,
    text: bool,
    block: bool,
    attrs: Option<Box<Attrs>>,
}

fn story_chunks<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<StoryChunk> {
    let mut chunks = Vec::new();
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let end = offset + out_len(&diff.insert);
        chunks.push(StoryChunk {
            start: offset,
            end,
            text: matches!(diff.insert, Out::Any(Any::String(_))),
            block: matches!(&diff.insert, Out::YMap(embed) if map_string(embed, txn, KIND_KEY)
                .is_some_and(|kind| crate::segments::is_block_embed(&kind))),
            attrs: diff.attributes.clone(),
        });
        offset = end;
    }
    chunks
}

fn chunks_within(chunks: &[StoryChunk], start: u32, end: u32) -> impl Iterator<Item = &StoryChunk> {
    let first = chunks.partition_point(|chunk| chunk.end <= start);
    chunks[first..]
        .iter()
        .take_while(move |chunk| chunk.start < end)
        .filter(move |chunk| chunk.end.min(end) > chunk.start.max(start))
}

/// [`clearing_attrs`] over `chunks`.
fn clearing_attrs_in(chunks: &[StoryChunk], start: u32, end: u32) -> Attrs {
    let mut keys: BTreeMap<String, ()> = BTreeMap::new();
    for chunk in chunks_within(chunks, start, end) {
        for (key, value) in chunk.attrs.iter().flat_map(|attrs| attrs.iter()) {
            if *value != Any::Null && FORMATTING_ATTRS.contains(&key.as_ref()) {
                keys.insert(key.to_string(), ());
            }
        }
    }
    keys.into_keys()
        .map(|key| (Arc::from(key.as_str()), Any::Null))
        .collect()
}

/// [`changes_text`] over `chunks`.
fn changes_text_in(chunks: &[StoryChunk], start: u32, end: u32, attrs: &Attrs) -> bool {
    chunks_within(chunks, start, end)
        .filter(|chunk| chunk.text)
        .any(|chunk| {
            let current = chunk.attrs.as_deref();
            attrs.iter().any(|(key, value)| {
                current
                    .and_then(|current| current.get(key))
                    .unwrap_or(&Any::Null)
                    != value
            })
        })
}

/// The spans of [`block_free_spans`] over `chunks`.
fn free_spans_in(chunks: &[StoryChunk], start: u32, end: u32, text_only: bool) -> Vec<(u32, u32)> {
    let mut spans = Vec::new();
    let mut from = start;
    for chunk in chunks_within(chunks, start, end) {
        if chunk.block || (text_only && !chunk.text) {
            let cut = chunk.start.max(start);
            if cut > from {
                spans.push((from, cut - from));
            }
            from = chunk.end.min(end);
        }
    }
    if end > from {
        spans.push((from, end - from));
    }
    spans
}

/// The paragraphs of a cell or content-control story: the range each spans, its own paragraph
/// mark included, and the paragraph's id.
fn story_paragraphs<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<(u32, u32, Option<String>)> {
    let mut paragraphs = Vec::new();
    let mut start = 0;
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let end = offset + out_len(&diff.insert);
        if let Out::YMap(embed) = &diff.insert
            && crate::is_pilcrow(embed, txn)
        {
            paragraphs.push((start, end, map_string(embed, txn, crate::PARA_ID)));
            start = end;
        }
        offset = end;
    }
    if offset > start {
        paragraphs.push((start, offset, None));
    }
    paragraphs
}

/// Whether setting `attrs` would change any text in `start..end`.
fn changes_text<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32, attrs: &Attrs) -> bool {
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let chunk_end = offset + out_len(&diff.insert);
        if chunk_end.min(end) > offset.max(start) && matches!(diff.insert, Out::Any(Any::String(_)))
        {
            let current = diff.attributes.as_deref();
            if attrs.iter().any(|(key, value)| {
                current
                    .and_then(|current| current.get(key))
                    .unwrap_or(&Any::Null)
                    != value
            }) {
                return true;
            }
        }
        offset = chunk_end;
        if offset >= end {
            break;
        }
    }
    false
}

/// The unprotected formatting set anywhere in `start..end`, each mapped to null.
fn clearing_attrs<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32) -> Attrs {
    let mut keys: BTreeMap<String, ()> = BTreeMap::new();
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let chunk_end = offset + out_len(&diff.insert);
        if chunk_end.min(end) > offset.max(start)
            && let Some(attrs) = diff.attributes.as_deref()
        {
            for (key, value) in attrs.iter() {
                if *value != Any::Null && FORMATTING_ATTRS.contains(&key.as_ref()) {
                    keys.insert(key.to_string(), ());
                }
            }
        }
        offset = chunk_end;
        if offset >= end {
            break;
        }
    }
    keys.into_keys()
        .map(|key| (Arc::from(key.as_str()), Any::Null))
        .collect()
}

impl EditingDoc {
    /// Toggles one simple format across a range.
    pub fn toggle_format(
        &self,
        ctx: &EditCtx,
        range: StoryRange,
        format: SimpleFormat,
    ) -> OpResult<Receipt> {
        self.toggle_format_keeping(ctx, range, format)
            .map(|(receipt, _)| receipt)
    }

    /// [`Self::toggle_format`], also counting the table-cell and content-control paragraphs it
    /// kept unchanged because saving them rewritten would lose content.
    pub fn toggle_format_keeping(
        &self,
        ctx: &EditCtx,
        mut range: StoryRange,
        format: SimpleFormat,
    ) -> OpResult<(Receipt, KeptUnchanged)> {
        let len = range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        crate::check_range(&story, &txn, range.start, len)?;
        (range.start, range.end) =
            crate::ops::code_point_range(&story, &txn, range.start, range.end);
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
        if ctx.is_suggesting() && refuses_suggestion(&story, &txn, range.start, range.end) {
            return Err(OpError::SuggestedFormattingOverContainer);
        }
        let mut keeper = Keeper::new(self);
        let revision_ids = self.apply_run_format_through(
            &mut txn,
            ctx,
            &story,
            &range,
            Some(&attrs),
            &mut keeper,
        )?;
        let loc_range =
            crate::op::loc_range_in_txn(&range.story, &story, &txn, range.start, range.end)?;
        Ok((
            Receipt {
                range: Some(loc_range),
                revision_ids,
                ..Receipt::default()
            },
            keeper.kept,
        ))
    }

    /// Applies a formatting delta across a range in one transaction.
    pub fn format_range(
        &self,
        ctx: &EditCtx,
        range: StoryRange,
        delta: &InlineFormatDelta,
    ) -> OpResult<Receipt> {
        self.format_range_keeping(ctx, range, delta)
            .map(|(receipt, _)| receipt)
    }

    /// [`Self::format_range`], also counting the table-cell and content-control paragraphs it
    /// kept unchanged because saving them rewritten would lose content.
    pub fn format_range_keeping(
        &self,
        ctx: &EditCtx,
        mut range: StoryRange,
        delta: &InlineFormatDelta,
    ) -> OpResult<(Receipt, KeptUnchanged)> {
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
        if ctx.is_suggesting() && refuses_suggestion(&story, &txn, range.start, range.end) {
            return Err(OpError::SuggestedFormattingOverContainer);
        }
        let mut keeper = Keeper::new(self);
        let revision_ids = self.apply_run_format_through(
            &mut txn,
            ctx,
            &story,
            &range,
            Some(&attrs),
            &mut keeper,
        )?;
        let loc_range =
            crate::op::loc_range_in_txn(&range.story, &story, &txn, range.start, range.end)?;
        let receipt = Receipt {
            range: Some(loc_range),
            revision_ids,
            ..Receipt::default()
        };
        Ok((receipt, keeper.kept))
    }

    /// The attributes that return each run of `start..end` of `story_id` to what its styles give,
    /// as seeding gives a run without direct formatting: the paragraph style `paragraph_style`
    /// and, unless `clear` is [`ClearMarks::All`], the run's own character style. Keys no style
    /// supplies are removed.
    pub(crate) fn style_reset_pieces<T: ReadTxn>(
        &self,
        txn: &T,
        story: &TextRef,
        story_id: &str,
        (start, end): (u32, u32),
        clear: &ClearMarks,
        paragraph_style: Option<&str>,
    ) -> OpResult<Vec<(u32, u32, Attrs)>> {
        let source = self.source_metadata();
        let mut pieces = Vec::new();
        for chunk in crate::ops::snapshot_range(story, txn, start, end) {
            let (from, to) = (chunk.start.max(start), chunk.end().min(end));
            if from >= to {
                continue;
            }
            let run_style = match clear {
                ClearMarks::All => None,
                ClearMarks::Keys(_) => chunk.attrs.get("runStyle").cloned(),
            };
            let base: BTreeMap<String, Any> = match &source {
                Some(source) => source
                    .seeded_run_attrs(story_id, paragraph_style, run_style.as_ref())
                    .map_err(OpError::InvalidFormatValue)?
                    .into_iter()
                    .collect(),
                None => BTreeMap::new(),
            };
            let keys: Vec<&str> = match clear {
                ClearMarks::Keys(keys) => keys.clone(),
                ClearMarks::All => FORMATTING_ATTRS
                    .iter()
                    .copied()
                    .chain(base.keys().map(String::as_str))
                    .filter(|key| !PROTECTED_ATTRS.contains(key))
                    .collect(),
            };
            let attrs: Attrs = keys
                .into_iter()
                .map(|key| (Arc::from(key), base.get(key).cloned().unwrap_or(Any::Null)))
                .collect();
            match pieces.last_mut() {
                Some((_, last_end, last)) if *last_end == from && *last == attrs => *last_end = to,
                _ => pieces.push((from, to, attrs)),
            }
        }
        Ok(pieces)
    }

    /// A host `formatText` over one non-empty range in one transaction: `clear` returns the
    /// named marks to the styles, then `set` writes its attributes. Returns the revision ids and
    /// the paragraphs kept unchanged because saving them rewritten would lose content.
    pub(crate) fn format_for_batch(
        &self,
        ctx: &EditCtx,
        mut range: StoryRange,
        set: &[(String, Any)],
        clear: Option<&ClearMarks>,
        paragraph_style: Option<&str>,
    ) -> OpResult<(Vec<String>, KeptUnchanged)> {
        let len = range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        crate::check_range(&story, &txn, range.start, len)?;
        (range.start, range.end) =
            crate::ops::code_point_range(&story, &txn, range.start, range.end);
        if ctx.is_suggesting() && refuses_suggestion(&story, &txn, range.start, range.end) {
            return Err(OpError::SuggestedFormattingOverContainer);
        }
        let set: Attrs = set
            .iter()
            .map(|(key, value)| (Arc::from(key.as_str()), value.clone()))
            .collect();
        let pieces = match clear {
            Some(clear) => self.style_reset_pieces(
                &txn,
                &story,
                &range.story,
                (range.start, range.end),
                clear,
                paragraph_style,
            )?,
            None => vec![(range.start, range.end, Attrs::new())],
        };
        let mut keeper = Keeper::new(self);
        let mut revision_ids: Vec<String> = Vec::new();
        for (start, end, mut attrs) in pieces {
            attrs.extend(set.iter().map(|(key, value)| (key.clone(), value.clone())));
            let piece = StoryRange::new(range.story.as_str(), start, end);
            for id in self.apply_run_format_through(
                &mut txn,
                ctx,
                &story,
                &piece,
                Some(&attrs),
                &mut keeper,
            )? {
                if !revision_ids.contains(&id) {
                    revision_ids.push(id);
                }
            }
        }
        Ok((revision_ids, keeper.kept))
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
        if ctx.is_suggesting() {
            let mut runs = Vec::new();
            let mut changed = false;
            for chunk in crate::ops::snapshot_range(&story, &txn, range.start, range.end) {
                let start = chunk.start.max(range.start);
                let end = chunk.end().min(range.end);
                if start >= end {
                    continue;
                }
                if [INS, DEL, RPR_CHANGE]
                    .iter()
                    .any(|key| chunk.attrs.contains_key(*key))
                {
                    return Err(OpError::OverlapsTrackedChange);
                }
                let crate::ops::ChunkKind::Text(text) = &chunk.kind else {
                    return Err(OpError::InvalidFormatValue(
                        "suggested hyperlinks require a text range".to_owned(),
                    ));
                };
                let text = String::from_utf16_lossy(
                    &text.encode_utf16().collect::<Vec<_>>()
                        [(start - chunk.start) as usize..(end - chunk.start) as usize],
                );
                let mut attrs = chunk.attrs.clone();
                changed |= attrs.get(HYPERLINK) != hyperlink.as_ref();
                attrs.remove(HYPERLINK);
                if let Some(value) = &hyperlink {
                    attrs.insert(HYPERLINK.to_owned(), value.clone());
                }
                runs.push(crate::RichRun {
                    text,
                    attrs: attrs.into_iter().collect(),
                });
            }
            if !changed {
                return Ok(Receipt::default());
            }
            let (receipt, at) = self.replace_range_rich_placed_in(&mut txn, ctx, range, &runs)?;
            story.format(
                &mut txn,
                at,
                len,
                Attrs::from([(Arc::from(HYPERLINK), hyperlink.unwrap_or(Any::Null))]),
            );
            return Ok(receipt);
        }
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
    pub fn clear_formatting(&self, ctx: &EditCtx, range: StoryRange) -> OpResult<Receipt> {
        self.clear_formatting_keeping(ctx, range)
            .map(|(receipt, _)| receipt)
    }

    /// [`Self::clear_formatting`], also counting the table-cell and content-control paragraphs
    /// it kept unchanged because saving them rewritten would lose content.
    pub fn clear_formatting_keeping(
        &self,
        ctx: &EditCtx,
        mut range: StoryRange,
    ) -> OpResult<(Receipt, KeptUnchanged)> {
        let len = range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        crate::check_range(&story, &txn, range.start, len)?;
        (range.start, range.end) =
            crate::ops::code_point_range(&story, &txn, range.start, range.end);
        if ctx.is_suggesting() && refuses_suggestion(&story, &txn, range.start, range.end) {
            return Err(OpError::SuggestedFormattingOverContainer);
        }
        let mut keeper = Keeper::new(self);
        let revision_ids =
            self.apply_run_format_through(&mut txn, ctx, &story, &range, None, &mut keeper)?;
        let loc_range =
            crate::op::loc_range_in_txn(&range.story, &story, &txn, range.start, range.end)?;
        let receipt = Receipt {
            range: Some(loc_range),
            revision_ids,
            ..Receipt::default()
        };
        Ok((receipt, keeper.kept))
    }

    /// [`Self::apply_run_format`] over `range`, leaving block embeds unmarked and formatting the
    /// text they hold, so a selection across a table formats its cells and not the table. A
    /// paragraph inside a block, or one of several the range spans, that saving rewritten would
    /// lose content of stays unchanged and is counted in `keeper`, as does a single paragraph
    /// whose fields saving rewritten would flatten; `delta` of `None` clears formatting instead.
    fn apply_run_format_through(
        &self,
        txn: &mut TransactionMut<'_>,
        ctx: &EditCtx,
        story: &TextRef,
        range: &StoryRange,
        delta: Option<&Attrs>,
        keeper: &mut Keeper<'_>,
    ) -> OpResult<Vec<String>> {
        let attrs = match delta {
            Some(delta) => delta.clone(),
            None => clearing_attrs(story, txn, range.start, range.end),
        };
        let (spans, children) = block_free_spans(story, txn, range.start, range.end, false);
        let paragraphs: Vec<_> = story_paragraphs(story, txn)
            .into_iter()
            .filter(|(start, end, _)| *end > range.start && *start < range.end)
            .collect();
        let mut revision_ids = Vec::new();
        if paragraphs.len() > 1 {
            let chunks = story_chunks(story, txn);
            let mut formatted = Vec::new();
            for (start, end, key) in paragraphs {
                let (from, to) = (start.max(range.start), end.min(range.end));
                let attrs = match delta {
                    Some(delta) => delta.clone(),
                    None => clearing_attrs_in(&chunks, from, to),
                };
                if attrs.is_empty() {
                    continue;
                }
                if changes_text_in(&chunks, from, to, &attrs)
                    && let Some(reason) = key.as_deref().and_then(|key| keeper.rewrite_loss(key))
                {
                    keeper.keep(reason);
                    continue;
                }
                for (at, len) in free_spans_in(&chunks, from, to, false) {
                    formatted.push((at, len, attrs.clone()));
                }
            }
            revision_ids.extend(self.apply_run_formats(
                txn,
                ctx,
                story,
                &range.story,
                formatted,
            )?);
        } else if let Some(key) = paragraphs.first().and_then(|(_, _, key)| key.as_deref())
            && changes_text(story, txn, range.start, range.end, &attrs)
            && keeper.rewrite_loss(key) == Some(KeptReason::Field)
        {
            keeper.keep(KeptReason::Field);
        } else {
            for (at, len) in spans {
                let span = StoryRange::new(range.story.clone(), at, at + len);
                revision_ids.extend(self.apply_run_format(
                    txn,
                    ctx,
                    story,
                    &span,
                    attrs.clone(),
                )?);
            }
        }
        for (child, _) in children {
            revision_ids.extend(self.format_story_paragraphs(txn, ctx, &child, delta, keeper)?);
        }
        revision_ids.dedup();
        Ok(revision_ids)
    }

    /// Formats or clears (`delta` of `None`) every paragraph of a cell or content-control story
    /// and of the blocks nested in it, keeping unchanged each one that saving rewritten would
    /// lose content of.
    fn format_story_paragraphs(
        &self,
        txn: &mut TransactionMut<'_>,
        ctx: &EditCtx,
        child: &str,
        delta: Option<&Attrs>,
        keeper: &mut Keeper<'_>,
    ) -> OpResult<Vec<String>> {
        let Ok(text) = story_ref(txn, child) else {
            return Ok(Vec::new());
        };
        let mut revision_ids = Vec::new();
        let chunks = story_chunks(&text, txn);
        let mut formatted = Vec::new();
        for (start, end, key) in story_paragraphs(&text, txn) {
            let attrs = match delta {
                Some(delta) => delta.clone(),
                None => clearing_attrs_in(&chunks, start, end),
            };
            if attrs.is_empty() || !changes_text_in(&chunks, start, end, &attrs) {
                continue;
            }
            if let Some(reason) = key.as_deref().and_then(|key| keeper.rewrite_loss(key)) {
                keeper.keep(reason);
                continue;
            }
            for (at, len) in free_spans_in(&chunks, start, end, true) {
                formatted.push((at, len, attrs.clone()));
            }
        }
        revision_ids.extend(self.apply_run_formats(txn, ctx, &text, child, formatted)?);
        let len = text.len(txn);
        for (nested, _) in block_free_spans(&text, txn, 0, len, true).1 {
            revision_ids.extend(self.format_story_paragraphs(txn, ctx, &nested, delta, keeper)?);
        }
        Ok(revision_ids)
    }

    /// Applies each `(at, len, attrs)` of `spans`, in document order, in one walk of `story`: a
    /// format call of its own would seek from the start of the story for every span.
    fn apply_run_formats(
        &self,
        txn: &mut TransactionMut<'_>,
        ctx: &EditCtx,
        story: &TextRef,
        story_id: &str,
        spans: Vec<(u32, u32, Attrs)>,
    ) -> OpResult<Vec<String>> {
        let one_by_one = ctx.is_suggesting()
            || spans
                .iter()
                .any(|(_, _, attrs)| matches!(attrs.get("language"), Some(Any::Map(_))));
        if one_by_one {
            let mut revision_ids = Vec::new();
            for (at, len, attrs) in spans {
                let span = StoryRange::new(story_id.to_owned(), at, at + len);
                revision_ids.extend(self.apply_run_format(txn, ctx, story, &span, attrs)?);
            }
            return Ok(revision_ids);
        }
        let mut deltas: Vec<Delta<In>> = Vec::with_capacity(spans.len() * 2);
        let mut cursor = 0;
        for (at, len, attrs) in spans {
            if attrs.is_empty() {
                continue;
            }
            if at > cursor {
                deltas.push(Delta::Retain(at - cursor, None));
            }
            deltas.push(Delta::Retain(len, Some(Box::new(attrs))));
            cursor = at + len;
        }
        story.apply_delta(txn, deltas);
        Ok(Vec::new())
    }

    fn apply_run_format(
        &self,
        txn: &mut TransactionMut<'_>,
        ctx: &EditCtx,
        story: &TextRef,
        range: &StoryRange,
        delta: Attrs,
    ) -> OpResult<Vec<String>> {
        if !ctx.is_suggesting() {
            if !delta.is_empty() {
                if let Some(language @ Any::Map(_)) = delta.get("language") {
                    let chunks = crate::ops::snapshot_range(story, txn, range.start, range.end);
                    for chunk in chunks {
                        let start = chunk.start.max(range.start);
                        let end = chunk.end().min(range.end);
                        if start >= end {
                            continue;
                        }
                        let mut patch = delta.clone();
                        if matches!(chunk.kind, crate::ops::ChunkKind::Text(_)) {
                            patch.insert(
                                Arc::from("language"),
                                merge_language(chunk.attrs.get("language"), language),
                            );
                        } else {
                            patch.remove("language");
                        }
                        if !patch.is_empty() {
                            story.format(txn, start, end - start, patch);
                        }
                    }
                } else {
                    story.format(txn, range.start, range.end - range.start, delta);
                }
            }
            return Ok(Vec::new());
        }
        let chunks = crate::ops::snapshot_range(story, txn, range.start, range.end);
        if holds_suggestion_container(&chunks, txn) {
            return Err(OpError::SuggestedFormattingOverContainer);
        }
        let mut revision_ids = Vec::new();
        let mut new_id = None;
        for chunk in chunks {
            let start = chunk.start.max(range.start);
            let end = chunk.end().min(range.end);
            if start >= end
                || matches!(chunk.kind, crate::ops::ChunkKind::Pilcrow(_))
                || matches!(&chunk.kind, crate::ops::ChunkKind::Embed(Some(map))
                    if crate::map_string(map, txn, crate::KIND_KEY).as_deref() == Some("field")
                        && crate::map_string(map, txn, "modelKind").as_deref() == Some("commentReference"))
            {
                continue;
            }
            let mut delta = delta.clone();
            if let Some(language @ Any::Map(_)) = delta.get("language").cloned() {
                delta.insert(
                    Arc::from("language"),
                    merge_language(chunk.attrs.get("language"), &language),
                );
            }
            let mut current = chunk.attrs.clone();
            for (key, value) in &delta {
                if *value == Any::Null {
                    current.remove(key.as_ref());
                } else {
                    current.insert(key.to_string(), value.clone());
                }
            }
            if current == chunk.attrs {
                continue;
            }
            let mut patch = delta.clone();
            let mut changes = match chunk.attrs.get(RPR_CHANGE) {
                Some(Any::Array(changes)) => changes.to_vec(),
                _ => Vec::new(),
            };
            if changes.is_empty() {
                let id = new_id.get_or_insert_with(|| self.next_id());
                let previous: HashMap<_, _> = chunk
                    .attrs
                    .iter()
                    .filter(|(key, _)| !matches!(key.as_str(), INS | DEL | HYPERLINK | RPR_CHANGE))
                    .map(|(key, value)| (key.clone(), value.clone()))
                    .collect();
                changes.push(Any::Map(Arc::new(HashMap::from([
                    ("type".to_owned(), Any::from("runPropertyChange")),
                    (
                        "info".to_owned(),
                        crate::revision_value(id, &ctx.revision_author()),
                    ),
                    (
                        "previousFormatting".to_owned(),
                        crate::seed::mark_run_formatting(previous.iter()).unwrap_or(Any::Null),
                    ),
                    (
                        "previousAttributes".to_owned(),
                        Any::Map(Arc::new(previous)),
                    ),
                ]))));
            }
            for change in &mut changes {
                if let Some((id, ..)) = crate::queries::revision_parts(change)
                    && !revision_ids.contains(&id)
                {
                    revision_ids.push(id);
                }
                if let Any::Map(values) = change {
                    let mut values = values.as_ref().clone();
                    values.insert(
                        "currentFormatting".to_owned(),
                        crate::seed::mark_run_formatting(current.iter()).unwrap_or(Any::Null),
                    );
                    *change = Any::Map(Arc::new(values));
                }
            }
            patch.insert(Arc::from(RPR_CHANGE), Any::Array(Arc::from(changes)));
            story.format(txn, start, end - start, patch);
        }
        Ok(revision_ids)
    }
}

pub(crate) fn previous_run_properties(attrs: &BTreeMap<String, Any>, change: &Any) -> Attrs {
    let mut patch: Attrs = attrs
        .keys()
        .filter(|key| !PROTECTED_ATTRS.contains(&key.as_str()))
        .map(|key| (Arc::from(key.as_str()), Any::Null))
        .collect();
    if let Any::Map(change) = change {
        let previous = match change.get("previousAttributes") {
            Some(Any::Map(previous)) => previous
                .iter()
                .map(|(key, value)| (key.clone(), value.clone()))
                .collect(),
            _ => change
                .get("previousFormatting")
                .map(crate::seed::mark_run_attrs)
                .unwrap_or_default(),
        };
        for (key, value) in previous {
            if !PROTECTED_ATTRS.contains(&key.as_str()) {
                patch.insert(Arc::from(key), value);
            }
        }
    }
    patch
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

#[cfg(test)]
#[path = "format_tests.rs"]
mod tests;
