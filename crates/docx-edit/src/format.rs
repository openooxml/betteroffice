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

/// The `w:dir` and `w:bdo` elements around a run, outermost first, as the parsed run holds them.
pub(crate) const DIRECTION_SCOPES: &str = "directionScopes";

pub(crate) const PROTECTED_ATTRS: [&str; 5] = [INS, DEL, HYPERLINK, RPR_CHANGE, DIRECTION_SCOPES];

const STRUCTURAL_ATTRS: [&str; 1] = ["fieldResult"];

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

pub(crate) fn merge_font_family(previous: Option<&Any>, patch: &Any, east_asian: bool) -> Any {
    let Any::Map(patch) = patch else {
        return patch.clone();
    };
    let mut fonts = match previous {
        Some(Any::Map(fonts)) => fonts.as_ref().clone(),
        _ => HashMap::new(),
    };
    fonts.extend(
        patch
            .iter()
            .map(|(key, value)| (key.clone(), value.clone())),
    );
    fonts.remove("asciiTheme");
    fonts.remove("hAnsiTheme");

    let east_asia = patch
        .get("eastAsia")
        .or_else(|| patch.get("ascii").filter(|_| east_asian));
    if let Some(east_asia) = east_asia {
        fonts.insert("eastAsia".to_owned(), east_asia.clone());
        fonts.remove("eastAsiaTheme");
    }
    Any::Map(Arc::new(fonts))
}

pub(crate) fn run_font_slot_flags(complex_script: Option<&Any>, fonts: Option<&Any>) -> (bool, bool) {
    let complex_script = match complex_script {
        Some(Any::Map(value)) => value.get("enabled") != Some(&Any::Bool(false)),
        Some(value) => *value != Any::Null && *value != Any::Bool(false),
        None => false,
    };
    let hint = matches!(fonts, Some(Any::Map(fonts))
        if matches!(fonts.get("hint"), Some(Any::String(hint)) if hint.as_ref() == "eastAsia"));
    (complex_script, hint)
}

fn font_pick_characters(
    text: &str,
    (complex_script, hint): (bool, bool),
) -> impl Iterator<Item = (char, bool)> + '_ {
    ooxml_text::font_slot_characters(text, complex_script, hint).map(move |(ch, slot)| {
        let east_asian = slot == 2
            || !complex_script
                && matches!(
                    ch as u32,
                    0x16fe0..=0x18dff | 0x1aff0..=0x1b2ff | 0x1f200..=0x1f2ff | 0x20000..=0x3ffff
                );
        (ch, east_asian)
    })
}

pub(crate) fn has_east_asian_font_pick_text(text: &str) -> bool {
    !text.is_ascii() && font_pick_characters(text, (false, false)).any(|(_, east_asian)| east_asian)
}

fn font_pick_segments(
    text: &str,
    from: u32,
    to: u32,
    flags: (bool, bool),
) -> impl Iterator<Item = (u32, u32, bool)> + '_ {
    let unsegmented = flags.0 || text.is_ascii();
    let mut single = Some((from, to, false));
    let mut offset = 0;
    let mut characters = font_pick_characters(text, flags)
        .map(move |(ch, east_asian)| {
            let start = offset;
            offset += ch.len_utf16() as u32;
            (start, offset, east_asian)
        })
        .skip_while(move |&(_, end, _)| end <= from)
        .take_while(move |&(start, _, _)| start < to)
        .peekable();
    std::iter::from_fn(move || {
        if unsegmented {
            return single.take();
        }
        let (start, mut end, east_asian) = characters.next()?;
        while let Some((_, next_end, _)) = characters.next_if(|&(_, _, ea)| ea == east_asian) {
            end = next_end;
        }
        Some((start.max(from), end.min(to), east_asian))
    })
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
                let fonts = HashMap::from([
                    ("ascii".into(), Any::from(patch.ascii.as_str())),
                    (
                        "hAnsi".into(),
                        Any::from(patch.h_ansi.as_deref().unwrap_or(&patch.ascii)),
                    ),
                ]);
                put("fontFamily", Any::Map(Arc::new(fonts)));
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
        chunk.attr_active("fieldResult")
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
        crate::story_text(txn, child).is_some_and(|text| {
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
    "typography",
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
                        if let Some(child) = crate::story_text(txn, &child)
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
    font_text: Option<Arc<str>>,
    block: bool,
    pilcrow: bool,
    attrs: Option<Box<Attrs>>,
}

fn story_chunks<T: ReadTxn>(story: &TextRef, txn: &T, merge_font: bool) -> Vec<StoryChunk> {
    let mut chunks = Vec::new();
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let end = offset + out_len(&diff.insert);
        let block = matches!(&diff.insert, Out::YMap(embed) if map_string(embed, txn, KIND_KEY)
            .is_some_and(|kind| crate::segments::is_block_embed(&kind)));
        let pilcrow = matches!(&diff.insert, Out::YMap(embed) if crate::is_pilcrow(embed, txn));
        chunks.push(StoryChunk {
            start: offset,
            end,
            text: matches!(diff.insert, Out::Any(Any::String(_))),
            font_text: match diff.insert {
                Out::Any(Any::String(text)) if merge_font => Some(text),
                _ => None,
            },
            block,
            pilcrow,
            attrs: diff.attributes,
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
fn changes_text_in(
    chunks: &[StoryChunk],
    start: u32,
    end: u32,
    attrs: &Attrs,
    font_merge: Option<FontMerge<'_>>,
) -> bool {
    changes_units_in(chunks, start, end, attrs, true, font_merge)
}

/// Whether setting `attrs` would change a unit in `start..end` that a pass writes: text only, or
/// with `text_only` false every unit but block embeds. A paragraph mark holds no run formatting
/// of its own, so one without a key counts as already holding what the pass writes for it.
fn changes_units_in(
    chunks: &[StoryChunk],
    start: u32,
    end: u32,
    attrs: &Attrs,
    text_only: bool,
    font_merge: Option<FontMerge<'_>>,
) -> bool {
    chunks_within(chunks, start, end)
        .filter(|chunk| if text_only { chunk.text } else { !chunk.block })
        .any(|chunk| {
            attrs_change_chunk(
                attrs,
                chunk.attrs.as_deref(),
                chunk.pilcrow,
                font_merge,
                chunk.font_text.as_deref().unwrap_or_default(),
                (
                    start.saturating_sub(chunk.start),
                    end.min(chunk.end) - chunk.start,
                ),
            )
        })
}

fn attrs_change_chunk(
    attrs: &Attrs,
    current: Option<&Attrs>,
    pilcrow: bool,
    font_merge: Option<FontMerge<'_>>,
    text: &str,
    (from, to): (u32, u32),
) -> bool {
    attrs.iter().any(|(key, value)| {
        let previous = current.and_then(|current| current.get(key));
        if pilcrow && previous.is_none() {
            return false;
        }
        let previous = previous.unwrap_or(&Any::Null);
        if let Some(font_merge) = font_merge
            && key.as_ref() == "fontFamily"
        {
            let (base, flags) = font_merge.chunk_fonts(
                Some(previous),
                current.and_then(|attrs| attrs.get("complexScript")),
            );
            font_pick_segments(text, from, to, flags).any(|(_, _, east_asian)| {
                *previous != merge_font_family(Some(base), value, east_asian)
            })
        } else {
            previous != value
        }
    })
}

fn merge_chunk_language(patch: &mut Attrs, previous: Option<&Any>, write_language: bool) {
    if let Some(language @ Any::Map(_)) = patch.get_mut("language") {
        if write_language {
            *language = merge_language(previous, language);
        } else {
            patch.remove("language");
        }
    }
}

fn chunk_format_pieces<'a>(
    patch: Attrs,
    fonts: Option<(&'a Any, (bool, bool))>,
    text: &'a str,
    chunk_start: u32,
    (start, end): (u32, u32),
) -> impl Iterator<Item = (u32, u32, Attrs)> + 'a {
    let (previous, flags) = fonts.unwrap_or((&Any::Null, (false, false)));
    let text = if fonts.is_some() { text } else { "" };
    font_pick_segments(text, start - chunk_start, end - chunk_start, flags).map(
        move |(from, to, east_asian)| {
            let mut attrs = patch.clone();
            if fonts.is_some() && let Some(fonts) = attrs.get_mut("fontFamily") {
                *fonts = merge_font_family(Some(previous), fonts, east_asian);
            }
            (chunk_start + from, chunk_start + to, attrs)
        },
    )
}

fn push_format_retain(
    deltas: &mut Vec<Delta<In>>,
    cursor: &mut u32,
    start: u32,
    end: u32,
    attrs: Attrs,
) {
    if attrs.is_empty() {
        return;
    }
    if start > *cursor {
        deltas.push(Delta::Retain(start - *cursor, None));
    }
    if let Some(Delta::Retain(len, Some(previous))) = deltas.last_mut()
        && previous.as_ref() == &attrs
    {
        *len += end - start;
    } else {
        deltas.push(Delta::Retain(end - start, Some(Box::new(attrs))));
    }
    *cursor = end;
}

fn apply_format_retains(story: &TextRef, txn: &mut TransactionMut<'_>, deltas: Vec<Delta<In>>) {
    match deltas.as_slice() {
        [] => {}
        [Delta::Retain(len, Some(attrs))] => {
            story.format(txn, 0, *len, attrs.as_ref().clone());
        }
        [Delta::Retain(at, None), Delta::Retain(len, Some(attrs))] => {
            story.format(txn, *at, *len, attrs.as_ref().clone());
        }
        _ => {
            story.apply_delta(txn, deltas);
        }
    }
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

/// A paragraph of a story: the range it spans, its own paragraph mark included, its id and its
/// paragraph style.
struct StoryParagraph {
    start: u32,
    end: u32,
    key: Option<String>,
    style: Option<String>,
}

/// The paragraphs of a story, in order.
fn story_paragraphs<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<StoryParagraph> {
    let mut paragraphs = Vec::new();
    let mut start = 0;
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let end = offset + out_len(&diff.insert);
        if let Out::YMap(embed) = &diff.insert
            && crate::is_pilcrow(embed, txn)
        {
            paragraphs.push(StoryParagraph {
                start,
                end,
                key: map_string(embed, txn, crate::PARA_ID),
                style: map_string(embed, txn, "pStyle"),
            });
            start = end;
        }
        offset = end;
    }
    if offset > start {
        paragraphs.push(StoryParagraph {
            start,
            end: offset,
            key: None,
            style: None,
        });
    }
    paragraphs
}

/// What a formatting pass writes over a range: the given attributes, or cleared character
/// formatting. A clear with the document's source returns each run to what its paragraph style,
/// and the table style around it, give; without a source it removes the attributes.
#[derive(Clone, Copy)]
enum RunDelta<'a> {
    Set {
        attrs: &'a Attrs,
        font_merge: Option<FontMerge<'a>>,
    },
    Clear(Option<&'a crate::seed::SourceMetadata>),
}

impl<'a> RunDelta<'a> {
    fn font_merge(self) -> Option<FontMerge<'a>> {
        match self {
            Self::Set { font_merge, .. } => font_merge,
            Self::Clear(_) => None,
        }
    }
}

#[derive(Clone, Copy)]
enum FontMerge<'a> {
    Existing,
    Reset(&'a Any),
}

impl<'a> FontMerge<'a> {
    fn chunk_fonts(
        self,
        current: Option<&'a Any>,
        complex_script: Option<&Any>,
    ) -> (&'a Any, (bool, bool)) {
        let base = match self {
            Self::Existing => current.unwrap_or(&Any::Null),
            Self::Reset(reset) => reset,
        };
        (base, run_font_slot_flags(complex_script, current))
    }
}

/// The run attributes the styles give an unformatted run of a paragraph styled `paragraph_style`
/// in `story_id`.
fn styled_run_base(
    source: &crate::seed::SourceMetadata,
    story_id: &str,
    paragraph_style: Option<&str>,
    run_style: Option<&Any>,
) -> OpResult<BTreeMap<String, Any>> {
    Ok(source
        .seeded_run_attrs(story_id, paragraph_style, run_style)
        .map_err(OpError::InvalidFormatValue)?
        .into_iter()
        .collect())
}

/// Every character-formatting attribute set back to `base`, removed where `base` has none. Only
/// [`FORMATTING_ATTRS`] are written, so script direction and every other attribute a clear keeps
/// stay as they are even where the styles give them a value.
fn reset_to(base: &BTreeMap<String, Any>) -> Attrs {
    FORMATTING_ATTRS
        .iter()
        .map(|key| {
            (
                Arc::from(*key),
                base.get(*key).cloned().unwrap_or(Any::Null),
            )
        })
        .collect()
}

/// The attributes `delta` writes over a paragraph styled `paragraph_style` in `story_id`;
/// `remove` gives the attributes a clear without the document's source removes.
fn paragraph_delta(
    delta: RunDelta<'_>,
    story_id: &str,
    paragraph_style: Option<&str>,
    remove: impl FnOnce() -> Attrs,
) -> OpResult<Attrs> {
    match delta {
        RunDelta::Set { attrs, .. } => Ok(attrs.clone()),
        RunDelta::Clear(Some(source)) => Ok(reset_to(&styled_run_base(
            source,
            story_id,
            paragraph_style,
            None,
        )?)),
        RunDelta::Clear(None) => Ok(remove()),
    }
}

/// Whether setting `attrs` would change any text in `start..end`.
fn changes_text<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    start: u32,
    end: u32,
    attrs: &Attrs,
    font_merge: Option<FontMerge<'_>>,
) -> bool {
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let chunk_end = offset + out_len(&diff.insert);
        if chunk_end.min(end) > offset.max(start)
            && let Out::Any(Any::String(text)) = diff.insert
            && attrs_change_chunk(
                attrs,
                diff.attributes.as_deref(),
                false,
                font_merge,
                &text,
                (start.saturating_sub(offset), end.min(chunk_end) - offset),
            )
        {
            return true;
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
        self.refuse_cell_write(&range.story)?;
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
            RunDelta::Set {
                attrs: &attrs,
                font_merge: None,
            },
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
        self.refuse_cell_write(&range.story)?;
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
            RunDelta::Set {
                attrs: &attrs,
                font_merge: (matches!(delta.font_family, Patch::Set(_))
                    && !delta.other.contains_key("fontFamily"))
                .then_some(FontMerge::Existing),
            },
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
    /// as seeding gives a run without direct formatting under `source`'s styles: the paragraph
    /// style `paragraph_style` and, unless `clear` is [`ClearMarks::All`], the run's own character
    /// style. Keys no style supplies are removed.
    pub(crate) fn style_reset_pieces<T: ReadTxn>(
        &self,
        txn: &T,
        story: &TextRef,
        story_id: &str,
        (start, end): (u32, u32),
        clear: &ClearMarks,
        paragraph_style: Option<&str>,
        source: Option<&crate::seed::SourceMetadata>,
    ) -> OpResult<Vec<(u32, u32, Attrs)>> {
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
            let base = match source {
                Some(source) => {
                    styled_run_base(source, story_id, paragraph_style, run_style.as_ref())?
                }
                None => BTreeMap::new(),
            };
            let attrs: Attrs = match clear {
                ClearMarks::Keys(keys) => keys
                    .iter()
                    .map(|key| {
                        (
                            Arc::from(*key),
                            base.get(*key).cloned().unwrap_or(Any::Null),
                        )
                    })
                    .collect(),
                ClearMarks::All => reset_to(&base),
            };
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
        source: Option<&crate::seed::SourceMetadata>,
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
        let merge_font = set.contains_key("fontFamily");
        let pieces = match clear {
            Some(clear) => self.style_reset_pieces(
                &txn,
                &story,
                &range.story,
                (range.start, range.end),
                clear,
                paragraph_style,
                source,
            )?,
            None => vec![(range.start, range.end, Attrs::new())],
        };
        let mut keeper = Keeper::new(self);
        let mut revision_ids: Vec<String> = Vec::new();
        for (start, end, mut attrs) in pieces {
            let reset_fonts = merge_font.then(|| attrs.remove("fontFamily")).flatten();
            let font_merge = merge_font.then(|| match &reset_fonts {
                Some(reset) => FontMerge::Reset(reset),
                None => FontMerge::Existing,
            });
            attrs.extend(set.iter().map(|(key, value)| (key.clone(), value.clone())));
            let piece = StoryRange::new(range.story.as_str(), start, end);
            for id in self.apply_run_format_through(
                &mut txn,
                ctx,
                &story,
                &piece,
                RunDelta::Set {
                    attrs: &attrs,
                    font_merge,
                },
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
        self.refuse_cell_write(&range.story)?;
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
            let chunks = crate::ops::snapshot_range(&story, &txn, range.start, range.end);
            if holds_suggestion_container(&chunks, &txn) {
                return Err(OpError::SuggestedFormattingOverContainer);
            }
            let mut runs = Vec::new();
            let mut changed = false;
            for chunk in chunks {
                let start = chunk.start.max(range.start);
                let end = chunk.end().min(range.end);
                if start >= end {
                    continue;
                }
                if [INS, DEL, RPR_CHANGE]
                    .iter()
                    .any(|key| chunk.attr_active(key))
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

    /// Clears character formatting while retaining tracked-change stamps and hyperlinks: each run
    /// returns to what its paragraph style, and the table style around it, give, and loses its
    /// character style. Without the opened document's styles the formatting is removed.
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
        self.refuse_cell_write(&range.story)?;
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
        let source = self.source_metadata();
        let mut keeper = Keeper::new(self);
        let revision_ids = self.apply_run_format_through(
            &mut txn,
            ctx,
            &story,
            &range,
            RunDelta::Clear(source.as_deref()),
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

    /// Formats `range`, leaving block embeds unmarked and formatting the
    /// text they hold, so a selection across a table formats its cells and not the table. A
    /// paragraph inside a block, or one of several the range spans, that saving rewritten would
    /// lose content of stays unchanged and is counted in `keeper`, as does a single paragraph
    /// whose fields saving rewritten would flatten.
    fn apply_run_format_through(
        &self,
        txn: &mut TransactionMut<'_>,
        ctx: &EditCtx,
        story: &TextRef,
        range: &StoryRange,
        delta: RunDelta<'_>,
        keeper: &mut Keeper<'_>,
    ) -> OpResult<Vec<String>> {
        self.prepare_table_styles(delta, &range.story);
        let (spans, children) = block_free_spans(story, txn, range.start, range.end, false);
        let paragraphs: Vec<_> = story_paragraphs(story, txn)
            .into_iter()
            .filter(|paragraph| paragraph.end > range.start && paragraph.start < range.end)
            .collect();
        let mut revision_ids = Vec::new();
        let font_merge = delta.font_merge();
        if paragraphs.len() > 1 {
            let chunks = story_chunks(story, txn, font_merge.is_some());
            let mut formatted = Vec::new();
            for paragraph in paragraphs {
                let (from, to) = (
                    paragraph.start.max(range.start),
                    paragraph.end.min(range.end),
                );
                let attrs =
                    paragraph_delta(delta, &range.story, paragraph.style.as_deref(), || {
                        clearing_attrs_in(&chunks, from, to)
                    })?;
                if attrs.is_empty()
                    || (matches!(delta, RunDelta::Clear(_))
                        && !changes_units_in(&chunks, from, to, &attrs, false, None))
                {
                    continue;
                }
                if changes_text_in(&chunks, from, to, &attrs, font_merge)
                    && let Some(reason) = paragraph
                        .key
                        .as_deref()
                        .and_then(|key| keeper.rewrite_loss(key))
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
                (Some(&chunks), font_merge),
            )?);
        } else {
            let paragraph = paragraphs.first();
            let attrs = paragraph_delta(
                delta,
                &range.story,
                paragraph.and_then(|paragraph| paragraph.style.as_deref()),
                || clearing_attrs(story, txn, range.start, range.end),
            )?;
            let unchanged = matches!(delta, RunDelta::Clear(_))
                && !changes_units_in(
                    &story_chunks(story, txn, false),
                    range.start,
                    range.end,
                    &attrs,
                    false,
                    None,
                );
            if let Some(key) = paragraph.and_then(|paragraph| paragraph.key.as_deref())
                && changes_text(story, txn, range.start, range.end, &attrs, font_merge)
                && keeper.rewrite_loss(key) == Some(KeptReason::Field)
            {
                keeper.keep(KeptReason::Field);
            } else if !unchanged {
                let formatted = spans
                    .into_iter()
                    .map(|(at, len)| (at, len, attrs.clone()))
                    .collect();
                revision_ids.extend(self.apply_run_formats(
                    txn,
                    ctx,
                    story,
                    &range.story,
                    formatted,
                    (None, font_merge),
                )?);
            }
        }
        for (child, _) in children {
            revision_ids.extend(self.format_story_paragraphs(txn, ctx, &child, delta, keeper)?);
        }
        revision_ids.dedup();
        Ok(revision_ids)
    }

    /// Formats or clears every paragraph of a cell or content-control story and of the blocks
    /// nested in it, keeping unchanged each one that saving rewritten would lose content of.
    fn format_story_paragraphs(
        &self,
        txn: &mut TransactionMut<'_>,
        ctx: &EditCtx,
        child: &str,
        delta: RunDelta<'_>,
        keeper: &mut Keeper<'_>,
    ) -> OpResult<Vec<String>> {
        let Some(text) = crate::story_text(txn, child) else {
            return Ok(Vec::new());
        };
        self.prepare_table_styles(delta, child);
        let mut revision_ids = Vec::new();
        let font_merge = delta.font_merge();
        let chunks = story_chunks(&text, txn, font_merge.is_some());
        let mut formatted = Vec::new();
        for StoryParagraph {
            start,
            end,
            key,
            style,
        } in story_paragraphs(&text, txn)
        {
            let attrs = paragraph_delta(delta, child, style.as_deref(), || {
                clearing_attrs_in(&chunks, start, end)
            })?;
            if attrs.is_empty() || !changes_text_in(&chunks, start, end, &attrs, font_merge) {
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
        revision_ids.extend(self.apply_run_formats(
            txn,
            ctx,
            &text,
            child,
            formatted,
            (Some(&chunks), font_merge),
        )?);
        let len = text.len(txn);
        for (nested, _) in block_free_spans(&text, txn, 0, len, true).1 {
            revision_ids.extend(self.format_story_paragraphs(txn, ctx, &nested, delta, keeper)?);
        }
        Ok(revision_ids)
    }

    /// Loads the table style context of `story` when it is a table cell, before a clear computes
    /// its runs.
    fn prepare_table_styles(&self, delta: RunDelta<'_>, story: &str) {
        if let RunDelta::Clear(Some(metadata)) = delta
            && let Some(index) = self.source_index()
        {
            metadata.prepare_patch_table_styles(story, &index);
        }
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
        (chunks, font_merge): (Option<&[StoryChunk]>, Option<FontMerge<'_>>),
    ) -> OpResult<Vec<String>> {
        if spans.is_empty() {
            return Ok(Vec::new());
        }
        if ctx.is_suggesting() {
            let mut revision_ids = Vec::new();
            for (at, len, attrs) in spans {
                let span = StoryRange::new(story_id.to_owned(), at, at + len);
                revision_ids
                    .extend(self.suggest_run_format(txn, ctx, story, &span, &attrs, font_merge)?);
            }
            return Ok(revision_ids);
        }
        let mut deltas: Vec<Delta<In>> = Vec::with_capacity(spans.len() * 2);
        let mut cursor = 0;
        let needs_chunks = |attrs: &Attrs| {
            font_merge.is_some() || matches!(attrs.get("language"), Some(Any::Map(_)))
        };
        let owned_chunks;
        let chunks = if chunks.is_none() && spans.iter().any(|(_, _, attrs)| needs_chunks(attrs)) {
            owned_chunks = story_chunks(story, txn, font_merge.is_some());
            Some(owned_chunks.as_slice())
        } else {
            chunks
        };
        for (at, len, attrs) in spans {
            if let Some(chunks) = chunks
                && needs_chunks(&attrs)
            {
                for chunk in chunks_within(chunks, at, at + len) {
                    let start = chunk.start.max(at);
                    let end = chunk.end.min(at + len);
                    let mut patch = attrs.clone();
                    let previous =
                        |key: &str| chunk.attrs.as_deref().and_then(|attrs| attrs.get(key));
                    merge_chunk_language(&mut patch, previous("language"), chunk.text);
                    let fonts = font_merge.map(|merge| {
                        merge.chunk_fonts(previous("fontFamily"), previous("complexScript"))
                    });
                    for (from, to, patch) in chunk_format_pieces(
                        patch,
                        fonts,
                        chunk.font_text.as_deref().unwrap_or_default(),
                        chunk.start,
                        (start, end),
                    ) {
                        push_format_retain(&mut deltas, &mut cursor, from, to, patch);
                    }
                }
            } else {
                push_format_retain(&mut deltas, &mut cursor, at, at + len, attrs);
            }
        }
        apply_format_retains(story, txn, deltas);
        Ok(Vec::new())
    }

    fn suggest_run_format(
        &self,
        txn: &mut TransactionMut<'_>,
        ctx: &EditCtx,
        story: &TextRef,
        range: &StoryRange,
        delta: &Attrs,
        font_merge: Option<FontMerge<'_>>,
    ) -> OpResult<Vec<String>> {
        let chunks = crate::ops::snapshot_range(story, txn, range.start, range.end);
        if holds_suggestion_container(&chunks, txn) {
            return Err(OpError::SuggestedFormattingOverContainer);
        }
        let mut revision_ids = Vec::new();
        let mut new_id = None;
        let mut deltas = Vec::new();
        let mut cursor = 0;
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
            let mut patch = delta.clone();
            merge_chunk_language(&mut patch, chunk.attrs.get("language"), true);
            let text = match &chunk.kind {
                crate::ops::ChunkKind::Text(text) => text.as_str(),
                _ => "",
            };
            let fonts = font_merge.map(|merge| {
                merge.chunk_fonts(chunk.attrs.get("fontFamily"), chunk.attrs.get("complexScript"))
            });
            let pieces = chunk_format_pieces(patch, fonts, text, chunk.start, (start, end));
            for (start, end, mut patch) in pieces {
                let mut current = chunk.attrs.clone();
                for (key, value) in &patch {
                    if *value == Any::Null {
                        current.remove(key.as_ref());
                    } else {
                        current.insert(key.to_string(), value.clone());
                    }
                }
                if current == chunk.attrs {
                    continue;
                }
                let mut changes = match chunk.attrs.get(RPR_CHANGE) {
                    Some(Any::Array(changes)) => changes.to_vec(),
                    _ => Vec::new(),
                };
                if changes.is_empty() {
                    let id = new_id.get_or_insert_with(|| self.next_id(txn));
                    let previous: HashMap<_, _> = chunk
                        .attrs
                        .iter()
                        .filter(|(key, _)| {
                            !matches!(key.as_str(), INS | DEL | HYPERLINK | RPR_CHANGE)
                        })
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
                push_format_retain(&mut deltas, &mut cursor, start, end, patch);
            }
        }
        apply_format_retains(story, txn, deltas);
        Ok(revision_ids)
    }
}

pub(crate) fn previous_run_properties(attrs: &BTreeMap<String, Any>, change: &Any) -> Attrs {
    let mut patch: Attrs = attrs
        .keys()
        .filter(|key| {
            !PROTECTED_ATTRS.contains(&key.as_str()) && !STRUCTURAL_ATTRS.contains(&key.as_str())
        })
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
            if !PROTECTED_ATTRS.contains(&key.as_str()) && !STRUCTURAL_ATTRS.contains(&key.as_str())
            {
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
