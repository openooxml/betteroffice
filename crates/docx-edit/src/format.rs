//! Inline formatting surface: `toggle_format`, `format_range`, `clear_formatting`
//! (op-contract §1 "Inline formatting").
//!
//! Attribute names and value shapes mirror the PM mark vocabulary so the render bridge's
//! `lower_attrs` stays mechanical: `bold`/`italic` are booleans, `underline` is `{style}`,
//! `strike` is `{double}`, `textColor` is `{rgb, themeColor}`, `highlight` is `{color}`,
//! `fontSize` is `{size, sizeCs}` (half-points), `fontFamily` is `{ascii, hAnsi}`, and
//! `superscript`/`subscript` are separate boolean attrs.

use std::collections::{BTreeMap, HashMap};
use std::sync::Arc;

use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{Any, Out, ReadTxn, Text, TextRef};

use crate::op::{OpError, OpResult, Receipt};
use crate::{DEL, EditCtx, EditingDoc, INS, StoryRange, out_len, story_ref};

/// The hyperlink text attribute; retained by [`EditingDoc::clear_formatting`].
pub const HYPERLINK: &str = "hyperlink";

/// Attributes that are never touched by formatting sweeps: tracked-change stamps and hyperlinks
/// (clear_formatting keeps hyperlinks — Word Ctrl+Space semantics, a deliberate improvement over
/// the PM-era strip-everything).
pub(crate) const PROTECTED_ATTRS: [&str; 3] = [INS, DEL, HYPERLINK];

/// Word's closed `w:highlight` palette, keyed by uppercase hex (port of
/// `packages/core/src/utils/highlightColors.ts` — an EXACT lookup, no nearest-color search).
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

/// Maps a hex highlight to Word's named palette; unmapped values pass through raw (parity with
/// the PM path, which stores a custom hex and serializes it as `w:shd` fill).
pub fn highlight_color_name(input: &str) -> String {
    let normalized = input.trim_start_matches('#').to_ascii_uppercase();
    HIGHLIGHT_HEX_TO_NAME
        .iter()
        .find(|(hex, _)| *hex == normalized)
        .map(|(_, name)| (*name).to_owned())
        .unwrap_or_else(|| input.to_owned())
}

/// Tri-state field patch (op-contract "tri-state `Patch<T>` per field").
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

/// The six simple toggles (PM `toggleMark` range semantics).
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

/// Tri-state inline formatting delta (port of `ApplyFormattingOptions.marks` in
/// `packages/core/src/prosemirror/applyFormatting.ts`). `Keep` = untouched; `Clear` = remove;
/// `Set` = write. The `other` bag carries the 14 passive run formats (`superscript`, `subscript`,
/// `allCaps`, `smallCaps`, `characterSpacing`, `emboss`, `imprint`, `textShadow`, `emphasisMark`,
/// `textOutline`, `hidden`, `rtl`, `textEffect`, `modernTextEffects`, plus `runStyle`) as opaque
/// attr values; `None` clears the key.
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

/// Formatting policy for [`EditingDoc::insert_text`] (op-contract §1).
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

/// True when every TEXT unit in the range carries a non-null `key` attribute (PM `toggleMark`
/// looks at inline text only; embeds neither veto nor satisfy the check). Returns false for a
/// range without any text unit so a toggle always applies.
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
    /// Toggles one simple format across a range: if EVERY text unit already carries it, remove
    /// it; otherwise add it (PM `toggleMark` range semantics). Superscript and subscript are
    /// mutually exclusive.
    pub fn toggle_format(
        &self,
        ctx: &EditCtx,
        range: StoryRange,
        format: SimpleFormat,
    ) -> OpResult<Receipt> {
        let len = range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        crate::check_range(&story, &txn, range.start, len)?;
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

    /// Applies a tri-state formatting delta across a range in one transaction (op-contract §1).
    ///
    /// S1 formats directly even in suggesting mode (`rPrChange` payloads arrive in S4 — flagged
    /// deviation 3), so the receipt carries no revision IDs.
    pub fn format_range(
        &self,
        ctx: &EditCtx,
        range: StoryRange,
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
        range: StoryRange,
        hyperlink: Option<Any>,
    ) -> OpResult<Receipt> {
        let len = range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        crate::check_range(&story, &txn, range.start, len)?;
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

    /// Removes every formatting attribute present on the range while KEEPING tracked-change
    /// stamps (`ins`/`del`) and hyperlinks (op-contract §1 — Word Ctrl+Space semantics).
    pub fn clear_formatting(&self, ctx: &EditCtx, range: StoryRange) -> OpResult<Receipt> {
        let len = range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        crate::check_range(&story, &txn, range.start, len)?;
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
