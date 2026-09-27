//! Read types shared by the DOCX read APIs: content locations, story selections, heading
//! classification and content-control metadata.

use serde::{Deserialize, Serialize};
use yrs::Any;

use crate::TextRange;

/// Where a piece of read content lives.
///
/// Paragraph, range, table and control anchors resolve against the session version (or the
/// snapshot) they were read at; table indices and control ids are not stable across saves. A
/// source-part anchor addresses retained source XML: the part, its SHA-256, and zero-based
/// element-child ordinals from the part's root element. It is provenance, not an edit target. An
/// unlocated anchor marks content of `story` with no location of its own, and says why.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum Anchor {
    Paragraph {
        story: String,
        para_id: String,
    },
    Range(TextRange),
    Table {
        story: String,
        table_index: u32,
    },
    Control {
        story: String,
        control_id: String,
    },
    SourcePart {
        part: String,
        part_sha256: String,
        path: Vec<u32>,
    },
    Unlocated {
        story: String,
        reason: UnlocatedReason,
    },
}

/// Why content has no location of its own.
#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum UnlocatedReason {
    /// Another paragraph of the story carries the same paragraph id.
    DuplicateParagraphId,
    /// The story is too large to read within the export's byte limit.
    StoryTooLarge,
    /// The story does not exist.
    MissingStory,
    /// Content the editing stream leaves out whose source XML can no longer be found.
    ProvenanceUnavailable,
}

/// A category of stories a read covers.
#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, Ord, PartialEq, PartialOrd, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum StorySelection {
    Body,
    Headers,
    Footers,
    Footnotes,
    Endnotes,
    Comments,
}

/// A heading's zero-based outline level (0..=8) and where it came from.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeadingInfo {
    pub outline_level: u8,
    pub source: OutlineSource,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum OutlineSource {
    /// The paragraph's own outline level.
    Direct,
    /// The paragraph style's outline level, inheritance included.
    Style { style_id: String },
    /// The document's default paragraph properties.
    DocumentDefault,
    /// No outline level anywhere; the style id is `Heading1`..`Heading9`.
    BuiltinStyleId { style_id: String },
}

/// A content control's identity and properties. `control_id` is the engine identity, distinct
/// from the authored `w:id` (`ooxml_id`), tag and alias, none of which need be unique.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ControlMetadata {
    pub control_id: String,
    pub ooxml_id: Option<String>,
    pub control_type: String,
    pub tag: Option<String>,
    pub alias: Option<String>,
    pub lock: Option<String>,
    pub showing_placeholder: bool,
    pub data_bound: bool,
}

/// The control id of a block content control: its child story.
pub(crate) fn block_control_id(child_story: &str) -> String {
    child_story.to_owned()
}

/// The control id of the `ordinal`-th inline control embed of a paragraph.
pub(crate) fn inline_control_id(story: &str, para_id: &str, ordinal: usize) -> String {
    format!("{story}|{para_id}|{ordinal}")
}

/// The control id of the `ordinal`-th control nested in another inline control's content.
pub(crate) fn nested_control_id(parent: &str, ordinal: usize) -> String {
    format!("{parent}|{ordinal}")
}

fn present(value: Option<&Any>) -> Option<&Any> {
    value.filter(|value| !matches!(value, Any::Null | Any::Undefined))
}

fn text(value: Option<&Any>) -> Option<String> {
    match present(value)? {
        Any::String(value) => Some(value.to_string()),
        Any::Number(value) if value.is_finite() && value.fract() == 0.0 => {
            Some(format!("{value:.0}"))
        }
        Any::Number(value) => Some(value.to_string()),
        Any::BigInt(value) => Some(value.to_string()),
        _ => None,
    }
}

/// Reads a content-control embed payload (`sdt` or `blockSdt`) through `field`.
pub(crate) fn control_metadata<'a>(
    control_id: String,
    field: impl Fn(&str) -> Option<&'a Any>,
) -> ControlMetadata {
    ControlMetadata {
        control_id,
        ooxml_id: text(field("id")),
        control_type: text(field("sdtType")).unwrap_or_else(|| "richText".to_owned()),
        tag: text(field("tag")),
        alias: text(field("alias")),
        lock: text(field("lock")),
        showing_placeholder: matches!(field("showingPlaceholder"), Some(Any::Bool(true))),
        data_bound: present(field("dataBinding")).is_some(),
    }
}
