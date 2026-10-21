//! Minimal patches for changed source paragraphs: an edit script in the seeded paragraph's
//! inline-unit coordinates, applied to the paragraph's source XML so that everything the edit
//! did not touch keeps its source bytes.

use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use serde_json::Value;

pub type Attrs = BTreeMap<String, Value>;

/// A changed source paragraph's edit script. Offsets are half-open UTF-16 offsets into the
/// seeded paragraph's inline units, excluding its pilcrow; every atom occupies one unit.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ParagraphPatchV1 {
    pub ordinal: u32,
    pub session_key: String,
    pub seed: SeedV1,
    pub text_edits: Vec<TextEditV1>,
    pub run_edits: Vec<RunEditV1>,
    pub para_deltas: Vec<AttrDeltaV1>,
    pub para_inherit: Attrs,
    pub result_sha256: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SeedV1 {
    pub units: Vec<UnitV1>,
    pub properties: Attrs,
    pub unit_sha256: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum UnitV1 {
    Text {
        text: String,
        attrs: Attrs,
    },
    Atom {
        atom_kind: String,
        payload: Attrs,
        attrs: Attrs,
    },
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TextEditV1 {
    pub from: u32,
    pub to: u32,
    pub insert: Vec<InsertedTextV1>,
    pub donor: DonorV1,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct InsertedTextV1 {
    pub text: String,
    pub attrs: Attrs,
    pub inherit: Attrs,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", deny_unknown_fields)]
pub enum DonorV1 {
    At { at: u32 },
    Pilcrow,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct RunEditV1 {
    pub from: u32,
    pub to: u32,
    pub deltas: Vec<AttrDeltaV1>,
    pub target: Attrs,
    pub inherit: Attrs,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct AttrDeltaV1 {
    pub path: Vec<String>,
    pub before: SlotV1,
    pub after: SlotV1,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "state", rename_all = "camelCase", deny_unknown_fields)]
pub enum SlotV1 {
    Missing,
    Present { value: Value },
}

/// A changed source paragraph the producer could not express as a patch.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct LegacyParagraphV1 {
    pub ordinal: u32,
    pub code: PatchErrorCode,
    pub detail: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "SCREAMING_SNAKE_CASE")]
pub enum PatchErrorCode {
    MissingSource,
    SourceDigestMismatch,
    InvalidPlan,
    SeedMismatch,
    AmbiguousAlignment,
    AmbiguousDonor,
    InvalidBoundary,
    UnsupportedKey,
    UnsupportedStructure,
    PropertyConflict,
    PatchConflict,
    ResultMismatch,
}

impl PatchErrorCode {
    /// Codes that describe edits v1 does not patch, as opposed to an inconsistent plan or source.
    pub fn is_unsupported(self) -> bool {
        matches!(
            self,
            Self::UnsupportedKey
                | Self::UnsupportedStructure
                | Self::AmbiguousAlignment
                | Self::AmbiguousDonor
        )
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ParagraphPatchError {
    pub part: String,
    pub ordinal: Option<u32>,
    pub code: PatchErrorCode,
    pub detail: String,
}

impl std::fmt::Display for ParagraphPatchError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let code = serde_json::to_value(self.code).ok();
        let code = code.as_ref().and_then(Value::as_str).unwrap_or("UNKNOWN");
        match self.ordinal {
            Some(ordinal) => write!(
                f,
                "DOCX_MINIMAL_PATCH [{code}] {} p#{ordinal}: {}",
                self.part, self.detail
            ),
            None => write!(
                f,
                "DOCX_MINIMAL_PATCH [{code}] {} p#?: {}",
                self.part, self.detail
            ),
        }
    }
}

impl std::error::Error for ParagraphPatchError {}
