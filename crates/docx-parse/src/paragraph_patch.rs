//! Minimal patches for changed source paragraphs: an edit script in the seeded paragraph's
//! inline-unit coordinates, applied to the paragraph's source XML so that everything the edit
//! did not touch keeps its source bytes.

use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use serde_json::Value;

pub type Attrs = BTreeMap<String, Value>;

pub const PARAGRAPH_PATCH_WRITER_VERSION: u32 = 1;

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
    #[serde(default)]
    pub mark_inherit: Attrs,
    pub result_sha256: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SeedV1 {
    pub units: Vec<UnitV1>,
    pub properties: Attrs,
    pub unit_sha256: String,
    #[serde(default)]
    pub inherit: InheritanceV1,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct InheritanceV1 {
    pub paragraph: Attrs,
    pub runs: Vec<Attrs>,
    pub paragraph_mark: Attrs,
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

pub fn canonical_json_v1(value: &Value) -> String {
    fn sorted(value: &Value) -> Value {
        match value {
            Value::Object(object) => {
                let entries: BTreeMap<_, _> = object
                    .iter()
                    .map(|(key, value)| (key.clone(), sorted(value)))
                    .collect();
                Value::Object(entries.into_iter().collect())
            }
            Value::Array(values) => Value::Array(values.iter().map(sorted).collect()),
            _ => value.clone(),
        }
    }
    serde_json::to_string(&sorted(value)).expect("JSON values are finite")
}

pub fn unit_sha256(units: &[UnitV1]) -> String {
    let units = canonical_units(units, false);
    let values: Vec<Value> = units
        .into_iter()
        .map(|unit| match unit {
            UnitV1::Text { text, .. } => serde_json::json!({"kind": "text", "text": text}),
            UnitV1::Atom {
                atom_kind, payload, ..
            } => {
                serde_json::json!({"kind": "atom", "atomKind": atom_kind, "payload": payload})
            }
        })
        .collect();
    sha256_json(&Value::Array(values))
}

pub fn result_sha256(units: &[UnitV1], properties: &Attrs) -> String {
    sha256_json(&serde_json::json!({
        "units": canonical_units(units, true),
        "properties": properties,
    }))
}

fn canonical_units(units: &[UnitV1], attributes: bool) -> Vec<UnitV1> {
    let mut result = Vec::new();
    for unit in units {
        let mut unit = unit.clone();
        let attrs = match &mut unit {
            UnitV1::Text { attrs, .. } | UnitV1::Atom { attrs, .. } => attrs,
        };
        if !attributes {
            attrs.clear();
        }
        if let UnitV1::Text { text, attrs } = &unit {
            if text.is_empty() {
                continue;
            }
            if let Some(UnitV1::Text {
                text: previous,
                attrs: previous_attrs,
            }) = result.last_mut()
                && previous_attrs == attrs
            {
                previous.push_str(text);
                continue;
            }
        }
        result.push(unit);
    }
    result
}

fn sha256_json(value: &Value) -> String {
    use sha2::{Digest, Sha256};
    format!("{:x}", Sha256::digest(canonical_json_v1(value).as_bytes()))
}

#[cfg(test)]
mod hash_tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn canonical_json_recursively_sorts_objects() {
        assert_eq!(
            canonical_json_v1(&json!({"z": [{"b": 2, "a": 1}], "a": null})),
            r#"{"a":null,"z":[{"a":1,"b":2}]}"#
        );
    }

    #[test]
    fn hashes_ignore_chunk_boundaries_but_keep_atoms_and_properties() {
        let text = |value: &str, attrs: Attrs| UnitV1::Text {
            text: value.to_owned(),
            attrs,
        };
        let plain = Attrs::new();
        let bold = BTreeMap::from([("bold".to_owned(), json!(true))]);
        let split = [text("a", bold.clone()), text("b", bold.clone())];
        let joined = [text("ab", bold)];
        assert_eq!(unit_sha256(&split), unit_sha256(&joined));
        assert_eq!(
            unit_sha256(&joined),
            unit_sha256(&[text("ab", plain.clone())])
        );
        assert_eq!(
            result_sha256(&split, &plain),
            result_sha256(&joined, &plain)
        );
        assert_ne!(
            result_sha256(&joined, &plain),
            result_sha256(&[text("ab", plain.clone())], &plain)
        );
        assert_ne!(
            result_sha256(&joined, &plain),
            result_sha256(
                &joined,
                &BTreeMap::from([("pStyle".to_owned(), json!("Heading1"))])
            )
        );
        let atom = |id: i32| UnitV1::Atom {
            atom_kind: "noteRef".to_owned(),
            payload: BTreeMap::from([("footnoteRefId".to_owned(), json!(id))]),
            attrs: Attrs::new(),
        };
        assert_ne!(unit_sha256(&[atom(1)]), unit_sha256(&[atom(2)]));
        assert_ne!(unit_sha256(&[atom(1)]), unit_sha256(&[text("1", plain)]));
    }
}
