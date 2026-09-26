//! The paragraph ID plan a save carries: IDs for source paragraphs outside
//! the edited stories, applied by occurrence to serialized parts, and parts
//! written as their source bytes with only paragraph IDs patched in place.

use std::collections::{BTreeMap, BTreeSet, HashMap};

use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::paragraph_identity::{
    paragraph_id_attributes, paragraph_occurrences, parse_paragraph_id,
    patch_paragraph_id_references, patch_paragraph_ids,
};
use crate::xml::ParseError;

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct S13ParagraphIds {
    /// IDs of source paragraphs by part and occurrence, applied to the parts
    /// this save serializes.
    #[serde(default)]
    pub assignments: Vec<S13ParagraphId>,
    /// Parts written as their source bytes with only paragraph IDs patched.
    #[serde(default)]
    pub patched_parts: Vec<S13PatchedPart>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct S13ParagraphId {
    pub part: String,
    pub ordinal: u32,
    pub para_id: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct S13PatchedPart {
    pub part: String,
    /// `[ordinal, paraId]` pairs.
    #[serde(default)]
    pub para_ids: Vec<(u32, String)>,
}

pub(crate) const COMMENTS_PART: &str = "word/comments.xml";
const COMMENT_COMPANIONS: [&str; 2] = ["word/commentsExtended.xml", "word/commentsIds.xml"];

impl S13ParagraphIds {
    pub(crate) fn validate(&self) -> Result<(), ParseError> {
        let values = self.assignments.iter().map(|id| id.para_id.as_str()).chain(
            self.patched_parts
                .iter()
                .flat_map(|part| part.para_ids.iter().map(|(_, id)| id.as_str())),
        );
        for value in values {
            if parse_paragraph_id(value).is_none() {
                return Err(ParseError::Canonical(format!(
                    "S13 package save: {value:?} is not a paragraph ID"
                )));
            }
        }
        Ok(())
    }

    /// Assignments grouped by part.
    pub(crate) fn assignments_by_part(&self) -> HashMap<&str, BTreeMap<u32, String>> {
        let mut parts: HashMap<&str, BTreeMap<u32, String>> = HashMap::new();
        for id in &self.assignments {
            parts
                .entry(id.part.as_str())
                .or_default()
                .insert(id.ordinal, id.para_id.clone());
        }
        parts
    }
}

/// Sets the assigned ID on every model paragraph whose `sourceOrdinal` it
/// names, and patches raw XML carried from the source part at its occurrences.
pub(crate) fn apply_assignments<T: Serialize + for<'de> Deserialize<'de>>(
    node: &mut T,
    ids: &BTreeMap<u32, String>,
) -> Result<(), ParseError> {
    let mut value =
        serde_json::to_value(&*node).map_err(|error| ParseError::Canonical(error.to_string()))?;
    assign(&mut value, ids);
    *node =
        serde_json::from_value(value).map_err(|error| ParseError::Canonical(error.to_string()))?;
    Ok(())
}

fn assign(value: &mut Value, ids: &BTreeMap<u32, String>) {
    match value {
        Value::Object(object) => {
            let ordinal = object
                .get("sourceOrdinal")
                .and_then(Value::as_u64)
                .and_then(|ordinal| u32::try_from(ordinal).ok());
            if let Some(ordinal) = ordinal {
                if object.get("type").and_then(Value::as_str) == Some("paragraph") {
                    if let Some(id) = ids.get(&ordinal) {
                        object.insert("paraId".to_owned(), Value::String(id.clone()));
                        object.remove("repeatedParaId");
                    }
                } else {
                    for key in ["xml", "verbatimXml"] {
                        if let Some(Value::String(xml)) = object.get_mut(key) {
                            patch_fragment(xml, ordinal, ids);
                        }
                    }
                }
            }
            for child in object.values_mut() {
                assign(child, ids);
            }
        }
        Value::Array(items) => {
            for item in items {
                assign(item, ids);
            }
        }
        _ => {}
    }
}

fn patch_fragment(xml: &mut String, base: u32, ids: &BTreeMap<u32, String>) {
    let Some(count) = paragraph_occurrences(xml).map(|occurrences| occurrences.len() as u32) else {
        return;
    };
    let local: BTreeMap<u32, String> = ids
        .range(base..base.saturating_add(count))
        .map(|(ordinal, id)| (ordinal - base, id.clone()))
        .collect();
    if !local.is_empty()
        && let Some(patched) = patch_paragraph_ids(xml, &local, false)
    {
        *xml = patched;
    }
}

/// Every valid paragraph ID a model subtree carries, raw XML included.
pub(crate) fn model_paragraph_ids<T: Serialize>(node: &T, ids: &mut BTreeSet<u32>) {
    if let Ok(value) = serde_json::to_value(node) {
        collect(&value, ids);
    }
}

fn collect(value: &Value, ids: &mut BTreeSet<u32>) {
    match value {
        Value::Object(object) => {
            for (key, child) in object {
                match (key.as_str(), child) {
                    ("paraId", Value::String(id)) => ids.extend(parse_paragraph_id(id)),
                    ("xml" | "verbatimXml", Value::String(xml)) => {
                        ids.extend(
                            paragraph_occurrences(xml)
                                .unwrap_or_default()
                                .iter()
                                .filter_map(|occurrence| occurrence.para_id.as_deref())
                                .filter_map(parse_paragraph_id),
                        );
                    }
                    _ => collect(child, ids),
                }
            }
        }
        Value::Array(items) => items.iter().for_each(|item| collect(item, ids)),
        _ => {}
    }
}

/// The source part with only the planned paragraph IDs patched in place.
pub(crate) fn patch_part(original: &[u8], ids: &[(u32, String)]) -> Option<Vec<u8>> {
    let xml = std::str::from_utf8(original).ok()?;
    let patches: BTreeMap<u32, String> = ids.iter().cloned().collect();
    patch_paragraph_ids(xml, &patches, true).map(String::into_bytes)
}

/// The comment parts with patched paragraph IDs, the companion references to
/// a changed ID renamed alongside. `None` when a changed ID the companions
/// reference is not unique among the comment paragraphs, so the reference
/// would be ambiguous.
pub(crate) fn patch_comment_parts(
    comments: &[u8],
    companions: &[(&'static str, Option<&[u8]>)],
    ids: &[(u32, String)],
) -> Option<Vec<(&'static str, Vec<u8>)>> {
    let xml = std::str::from_utf8(comments).ok()?;
    let occurrences = paragraph_occurrences(xml)?;
    let mut counts: HashMap<u32, usize> = HashMap::new();
    for id in occurrences
        .iter()
        .filter_map(|occurrence| occurrence.para_id.as_deref())
        .filter_map(parse_paragraph_id)
    {
        *counts.entry(id).or_default() += 1;
    }
    let referenced: BTreeSet<u32> = companions
        .iter()
        .filter_map(|(_, bytes)| *bytes)
        .flat_map(paragraph_id_attributes)
        .collect();
    let mut renames = HashMap::new();
    for (ordinal, id) in ids {
        let old = occurrences
            .get(*ordinal as usize)?
            .para_id
            .as_deref()
            .and_then(parse_paragraph_id);
        if let Some(old) = old
            && parse_paragraph_id(id) != Some(old)
            && referenced.contains(&old)
        {
            if counts.get(&old) != Some(&1) {
                return None;
            }
            renames.insert(old, id.clone());
        }
    }
    let mut parts = vec![(COMMENTS_PART, patch_part(comments, ids)?)];
    for (path, bytes) in companions {
        if let Some(bytes) = bytes {
            let patched =
                patch_paragraph_id_references(std::str::from_utf8(bytes).ok()?, &renames)?;
            if patched.as_bytes() != *bytes {
                parts.push((*path, patched.into_bytes()));
            }
        }
    }
    Some(parts)
}

pub(crate) fn comment_companions() -> [&'static str; 2] {
    COMMENT_COMPANIONS
}
