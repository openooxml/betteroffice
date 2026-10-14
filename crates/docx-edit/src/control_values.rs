//! Text content controls hold their text as content, never as an authored `value`. Local
//! authoring may not introduce, replace or move such a value, and retyping a valued control drops
//! it. Collaboration updates integrate whatever values they carry: the save projection ignores a
//! text control's value and keeps its content, so a value can only be stale, never authoritative.

use std::collections::HashMap;

#[cfg(feature = "wasm")]
use yrs::Transact;
use yrs::{Any, Map, MapRef, Out, ReadTxn};

#[cfg(feature = "wasm")]
use crate::{EditError, EditResult, EditingDoc};
use crate::{KIND_KEY, OpError, OpResult};

/// The payload keys that decide whether a control is a valued text control.
const VALUE_KEYS: [&str; 5] = ["value", "sdtType", "propertiesJson", KIND_KEY, "content"];

/// The longest parsed-properties JSON the save projection reads.
const MAX_PROPERTIES_JSON: usize = 1_000_000;

fn present(value: Option<&Any>) -> bool {
    value.is_some_and(|value| !matches!(value, Any::Null | Any::Undefined))
}

fn text(value: Option<&Any>) -> Option<&str> {
    match value? {
        Any::String(value) => Some(value),
        _ => None,
    }
}

/// The type the save projection gives a control embed of `kind`: an inline control's parsed
/// properties win over its flat `sdtType`, a block control reads only the flat one, and an untyped
/// control is rich text.
pub(crate) fn resolved_type<'a>(
    kind: Option<&str>,
    field: impl Fn(&str) -> Option<&'a Any>,
) -> String {
    let parsed = (kind == Some("sdt"))
        .then(|| text(field("propertiesJson")))
        .flatten()
        .filter(|json| json.len() <= MAX_PROPERTIES_JSON)
        .and_then(|json| serde_json::from_str::<serde_json::Value>(json).ok())
        .and_then(|properties| properties.get("sdtType")?.as_str().map(str::to_owned));
    parsed
        .or_else(|| text(field("sdtType")).map(str::to_owned))
        .unwrap_or_else(|| "richText".to_owned())
}

pub(crate) fn is_text_type(value: &str) -> bool {
    matches!(value, "plainText" | "richText")
}

fn kind(fields: &HashMap<String, Any>) -> Option<&str> {
    text(fields.get(KIND_KEY))
}

fn type_of(fields: &HashMap<String, Any>) -> String {
    resolved_type(kind(fields), |key| fields.get(key))
}

fn is_control(fields: &HashMap<String, Any>) -> bool {
    matches!(kind(fields), Some("sdt" | "blockSdt"))
}

/// Whether a control payload is a text control carrying an authored value.
fn valued_text(fields: &HashMap<String, Any>) -> bool {
    is_control(fields) && present(fields.get("value")) && is_text_type(&type_of(fields))
}

/// The valued text controls nested in an inline control's content: position, value and type.
fn nested_values(content: Option<&Any>, path: &mut Vec<usize>, out: &mut Vec<(Vec<usize>, Any)>) {
    let Some(Any::Array(items)) = content else {
        return;
    };
    for (index, item) in items.iter().enumerate() {
        let Any::Map(item) = item else { continue };
        let Some(Any::Map(payload)) = item.get("payload") else {
            continue;
        };
        path.push(index);
        let mut fields: HashMap<String, Any> = payload
            .iter()
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect();
        if let Some(kind) = item.get("kind") {
            fields.insert(KIND_KEY.to_owned(), kind.clone());
        }
        if valued_text(&fields) {
            out.push((
                path.clone(),
                Any::Array(std::sync::Arc::from([
                    fields.get("value").cloned().unwrap_or(Any::Null),
                    Any::from(type_of(&fields)),
                ])),
            ));
        }
        nested_values(payload.get("content"), path, out);
        path.pop();
    }
}

fn nested(content: Option<&Any>) -> Vec<(Vec<usize>, Any)> {
    let mut out = Vec::new();
    nested_values(content, &mut Vec::new(), &mut out);
    out
}

/// Whether `after` keeps every valued text control `before` had as it was and adds none: same
/// value, kind and type, and for nested controls the same place.
fn keeps_values(before: &HashMap<String, Any>, after: &HashMap<String, Any>) -> bool {
    if valued_text(after)
        && !(valued_text(before)
            && before.get("value") == after.get("value")
            && kind(before) == kind(after)
            && type_of(before) == type_of(after))
    {
        return false;
    }
    let previous = nested(before.get("content"));
    nested(after.get("content"))
        .iter()
        .all(|entry| previous.contains(entry))
}

/// The value-relevant fields of a map as it is now.
fn fields_of<T: ReadTxn>(map: &MapRef, txn: &T) -> HashMap<String, Any> {
    VALUE_KEYS
        .into_iter()
        .filter_map(|key| match map.get(txn, key) {
            Some(Out::Any(value)) => Some((key.to_owned(), value)),
            _ => None,
        })
        .collect()
}

/// What a local write of `entries` to a control whose current fields are `before` must also do:
/// `Ok(true)` when retyping a valued control must drop its value in the same transaction.
fn plan_write(before: &HashMap<String, Any>, entries: &[(&str, &Any)]) -> OpResult<bool> {
    let mut after = before.clone();
    for (key, value) in entries {
        if VALUE_KEYS.contains(key) {
            if matches!(value, Any::Null) {
                after.remove(*key);
            } else {
                after.insert((*key).to_owned(), (*value).clone());
            }
        }
    }
    let retyped = present(before.get("value"))
        && !entries.iter().any(|(key, _)| *key == "value")
        && (type_of(before) != type_of(&after) || kind(before) != kind(&after));
    if retyped {
        after.remove("value");
    }
    if !keeps_values(before, &after) {
        return Err(OpError::TextControlValue);
    }
    Ok(retyped)
}

/// Refuses writing `entries` onto the embed `map` when that would introduce, replace or move a
/// text control's authored value; `Ok(true)` when the write retypes a valued control, whose
/// value the caller then drops in the same transaction.
pub(crate) fn guard_embed_write<'a, T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    entries: impl IntoIterator<Item = (&'a str, &'a Any)>,
) -> OpResult<bool> {
    let entries: Vec<(&str, &Any)> = entries.into_iter().collect();
    if !entries.iter().any(|(key, _)| VALUE_KEYS.contains(key)) {
        return Ok(false);
    }
    plan_write(&fields_of(map, txn), &entries)
}

/// Refuses inserting an embed of `kind` whose `payload` carries a text control's authored value.
pub(crate) fn guard_embed_insert<'a>(
    kind: &str,
    payload: impl IntoIterator<Item = (&'a str, &'a Any)>,
) -> OpResult<()> {
    let mut fields: HashMap<String, Any> = payload
        .into_iter()
        .filter(|(key, _)| VALUE_KEYS.contains(key))
        .map(|(key, value)| (key.to_owned(), value.clone()))
        .collect();
    fields.insert(KIND_KEY.to_owned(), Any::from(kind));
    if !keeps_values(&HashMap::new(), &fields) {
        return Err(OpError::TextControlValue);
    }
    Ok(())
}

#[cfg(feature = "wasm")]
impl EditingDoc {
    /// Whether the embed at `raw` in `story` is a plain- or rich-text content control.
    pub(crate) fn text_control_at(&self, story: &str, raw: u32) -> OpResult<bool> {
        let txn = self.yrs_doc().transact();
        let text = crate::story_ref(&txn, story)?;
        let map = crate::ops::embed::embed_map_at(&text, &txn, raw)?;
        let fields = fields_of(&map, &txn);
        Ok(is_control(&fields) && is_text_type(&type_of(&fields)))
    }

    /// The text route of the legacy value setters: fills the text control whose embed sits at
    /// `raw` in `story` as one version-checked batch step and one undo step.
    pub(crate) fn fill_text_control_at(
        &self,
        story: &str,
        raw: u32,
        text: &str,
        history: &crate::UndoSession,
    ) -> EditResult<Result<crate::EditApplication, crate::EditRefusal>> {
        use crate::content_controls::{ContentControlSelector, Inventory, Site};
        let control_id = {
            let txn = self.yrs_doc().transact();
            let inventory = Inventory::build(self, &txn)
                .map_err(|failure| EditError::InvalidUpdate(failure.message))?;
            let position = inventory
                .records
                .iter()
                .position(|record| {
                    record.story == story
                        && matches!(record.site, Site::Inline { raw: at, .. } | Site::Block { raw: at, .. } if at == raw)
                })
                .ok_or_else(|| {
                    EditError::InvalidUpdate(format!(
                        "no content control sits at {story}:{raw}"
                    ))
                })?;
            let canonical = inventory
                .companions
                .iter()
                .find(|(_, copies)| copies.contains(&position))
                .map_or(position, |(control, _)| *control);
            inventory.records[canonical].id().to_owned()
        };
        let request = crate::EditRequest {
            expect_version: self.version(),
            source: crate::EditSource::Host,
            history: crate::EditHistory::Separate,
            steps: vec![crate::EditStep::new(
                crate::EditOperation::SetContentControlText {
                    target: ContentControlSelector::Id { control_id },
                    text: text.to_owned(),
                },
            )],
        };
        self.apply_edits(&request, history)
    }
}
