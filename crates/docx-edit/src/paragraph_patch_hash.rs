use std::collections::BTreeMap;

use docx_parse::paragraph_patch::{Attrs, UnitV1};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

fn canonical_units(units: &[UnitV1]) -> Vec<UnitV1> {
    let mut result = Vec::new();
    for unit in units {
        match unit {
            UnitV1::Text { text, .. } if text.is_empty() => {}
            UnitV1::Text { text, attrs } => {
                if let Some(UnitV1::Text {
                    text: previous,
                    attrs: previous_attrs,
                }) = result.last_mut()
                    && previous_attrs == attrs
                {
                    previous.push_str(text);
                } else {
                    result.push(unit.clone());
                }
            }
            UnitV1::Atom { .. } => result.push(unit.clone()),
        }
    }
    result
}

pub(crate) fn canonical_json_v1(value: &Value) -> String {
    fn ordered(value: &Value) -> Value {
        match value {
            Value::Object(map) => {
                let sorted: BTreeMap<_, _> = map
                    .iter()
                    .map(|(key, value)| (key, ordered(value)))
                    .collect();
                Value::Object(
                    sorted
                        .into_iter()
                        .map(|(key, value)| (key.clone(), value))
                        .collect(),
                )
            }
            Value::Array(values) => Value::Array(values.iter().map(ordered).collect()),
            value => value.clone(),
        }
    }
    serde_json::to_string(&ordered(value)).expect("JSON value")
}

fn digest(value: &Value) -> String {
    format!("{:x}", Sha256::digest(canonical_json_v1(value).as_bytes()))
}

pub(crate) fn unit_sha256(units: &[UnitV1]) -> String {
    let bare: Vec<_> = units
        .iter()
        .map(|unit| match unit {
            UnitV1::Text { text, .. } => UnitV1::Text {
                text: text.clone(),
                attrs: Attrs::new(),
            },
            UnitV1::Atom {
                atom_kind, payload, ..
            } => UnitV1::Atom {
                atom_kind: atom_kind.clone(),
                payload: payload.clone(),
                attrs: Attrs::new(),
            },
        })
        .collect();
    let bare = canonical_units(&bare);
    let values: Vec<_> = bare
        .iter()
        .map(|unit| match unit {
            UnitV1::Text { text, .. } => json!({ "kind": "text", "text": text }),
            UnitV1::Atom {
                atom_kind, payload, ..
            } => json!({ "kind": "atom", "atomKind": atom_kind, "payload": payload }),
        })
        .collect();
    digest(&json!(values))
}

pub(crate) fn result_sha256(units: &[UnitV1], properties: &Attrs) -> String {
    digest(&json!({ "units": canonical_units(units), "properties": properties }))
}
