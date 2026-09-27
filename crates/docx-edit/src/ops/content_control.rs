//! Filling a text content control with plain text: the text rules, the canonical content that
//! replaces an inline control's frozen content or a block control's paragraphs, and the property
//! changes that clear its placeholder state.

use std::collections::HashMap;
use std::sync::Arc;

use serde_json::Value;
use yrs::types::Attrs;
use yrs::{Any, Map, Text};

use crate::format::PROTECTED_ATTRS;
use crate::op::{OpError, OpResult};
use crate::ops::embed::embed_map_at;
use crate::ops::paragraph::ParagraphRecord;
use crate::{EditCtx, EditingDoc, KIND_KEY, ParagraphId, insertion_attrs, story_ref};

/// Why fill text is refused.
#[derive(Debug, Eq, PartialEq)]
pub(crate) enum TextRefusal {
    Invalid(String),
    Multiline,
}

fn unsupported_char(ch: char) -> bool {
    matches!(ch,
        '\u{0}'..='\u{8}'
        | '\u{B}'
        | '\u{C}'
        | '\u{E}'..='\u{1F}'
        | '\u{7F}'..='\u{9F}'
        | '\u{2028}'
        | '\u{2029}'
        | '\u{FFFC}'
        | '\u{FFFE}'
        | '\u{FFFF}')
}

/// Normalizes CRLF to LF and refuses text a control cannot hold: a bare CR, characters XML or
/// the editing stream cannot carry, and line breaks a single-line control does not accept.
pub(crate) fn fill_text(text: &str, allow_breaks: bool) -> Result<String, TextRefusal> {
    let text = text.replace("\r\n", "\n");
    if text.contains('\r') {
        return Err(TextRefusal::Invalid(
            "a bare carriage return is not a line break; use LF or CRLF".to_owned(),
        ));
    }
    if let Some(ch) = text.chars().find(|ch| unsupported_char(*ch)) {
        return Err(TextRefusal::Invalid(format!(
            "U+{:04X} cannot be written into a content control",
            ch as u32
        )));
    }
    if !allow_breaks && text.contains('\n') {
        return Err(TextRefusal::Multiline);
    }
    Ok(text)
}

fn item(kind: &str, attrs: &Any, text: Option<&str>) -> Any {
    let mut entries = HashMap::from([
        ("kind".to_owned(), Any::from(kind)),
        ("attrs".to_owned(), attrs.clone()),
    ]);
    match text {
        Some(text) => {
            entries.insert("text".to_owned(), Any::from(text));
        }
        None if kind != "tab" => {
            entries.insert("payload".to_owned(), Any::Map(Arc::new(HashMap::new())));
        }
        None => {}
    }
    Any::Map(Arc::new(entries))
}

/// The frozen content of an inline control holding `text`: text, tab and line-break units that
/// all carry `attrs`.
pub(crate) fn inline_content(text: &str, attrs: &[(String, Any)]) -> Any {
    let attrs = Any::Map(Arc::new(
        attrs
            .iter()
            .filter(|(key, _)| !PROTECTED_ATTRS.contains(&key.as_str()))
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect(),
    ));
    let mut items = Vec::new();
    for (line_index, line) in text.split('\n').enumerate() {
        if line_index > 0 {
            items.push(item("break", &attrs, None));
        }
        for (segment_index, segment) in line.split('\t').enumerate() {
            if segment_index > 0 {
                items.push(item("tab", &attrs, None));
            }
            if !segment.is_empty() {
                items.push(item("text", &attrs, Some(segment)));
            }
        }
    }
    Any::Array(Arc::from(items))
}

/// The formatting attributes of a content item or chunk, tracked-change stamps and nulls aside.
pub(crate) fn formatting(attrs: &HashMap<String, Any>) -> Vec<(String, Any)> {
    let mut formatting: Vec<(String, Any)> = attrs
        .iter()
        .filter(|(key, value)| {
            !matches!(key.as_str(), crate::INS | crate::DEL) && **value != Any::Null
        })
        .map(|(key, value)| (key.clone(), value.clone()))
        .collect();
    formatting.sort_by(|left, right| left.0.cmp(&right.0));
    formatting
}

/// The payload changes that clear a control's placeholder state and any obsolete authored
/// value, keeping every other property as captured. `Any::Null` removes a key.
pub(crate) fn property_patch(payload: &HashMap<String, Any>) -> Result<Vec<(String, Any)>, String> {
    let mut patch = Vec::new();
    if matches!(payload.get("showingPlaceholder"), Some(Any::Bool(true))) {
        patch.push(("showingPlaceholder".to_owned(), Any::Bool(false)));
    }
    let raw = match payload.get("rawPropertiesXml") {
        Some(Any::String(raw)) => {
            docx_parse::clear_showing_placeholder_xml(raw).map_err(|error| {
                format!("the control's captured properties are unreadable: {error}")
            })?
        }
        _ => None,
    };
    if let Some(raw) = &raw {
        patch.push(("rawPropertiesXml".to_owned(), Any::from(raw.as_str())));
    }
    if let Some(Any::String(json)) = payload.get("propertiesJson") {
        let mut properties: Value = serde_json::from_str(json)
            .map_err(|error| format!("the control's properties are unreadable: {error}"))?;
        let mut changed = false;
        if let Some(object) = properties.as_object_mut() {
            if object.remove("showingPlaceholder").is_some() {
                changed = true;
            }
            if let Some(Value::Object(state)) = object.get_mut("controlState")
                && state.get("placeholder") == Some(&Value::Bool(true))
            {
                state.insert("placeholder".to_owned(), Value::Bool(false));
                changed = true;
            }
            if let Some(raw) = &raw {
                object.insert("rawPropertiesXml".to_owned(), Value::String(raw.clone()));
                changed = true;
            }
        }
        if changed {
            patch.push((
                "propertiesJson".to_owned(),
                Any::from(serde_json::to_string(&properties).map_err(|error| error.to_string())?),
            ));
        }
    }
    Ok(patch)
}

/// One surviving paragraph of a block control whose inline content is replaced.
pub(crate) struct ParagraphFill {
    pub node_start: u32,
    pub pilcrow: u32,
    pub text: String,
    pub attrs: Vec<(String, Any)>,
}

/// A planned fill, against the captured state.
pub(crate) enum ControlFill {
    Inline {
        raw: u32,
        content: Option<Any>,
        patch: Vec<(String, Any)>,
        /// The control carries an authored value, which the commit drops once the fill is in,
        /// outside undo history.
        drop_value: bool,
    },
    Block {
        parent: String,
        raw: u32,
        patch: Vec<(String, Any)>,
        drop_value: bool,
        child: String,
        /// Surviving paragraphs whose text changes, in story order.
        paragraphs: Vec<ParagraphFill>,
        /// Trailing paragraphs `[start, end)` that go.
        remove: Option<(u32, u32)>,
        /// Paragraphs added after the last surviving one, at this story index.
        insert: Option<(u32, Vec<ParagraphRecord>)>,
    },
}

impl EditingDoc {
    fn patch_embed(
        &self,
        story_id: &str,
        raw: u32,
        kind: &str,
        entries: &[(String, Any)],
    ) -> OpResult<()> {
        let mut txn = self.transact_for(&EditCtx::local(String::new(), String::new()));
        let story = story_ref(&txn, story_id)?;
        let map = embed_map_at(&story, &txn, raw)?;
        if crate::map_string(&map, &txn, KIND_KEY).as_deref() != Some(kind) {
            return Err(OpError::OutOfBounds {
                index: raw,
                len: story.len(&txn),
            });
        }
        for (key, value) in entries {
            if *value == Any::Null {
                map.remove(&mut txn, key);
            } else {
                map.insert(&mut txn, key.clone(), value.clone());
            }
        }
        Ok(())
    }

    /// Applies a fill planned against this state; returns the ids of the paragraphs it adds.
    pub(crate) fn apply_control_fill(
        &self,
        story_id: &str,
        fill: &ControlFill,
    ) -> OpResult<Vec<ParagraphId>> {
        match fill {
            ControlFill::Inline {
                raw,
                content,
                patch,
                ..
            } => {
                let mut entries = patch.clone();
                if let Some(content) = content {
                    entries.push(("content".to_owned(), content.clone()));
                }
                self.patch_embed(story_id, *raw, "sdt", &entries)?;
                Ok(Vec::new())
            }
            ControlFill::Block {
                parent,
                raw,
                patch,
                child,
                paragraphs,
                remove,
                insert,
                ..
            } => {
                if !patch.is_empty() {
                    self.patch_embed(parent, *raw, "blockSdt", patch)?;
                }
                let mut minted = Vec::new();
                if let Some((start, end)) = remove {
                    self.remove_paragraph_records(child, *start, *end)?;
                }
                if let Some((at, records)) = insert {
                    minted = self.insert_paragraph_records(child, *at, records)?;
                }
                let mut txn = self.transact_for(&EditCtx::local(String::new(), String::new()));
                let story = story_ref(&txn, child)?;
                for paragraph in paragraphs.iter().rev() {
                    let len = paragraph.pilcrow - paragraph.node_start;
                    if len > 0 {
                        story.remove_range(&mut txn, paragraph.node_start, len);
                    }
                    if !paragraph.text.is_empty() {
                        let mut attrs: Attrs = paragraph
                            .attrs
                            .iter()
                            .filter(|(key, _)| !PROTECTED_ATTRS.contains(&key.as_str()))
                            .map(|(key, value)| (Arc::from(key.as_str()), value.clone()))
                            .collect();
                        attrs.extend(insertion_attrs(None, None));
                        story.insert_with_attributes(
                            &mut txn,
                            paragraph.node_start,
                            &paragraph.text,
                            attrs,
                        );
                    }
                }
                Ok(minted)
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fill_text_normalizes_and_refuses() {
        assert_eq!(fill_text("a\r\nb", true).unwrap(), "a\nb");
        assert_eq!(fill_text("a\tb", false).unwrap(), "a\tb");
        assert!(matches!(
            fill_text("a\rb", true),
            Err(TextRefusal::Invalid(_))
        ));
        assert!(matches!(
            fill_text("a\u{1}b", true),
            Err(TextRefusal::Invalid(_))
        ));
        assert!(matches!(
            fill_text("a\u{FFFC}", true),
            Err(TextRefusal::Invalid(_))
        ));
        assert_eq!(fill_text("a\nb", false), Err(TextRefusal::Multiline));
        assert_eq!(fill_text("Ünïcödé 😀", false).unwrap(), "Ünïcödé 😀");
    }

    #[test]
    fn inline_content_splits_tabs_and_breaks() {
        let attrs = vec![("bold".to_owned(), Any::Bool(true))];
        let Any::Array(items) = inline_content("a\tb\n\tc", &attrs) else {
            panic!("an array");
        };
        let kinds: Vec<String> = items
            .iter()
            .map(|item| match item {
                Any::Map(map) => match map.get("kind") {
                    Some(Any::String(kind)) => kind.to_string(),
                    _ => String::new(),
                },
                _ => String::new(),
            })
            .collect();
        assert_eq!(kinds, ["text", "tab", "text", "break", "tab", "text"]);
        let Any::Array(empty) = inline_content("", &attrs) else {
            panic!("an array");
        };
        assert!(empty.is_empty());
    }

    #[test]
    fn property_patch_clears_every_placeholder_projection() {
        let raw = "<w:sdtPr><w:tag w:val=\"t\"/><w:showingPlcHdr/><w:text/></w:sdtPr>";
        let payload = HashMap::from([
            ("showingPlaceholder".to_owned(), Any::Bool(true)),
            ("rawPropertiesXml".to_owned(), Any::from(raw)),
            (
                "propertiesJson".to_owned(),
                Any::from(
                    serde_json::json!({
                        "sdtType": "plainText",
                        "showingPlaceholder": true,
                        "controlState": {"placeholder": true},
                        "rawPropertiesXml": raw,
                    })
                    .to_string(),
                ),
            ),
            ("value".to_owned(), Any::from("legacy")),
        ]);
        let patch: HashMap<String, Any> = property_patch(&payload).unwrap().into_iter().collect();
        assert_eq!(patch.get("showingPlaceholder"), Some(&Any::Bool(false)));
        let cleared = "<w:sdtPr><w:tag w:val=\"t\"/><w:text/></w:sdtPr>";
        assert_eq!(patch.get("rawPropertiesXml"), Some(&Any::from(cleared)));
        let Some(Any::String(json)) = patch.get("propertiesJson") else {
            panic!("properties are patched");
        };
        let json: Value = serde_json::from_str(json).unwrap();
        assert_eq!(json.get("showingPlaceholder"), None);
        assert_eq!(json["controlState"]["placeholder"], Value::Bool(false));
        assert_eq!(json["rawPropertiesXml"], Value::String(cleared.to_owned()));
        assert_eq!(patch.get("value"), None);
        let settled = HashMap::from([("rawPropertiesXml".to_owned(), Any::from(cleared))]);
        assert!(property_patch(&settled).unwrap().is_empty());
    }
}
