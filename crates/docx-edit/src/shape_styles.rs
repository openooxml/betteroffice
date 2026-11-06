use std::sync::Arc;

use serde_json::Value;
use sha2::{Digest, Sha256};
use yrs::{Any, Doc, Map, Out, ReadTxn, Transact};

use crate::{EditingDoc, StyleResolver};

pub(crate) const SHAPE_STYLES: &str = "shapeStyles";
const RESOLVER: &str = "resolver";
const MAX_RESOLVER_BYTES: usize = 4 * 1024 * 1024;

#[derive(Default)]
pub(crate) struct ResolverCache {
    json: Option<Arc<str>>,
    resolver: Option<Arc<StyleResolver>>,
}

impl EditingDoc {
    pub(crate) fn shape_style_resolver(&self) -> Option<Arc<StyleResolver>> {
        let txn = self.yrs_doc().transact();
        let json = txn
            .get_map(SHAPE_STYLES)
            .and_then(|root| root.get(&txn, RESOLVER))
            .and_then(|value| match value {
                Out::Any(Any::String(json)) => Some(json),
                _ => None,
            });
        let mut cache = self.shape_styles.lock().unwrap();
        let unchanged = match (&cache.json, &json) {
            (Some(previous), Some(current)) => {
                Arc::ptr_eq(previous, current) || previous == current
            }
            (None, None) => true,
            _ => false,
        };
        if !unchanged {
            cache.resolver = json
                .as_deref()
                .filter(|json| json.len() <= MAX_RESOLVER_BYTES)
                .and_then(|json| serde_json::from_str(json).ok())
                .map(Arc::new);
        }
        cache.json = json;
        cache.resolver.clone()
    }
}

fn has_text_box_content(value: &Value, legacy: bool) -> bool {
    match value {
        Value::Array(values) => values
            .iter()
            .any(|value| has_text_box_content(value, false)),
        Value::Object(values) => {
            let content = (value.get("type").and_then(Value::as_str) == Some("shape"))
                .then(|| value.pointer("/textBody/content"))
                .flatten()
                .or_else(|| {
                    (legacy || value.get("type").and_then(Value::as_str) == Some("textBox"))
                        .then(|| value.get("content"))
                        .flatten()
                });
            content.is_some_and(|content| {
                content
                    .as_array()
                    .is_some_and(|content| !content.is_empty())
            }) || values
                .iter()
                .any(|(key, child)| has_text_box_content(child, key == "textBox"))
        }
        _ => false,
    }
}

pub(crate) fn for_package(package: &Value, styles: &StyleResolver) -> Option<StyleResolver> {
    has_text_box_content(package, false).then(|| styles.for_shape_styles())
}

pub(crate) fn seed(document: &EditingDoc, styles: &StyleResolver) {
    let json = serde_json::to_string(styles).expect("shape style resolver JSON");
    if json.len() > MAX_RESOLVER_BYTES {
        return;
    }
    let digest = Sha256::digest(json.as_bytes());
    let client = u64::from_le_bytes(digest[..8].try_into().unwrap()) & ((1 << 53) - 1);
    let seed = Doc::with_client_id(client);
    let root = seed.get_or_insert_map(SHAPE_STYLES);
    let mut txn = seed.transact_mut();
    root.insert(&mut txn, RESOLVER, json.as_str());
    let update = txn.encode_update_v1();
    drop(txn);
    document
        .apply_verbatim_v1(&update)
        .expect("shape style seed applies");
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeSet;

    use super::*;
    use serde_json::json;

    fn package() -> Value {
        json!({"type": "shape", "textBody": {"content": [{"type": "paragraph",
            "formatting": {"styleId": "Box", "runProperties": {"styleId": "Mark"}}, "content": [{"type": "run",
                "formatting": {"styleId": "Accent"}}]}]}})
    }

    #[test]
    fn resolver_keeps_all_text_styles_is_deterministic_cached_and_outside_undo() {
        let styles = StyleResolver::new(Some(
            &json!({"docDefaults": {"rPr": {"fontSize": 22}}, "styles": [
                {"styleId": "Normal", "type": "paragraph", "default": true},
                {"styleId": "DefaultCharacter", "type": "character", "default": true},
                {"styleId": "Base", "type": "paragraph"},
                {"styleId": "Box", "type": "paragraph", "basedOn": "Base", "rPr": {"bold": true}},
                {"styleId": "Accent", "type": "character"},
                {"styleId": "Mark", "type": "character", "basedOn": "MarkBase"},
                {"styleId": "MarkBase", "type": "character"},
                {"styleId": "Unused", "type": "paragraph"},
                {"styleId": "Table", "type": "table", "default": true}
            ]}),
        ));
        let doc = EditingDoc::new(75367);
        let undo = doc.undo_manager();
        let stories = doc.stories_changed_since(0);
        seed(&doc, &for_package(&package(), &styles).unwrap());
        doc.mark_seeded();
        assert!(doc.unchanged_since_seeded());
        assert_eq!(doc.stories_changed_since(0), stories);
        assert_eq!(undo.undo_depth(), 0);
        let txn = doc.yrs_doc().transact();
        let map = txn.get_map(SHAPE_STYLES).unwrap();
        assert_eq!(map.len(&txn), 1);
        let json = crate::map_string(&map, &txn, RESOLVER).unwrap();
        let value: Value = serde_json::from_str(&json).unwrap();
        let ids: BTreeSet<_> = value["styles"]
            .as_object()
            .unwrap()
            .keys()
            .map(String::as_str)
            .collect();
        assert_eq!(
            ids,
            BTreeSet::from([
                "Accent",
                "Base",
                "Box",
                "DefaultCharacter",
                "Mark",
                "MarkBase",
                "Normal",
                "Unused"
            ])
        );
        assert_eq!(value["doc_defaults"]["rPr"]["fontSize"], 22);
        assert!(value["table_paragraph_formatting"].is_null());
        drop(txn);
        let first = doc.shape_style_resolver().unwrap();
        assert!(Arc::ptr_eq(&first, &doc.shape_style_resolver().unwrap()));
        let repeat = EditingDoc::new(75368);
        seed(&repeat, &for_package(&package(), &styles).unwrap());
        assert_eq!(
            doc.encode_state_as_update_v1(),
            repeat.encode_state_as_update_v1()
        );
        let mut changed = value;
        changed["doc_defaults"]["rPr"]["fontSize"] = json!(30);
        {
            let mut txn = doc.transact_for(&crate::EditCtx::local("", ""));
            map.insert(&mut txn, RESOLVER, changed.to_string());
        }
        let second = doc.shape_style_resolver().unwrap();
        assert!(!Arc::ptr_eq(&first, &second));
        assert!(Arc::ptr_eq(&second, &doc.shape_style_resolver().unwrap()));
        {
            let mut txn = doc.transact_for(&crate::EditCtx::local("", ""));
            map.remove(&mut txn, RESOLVER);
        }
        assert!(doc.shape_style_resolver().is_none());
        assert!(!doc.unchanged_since_seeded());
        assert_eq!(undo.undo_depth(), 0);
    }

    #[test]
    fn retained_comments_group_children_and_legacy_content_trigger_style_seeding() {
        let styles = StyleResolver::new(Some(&json!({"styles": [
            {"styleId": "Box", "type": "paragraph", "rPr": {"fontSize": 30}},
            {"styleId": "Accent", "type": "character", "rPr": {"italic": true}}
        ]})));
        for package in [
            json!({"document": {"comments": [{"body": [package()]}]}}),
            json!({"headerEntries": [["header", {"content": [{"children": [package()]}]}]]}),
            json!({"footnotes": [{"content": [{"textBox": {"content": package()["textBody"]["content"]}}]}]}),
            json!({"endnotes": [{"content": [{"fieldResult": [package()]}]}]}),
        ] {
            let resolver = for_package(&package, &styles).unwrap();
            let json = serde_json::to_value(resolver).unwrap();
            assert!(json["styles"].get("Box").is_some());
            assert!(json["styles"].get("Accent").is_some());
        }
    }

    #[test]
    fn absent_text_content_does_not_write_any_state() {
        let doc = EditingDoc::new(75369);
        let before = doc.encode_state_as_update_v1();
        let package = json!({"type": "shape", "textBody": {"content": []}});
        let styles = for_package(&package, &StyleResolver::default());
        assert!(styles.is_none());
        if let Some(styles) = styles {
            seed(&doc, &styles);
        }
        assert_eq!(doc.encode_state_as_update_v1(), before);
        assert!(doc.shape_style_resolver().is_none());
    }

    #[test]
    fn equal_entries_reuse_the_resolver_and_cache_the_crdt_string_without_copying() {
        let doc = EditingDoc::new(75378);
        seed(&doc, &StyleResolver::default());
        let first = doc.shape_style_resolver().unwrap();
        let current = || {
            let txn = doc.yrs_doc().transact();
            let Out::Any(Any::String(json)) = txn
                .get_map(SHAPE_STYLES)
                .unwrap()
                .get(&txn, RESOLVER)
                .unwrap()
            else {
                panic!("resolver string")
            };
            json
        };
        let original = current();
        assert!(Arc::ptr_eq(
            doc.shape_styles.lock().unwrap().json.as_ref().unwrap(),
            &original
        ));
        {
            let mut txn = doc.yrs_doc().transact_mut();
            txn.get_map(SHAPE_STYLES)
                .unwrap()
                .insert(&mut txn, RESOLVER, original.to_string());
        }
        let replacement = current();
        assert!(!Arc::ptr_eq(&original, &replacement));
        assert!(Arc::ptr_eq(&first, &doc.shape_style_resolver().unwrap()));
        assert!(Arc::ptr_eq(
            doc.shape_styles.lock().unwrap().json.as_ref().unwrap(),
            &replacement
        ));
        assert!(Arc::ptr_eq(&first, &doc.shape_style_resolver().unwrap()));
    }

    #[test]
    fn tables_over_the_read_budget_are_never_written() {
        let doc = EditingDoc::new(75386);
        let before = doc.encode_state_as_update_v1();
        let big = StyleResolver::new(Some(&serde_json::json!({"styles": [
            {"type": "paragraph", "styleId": "x".repeat(MAX_RESOLVER_BYTES)}
        ]})));
        seed(&doc, &big);
        assert_eq!(doc.encode_state_as_update_v1(), before);
        assert!(doc.shape_style_resolver().is_none());
    }

    #[test]
    fn oversized_and_invalid_entries_fall_back_and_cache_the_rejection() {
        assert_eq!(
            MAX_RESOLVER_BYTES,
            docx_parse::ParseLimits::default().max_attribute_bytes
        );
        let doc = EditingDoc::new(75379);
        let valid = serde_json::to_string(&StyleResolver::default()).unwrap();
        let boundary = format!("{}{}", " ".repeat(MAX_RESOLVER_BYTES - valid.len()), valid);
        for (json, accepted) in [
            (valid.clone(), true),
            (boundary.clone(), true),
            (format!(" {boundary}"), false),
            ("{".to_owned(), false),
            ("{}".to_owned(), false),
            ("null".to_owned(), false),
        ] {
            {
                let mut txn = doc.yrs_doc().transact_mut();
                txn.get_map(SHAPE_STYLES)
                    .unwrap()
                    .insert(&mut txn, RESOLVER, json);
            }
            assert_eq!(doc.shape_style_resolver().is_some(), accepted);
            let cached = doc.shape_styles.lock().unwrap().json.clone().unwrap();
            assert_eq!(doc.shape_style_resolver().is_some(), accepted);
            assert!(Arc::ptr_eq(
                &cached,
                doc.shape_styles.lock().unwrap().json.as_ref().unwrap()
            ));
        }
    }
}
