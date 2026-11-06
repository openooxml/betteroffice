use std::collections::BTreeSet;
use std::sync::Arc;

use serde_json::Value;
use sha2::{Digest, Sha256};
use yrs::{Doc, Map, ReadTxn, Transact};

use crate::{EditingDoc, StyleResolver};

pub(crate) const SHAPE_STYLES: &str = "shapeStyles";
const RESOLVER: &str = "resolver";

#[derive(Default)]
pub(crate) struct ResolverCache {
    json: Option<String>,
    resolver: Option<Arc<StyleResolver>>,
}

impl EditingDoc {
    pub(crate) fn shape_style_resolver(&self) -> Option<Arc<StyleResolver>> {
        let txn = self.yrs_doc().transact();
        let json = txn
            .get_map(SHAPE_STYLES)
            .and_then(|root| crate::map_string(&root, &txn, RESOLVER));
        let mut cache = self.shape_styles.lock().unwrap();
        if cache.json != json {
            cache.resolver = json
                .as_deref()
                .and_then(|json| serde_json::from_str(json).ok())
                .map(Arc::new);
            cache.json = json;
        }
        cache.resolver.clone()
    }
}

fn references(value: &Value, styles: &mut BTreeSet<String>) {
    match value {
        Value::Array(values) => {
            for value in values {
                references(value, styles);
            }
        }
        Value::Object(values) => {
            if matches!(
                value.get("type").and_then(Value::as_str),
                Some("paragraph" | "run" | "simpleField" | "complexField")
            ) {
                for path in ["/formatting/styleId", "/formatting/runProperties/styleId"] {
                    if let Some(id) = value.pointer(path).and_then(Value::as_str) {
                        styles.insert(id.to_owned());
                    }
                }
            }
            for (key, child) in values {
                if key != "formatting" {
                    references(child, styles);
                }
            }
        }
        _ => {}
    }
}

fn collect(value: &Value, legacy: bool, styles: &mut BTreeSet<String>) -> bool {
    match value {
        Value::Array(values) => {
            let mut found = false;
            for value in values {
                found |= collect(value, false, styles);
            }
            found
        }
        Value::Object(values) => {
            let content = (value.get("type").and_then(Value::as_str) == Some("shape"))
                .then(|| value.pointer("/textBody/content"))
                .flatten()
                .or_else(|| {
                    (legacy || value.get("type").and_then(Value::as_str) == Some("textBox"))
                        .then(|| value.get("content"))
                        .flatten()
                });
            let mut found = false;
            if let Some(content) = content.filter(|content| {
                content
                    .as_array()
                    .is_some_and(|content| !content.is_empty())
            }) {
                references(content, styles);
                found = true;
            }
            for (key, child) in values {
                found |= collect(child, key == "textBox", styles);
            }
            found
        }
        _ => false,
    }
}

pub(crate) fn for_package(package: &Value, styles: &StyleResolver) -> Option<StyleResolver> {
    let mut referenced = BTreeSet::new();
    collect(package, false, &mut referenced).then(|| styles.for_shape_styles(&referenced))
}

pub(crate) fn seed(document: &EditingDoc, styles: &StyleResolver) {
    let json = serde_json::to_string(styles).expect("shape style resolver JSON");
    let digest = Sha256::digest(json.as_bytes());
    let client = u64::from_le_bytes(digest[..8].try_into().unwrap()) & ((1 << 53) - 1);
    let seed = Doc::with_client_id(client);
    let root = seed.get_or_insert_map(SHAPE_STYLES);
    let mut txn = seed.transact_mut();
    root.insert(&mut txn, RESOLVER, json);
    let update = txn.encode_update_v1();
    drop(txn);
    document
        .apply_verbatim_v1(&update)
        .expect("shape style seed applies");
}

pub(crate) fn only_style_changes(txn: &yrs::TransactionMut<'_>) -> bool {
    !txn.changed_parent_types().is_empty()
        && txn.changed_parent_types().iter().all(|branch| {
            matches!(branch.id(), yrs::BranchID::Root(name) if name.as_ref() == SHAPE_STYLES)
        })
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn package() -> Value {
        json!({"type": "shape", "textBody": {"content": [{"type": "paragraph",
            "formatting": {"styleId": "Box", "runProperties": {"styleId": "Mark"}}, "content": [{"type": "run",
                "formatting": {"styleId": "Accent"}}]}]}})
    }

    #[test]
    fn resolver_is_pruned_deterministic_cached_and_outside_authored_history() {
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
        doc.mark_seeded();
        let authored = doc
            .authored_epoch
            .load(std::sync::atomic::Ordering::Relaxed);
        let stories = doc.stories_changed_since(0);
        seed(&doc, &for_package(&package(), &styles).unwrap());
        assert!(doc.unchanged_since_seeded());
        assert_eq!(
            doc.authored_epoch
                .load(std::sync::atomic::Ordering::Relaxed),
            authored
        );
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
                "Normal"
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
        assert!(doc.unchanged_since_seeded());
        assert_eq!(undo.undo_depth(), 0);
    }

    #[test]
    fn retained_comments_group_children_and_legacy_content_collect_style_references() {
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
}
