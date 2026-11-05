use serde_json::{Map as JsonMap, Value};
use yrs::{Map, ReadTxn, Transact};

use crate::{EditCtx, EditingDoc, map_string};

pub(crate) const STYLE_LANGUAGES: &str = "styleLanguages";
const SLOT_SEPARATOR: char = '\0';

impl EditingDoc {
    pub fn style_languages_json(&self) -> String {
        let txn = self.yrs_doc().transact();
        let mut entries = JsonMap::new();
        if let Some(map) = txn.get_map(STYLE_LANGUAGES) {
            for key in map.keys(&txn) {
                if key.contains(SLOT_SEPARATOR) {
                    continue;
                }
                if let Some(value) = map_string(&map, &txn, key)
                    .and_then(|json| serde_json::from_str::<JsonMap<String, Value>>(&json).ok())
                {
                    entries.insert(key.to_owned(), Value::Object(value));
                }
            }
            for key in map.keys(&txn) {
                let Some((style_id, slot)) = key.split_once(SLOT_SEPARATOR) else {
                    continue;
                };
                if !matches!(slot, "lang" | "eastAsia" | "bidi") {
                    continue;
                }
                if let Some(value) = map_string(&map, &txn, key)
                    .and_then(|json| serde_json::from_str::<Value>(&json).ok())
                {
                    entries
                        .entry(style_id.to_owned())
                        .or_insert_with(|| Value::Object(JsonMap::new()))
                        .as_object_mut()
                        .expect("style language slots are an object")
                        .insert(slot.to_owned(), value);
                }
            }
        }
        Value::Object(entries).to_string()
    }

    pub fn set_style_language(
        &self,
        ctx: &EditCtx,
        style_id: Option<&str>,
        slots_json: &str,
    ) -> Result<bool, String> {
        let slots: Value = serde_json::from_str(slots_json).map_err(|error| error.to_string())?;
        crate::LanguagePatch::parse(&slots)?;
        let slots = slots.as_object().expect("validated object");
        if slots.is_empty() {
            return Err("supply at least one language slot".to_owned());
        }
        let style_id = style_id.unwrap_or_default();
        if style_id.contains(SLOT_SEPARATOR) {
            return Err("invalid style id".to_owned());
        }
        let mut txn = self.transact_for(ctx);
        let map = txn
            .get_map(STYLE_LANGUAGES)
            .expect("style languages root is declared");
        let mut changed = false;
        for (slot, value) in slots {
            let key = format!("{style_id}{SLOT_SEPARATOR}{slot}");
            let json = value.to_string();
            if map_string(&map, &txn, &key).as_deref() == Some(json.as_str()) {
                continue;
            }
            map.insert(&mut txn, key, json);
            changed = true;
        }
        Ok(changed)
    }

    pub fn patch_style_languages_xml(&self, source: &str) -> Result<String, String> {
        docx_parse::serializer::patch_style_languages(source, &self.style_languages_json())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sync(left: &EditingDoc, right: &EditingDoc) {
        let left_update = left.encode_state_as_update_v1();
        let right_update = right.encode_state_as_update_v1();
        left.apply_update_v1(&right_update).unwrap();
        right.apply_update_v1(&left_update).unwrap();
    }

    #[test]
    fn concurrent_slots_merge_and_same_slot_converges() {
        let left = EditingDoc::new(96);
        let right = EditingDoc::new(97);
        let ctx = EditCtx::local("", "");
        left.set_style_language(
            &ctx,
            Some("Normal"),
            r#"{"lang":"en-US","eastAsia":"ja-JP"}"#,
        )
        .unwrap();
        sync(&left, &right);
        left.set_style_language(&ctx, Some("Normal"), r#"{"lang":"de-DE"}"#)
            .unwrap();
        right
            .set_style_language(&ctx, Some("Normal"), r#"{"eastAsia":"zh-CN"}"#)
            .unwrap();
        sync(&left, &right);
        assert_eq!(
            left.style_languages_json(),
            r#"{"Normal":{"eastAsia":"zh-CN","lang":"de-DE"}}"#
        );
        assert_eq!(left.style_languages_json(), right.style_languages_json());

        left.set_style_language(&ctx, Some("Normal"), r#"{"lang":null}"#)
            .unwrap();
        right
            .set_style_language(&ctx, Some("Normal"), r#"{"lang":"fr-FR"}"#)
            .unwrap();
        sync(&left, &right);
        assert_eq!(left.style_languages_json(), right.style_languages_json());
        let slots: Value = serde_json::from_str(&left.style_languages_json()).unwrap();
        assert!(slots["Normal"]["lang"].is_null() || slots["Normal"]["lang"] == "fr-FR");
        assert_eq!(slots["Normal"]["eastAsia"], "zh-CN");
    }

    #[test]
    fn slots_merge_are_versioned_and_undoable() {
        let doc = EditingDoc::new(94);
        let mut undo = doc.undo_manager();
        let ctx = EditCtx::local("", "");
        let before = doc.version();
        assert!(
            doc.set_style_language(
                &ctx,
                Some("Normal"),
                r#"{"lang":"de-DE","eastAsia":"ja-JP"}"#
            )
            .unwrap()
        );
        assert_ne!(doc.version(), before);
        undo.add_undo_barrier();
        assert!(
            doc.set_style_language(&ctx, Some("Normal"), r#"{"lang":null}"#)
                .unwrap()
        );
        assert_eq!(
            doc.style_languages_json(),
            r#"{"Normal":{"eastAsia":"ja-JP","lang":null}}"#
        );
        let version = doc.version();
        assert!(
            !doc.set_style_language(&ctx, Some("Normal"), r#"{"lang":null}"#)
                .unwrap()
        );
        assert_eq!(doc.version(), version);
        assert!(undo.undo());
        assert_eq!(
            doc.style_languages_json(),
            r#"{"Normal":{"eastAsia":"ja-JP","lang":"de-DE"}}"#
        );
        assert!(undo.redo());
        let peer = EditingDoc::new(95);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(peer.style_languages_json(), doc.style_languages_json());
        let before = doc.encode_state_as_update_v1();
        for json in ["{", "{}", r#"{"lang":"bad_tag"}"#, r#"{"unknown":"de-DE"}"#] {
            assert!(doc.set_style_language(&ctx, None, json).is_err());
            assert_eq!(doc.encode_state_as_update_v1(), before);
        }
    }
}
