use crate::{EditCtx, EditingDoc, map_string};
use serde_json::{Map as JsonMap, Value};
use yrs::{Map, MapPrelim, ReadTxn, Text, Transact};

pub(crate) const DOCUMENT_PROPERTIES: &str = "documentProperties";

impl EditingDoc {
    pub fn document_properties_json(&self) -> String {
        let txn = self.yrs_doc().transact();
        let mut values = JsonMap::new();
        if let Some(properties) = txn.get_map(DOCUMENT_PROPERTIES) {
            for key in properties.keys(&txn) {
                if let Some(value) = map_string(&properties, &txn, key)
                    .and_then(|json| serde_json::from_str::<Value>(&json).ok())
                {
                    values.insert(key.to_owned(), value);
                }
            }
        }
        Value::Object(values).to_string()
    }

    pub fn set_document_property(
        &self,
        ctx: &EditCtx,
        key: &str,
        json: &str,
    ) -> Result<bool, String> {
        if key != "metadata" {
            return Err("unknown document property".to_owned());
        }
        let value: Value = serde_json::from_str(json).map_err(|error| error.to_string())?;
        if !value.is_object() {
            return Err("document metadata must be an object".to_owned());
        }
        let json = value.to_string();
        let aliases = self.header_footer_aliases();
        let mut txn = self.transact_for(ctx);
        let properties = txn
            .get_map(DOCUMENT_PROPERTIES)
            .expect("document properties root is declared");
        if map_string(&properties, &txn, key).as_deref() == Some(json.as_str()) {
            return Ok(false);
        }
        let stories = txn
            .get_map(crate::STORIES)
            .expect("stories root is declared");
        for kind in ["headers", "footers"] {
            for entry in value
                .get(kind)
                .and_then(Value::as_array)
                .into_iter()
                .flatten()
            {
                if let Some(id) = entry.get(0).and_then(Value::as_str) {
                    let canonical = aliases
                        .iter()
                        .find_map(|(alias, canonical)| (alias == id).then_some(canonical.as_str()))
                        .unwrap_or(id);
                    let story_id = format!("hf:{canonical}");
                    if !crate::story_id_taken(&stories, &txn, &story_id) {
                        let story = crate::insert_story(&mut txn, &stories, story_id);
                        let pilcrow = story.insert_embed_with_attributes(
                            &mut txn,
                            0,
                            MapPrelim::default(),
                            crate::insertion_attrs(None, None),
                        );
                        crate::write_pilcrow_properties(
                            &pilcrow,
                            &mut txn,
                            &self.next_id(),
                            "Normal",
                            "left",
                        );
                    }
                }
            }
        }
        properties.insert(&mut txn, key, json);
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use crate::{EditCtx, EditingDoc};

    #[test]
    fn property_writes_are_versioned_idempotent_and_undoable() {
        let doc = EditingDoc::new(91);
        doc.create_story("body", "Text", "Normal", "left").unwrap();
        let mut undo = doc.undo_manager();
        let version = doc.version();
        let ctx = EditCtx::local("", "");
        assert!(
            doc.set_document_property(&ctx, "metadata", r#"{"marginLeft":1440}"#)
                .unwrap()
        );
        assert_ne!(doc.version(), version);
        let after = doc.version();
        assert!(
            !doc.set_document_property(&ctx, "metadata", r#"{"marginLeft":1440}"#)
                .unwrap()
        );
        assert_eq!(doc.version(), after);
        assert_eq!(
            doc.document_properties_json(),
            r#"{"metadata":{"marginLeft":1440}}"#
        );
        assert!(undo.undo());
        assert_eq!(doc.document_properties_json(), "{}");
        assert!(undo.redo());
        assert_eq!(
            doc.document_properties_json(),
            r#"{"metadata":{"marginLeft":1440}}"#
        );
        let peer = EditingDoc::new(92);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(
            peer.document_properties_json(),
            doc.document_properties_json()
        );
    }

    #[test]
    fn malformed_property_json_writes_nothing() {
        let doc = EditingDoc::new(93);
        let before = doc.encode_state_as_update_v1();
        for value in ["{", "NaN", "undefined"] {
            assert!(
                doc.set_document_property(&EditCtx::local("", ""), "metadata", value)
                    .is_err()
            );
            assert_eq!(doc.encode_state_as_update_v1(), before);
        }
    }

    #[test]
    fn property_changes_keep_header_aliases_shared() {
        let doc = EditingDoc::new(95);
        doc.create_story("hf:rId7", "Header", "Normal", "left")
            .unwrap();
        doc.set_header_footer_aliases(
            r#"[{"isHeader":true,"partPath":"word/header.xml","relationshipIds":["rId7","rId9"]}]"#,
        )
        .unwrap();
        doc.set_document_property(
            &EditCtx::local("", ""),
            "metadata",
            r#"{"headers":[["rId7",{}],["rId9",{}]]}"#,
        )
        .unwrap();
        assert_eq!(doc.header_footer_story("rId9"), "hf:rId7");
        assert!(doc.paragraphs("hf:rId9").is_err());
        assert_eq!(doc.paragraphs("hf:rId7").unwrap()[0].text, "Header");
    }

    #[test]
    fn header_creation_and_metadata_share_one_undo_step() {
        let doc = EditingDoc::new(94);
        let mut undo = doc.undo_manager();
        let ctx = EditCtx::local("", "");
        doc.set_document_property(&ctx, "metadata", r#"{"headers":[["rIdWatermark",{}]]}"#)
            .unwrap();
        assert!(doc.paragraphs("hf:rIdWatermark").is_ok());
        assert_eq!(undo.undo_depth(), 1);
        assert!(undo.undo());
        assert!(!doc.paragraphs("hf:rIdWatermark").is_ok());
        assert_eq!(doc.document_properties_json(), "{}");
        assert!(undo.redo());
        assert!(doc.paragraphs("hf:rIdWatermark").is_ok());
    }
}
