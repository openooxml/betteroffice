use std::collections::HashSet;

use docx_parse::HeaderFooterAliasGroup;
use yrs::{Map, ReadTxn, Transact};

use crate::{EditCtx, EditingDoc, map_string};

const HF_ALIASES: &str = "hfAliases";

pub(crate) fn story_id(relationship_id: &str) -> String {
    format!("hf:{relationship_id}")
}

fn parse_groups(json: &str) -> Result<Vec<HeaderFooterAliasGroup>, String> {
    let groups: Vec<HeaderFooterAliasGroup> =
        serde_json::from_str(json).map_err(|error| error.to_string())?;
    let mut ids = HashSet::new();
    for group in &groups {
        if group.relationship_ids.len() < 2 {
            return Err(
                "a header/footer alias group needs at least two relationship ids".to_owned(),
            );
        }
        for id in &group.relationship_ids {
            if id.is_empty() || !ids.insert(id) {
                return Err(
                    "header/footer relationship ids must be non-empty and unique".to_owned(),
                );
            }
        }
    }
    Ok(groups)
}

impl EditingDoc {
    #[doc(hidden)]
    pub fn set_header_footer_aliases(&self, json: &str) -> Result<(), String> {
        let groups = parse_groups(json)?;
        if groups.is_empty() {
            return Ok(());
        }
        let json = serde_json::to_string(&groups).expect("header/footer alias groups serialize");
        let mut txn = self.transact_for(&EditCtx::system(""));
        let session = txn
            .get_map(crate::identity::SESSION)
            .expect("session root is declared by EditingDoc::new");
        if map_string(&session, &txn, HF_ALIASES).as_deref() != Some(json.as_str()) {
            session.insert(&mut txn, HF_ALIASES, json);
        }
        Ok(())
    }

    /// The content story for a header/footer relationship in the current state.
    #[doc(hidden)]
    pub fn header_footer_story(&self, relationship_id: &str) -> String {
        self.header_footer_aliases()
            .into_iter()
            .find_map(|(alias, canonical)| (alias == relationship_id).then(|| story_id(&canonical)))
            .unwrap_or_else(|| story_id(relationship_id))
    }

    /// Active header/footer aliases paired with their canonical relationship ids.
    #[doc(hidden)]
    pub fn header_footer_aliases(&self) -> Vec<(String, String)> {
        let txn = self.yrs_doc().transact();
        let Some(groups) = txn
            .get_map(crate::identity::SESSION)
            .and_then(|session| map_string(&session, &txn, HF_ALIASES))
            .and_then(|json| parse_groups(&json).ok())
        else {
            return Vec::new();
        };
        let stories = txn.get_map(crate::STORIES);
        let mut aliases = Vec::new();
        for group in groups {
            let canonical = &group.relationship_ids[0];
            let inactive = group.relationship_ids[1..].iter().any(|alias| {
                let root = story_id(alias);
                let prefix = format!("{root}:");
                stories.as_ref().is_some_and(|stories| {
                    stories
                        .keys(&txn)
                        .any(|key| key == root || key.starts_with(&prefix))
                })
            });
            if !inactive {
                aliases.extend(
                    group.relationship_ids[1..]
                        .iter()
                        .map(|alias| (alias.clone(), canonical.clone())),
                );
            }
        }
        aliases
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn group(ids: &[&str]) -> serde_json::Value {
        json!({"isHeader": true, "partPath": "word/header1.xml", "relationshipIds": ids})
    }

    #[test]
    fn malformed_groups_write_nothing() {
        let doc = EditingDoc::new(41);
        let before = doc.encode_state_as_update_v1();
        for malformed in [
            "{",
            "{}",
            "null",
            r#"[{"isHeader":true,"partPath":"word/header1.xml"}]"#,
            r#"[{"isHeader":"true","partPath":"word/header1.xml","relationshipIds":["rId7","rId9"]}]"#,
        ] {
            assert!(doc.set_header_footer_aliases(malformed).is_err());
            assert_eq!(doc.encode_state_as_update_v1(), before);
        }
        for groups in [
            json!([group(&[])]),
            json!([group(&["rId7"])]),
            json!([group(&["rId7", ""])]),
            json!([group(&["", "rId9"])]),
            json!([group(&["rId7", "rId7"])]),
            json!([group(&["rId7", "rId9"]), group(&["rId9", "rId11"])]),
            json!([group(&["rId7", "rId9"]), group(&["rId7", "rId11"])]),
        ] {
            assert!(doc.set_header_footer_aliases(&groups.to_string()).is_err());
            assert_eq!(doc.encode_state_as_update_v1(), before);
        }
    }

    #[test]
    fn empty_groups_are_a_noop_and_metadata_is_outside_undo() {
        let doc = EditingDoc::new(41);
        let before = doc.encode_state_as_update_v1();
        doc.set_header_footer_aliases("[]").unwrap();
        assert_eq!(doc.encode_state_as_update_v1(), before);
        let txn = doc.yrs_doc().transact();
        assert!(
            !txn.get_map(crate::identity::SESSION)
                .unwrap()
                .contains_key(&txn, HF_ALIASES)
        );
        drop(txn);
        let undo = doc.undo_manager();
        let groups = json!([group(&["rId7", "rId9"])]).to_string();
        doc.set_header_footer_aliases(&groups).unwrap();
        assert_eq!(doc.header_footer_story("rId9"), "hf:rId7");
        assert!(!undo.can_undo());
        let txn = doc.yrs_doc().transact();
        let stored = map_string(
            &txn.get_map(crate::identity::SESSION).unwrap(),
            &txn,
            HF_ALIASES,
        )
        .unwrap();
        assert_eq!(
            serde_json::from_str::<serde_json::Value>(&stored).unwrap(),
            serde_json::from_str::<serde_json::Value>(&groups).unwrap()
        );
        drop(txn);
        let seeded = doc.encode_state_as_update_v1();
        doc.set_header_footer_aliases(&groups).unwrap();
        doc.set_header_footer_aliases("[]").unwrap();
        assert_eq!(doc.encode_state_as_update_v1(), seeded);
        assert_eq!(doc.header_footer_story("rId9"), "hf:rId7");
    }
}
