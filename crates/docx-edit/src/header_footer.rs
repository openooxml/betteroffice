use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use docx_parse::HeaderFooterAliasGroup;
use yrs::{Any, Map, Out, ReadTxn, Transact};

use crate::{EditCtx, EditingDoc, map_string};

const HF_ALIASES: &str = "hfAliases";

pub(crate) fn story_id(relationship_id: &str) -> String {
    format!("hf:{relationship_id}")
}

#[cfg(test)]
thread_local! {
    pub(crate) static ALIAS_PARSES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

fn parse_groups(json: &str) -> Result<Vec<HeaderFooterAliasGroup>, String> {
    #[cfg(test)]
    ALIAS_PARSES.set(ALIAS_PARSES.get() + 1);
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

#[derive(Default)]
pub(crate) struct ProjectionAliases {
    declaration: Option<Arc<str>>,
    groups: Vec<ProjectionAliasGroup>,
    by_root: HashMap<String, usize>,
}

struct ProjectionAliasGroup {
    roots: Arc<[String]>,
    activity: Option<(u64, bool)>,
}

impl ProjectionAliasGroup {
    fn active(&mut self, revisions: &crate::StoryRevisions) -> bool {
        let revision = revisions.projection_membership;
        if let Some((previous, active)) = self.activity
            && previous == revision
        {
            return active;
        }
        let active = self.roots[1..]
            .iter()
            .all(|alias| !revisions.live_projection_roots.contains_key(alias));
        self.activity = Some((revision, active));
        active
    }
}

impl ProjectionAliases {
    fn refresh<T: ReadTxn>(&mut self, txn: &T) {
        let declaration = txn.get_map(crate::identity::SESSION).and_then(|session| {
            match session.get(txn, HF_ALIASES) {
                Some(Out::Any(Any::String(value))) => Some(value),
                _ => None,
            }
        });
        let same = match (&self.declaration, &declaration) {
            (Some(previous), Some(current)) => Arc::ptr_eq(previous, current),
            (None, None) => true,
            _ => false,
        };
        if !same {
            self.by_root.clear();
            self.groups.clear();
            if let Some(groups) = declaration
                .as_deref()
                .and_then(|json| parse_groups(json).ok())
            {
                for group in groups {
                    let mut roots: Vec<_> = group
                        .relationship_ids
                        .iter()
                        .map(|relationship| story_id(relationship))
                        .collect();
                    roots[1..].sort();
                    for root in &roots {
                        self.by_root.insert(root.clone(), self.groups.len());
                    }
                    self.groups.push(ProjectionAliasGroup {
                        roots: roots.into(),
                        activity: None,
                    });
                }
            }
            self.declaration = declaration;
        }
    }

    pub(crate) fn group<T: ReadTxn>(
        &mut self,
        doc: &EditingDoc,
        txn: &T,
        root: &str,
    ) -> Option<Arc<[String]>> {
        root.strip_prefix("hf:")?;
        self.refresh(txn);
        let group = &mut self.groups[*self.by_root.get(root)?];
        let revisions = doc.story_revisions.lock().unwrap();
        group.active(&revisions).then(|| Arc::clone(&group.roots))
    }

    pub(crate) fn active_aliases<T: ReadTxn>(
        &mut self,
        doc: &EditingDoc,
        txn: &T,
    ) -> BTreeMap<String, String> {
        self.refresh(txn);
        let revisions = doc.story_revisions.lock().unwrap();
        let mut aliases = BTreeMap::new();
        for group in &mut self.groups {
            if group.active(&revisions) {
                aliases.extend(
                    group.roots[1..]
                        .iter()
                        .cloned()
                        .map(|alias| (alias, group.roots[0].clone())),
                );
            }
        }
        aliases
    }
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
        self.header_footer_aliases_in(&txn)
    }

    pub(crate) fn header_footer_aliases_in<T: ReadTxn>(&self, txn: &T) -> Vec<(String, String)> {
        let Some(groups) = txn
            .get_map(crate::identity::SESSION)
            .and_then(|session| map_string(&session, txn, HF_ALIASES))
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
                    crate::live_stories(stories, txn)
                        .any(|(key, _)| key == root || key.starts_with(&prefix))
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
