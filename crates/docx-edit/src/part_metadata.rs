use crate::{EditCtx, EditingDoc, map_string};
use serde_json::Value;
use yrs::types::text::YChange;
use yrs::{Any, Out};
use yrs::{Map, MapPrelim, ReadTxn, Text, TextPrelim, Transact};
pub(crate) const PART_METADATA: &str = "partMetadata";

impl EditingDoc {
    pub fn part_metadata_json(&self) -> String {
        let txn = self.yrs_doc().transact();
        txn.get_map(PART_METADATA)
            .and_then(|root| map_string(&root, &txn, "metadata"))
            .unwrap_or_else(|| "{}".to_owned())
    }

    pub fn set_part_metadata(
        &self,
        ctx: &EditCtx,
        json: &str,
        remove: &[String],
    ) -> Result<bool, String> {
        let value: Value = serde_json::from_str(json).map_err(|error| error.to_string())?;
        if !value.get("body").is_some_and(Value::is_object)
            || remove.iter().any(|id| {
                !["hf:", "fn:", "en:"]
                    .iter()
                    .any(|prefix| id.starts_with(prefix))
            })
        {
            return Err("invalid part metadata".to_owned());
        }
        let aliases = self.header_footer_aliases();
        let mut required = Vec::new();
        for kind in ["headers", "footers"] {
            for entry in value
                .get(kind)
                .and_then(Value::as_array)
                .into_iter()
                .flatten()
            {
                let id = entry
                    .get(0)
                    .and_then(Value::as_str)
                    .filter(|id| !id.is_empty())
                    .ok_or("invalid part id")?;
                let canonical = aliases
                    .iter()
                    .find_map(|(alias, canonical)| (alias == id).then_some(canonical.as_str()))
                    .unwrap_or(id);
                required.push(format!("hf:{canonical}"));
            }
        }
        for (kind, prefix) in [("footnotes", "fn"), ("endnotes", "en")] {
            for note in value
                .get(kind)
                .and_then(Value::as_array)
                .into_iter()
                .flatten()
            {
                let id = note
                    .get("id")
                    .and_then(Value::as_i64)
                    .ok_or("invalid note id")?;
                required.push(format!("{prefix}:{id}"));
            }
        }
        if remove.iter().any(|id| required.contains(id)) {
            return Err("part is still referenced".to_owned());
        }
        let json = value.to_string();
        let mut txn = self.transact_for(ctx);
        let root = txn
            .get_map(PART_METADATA)
            .expect("part metadata root is declared");
        if map_string(&root, &txn, "metadata").as_deref() == Some(json.as_str())
            && remove.is_empty()
        {
            return Ok(false);
        }
        let stories = txn
            .get_map(crate::STORIES)
            .expect("stories root is declared");
        let all: Vec<_> = stories
            .iter(&txn)
            .filter_map(|(_, value)| match value {
                Out::YText(story) => Some(story),
                _ => None,
            })
            .collect();
        for story in all {
            let mut offset = 0;
            let mut refs = Vec::new();
            for diff in story.diff(&txn, YChange::identity) {
                if let Out::YMap(map) = &diff.insert
                    && map_string(map, &txn, crate::KIND_KEY).as_deref() == Some("noteRef")
                {
                    for (key, prefix) in [("footnoteRefId", "fn"), ("endnoteRefId", "en")] {
                        let id = match map.get(&txn, key) {
                            Some(Out::Any(Any::Number(id))) => Some(id as i64),
                            Some(Out::Any(Any::String(id))) => id.parse::<i64>().ok(),
                            _ => None,
                        };
                        if id.is_some_and(|id| remove.contains(&format!("{prefix}:{id}"))) {
                            refs.push(offset);
                        }
                    }
                }
                offset += crate::out_len(&diff.insert);
            }
            refs.sort_unstable();
            refs.dedup();
            for at in refs.into_iter().rev() {
                story.remove_range(&mut txn, at, 1);
            }
        }
        for id in required {
            if !stories.contains_key(&txn, &id) {
                let story = stories.insert(&mut txn, id, TextPrelim::new(""));
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
        let deleted: Vec<_> = stories
            .keys(&txn)
            .filter(|id| {
                remove
                    .iter()
                    .any(|root| *id == root || id.starts_with(&format!("{root}:")))
            })
            .map(str::to_owned)
            .collect();
        for id in deleted {
            stories.remove(&mut txn, &id);
        }
        root.insert(&mut txn, "metadata", json);
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use crate::{EditCtx, EditingDoc};
    #[test]
    fn part_metadata_and_story_lifecycle_share_history() {
        let doc = EditingDoc::new(111);
        doc.create_story("body", "Body", "Normal", "left").unwrap();
        let mut undo = doc.undo_manager();
        let ctx = EditCtx::local("", "");
        let json = r#"{"body":{},"headers":[["rIdNew",{"type":"header","hdrFtrType":"default","content":[]}]]}"#;
        assert!(doc.set_part_metadata(&ctx, json, &[]).unwrap());
        assert!(doc.paragraphs("hf:rIdNew").is_ok());
        assert_eq!(undo.undo_depth(), 1);
        assert!(undo.undo());
        assert!(doc.paragraphs("hf:rIdNew").is_err());
        assert_eq!(doc.part_metadata_json(), "{}");
        assert!(undo.redo());
        undo.add_undo_barrier();
        assert!(
            doc.set_part_metadata(
                &ctx,
                r#"{"body":{},"headers":[]}"#,
                &["hf:rIdNew".to_owned()]
            )
            .unwrap()
        );
        assert!(doc.paragraphs("hf:rIdNew").is_err());
        assert!(undo.undo());
        assert!(doc.paragraphs("hf:rIdNew").is_ok());
    }
}
