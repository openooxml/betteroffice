use crate::{EditCtx, EditingDoc, map_string};
use serde_json::{Map as JsonMap, Value};
use yrs::types::text::YChange;
use yrs::{Any, Map, Out, ReadTxn, Text, Transact};

pub(crate) const SECTION_MARGINS: &str = "sectionMargins";

impl EditingDoc {
    pub fn section_margins_json(&self) -> String {
        let txn = self.yrs_doc().transact();
        let margins = txn
            .get_map(SECTION_MARGINS)
            .expect("section margins root is declared");
        let mut values = JsonMap::new();
        for key in margins.keys(&txn) {
            if let Some(value) = map_string(&margins, &txn, key)
                .and_then(|json| serde_json::from_str::<Value>(&json).ok())
            {
                values.insert(key.to_owned(), value);
            }
        }
        Value::Object(values).to_string()
    }

    pub(crate) fn final_section_properties(&self) -> Option<Value> {
        let parts: Value = serde_json::from_str(&self.part_metadata_json()).ok()?;
        let properties: Value = serde_json::from_str(&self.document_properties_json()).ok()?;
        let base = parts
            .pointer("/body/finalSectionProperties")
            .filter(|value| value.is_object())
            .cloned()
            .or_else(|| {
                properties
                    .pointer("/metadata/body/finalSectionProperties")
                    .filter(|value| value.is_object())
                    .cloned()
            })
            .or_else(|| {
                self.source_metadata()
                    .and_then(|source| source.read().final_section.clone())
            });
        let overrides: Value = serde_json::from_str(&self.section_margins_json()).ok()?;
        let patch = overrides.get("final").and_then(Value::as_object);
        if patch.is_none() {
            return base;
        }
        let mut result = base
            .and_then(|value| value.as_object().cloned())
            .unwrap_or_default();
        result.extend(patch?.clone());
        Some(Value::Object(result))
    }

    pub fn set_section_margins(
        &self,
        ctx: &EditCtx,
        section: &str,
        json: &str,
    ) -> Result<bool, String> {
        let patch: Value = serde_json::from_str(json).map_err(|error| error.to_string())?;
        let patch = patch
            .as_object()
            .ok_or("section margins must be an object")?;
        if section.is_empty()
            || !patch.iter().all(|(key, value)| {
                [
                    "marginTop",
                    "marginBottom",
                    "marginLeft",
                    "marginRight",
                    "headerDistance",
                    "footerDistance",
                    "gutter",
                ]
                .contains(&key.as_str())
                    && value
                        .as_f64()
                        .is_some_and(|number| number.is_finite() && number >= 0.0)
            })
        {
            return Err("invalid section margins".to_owned());
        }
        let mut txn = self.transact_for(ctx);
        let target = if section == "final" {
            None
        } else {
            let story = crate::story_ref(&txn, "body").map_err(|error| error.to_string())?;
            Some(
                story
                    .diff(&txn, YChange::identity)
                    .into_iter()
                    .find_map(|diff| {
                        if let Out::YMap(map) = diff.insert
                            && crate::is_pilcrow(&map, &txn)
                            && map_string(&map, &txn, crate::PARA_ID).as_deref() == Some(section)
                            && map.get(&txn, "sectPr").is_some()
                        {
                            return Some(map);
                        }
                        None
                    })
                    .ok_or("unknown section")?,
            )
        };
        let margins = txn
            .get_map(SECTION_MARGINS)
            .expect("section margins root is declared");
        let mut current = map_string(&margins, &txn, section)
            .and_then(|json| serde_json::from_str::<Value>(&json).ok())
            .and_then(|value| value.as_object().cloned())
            .unwrap_or_default();
        if patch
            .iter()
            .all(|(key, value)| current.get(key) == Some(value))
        {
            return Ok(false);
        }
        current.extend(patch.clone());
        if let Some(map) = target {
            let mut properties = map
                .get(&txn, "sectPr")
                .and_then(|value| match value {
                    Out::Any(value) => serde_json::to_value(value).ok(),
                    _ => None,
                })
                .and_then(|value| value.as_object().cloned())
                .unwrap_or_default();
            properties.extend(patch.clone());
            let value = Any::from_json(&Value::Object(properties).to_string())
                .map_err(|error| error.to_string())?;
            map.insert(&mut txn, "sectPr", value);
        }
        margins.insert(&mut txn, section, Value::Object(current).to_string());
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use crate::{EditCtx, EditingDoc};

    #[test]
    fn section_margins_merge_replicate_and_undo_without_changing_text() {
        let doc = EditingDoc::new(101);
        doc.create_story("body", "Text", "Normal", "left").unwrap();
        let mut undo = doc.undo_manager();
        let ctx = EditCtx::local("", "");
        assert!(
            doc.set_section_margins(&ctx, "final", r#"{"marginLeft":1440}"#)
                .unwrap()
        );
        undo.add_undo_barrier();
        assert!(
            doc.set_section_margins(&ctx, "final", r#"{"marginRight":720}"#)
                .unwrap()
        );
        assert_eq!(
            doc.section_margins_json(),
            r#"{"final":{"marginLeft":1440,"marginRight":720}}"#
        );
        let version = doc.version();
        assert!(
            !doc.set_section_margins(&ctx, "final", r#"{"marginRight":720}"#)
                .unwrap()
        );
        assert_eq!(doc.version(), version);
        assert!(undo.undo());
        assert_eq!(
            doc.section_margins_json(),
            r#"{"final":{"marginLeft":1440}}"#
        );
        assert_eq!(doc.paragraphs("body").unwrap()[0].text, "Text");
        let peer = EditingDoc::new(102);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(peer.section_margins_json(), doc.section_margins_json());
    }

    #[test]
    fn final_section_reads_combine_metadata_and_margin_overrides() {
        let doc = EditingDoc::new(104);
        let ctx = EditCtx::local("", "");
        doc.set_part_metadata(
            &ctx,
            r#"{"body":{"finalSectionProperties":{"pageWidth":10000,"marginLeft":1440}}}"#,
            &[],
        )
        .unwrap();
        doc.set_section_margins(&ctx, "final", r#"{"marginLeft":2880}"#)
            .unwrap();
        let properties = doc.final_section_properties().unwrap();
        assert_eq!(properties["pageWidth"], 10000);
        assert_eq!(properties["marginLeft"], 2880);
    }

    #[test]
    fn invalid_margins_never_mutate() {
        let doc = EditingDoc::new(103);
        let before = doc.encode_state_as_update_v1();
        for value in [
            r#"{"marginLeft":-1}"#,
            r#"{"marginRight":"720"}"#,
            r#"{"headerReferences":[]}"#,
        ] {
            assert!(
                doc.set_section_margins(&EditCtx::local("", ""), "final", value)
                    .is_err()
            );
            assert_eq!(doc.encode_state_as_update_v1(), before);
        }
    }
}
