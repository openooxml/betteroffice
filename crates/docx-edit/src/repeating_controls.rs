use crate::content_controls::{ControlRecord, Inventory, Site};
use crate::{EditCtx, EditingDoc, SegmentContent, StorySegment};
use serde_json::Value;
use std::collections::{BTreeMap, HashMap, VecDeque};
use std::sync::Arc;
use yrs::{Any, Map, MapPrelim, ReadTxn, Text, TextPrelim, Transact};

fn item(record: &ControlRecord) -> bool {
    record.control.metadata.control_type == "repeatingSectionItem"
        || matches!(record.payload.get("rawPropertiesXml"), Some(Any::String(xml)) if xml.contains("repeatingSectionItem"))
}

fn references(value: &Any, out: &mut Vec<String>) {
    match value {
        Any::Map(entries) => {
            for (key, value) in entries.iter() {
                if key == "story"
                    && let Any::String(story) = value
                {
                    out.push(story.to_string());
                } else {
                    references(value, out);
                }
            }
        }
        Any::Array(entries) => {
            for value in entries.iter() {
                references(value, out);
            }
        }
        _ => {}
    }
}

fn rewrite(value: &Any, stories: &HashMap<String, String>) -> Any {
    match value {
        Any::Map(entries) => Any::Map(Arc::new(
            entries
                .iter()
                .map(|(key, value)| {
                    let value = if key == "story"
                        && let Any::String(story) = value
                    {
                        Any::from(
                            stories
                                .get(story.as_ref())
                                .map_or(story.as_ref(), String::as_str),
                        )
                    } else {
                        rewrite(value, stories)
                    };
                    (key.clone(), value)
                })
                .collect(),
        )),
        Any::Array(entries) => Any::Array(Arc::from(
            entries
                .iter()
                .map(|value| rewrite(value, stories))
                .collect::<Vec<_>>(),
        )),
        _ => value.clone(),
    }
}

fn fresh_payload(
    kind: &str,
    payload: &BTreeMap<String, Any>,
    stories: &HashMap<String, String>,
    next_id: &mut i64,
) -> BTreeMap<String, Any> {
    let mut value: BTreeMap<_, _> = payload
        .iter()
        .map(|(key, value)| {
            let value = if key == "story"
                && let Any::String(story) = value
            {
                Any::from(
                    stories
                        .get(story.as_ref())
                        .map_or(story.as_ref(), String::as_str),
                )
            } else {
                rewrite(value, stories)
            };
            (key.clone(), value)
        })
        .collect();
    if matches!(kind, "sdt" | "blockSdt") {
        *next_id += 1;
        value.insert("id".to_owned(), Any::Number(*next_id as f64));
        if let Some(Any::String(json)) = value.get("propertiesJson")
            && let Ok(mut props) = serde_json::from_str::<Value>(json)
            && let Some(props) = props.as_object_mut()
        {
            props.insert("id".to_owned(), Value::from(*next_id));
            value.insert(
                "propertiesJson".to_owned(),
                Any::from(Value::Object(props.clone()).to_string()),
            );
        }
    }
    value
}

impl EditingDoc {
    pub fn repeat_control_at(
        &self,
        ctx: &EditCtx,
        story: &str,
        raw: u32,
        remove: bool,
    ) -> Result<bool, String> {
        let (record, mut control_id) = {
            let txn = self.yrs_doc().transact();
            let inventory = Inventory::build(self, &txn).map_err(|failure| failure.message)?;
            let record = inventory
                .records
                .iter()
                .find(|record| {
                    record.story == story
                        && matches!(record.site, Site::Block { raw: index, .. } if index == raw)
                })
                .ok_or("unknown-control")?;
            if !item(record) {
                return Err("not-repeating-item".to_owned());
            }
            if !record.control.effective_lock.known
                || record.control.effective_lock.content
                || record.control.effective_lock.control
                || record.control.metadata.data_bound
            {
                return Err("locked-target".to_owned());
            }
            if record.stamped {
                return Err("tracked-revision-conflict".to_owned());
            }
            let parent = record
                .control
                .parent_control_id
                .as_ref()
                .ok_or("missing-repeating-parent")?;
            let container = inventory
                .records
                .iter()
                .find(|entry| entry.id() == parent)
                .ok_or("missing-repeating-parent")?;
            if container.control.metadata.data_bound {
                return Err("locked-target".to_owned());
            }
            if container.control.metadata.control_type != "repeatingSection"
                && !matches!(container.payload.get("rawPropertiesXml"), Some(Any::String(xml)) if xml.contains("repeatingSection"))
            {
                return Err("missing-repeating-parent".to_owned());
            }
            if remove
                && inventory
                    .records
                    .iter()
                    .filter(|entry| {
                        item(entry) && entry.control.parent_control_id.as_ref() == Some(parent)
                    })
                    .count()
                    <= 1
            {
                return Err("last-repeating-item".to_owned());
            }
            let max_id = inventory
                .records
                .iter()
                .filter_map(|record| {
                    record
                        .control
                        .metadata
                        .ooxml_id
                        .as_ref()?
                        .parse::<i64>()
                        .ok()
                })
                .max()
                .unwrap_or(0);
            (record.clone(), max_id)
        };
        let Site::Block { child, .. } = &record.site else {
            return Err("unsupported-control".to_owned());
        };
        let mut graph = BTreeMap::<String, Vec<StorySegment>>::new();
        let mut queue = VecDeque::from([child.clone()]);
        let mut total = 0;
        while let Some(id) = queue.pop_front() {
            if graph.contains_key(&id) {
                continue;
            }
            if id == story || graph.len() >= 10_000 {
                return Err("invalid-repeating-tree".to_owned());
            }
            let segments = self
                .story_segments(&id)
                .map_err(|error| error.to_string())?;
            total += segments.len();
            if total > 100_000 {
                return Err("limit-exceeded".to_owned());
            }
            for segment in &segments {
                if segment.attributes.iter().any(|(key, value)| {
                    ["ins", "del"].contains(&key.as_str()) && *value != Any::Null
                }) {
                    return Err("tracked-revision-conflict".to_owned());
                }
                if let SegmentContent::OtherEmbed { payload, .. } = &segment.content {
                    let mut children = Vec::new();
                    for value in payload.values() {
                        references(value, &mut children);
                    }
                    if let Some(Any::String(child)) = payload.get("story") {
                        children.push(child.to_string());
                    }
                    queue.extend(children);
                }
            }
            graph.insert(id, segments);
        }
        if control_id
            .checked_add(total as i64 + 1)
            .is_none_or(|last| last > i32::MAX as i64)
        {
            return Err("control-id-exhausted".to_owned());
        }
        let names: HashMap<String, String> = graph
            .keys()
            .map(|id| (id.clone(), format!("{story}:repeat:{}", self.next_id())))
            .collect();
        let root_payload: BTreeMap<_, _> = record
            .payload
            .iter()
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect();
        let root_payload = fresh_payload("blockSdt", &root_payload, &names, &mut control_id);
        let mut txn = self.transact_for(ctx);
        let parent = crate::story_ref(&txn, story).map_err(|error| error.to_string())?;
        let root = txn
            .get_map(crate::STORIES)
            .expect("stories root is declared");
        if remove {
            parent.remove_range(&mut txn, raw, 1);
            for id in graph.keys() {
                root.remove(&mut txn, id);
            }
            return Ok(true);
        }
        for (id, segments) in graph {
            let copied = root.insert(&mut txn, names[&id].clone(), TextPrelim::new(""));
            let mut offset = 0;
            for segment in segments {
                let attrs = segment
                    .attributes
                    .into_iter()
                    .map(|(key, value)| (Arc::from(key), value))
                    .collect();
                match segment.content {
                    SegmentContent::Text(text) => {
                        copied.insert_with_attributes(&mut txn, offset, &text, attrs);
                        offset += text.encode_utf16().count() as u32;
                    }
                    SegmentContent::Pilcrow(properties) => {
                        let mut values: BTreeMap<_, _> = properties
                            .values
                            .into_iter()
                            .filter(|(key, _)| !crate::is_identity_key(key))
                            .collect();
                        values.insert(crate::KIND_KEY.to_owned(), Any::from(crate::PILCROW_KIND));
                        values.insert(crate::PARA_ID.to_owned(), Any::from(self.next_id()));
                        copied.insert_embed_with_attributes(
                            &mut txn,
                            offset,
                            MapPrelim::from_iter(values),
                            attrs,
                        );
                        offset += 1;
                    }
                    SegmentContent::OtherEmbed { kind, payload } => {
                        let mut values = fresh_payload(&kind, &payload, &names, &mut control_id);
                        values.insert(crate::KIND_KEY.to_owned(), Any::from(kind));
                        copied.insert_embed_with_attributes(
                            &mut txn,
                            offset,
                            MapPrelim::from_iter(values),
                            attrs,
                        );
                        offset += 1;
                    }
                }
            }
        }
        let copied = parent.insert_embed_with_attributes(
            &mut txn,
            raw + 1,
            MapPrelim::default(),
            crate::insertion_attrs(None, None),
        );
        copied.insert(&mut txn, crate::KIND_KEY, "blockSdt");
        for (key, value) in root_payload {
            copied.insert(&mut txn, key, value);
        }
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use crate::{Any, EditCtx, EditingDoc, RawOp};

    #[test]
    fn repeating_item_clones_stories_and_has_separate_history() {
        let doc = EditingDoc::new(121);
        let ctx = EditCtx::local("", "");
        doc.create_story("body", "Body", "Normal", "left").unwrap();
        doc.create_story("body:repeat", "", "Normal", "left")
            .unwrap();
        doc.create_story("body:repeat:item", "Item", "Normal", "left")
            .unwrap();
        doc.apply_raw_ops(
            "body",
            vec![RawOp::InsertEmbed {
                index: 0,
                kind: "blockSdt".into(),
                payload: vec![
                    ("sdtType".into(), Any::from("repeatingSection")),
                    ("story".into(), Any::from("body:repeat")),
                ],
                attrs: Default::default(),
            }],
            &ctx,
        )
        .unwrap();
        doc.apply_raw_ops(
            "body:repeat",
            vec![RawOp::InsertEmbed {
                index: 0,
                kind: "blockSdt".into(),
                payload: vec![
                    ("sdtType".into(), Any::from("repeatingSectionItem")),
                    ("story".into(), Any::from("body:repeat:item")),
                ],
                attrs: Default::default(),
            }],
            &ctx,
        )
        .unwrap();
        let mut undo = doc.undo_manager();
        assert!(
            doc.repeat_control_at(&ctx, "body:repeat", 0, false)
                .unwrap()
        );
        let items: Vec<_> = doc
            .story_segments("body:repeat")
            .unwrap()
            .into_iter()
            .filter_map(|segment| {
                if let crate::SegmentContent::OtherEmbed { payload, .. } = segment.content
                    && let Some(Any::String(story)) = payload.get("story")
                {
                    return Some(story.to_string());
                }
                None
            })
            .collect();
        assert_eq!(items.len(), 2);
        assert_ne!(items[0], items[1]);
        assert_eq!(
            doc.paragraphs(&items[0]).unwrap()[0].text,
            doc.paragraphs(&items[1]).unwrap()[0].text
        );
        assert_ne!(
            doc.paragraphs(&items[0]).unwrap()[0].para_id,
            doc.paragraphs(&items[1]).unwrap()[0].para_id
        );
        assert_eq!(undo.undo_depth(), 1);
        assert!(undo.undo());
        assert!(undo.redo());
        undo.add_undo_barrier();
        assert!(doc.repeat_control_at(&ctx, "body:repeat", 1, true).unwrap());
        assert!(doc.paragraphs(&items[0]).is_ok());
        assert!(doc.paragraphs(&items[1]).is_err());
        assert!(undo.undo());
        assert!(doc.paragraphs(&items[1]).is_ok());
    }
}
