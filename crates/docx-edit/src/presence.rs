#![cfg_attr(not(feature = "wasm"), allow(dead_code))]

use std::sync::{Arc, Mutex};

use yrs::types::{DeepObservable, Delta, Event, PathSegment};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{Any, Assoc, GetString, ID, Out, ReadTxn, StickyIndex, Transact};

use crate::EditingDoc;
use crate::sticky_positions::sticky_index;

#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct TypingInference {
    pub client_id: u64,
    pub story: String,
    pub para_id: String,
    pub end_offset: u32,
}

pub(crate) fn encode_sticky(index: &StickyIndex) -> Vec<u8> {
    index.encode_v1()
}

pub(crate) fn resolve_sticky_selection(
    doc: &EditingDoc,
    story: &str,
    anchor: &[u8],
    head: &[u8],
) -> Result<(u32, u32), String> {
    let collapsed = anchor == head;
    let anchor = StickyIndex::decode_v1(anchor).map_err(|error| error.to_string())?;
    let head = if collapsed {
        anchor.clone()
    } else {
        StickyIndex::decode_v1(head).map_err(|error| error.to_string())?
    };
    doc.resolve_cached_sticky_selection(story, &anchor, &head)
}

pub(crate) fn apply_update_with_typing_inference(
    doc: &EditingDoc,
    bytes: &[u8],
) -> Result<Option<TypingInference>, String> {
    let update = crate::decode_peer_update_v1(bytes).map_err(|error| error.to_string())?;
    let insertions = update.insertions(false);
    let mut clients = insertions.client_ids();
    let client_id = clients.next().filter(|_| clients.next().is_none());
    let changes = Arc::new(Mutex::new(Vec::<InsertedContent>::new()));
    // The inference only consumes story text insertions, and a block that does
    // not extend the local state vector can never integrate — empty,
    // already-known and delete-only updates skip watching entirely.
    let subscription = if client_id
        .is_some_and(|_| update.extends(&doc.yrs_doc().transact().state_vector()))
    {
        let txn = doc.yrs_doc().transact();
        let subscription = txn.get_map(super::STORIES).map(|stories| {
            let changes = Arc::clone(&changes);
            stories.observe_deep(move |txn, events| {
                for event in events.iter() {
                    // A story created by this update arrives as a map
                    // insert carrying a prefilled text branch.
                    if let Event::Map(event) = event {
                        for (story, change) in event.keys(txn).iter() {
                            let yrs::types::EntryChange::Inserted(Out::YText(text)) = change else {
                                continue;
                            };
                            let end_index = text.get_string(txn).encode_utf16().count() as u32;
                            if let Some(id) = sticky_index(text, txn, end_index, Assoc::Before)
                                .and_then(|sticky| sticky.id().copied())
                            {
                                changes.lock().unwrap().push(InsertedContent {
                                    id,
                                    story: story.to_string(),
                                    end_index,
                                    is_text: true,
                                });
                            }
                        }
                        continue;
                    }
                    let Event::Text(event) = event else {
                        continue;
                    };
                    let Some(story) = event.path().front().and_then(|segment| match segment {
                        PathSegment::Key(key) => Some(key.to_string()),
                        PathSegment::Index(_) => None,
                    }) else {
                        continue;
                    };
                    let mut index = 0_u32;
                    for delta in event.delta(txn) {
                        match delta {
                            Delta::Retain(length, _) => index += length,
                            Delta::Deleted(_) => {}
                            Delta::Inserted(value, _) => {
                                let (length, is_text) = match value {
                                    Out::Any(Any::String(text)) => {
                                        (text.encode_utf16().count() as u32, true)
                                    }
                                    _ => (1, false),
                                };
                                let end_index = index + length;
                                if let Some(id) =
                                    sticky_index(event.target(), txn, end_index, Assoc::Before)
                                        .and_then(|sticky| sticky.id().copied())
                                {
                                    changes.lock().unwrap().push(InsertedContent {
                                        id,
                                        story: story.clone(),
                                        end_index,
                                        is_text,
                                    });
                                }
                                index = end_index;
                            }
                        }
                    }
                }
            })
        });
        drop(txn);
        subscription
    } else {
        None
    };

    doc.integrate_update(update, crate::UpdateOrigin::Remote)
        .map_err(|error| error.to_string())?;
    drop(subscription);
    let Some(client_id) = client_id else {
        return Ok(None);
    };
    let candidate = changes
        .lock()
        .unwrap()
        .iter()
        .filter(|change| change.id.client == client_id && insertions.contains(&change.id))
        .max_by_key(|change| change.id.clock)
        .cloned();
    let Some(candidate) = candidate.filter(|candidate| candidate.is_text) else {
        return Ok(None);
    };
    let Ok((para_id, end_offset)) = index_loc(doc, &candidate.story, candidate.end_index) else {
        return Ok(None);
    };
    Ok(Some(TypingInference {
        client_id: client_id.get(),
        story: candidate.story,
        para_id,
        end_offset,
    }))
}

#[derive(Clone)]
struct InsertedContent {
    id: ID,
    story: String,
    end_index: u32,
    is_text: bool,
}

fn index_loc(doc: &EditingDoc, story: &str, index: u32) -> Result<(String, u32), String> {
    doc.paragraph_index(story)
        .map_err(|error| error.to_string())?
        .para_at(index)
        .map(|para| (para.para_id.to_string(), index.saturating_sub(para.start)))
        .ok_or_else(|| format!("selection index {index} does not resolve in story {story:?}"))
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use super::*;
    use yrs::{Any, Assoc, IndexedSequence, Text};

    use crate::{EditCtx, FormatPolicy, Position, StoryRange, story_ref};

    fn seeded(client_id: u64) -> (EditingDoc, String) {
        let doc = EditingDoc::new(client_id);
        let para_id = doc.create_story("body", "hello", "Normal", "left").unwrap();
        (doc, para_id)
    }

    #[test]
    fn sticky_selection_round_trips_through_binary_encoding() {
        let (doc, _) = seeded(1);
        let txn = doc.yrs_doc().transact();
        let story = story_ref(&txn, "body").unwrap();
        let anchor = story.sticky_index(&txn, 1, Assoc::After).unwrap();
        let head = story.sticky_index(&txn, 4, Assoc::After).unwrap();
        drop(txn);

        assert_eq!(
            resolve_sticky_selection(&doc, "body", &encode_sticky(&anchor), &encode_sticky(&head))
                .unwrap(),
            (1, 4)
        );
    }

    #[test]
    fn collapsed_sticky_selection_matches_one_offset_and_checks_the_story() {
        let (doc, _) = seeded(1);
        doc.create_story("hf:rId1", "header", "Normal", "left")
            .unwrap();
        let positions = {
            let txn = doc.yrs_doc().transact();
            let story = story_ref(&txn, "body").unwrap();
            let len = story.len(&txn);
            [0, 1, len / 2, len - 1]
                .map(|index| encode_sticky(&story.sticky_index(&txn, index, Assoc::After).unwrap()))
        };
        doc.insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 1),
            "xy",
            FormatPolicy::Inherit,
        )
        .unwrap();
        doc.delete_range(&EditCtx::local("", ""), StoryRange::new("body", 3, 4))
            .unwrap();
        for position in positions {
            let expected = StickyIndex::decode_v1(&position)
                .unwrap()
                .get_offset(&doc.yrs_doc().transact())
                .unwrap()
                .index;
            assert_eq!(
                resolve_sticky_selection(&doc, "body", &position, &position).unwrap(),
                (expected, expected)
            );
            assert_eq!(
                resolve_sticky_selection(&doc, "hf:rId1", &position, &position).unwrap_err(),
                "selection sticky index belongs to another story"
            );
        }
        assert!(resolve_sticky_selection(&doc, "body", &[], &[]).is_err());
    }

    #[test]
    fn sticky_selection_preserves_endpoint_errors_and_story_validation() {
        let (doc, _) = seeded(1);
        doc.create_story("header:rId7", "header", "Normal", "left")
            .unwrap();
        let anchor = doc
            .cached_sticky_index("body", 1, Assoc::After)
            .unwrap()
            .unwrap();
        let foreign = doc
            .cached_sticky_index("header:rId7", 1, Assoc::After)
            .unwrap()
            .unwrap();
        let missing = StickyIndex::new(
            yrs::IndexScope::Relative(ID::new(yrs::ClientID::new(99), 0)),
            Assoc::After,
        );
        for (anchor, head, expected) in [
            (
                &anchor,
                &foreign,
                "selection sticky index belongs to another story",
            ),
            (
                &foreign,
                &foreign,
                "selection sticky index belongs to another story",
            ),
            (&foreign, &missing, "selection head no longer resolves"),
            (&missing, &foreign, "selection anchor no longer resolves"),
            (&missing, &missing, "selection anchor no longer resolves"),
        ] {
            assert_eq!(
                resolve_sticky_selection(
                    &doc,
                    "body",
                    &encode_sticky(anchor),
                    &encode_sticky(head)
                ),
                Err(expected.to_owned())
            );
        }
    }

    #[test]
    fn infers_the_end_of_the_last_remote_text_insertion() {
        let (baseline, para_id) = seeded(1);
        let state = baseline.encode_state_as_update_v1();
        let writer = EditingDoc::new(7);
        writer.apply_update_v1(&state).unwrap();
        let reader = EditingDoc::new(9);
        reader.apply_update_v1(&state).unwrap();
        let reader_state = reader.encode_state_vector_v1();

        writer
            .insert_text(
                &EditCtx::local("", ""),
                Position::new("body", 2),
                "XY",
                FormatPolicy::Plain,
            )
            .unwrap();
        let update = writer.encode_diff_v1(&reader_state).unwrap();

        assert_eq!(
            apply_update_with_typing_inference(&reader, &update).unwrap(),
            Some(TypingInference {
                client_id: 7,
                story: "body".to_owned(),
                para_id,
                end_offset: 4,
            })
        );
    }

    #[test]
    fn infers_formatted_text_when_format_markers_follow_the_content() {
        let (baseline, para_id) = seeded(1);
        let state = baseline.encode_state_as_update_v1();
        let writer = EditingDoc::new(7);
        writer.apply_update_v1(&state).unwrap();
        let reader = EditingDoc::new(9);
        reader.apply_update_v1(&state).unwrap();
        let reader_state = reader.encode_state_vector_v1();

        writer
            .insert_text(
                &EditCtx::local("", ""),
                Position::new("body", 2),
                "XY",
                FormatPolicy::Explicit(BTreeMap::from([("bold".to_owned(), Any::Bool(true))])),
            )
            .unwrap();
        let update = writer.encode_diff_v1(&reader_state).unwrap();

        assert_eq!(
            apply_update_with_typing_inference(&reader, &update).unwrap(),
            Some(TypingInference {
                client_id: 7,
                story: "body".to_owned(),
                para_id,
                end_offset: 4,
            })
        );
    }

    #[test]
    fn infers_typing_inside_a_story_created_by_the_same_update() {
        let (baseline, _) = seeded(1);
        let state = baseline.encode_state_as_update_v1();
        let writer = EditingDoc::new(7);
        writer.apply_update_v1(&state).unwrap();
        let reader = EditingDoc::new(9);
        reader.apply_update_v1(&state).unwrap();
        let reader_state = reader.encode_state_vector_v1();

        let para_id = writer
            .create_story("notes", "typed", "Normal", "left")
            .unwrap();
        let update = writer.encode_diff_v1(&reader_state).unwrap();

        assert_eq!(
            apply_update_with_typing_inference(&reader, &update).unwrap(),
            Some(TypingInference {
                client_id: 7,
                story: "notes".to_owned(),
                para_id,
                end_offset: 5,
            })
        );
    }

    #[test]
    fn deletion_only_update_has_no_typing_inference() {
        let (baseline, _) = seeded(1);
        let state = baseline.encode_state_as_update_v1();
        let writer = EditingDoc::new(7);
        writer.apply_update_v1(&state).unwrap();
        let reader = EditingDoc::new(9);
        reader.apply_update_v1(&state).unwrap();
        let reader_state = reader.encode_state_vector_v1();

        writer
            .delete_range(&EditCtx::local("", ""), StoryRange::new("body", 1, 2))
            .unwrap();
        let update = writer.encode_diff_v1(&reader_state).unwrap();

        assert_eq!(
            apply_update_with_typing_inference(&reader, &update).unwrap(),
            None
        );
    }
}
