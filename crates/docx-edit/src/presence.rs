#![cfg_attr(not(feature = "wasm"), allow(dead_code))]

use std::sync::{Arc, Mutex};

use yrs::branch::{Branch, BranchPtr};
use yrs::types::Delta;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{
    Any, Assoc, ID, IndexedSequence, Map, Observable, Out, ReadTxn, StickyIndex, TextRef, Transact,
    Update,
};

use crate::{EditingDoc, SegmentContent, story_ref};

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
    let anchor = StickyIndex::decode_v1(anchor).map_err(|error| error.to_string())?;
    let head = StickyIndex::decode_v1(head).map_err(|error| error.to_string())?;
    let txn = doc.yrs_doc().transact();
    let expected = story_ref(&txn, story).map_err(|error| error.to_string())?;
    let anchor = anchor
        .get_offset(&txn)
        .ok_or_else(|| "selection anchor no longer resolves".to_owned())?;
    let head = head
        .get_offset(&txn)
        .ok_or_else(|| "selection head no longer resolves".to_owned())?;
    let expected = BranchPtr::from(<TextRef as AsRef<Branch>>::as_ref(&expected));
    if anchor.branch != expected || head.branch != expected {
        return Err("selection sticky index belongs to another story".to_owned());
    }
    Ok((anchor.index, head.index))
}

pub(crate) fn apply_update_with_typing_inference(
    doc: &EditingDoc,
    bytes: &[u8],
) -> Result<Option<TypingInference>, String> {
    let update = Update::decode_v1(bytes).map_err(|error| error.to_string())?;
    let insertions = update.insertions(false);
    let mut clients = insertions.client_ids();
    let client_id = clients.next().filter(|_| clients.next().is_none());
    let changes = Arc::new(Mutex::new(Vec::<InsertedContent>::new()));
    let subscriptions = if client_id.is_some() {
        let txn = doc.yrs_doc().transact();
        txn.get_map(super::STORIES)
            .map(|stories| {
                stories
                    .iter(&txn)
                    .filter_map(|(story, value)| match value {
                        Out::YText(text) => {
                            let story = story.to_string();
                            let changes = Arc::clone(&changes);
                            Some(text.observe(move |txn, event| {
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
                                            if let Some(id) = event
                                                .target()
                                                .sticky_index(txn, end_index, Assoc::Before)
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
                            }))
                        }
                        _ => None,
                    })
                    .collect()
            })
            .unwrap_or_default()
    } else {
        Vec::new()
    };

    doc.yrs_doc()
        .transact_mut()
        .apply_update(update)
        .map_err(|error| error.to_string())?;
    drop(subscriptions);
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
    let mut cursor = 0_u32;
    let mut para_start = 0_u32;
    for segment in doc
        .story_segments(story)
        .map_err(|error| error.to_string())?
    {
        match segment.content {
            SegmentContent::Text(text) => cursor += text.encode_utf16().count() as u32,
            SegmentContent::Pilcrow(properties) => {
                if index <= cursor {
                    return Ok((properties.para_id, index.saturating_sub(para_start)));
                }
                cursor += 1;
                para_start = cursor;
            }
            SegmentContent::OtherEmbed { .. } => cursor += 1,
        }
    }
    Err(format!(
        "selection index {index} does not resolve in story {story:?}"
    ))
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use super::*;
    use yrs::{Any, Assoc, IndexedSequence};

    use crate::{EditCtx, FormatPolicy, Position, StoryRange};

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
