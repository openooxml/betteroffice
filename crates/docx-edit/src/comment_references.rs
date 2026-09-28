//! Comment reference embeds follow their comment's anchors.

use std::collections::{BTreeMap, BTreeSet};
use std::sync::{Arc, Mutex};

use yrs::types::text::YChange;
use yrs::types::{Attrs, EntryChange, Event, PathSegment};
use yrs::{
    Any, DeepObservable, Map, MapPrelim, MapRef, Out, ReadTxn, Subscription, Text, TextRef,
    Transact, TransactionMut,
};

use crate::{COMMENTS, EditingDoc, STORIES, decode_anchor, map_string, out_len};

/// Records the comments an update reanchors while it integrates: a replaced comment or changed
/// anchors, never a comment the update first brings in.
pub(crate) struct CommentWatch {
    reanchored: Arc<Mutex<BTreeSet<String>>>,
    _subscription: Option<Subscription>,
}

impl CommentWatch {
    pub(crate) fn new(doc: &EditingDoc) -> Self {
        let reanchored = Arc::new(Mutex::new(BTreeSet::new()));
        let comments = doc.yrs_doc().transact().get_map(COMMENTS);
        let subscription = comments.map(|comments| {
            let reanchored = Arc::clone(&reanchored);
            comments.observe_deep(move |txn, events| {
                let mut reanchored = reanchored.lock().unwrap();
                for event in events.iter() {
                    let Event::Map(event) = event else {
                        continue;
                    };
                    let path = event.path();
                    match path.front() {
                        Some(PathSegment::Key(id)) => {
                            if path.len() == 1 && event.keys(txn).contains_key("anchors") {
                                reanchored.insert(id.to_string());
                            }
                        }
                        _ => {
                            for (id, change) in event.keys(txn).iter() {
                                if matches!(change, EntryChange::Updated(..)) {
                                    reanchored.insert(id.to_string());
                                }
                            }
                        }
                    }
                }
            })
        });
        Self {
            reanchored,
            _subscription: subscription,
        }
    }

    pub(crate) fn take(self) -> BTreeSet<String> {
        std::mem::take(&mut self.reanchored.lock().unwrap())
    }
}

/// The comment a reference embed belongs to.
fn referenced_comment<T: ReadTxn>(map: &MapRef, txn: &T) -> Option<String> {
    if map_string(map, txn, "modelKind").as_deref() != Some("commentReference") {
        return None;
    }
    match map.get(txn, "commentId")? {
        Out::Any(Any::Number(id)) => Some(id.to_string()),
        Out::Any(Any::BigInt(id)) => Some(id.to_string()),
        Out::Any(Any::String(id)) => Some(id.to_string()),
        _ => None,
    }
}

/// Where the comment's ranges end, per story: the end of its last range there.
fn range_ends<T: ReadTxn>(txn: &T, comment_id: &str) -> BTreeMap<String, u32> {
    let mut ends = BTreeMap::new();
    let Some(Out::YMap(comment)) = txn
        .get_map(COMMENTS)
        .and_then(|comments| comments.get(txn, comment_id))
    else {
        return ends;
    };
    let Some(Out::Any(Any::Array(anchors))) = comment.get(txn, "anchors") else {
        return ends;
    };
    for anchor in anchors.iter().filter_map(|value| decode_anchor(value).ok()) {
        if let Some(end) = anchor.end.get_offset(txn) {
            let entry = ends.entry(anchor.story).or_insert(end.index);
            *entry = (*entry).max(end.index);
        }
    }
    ends
}

/// Leaves each comment one reference embed right after its range in each story holding a range,
/// and none elsewhere. With `place` false, as a reanchoring remote update needs, references are
/// only removed, and Save writes the reference beside the range; with it, a moved comment's
/// reference is also added where its new range lacks one. A comment whose anchors no longer
/// resolve keeps its references.
pub(crate) fn reconcile(txn: &mut TransactionMut<'_>, comment_ids: &BTreeSet<String>, place: bool) {
    let ends: BTreeMap<&str, BTreeMap<String, u32>> = comment_ids
        .iter()
        .map(|id| (id.as_str(), range_ends(txn, id)))
        .filter(|(_, ends)| !ends.is_empty())
        .collect();
    if ends.is_empty() {
        return;
    }
    let stories: Vec<(String, TextRef)> = txn
        .get_map(STORIES)
        .expect("stories root is declared by EditingDoc::new")
        .iter(txn)
        .filter_map(|(id, value)| Some((id.to_owned(), value.cast::<TextRef>().ok()?)))
        .collect();
    let mut found: BTreeMap<&str, Vec<(usize, u32)>> = BTreeMap::new();
    let mut payloads: BTreeMap<&str, Vec<(String, Any)>> = BTreeMap::new();
    for (index, (_, story)) in stories.iter().enumerate() {
        let mut offset = 0;
        for diff in story.diff(txn, YChange::identity) {
            if let Out::YMap(map) = &diff.insert
                && let Some(id) = referenced_comment(map, txn)
                && let Some((&id, _)) = ends.get_key_value(id.as_str())
            {
                found.entry(id).or_default().push((index, offset));
                payloads.entry(id).or_insert_with(|| {
                    map.iter(txn)
                        .filter_map(|(key, value)| match value {
                            Out::Any(value) => Some((key.to_owned(), value)),
                            _ => None,
                        })
                        .collect()
                });
            }
            offset += out_len(&diff.insert);
        }
    }
    let mut removed = vec![Vec::new(); stories.len()];
    let mut missing = Vec::new();
    for (&id, story_ends) in &ends {
        let references = found.get(id).map(Vec::as_slice).unwrap_or_default();
        let mut kept = BTreeSet::new();
        for &(index, offset) in references {
            if story_ends.get(&stories[index].0) != Some(&offset) || !kept.insert(index) {
                removed[index].push(offset);
            }
        }
        if place && payloads.contains_key(id) {
            missing.extend(
                stories
                    .iter()
                    .enumerate()
                    .filter(|(index, (story, _))| {
                        story_ends.contains_key(story) && !kept.contains(index)
                    })
                    .map(|(index, _)| (id, index)),
            );
        }
    }
    for (offsets, (_, story)) in removed.iter_mut().zip(&stories) {
        offsets.sort_unstable();
        for &offset in offsets.iter().rev() {
            story.remove_range(txn, offset, 1);
        }
    }
    for (id, index) in missing {
        let (story_id, story) = &stories[index];
        let Some(&end) = range_ends(txn, id).get(story_id) else {
            continue;
        };
        let embed =
            story.insert_embed_with_attributes(txn, end, MapPrelim::default(), Attrs::new());
        for (key, value) in &payloads[id] {
            embed.insert(txn, key.as_str(), value.clone());
        }
    }
}
