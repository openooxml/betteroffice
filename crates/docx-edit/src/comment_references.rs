//! Comment reference embeds follow their comment's anchors.

use std::collections::{BTreeMap, BTreeSet};
use std::sync::{Arc, Mutex};

use yrs::types::text::YChange;
use yrs::types::{Attrs, Delta, EntryChange, Event, PathSegment};
use yrs::{
    Any, DeepObservable, Map, MapPrelim, MapRef, Out, ReadTxn, Subscription, Text, TextRef,
    Transact, TransactionMut,
};

use crate::{COMMENTS, EditingDoc, STORIES, decode_anchor, map_string, out_len};

/// Records the comments an update changes while it integrates, and which of them it reanchors:
/// a replaced comment, changed anchors, or a reference the update inserts into an existing story
/// (a losing concurrent move changes no anchors here but still moves its reference), never a
/// comment or story the update first brings in.
pub(crate) struct CommentWatch {
    changed: Arc<Mutex<(BTreeSet<String>, BTreeSet<String>)>>,
    _subscriptions: Vec<Subscription>,
}

impl CommentWatch {
    pub(crate) fn new(doc: &EditingDoc) -> Self {
        let changed = Arc::new(Mutex::new((BTreeSet::new(), BTreeSet::new())));
        let (comments, stories) = {
            let txn = doc.yrs_doc().transact();
            (txn.get_map(COMMENTS), txn.get_map(STORIES))
        };
        let mut subscriptions = Vec::new();
        subscriptions.extend(comments.map(|comments| {
            let changed = Arc::clone(&changed);
            comments.observe_deep(move |txn, events| {
                let (changed, reanchored) = &mut *changed.lock().unwrap();
                for event in events.iter() {
                    let Event::Map(event) = event else {
                        continue;
                    };
                    let path = event.path();
                    match path.front() {
                        Some(PathSegment::Key(id)) => {
                            changed.insert(id.to_string());
                            if path.len() == 1 && event.keys(txn).contains_key("anchors") {
                                reanchored.insert(id.to_string());
                            }
                        }
                        _ => {
                            for (id, change) in event.keys(txn).iter() {
                                changed.insert(id.to_string());
                                if matches!(change, EntryChange::Updated(..)) {
                                    reanchored.insert(id.to_string());
                                }
                            }
                        }
                    }
                }
            })
        }));
        subscriptions.extend(stories.map(|stories| {
            let changed = Arc::clone(&changed);
            stories.observe_deep(move |txn, events| {
                let (changed, reanchored) = &mut *changed.lock().unwrap();
                for event in events.iter() {
                    let Event::Text(event) = event else {
                        continue;
                    };
                    for delta in event.delta(txn) {
                        if let Delta::Inserted(Out::YMap(map), _) = delta
                            && let Some(id) = referenced_comment(map, txn)
                        {
                            changed.insert(id.clone());
                            reanchored.insert(id);
                        }
                    }
                }
            })
        }));
        Self {
            changed,
            _subscriptions: subscriptions,
        }
    }

    /// The changed comments, and the subset whose anchors changed or that were replaced.
    pub(crate) fn take(self) -> (BTreeSet<String>, BTreeSet<String>) {
        std::mem::take(&mut self.changed.lock().unwrap())
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
/// and none elsewhere. With `place` false, as a remote update needs, references are only removed:
/// duplicates, and for a comment in `reanchored` a lone one away from its range end, which Save
/// then writes beside the range. With it, a moved comment's reference is also added where its new
/// range lacks one. A comment whose anchors no longer resolve keeps its references.
pub(crate) fn reconcile(
    txn: &mut TransactionMut<'_>,
    comment_ids: &BTreeSet<String>,
    reanchored: &BTreeSet<String>,
    place: bool,
) {
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
        if !place && references.len() <= 1 && !reanchored.contains(id) {
            continue;
        }
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
