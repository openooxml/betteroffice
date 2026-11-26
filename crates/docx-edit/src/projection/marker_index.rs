use std::collections::{BTreeSet, HashMap};
use std::sync::{Arc, Mutex};

use yrs::types::{Event, PathSegment};
use yrs::updates::decoder::Decode;
use yrs::{Any, DeepObservable, Doc, Map, Out, ReadTxn, StickyIndex, Subscription};

use crate::{COMMENTS, RAW_INLINES, decode_anchor, projection::projection_root, sticky_offset};

#[derive(Default)]
struct Membership {
    by_id: HashMap<String, BTreeSet<String>>,
    by_story: HashMap<String, BTreeSet<String>>,
}

impl Membership {
    fn replace(&mut self, id: &str, stories: BTreeSet<String>) -> BTreeSet<String> {
        let previous = self.by_id.remove(id).unwrap_or_default();
        for story in previous.difference(&stories) {
            if let Some(ids) = self.by_story.get_mut(story) {
                ids.remove(id);
                if ids.is_empty() {
                    self.by_story.remove(story);
                }
            }
        }
        for story in stories.difference(&previous) {
            self.by_story.entry(story.clone()).or_default().insert(id.to_owned());
        }
        let affected = previous.union(&stories).cloned().collect();
        if !stories.is_empty() {
            self.by_id.insert(id.to_owned(), stories);
        }
        affected
    }
}

#[derive(Default)]
pub(crate) struct MarkerIndex {
    comments: Membership,
    raw: Membership,
    revision: u64,
    roots: HashMap<String, u64>,
}

impl MarkerIndex {
    pub(crate) fn watch(doc: &Doc) -> (Arc<Mutex<Self>>, [Subscription; 2]) {
        let index = Arc::new(Mutex::new(Self::default()));
        let subscriptions = [COMMENTS, RAW_INLINES].map(|name| {
            let index = Arc::clone(&index);
            let map = doc.get_or_insert_map(name);
            let values = map.clone();
            map.observe_deep(move |txn, events| {
                let mut ids = BTreeSet::new();
                for event in events.iter() {
                    let Event::Map(event) = event else { continue };
                    match event.path().front() {
                        None => ids.extend(event.keys(txn).keys().cloned()),
                        Some(PathSegment::Key(id)) => {
                            ids.insert(Arc::clone(id));
                        }
                        Some(PathSegment::Index(_)) => {}
                    }
                }
                let mut index = index.lock().unwrap();
                index.revision += 1;
                let revision = index.revision;
                for id in ids {
                    let value = values.get(txn, &id);
                    let affected = if name == COMMENTS {
                        let stories = value.and_then(|value| value.cast::<yrs::MapRef>().ok())
                            .and_then(|comment| match comment.get(txn, "anchors") {
                                Some(Out::Any(Any::Array(anchors))) => Some(anchors),
                                _ => None,
                            })
                            .and_then(|anchors| anchors.iter().map(|anchor| {
                                decode_anchor(anchor).map(|anchor| anchor.story)
                            }).collect::<Result<BTreeSet<_>, _>>().ok())
                            .unwrap_or_default();
                        index.comments.replace(&id, stories)
                    } else {
                        let stories = raw_mark(value).and_then(|mark| match mark.get("story") {
                            Some(Any::String(story)) => Some(story.to_string()),
                            _ => None,
                        }).into_iter().collect();
                        index.raw.replace(&id, stories)
                    };
                    for root in affected.iter().filter_map(|story| projection_root(story)) {
                        index.roots.insert(root.to_owned(), revision);
                    }
                }
            })
        });
        (index, subscriptions)
    }

    pub(super) fn revision(&self, root: &str) -> u64 {
        self.roots.get(root).copied().unwrap_or_default()
    }

    pub(super) fn has_comments(&self, story: &str) -> bool {
        self.comments.by_story.contains_key(story)
    }

    pub(super) fn raw_offsets<T: ReadTxn>(&self, txn: &T, story: &str) -> Vec<usize> {
        let Some(ids) = self.raw.by_story.get(story) else { return Vec::new() };
        let Some(marks) = txn.get_map(RAW_INLINES) else { return Vec::new() };
        let mut offsets: Vec<_> = ids.iter().filter_map(|id| {
            let mark = raw_mark(marks.get(txn, id))?;
            let (Some(Any::Buffer(at)), Some(Any::String(_)), Some(Any::Number(_))) =
                (mark.get("at"), mark.get("xml"), mark.get("ordinal"))
            else {
                return None;
            };
            let anchor = StickyIndex::decode_v1(at).ok()?;
            Some(sticky_offset(&anchor, txn)?.index as usize)
        }).collect();
        offsets.sort_unstable();
        offsets
    }
}

fn raw_mark(value: Option<Out>) -> Option<Arc<HashMap<String, Any>>> {
    match value? {
        Out::Any(Any::Map(mark)) => Some(mark),
        _ => None,
    }
}
