use std::cell::RefCell;
use std::collections::{HashMap, HashSet};

use yrs::{Any, Array, ArrayRef, Map, MapRef, Out, ReadTxn};

use crate::deck::{
    live_shape_order, map_string_array, required_map, shape_ref, slide_ref, slide_shape_order,
};
use crate::proposal_diff::story_owner;
use crate::{EditResult, SHAPES, SLIDE_ORDER};

/// The slide and the shape ids from a top-level shape down to the owner of a story, as last
/// found by the full walk.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct OwnerPath {
    slide_id: String,
    shape_ids: Vec<String>,
}

/// Per-session memo of story owners; an entry answers only after `path_holds` re-checks it
/// against the current document, so remote updates, undo and redo need no invalidation.
#[derive(Default)]
pub(crate) struct StoryOwners {
    paths: RefCell<HashMap<String, OwnerPath>>,
}

impl StoryOwners {
    pub(crate) fn owner<T: ReadTxn>(
        &self,
        txn: &T,
        story_id: &str,
    ) -> EditResult<Option<(String, String)>> {
        if let Some(path) = self.paths.borrow().get(story_id)
            && path_holds(txn, story_id, path)
            && let Some(shape_id) = path.shape_ids.last()
        {
            return Ok(Some((path.slide_id.clone(), shape_id.clone())));
        }
        let owner = story_owner(txn, story_id)?;
        let path = owner
            .as_ref()
            .and_then(|(slide_id, shape_id)| owner_path(txn, slide_id, shape_id));
        let mut paths = self.paths.borrow_mut();
        match path {
            Some(path) => {
                paths.insert(story_id.to_owned(), path);
            }
            None => {
                paths.remove(story_id);
            }
        }
        Ok(owner)
    }
}

fn path_holds<T: ReadTxn>(txn: &T, story_id: &str, path: &OwnerPath) -> bool {
    let Some(top) = path.shape_ids.first() else {
        return false;
    };
    let Ok(slide) = slide_ref(txn, &path.slide_id) else {
        return false;
    };
    let slide_listed = txn
        .get_array(SLIDE_ORDER)
        .is_some_and(|order| order_lists(&order, txn, &path.slide_id));
    if !slide_listed
        || !slide_shape_order(&slide, txn).is_ok_and(|order| order_lists(&order, txn, top))
    {
        return false;
    }
    let mut parent: Option<MapRef> = None;
    for shape_id in &path.shape_ids {
        if let Some(parent) = &parent
            && !array_lists(parent, txn, "children", shape_id)
        {
            return false;
        }
        let Ok(shape) = shape_ref(txn, shape_id) else {
            return false;
        };
        parent = Some(shape);
    }
    parent.is_some_and(|owner| array_lists(&owner, txn, "textStories", story_id))
}

fn order_lists<T: ReadTxn>(order: &ArrayRef, txn: &T, id: &str) -> bool {
    order
        .iter(txn)
        .any(|value| matches!(value, Out::Any(Any::String(entry)) if &*entry == id))
}

fn array_lists<T: ReadTxn>(map: &MapRef, txn: &T, key: &str, id: &str) -> bool {
    match map.get(txn, key) {
        Some(Out::Any(Any::Array(values))) => values
            .iter()
            .any(|value| matches!(value, Any::String(entry) if &**entry == id)),
        _ => false,
    }
}

/// The first path to `shape_id` in the order `story_owner` walks the slide.
fn owner_path<T: ReadTxn>(txn: &T, slide_id: &str, shape_id: &str) -> Option<OwnerPath> {
    let slide = slide_ref(txn, slide_id).ok()?;
    let shapes = required_map(txn, SHAPES).ok()?;
    let order = slide_shape_order(&slide, txn).ok()?;
    let mut pending: Vec<Vec<String>> = live_shape_order(&order, txn)
        .ok()?
        .into_iter()
        .rev()
        .map(|id| vec![id])
        .collect();
    let mut seen = HashSet::new();
    while let Some(path) = pending.pop() {
        let id = path.last()?.clone();
        if id == shape_id {
            return Some(OwnerPath {
                slide_id: slide_id.to_owned(),
                shape_ids: path,
            });
        }
        if !seen.insert(id.clone()) {
            continue;
        }
        let shape = shapes.get(txn, &id)?.cast::<MapRef>().ok()?;
        let children = map_string_array(&shape, txn, "children").ok()?;
        for child in children.into_iter().rev() {
            let mut next = path.clone();
            next.push(child);
            pending.push(next);
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use yrs::Transact;

    use super::*;
    use crate::deck::string_array_ref;
    use crate::{DeckSession, EditCtx, ShapeDraft, ShapeRect, TextStyle};

    fn session() -> (DeckSession, Vec<String>, String, String) {
        let mut package = pptx_parse::PptxPackage::default();
        package.presentation.width_emu = 9_144_000;
        package.presentation.height_emu = 6_858_000;
        let session =
            DeckSession::from_package_with_fingerprint(package, "owner-paths".to_owned(), 43)
                .unwrap();
        let context = EditCtx::local("test");
        let slides: Vec<String> = (0..2)
            .map(|index| {
                session
                    .insert_slide(&context, index, None)
                    .unwrap()
                    .slide_id
            })
            .collect();
        let shape = session
            .add_text_box(
                &context,
                &slides[0],
                &ShapeDraft {
                    name: "Box".to_owned(),
                    rect: ShapeRect {
                        x: 0,
                        y: 0,
                        width: 1000,
                        height: 1000,
                    },
                    text: "text".to_owned(),
                    style: TextStyle::default(),
                },
            )
            .unwrap()
            .shape_id;
        let story = format!("story:{shape}:0");
        (session, slides, shape, story)
    }

    fn cached(session: &DeckSession, story: &str) -> OwnerPath {
        let txn = session.doc.transact();
        assert!(session.story_owners.owner(&txn, story).unwrap().is_some());
        session
            .story_owners
            .paths
            .borrow()
            .get(story)
            .cloned()
            .unwrap()
    }

    fn holds(session: &DeckSession, story: &str, path: &OwnerPath) -> bool {
        path_holds(&session.doc.transact(), story, path)
    }

    fn owner(session: &DeckSession, story: &str) -> Option<(String, String)> {
        let txn = session.doc.transact();
        session.story_owners.owner(&txn, story).unwrap()
    }

    fn move_shape_to(session: &DeckSession, from: &str, to: &str, shape: &str) {
        let mut txn = session.transact_for(&EditCtx::local("test"));
        let source = slide_shape_order(&slide_ref(&txn, from).unwrap(), &txn).unwrap();
        let index = string_array_ref(&source, &txn)
            .iter()
            .position(|id| id == shape)
            .unwrap();
        source.remove(&mut txn, index as u32);
        let target = slide_shape_order(&slide_ref(&txn, to).unwrap(), &txn).unwrap();
        target.push_back(&mut txn, shape);
    }

    #[test]
    fn a_cached_path_holds_while_nothing_changed() {
        let (session, slides, shape, story) = session();
        let path = cached(&session, &story);
        assert_eq!(path.slide_id, slides[0]);
        assert_eq!(path.shape_ids, vec![shape]);
        assert!(holds(&session, &story, &path));
    }

    #[test]
    fn a_cached_path_is_rejected_once_its_shape_is_removed() {
        let (session, slides, shape, story) = session();
        let path = cached(&session, &story);
        session
            .remove_shape(&EditCtx::local("test"), &slides[0], &shape)
            .unwrap();
        assert!(!holds(&session, &story, &path));
        assert_eq!(owner(&session, &story), None);
    }

    #[test]
    fn a_cached_path_is_rejected_once_its_slide_is_deleted() {
        let (session, slides, _, story) = session();
        let path = cached(&session, &story);
        session
            .delete_slide(&EditCtx::local("test"), &slides[0])
            .unwrap();
        assert!(!holds(&session, &story, &path));
        assert_eq!(owner(&session, &story), None);
    }

    #[test]
    fn a_cached_path_is_rejected_once_its_shape_moves_to_another_slide() {
        let (session, slides, shape, story) = session();
        let path = cached(&session, &story);
        move_shape_to(&session, &slides[0], &slides[1], &shape);
        assert!(!holds(&session, &story, &path));
        assert_eq!(
            owner(&session, &story),
            Some((slides[1].clone(), shape.clone()))
        );
        assert_eq!(cached(&session, &story).slide_id, slides[1]);
    }

    #[test]
    fn a_cached_path_is_rejected_once_its_shape_leaves_its_group() {
        let (session, slides, shape, story) = session();
        let context = EditCtx::local("test");
        let group = session
            .add_text_box(
                &context,
                &slides[0],
                &ShapeDraft {
                    name: "Group".to_owned(),
                    rect: ShapeRect {
                        x: 0,
                        y: 0,
                        width: 1000,
                        height: 1000,
                    },
                    text: "group".to_owned(),
                    style: TextStyle::default(),
                },
            )
            .unwrap()
            .shape_id;
        {
            let mut txn = session.transact_for(&context);
            let order = slide_shape_order(&slide_ref(&txn, &slides[0]).unwrap(), &txn).unwrap();
            let index = string_array_ref(&order, &txn)
                .iter()
                .position(|id| *id == shape)
                .unwrap();
            order.remove(&mut txn, index as u32);
            let entry = shape_ref(&txn, &group).unwrap();
            entry.insert(
                &mut txn,
                "children",
                Any::Array(std::sync::Arc::from(vec![Any::from(shape.as_str())])),
            );
        }
        let path = cached(&session, &story);
        assert_eq!(path.shape_ids, vec![group.clone(), shape.clone()]);
        {
            let mut txn = session.transact_for(&context);
            let entry = shape_ref(&txn, &group).unwrap();
            entry.insert(
                &mut txn,
                "children",
                Any::Array(std::sync::Arc::from(Vec::<Any>::new())),
            );
        }
        assert!(!holds(&session, &story, &path));
        assert_eq!(owner(&session, &story), None);
    }
}
