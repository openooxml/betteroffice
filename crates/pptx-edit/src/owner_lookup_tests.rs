use std::collections::BTreeSet;
use std::sync::Arc;

use yrs::{Any, Array, ArrayRef, Map, MapRef, Transact, TransactionMut};

use crate::deck::{
    live_shape_order, live_slide_ids, map_string_array, required_map, shape_ref, slide_ref,
    slide_shape_order, string_array_ref,
};
use crate::proposal_diff::story_owner;
use crate::story_owners::{LookupCounts, find_path};
use crate::{DeckSession, EditCtx, EditError, SHAPES, ShapeDraft, ShapeRect, TextStyle};

struct SplitMix(u64);

impl SplitMix {
    fn next(&mut self) -> u64 {
        self.0 = self.0.wrapping_add(0x9E37_79B9_7F4A_7C15);
        let mut value = self.0;
        value = (value ^ (value >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        value = (value ^ (value >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        value ^ (value >> 31)
    }

    fn below(&mut self, bound: usize) -> usize {
        (self.next() % bound as u64) as usize
    }

    fn pick<T: Clone>(&mut self, items: &[T]) -> Option<T> {
        if items.is_empty() {
            None
        } else {
            Some(items[self.below(items.len())].clone())
        }
    }
}

fn empty_session(client_id: u64) -> DeckSession {
    let mut package = pptx_parse::PptxPackage::default();
    package.presentation.width_emu = 9_144_000;
    package.presentation.height_emu = 6_858_000;
    DeckSession::from_package_with_fingerprint(package, "owner-lookup".to_owned(), client_id)
        .unwrap()
}

fn add_box(session: &DeckSession, slide: &str) -> String {
    session
        .add_text_box(
            &EditCtx::local("test"),
            slide,
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
        .shape_id
}

fn sync(from: &DeckSession, to: &DeckSession) {
    let diff = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.apply_update_v1(&diff).unwrap();
}

#[derive(Clone)]
struct Placed {
    slide: String,
    shape: String,
    parent: Option<String>,
    children: Vec<String>,
}

fn placed_shapes(session: &DeckSession) -> Vec<Placed> {
    let txn = session.doc.transact();
    let mut placed = Vec::new();
    for slide_id in live_slide_ids(&txn).unwrap() {
        let slide = slide_ref(&txn, &slide_id).unwrap();
        let order = slide_shape_order(&slide, &txn).unwrap();
        let mut pending: Vec<(String, Option<String>)> = live_shape_order(&order, &txn)
            .unwrap()
            .into_iter()
            .map(|id| (id, None))
            .collect();
        while let Some((shape_id, parent)) = pending.pop() {
            let Ok(shape) = shape_ref(&txn, &shape_id) else {
                continue;
            };
            let children = map_string_array(&shape, &txn, "children").unwrap();
            pending.extend(
                children
                    .iter()
                    .map(|child| (child.clone(), Some(shape_id.clone()))),
            );
            placed.push(Placed {
                slide: slide_id.clone(),
                shape: shape_id,
                parent,
                children,
            });
        }
    }
    placed
}

fn walk_finds_owner(session: &DeckSession, story: &str) -> bool {
    let txn = session.doc.transact();
    story_owner(&txn, story).ok().flatten().is_some()
}

fn order_of(txn: &TransactionMut<'_>, slide: &str) -> ArrayRef {
    slide_shape_order(&slide_ref(txn, slide).unwrap(), txn).unwrap()
}

fn remove_entries(order: &ArrayRef, txn: &mut TransactionMut<'_>, id: &str) {
    let entries = string_array_ref(order, &*txn);
    for (index, entry) in entries.iter().enumerate().rev() {
        if entry == id {
            order.remove(txn, index as u32);
        }
    }
}

fn set_children(shape: &MapRef, txn: &mut TransactionMut<'_>, ids: &[String]) {
    let values: Vec<Any> = ids.iter().map(|id| Any::from(id.as_str())).collect();
    shape.insert(txn, "children", Any::Array(Arc::from(values)));
}

fn ancestors(placed: &[Placed], shape: &Placed) -> Vec<String> {
    let mut found = Vec::new();
    let mut parent = shape.parent.clone();
    while let Some(id) = parent {
        parent = placed
            .iter()
            .find(|candidate| candidate.slide == shape.slide && candidate.shape == id)
            .and_then(|candidate| candidate.parent.clone());
        found.push(id);
    }
    found
}

fn group_into(session: &DeckSession, placed: &[Placed], rng: &mut SplitMix) {
    let Some(container) = rng.pick(placed) else {
        return;
    };
    let above = ancestors(placed, &container);
    let members: Vec<Placed> = placed
        .iter()
        .filter(|candidate| {
            candidate.slide == container.slide
                && candidate.parent.is_none()
                && candidate.shape != container.shape
                && !above.contains(&candidate.shape)
        })
        .cloned()
        .collect();
    let Some(member) = rng.pick(&members) else {
        return;
    };
    nest(session, &container.slide, &container.shape, &member.shape);
}

// No grouping API exists, so nest, group_into, ungroup and move_out edit the document directly.
fn nest(session: &DeckSession, slide: &str, container: &str, member: &str) {
    let mut txn = session.transact_for(&EditCtx::local("test"));
    let order = order_of(&txn, slide);
    remove_entries(&order, &mut txn, member);
    let shape = shape_ref(&txn, container).unwrap();
    let mut children = map_string_array(&shape, &txn, "children").unwrap();
    children.push(member.to_owned());
    set_children(&shape, &mut txn, &children);
}

fn ungroup(session: &DeckSession, placed: &[Placed], rng: &mut SplitMix) {
    let groups: Vec<Placed> = placed
        .iter()
        .filter(|candidate| candidate.parent.is_none() && !candidate.children.is_empty())
        .cloned()
        .collect();
    let Some(group) = rng.pick(&groups) else {
        return;
    };
    let mut txn = session.transact_for(&EditCtx::local("test"));
    let order = order_of(&txn, &group.slide);
    remove_entries(&order, &mut txn, &group.shape);
    for child in &group.children {
        order.push_back(&mut txn, child.as_str());
    }
    let shape = shape_ref(&txn, &group.shape).unwrap();
    set_children(&shape, &mut txn, &[]);
    shape.remove(&mut txn, "id");
}

fn move_out(session: &DeckSession, placed: &[Placed], rng: &mut SplitMix) {
    let nested: Vec<Placed> = placed
        .iter()
        .filter(|candidate| candidate.parent.is_some())
        .cloned()
        .collect();
    let Some(member) = rng.pick(&nested) else {
        return;
    };
    let parent_id = member.parent.clone().unwrap();
    let Some(parent) = placed
        .iter()
        .find(|candidate| candidate.slide == member.slide && candidate.shape == parent_id)
    else {
        return;
    };
    let rest: Vec<String> = parent
        .children
        .iter()
        .filter(|id| **id != member.shape)
        .cloned()
        .collect();
    let mut txn = session.transact_for(&EditCtx::local("test"));
    let shape = shape_ref(&txn, &parent.shape).unwrap();
    set_children(&shape, &mut txn, &rest);
    let order = order_of(&txn, &member.slide);
    order.push_back(&mut txn, member.shape.as_str());
}

const MAX_SLIDES: usize = 8;
const MAX_SHAPES: usize = 40;
const STEADY: [usize; 7] = [0, 1, 5, 6, 7, 8, 9];

fn below_cap(slides: usize, shapes: usize) -> bool {
    slides < MAX_SLIDES && shapes < MAX_SHAPES
}

fn over_cap(session: &DeckSession) -> bool {
    let slides = live_slide_ids(&session.doc.transact()).unwrap().len();
    slides > MAX_SLIDES || placed_shapes(session).len() > MAX_SHAPES
}

fn step(rng: &mut SplitMix, local: &DeckSession, peer: &DeckSession, stories: &[String]) -> bool {
    let context = EditCtx::local("test");
    let slides = live_slide_ids(&local.doc.transact()).unwrap();
    let placed = placed_shapes(local);
    let choice = if below_cap(slides.len(), placed.len()) {
        rng.below(11)
    } else {
        rng.pick(&STEADY).unwrap()
    };
    match choice {
        0 => {
            if let Some(slide) = rng.pick(&slides) {
                local.delete_slide(&context, &slide).unwrap();
            }
        }
        1 => {
            let tops: Vec<Placed> = placed
                .iter()
                .filter(|candidate| candidate.parent.is_none())
                .cloned()
                .collect();
            if let Some(shape) = rng.pick(&tops) {
                local
                    .remove_shape(&context, &shape.slide, &shape.shape)
                    .unwrap();
            }
        }
        2 => {
            if local.undo() && over_cap(local) {
                local.redo();
            }
        }
        3 => {
            if local.redo() && over_cap(local) {
                local.undo();
            }
        }
        4 => {
            if let Some(slide) = rng.pick(&slides)
                && placed.len() + placed.iter().filter(|shape| shape.slide == slide).count()
                    <= MAX_SHAPES
            {
                local
                    .duplicate_slide(&context, &slide, slides.len() as u32)
                    .unwrap();
            }
        }
        5 => group_into(local, &placed, rng),
        6 => ungroup(local, &placed, rng),
        7 => move_out(local, &placed, rng),
        8 => {
            sync(local, peer);
            let remote = live_slide_ids(&peer.doc.transact()).unwrap();
            if let Some(slide) = rng.pick(&remote) {
                peer.delete_slide(&context, &slide).unwrap();
                sync(peer, local);
            }
        }
        9 => {
            if let Some(story) = rng.pick(stories) {
                let owned = walk_finds_owner(local, &story);
                let written = local.insert_text(&context, &story, 0, "x", &TextStyle::default());
                assert_eq!(written.is_ok(), owned, "write to {story}: {written:?}");
                if !owned {
                    assert!(matches!(written, Err(EditError::StoryNotFound(_))));
                }
            }
        }
        _ => {
            if let Some(slide) = rng.pick(&slides) {
                add_box(local, &slide);
            }
        }
    }
    local.add_undo_barrier();
    matches!(choice, 4 | 10)
}

fn discover(session: &DeckSession, seen: &mut BTreeSet<String>, stories: &mut Vec<String>) {
    let found: Vec<String> = {
        let txn = session.doc.transact();
        required_map(&txn, SHAPES)
            .unwrap()
            .iter(&txn)
            .filter_map(|(_, value)| value.cast::<MapRef>().ok())
            .flat_map(|shape| map_string_array(&shape, &txn, "textStories").unwrap_or_default())
            .collect()
    };
    for story in found {
        if seen.insert(story.clone()) {
            stories.push(story);
        }
    }
}

fn resolves(session: &DeckSession, story: &str) -> bool {
    session
        .anchor_caret(story, 0)
        .ok()
        .and_then(|anchor| session.resolve_caret_anchor(&anchor))
        .is_some()
}

fn grouped_deck(session: &DeckSession) {
    let context = EditCtx::local("test");
    for index in 0..3 {
        let slide = session
            .insert_slide(&context, index, None)
            .unwrap()
            .slide_id;
        let boxes: Vec<String> = (0..3).map(|_| add_box(session, &slide)).collect();
        if index == 2 {
            nest(session, &slide, &boxes[1], &boxes[2]);
        }
        nest(session, &slide, &boxes[0], &boxes[1]);
    }
    session.add_undo_barrier();
}

#[test]
fn resolving_a_story_agrees_with_the_full_walk_after_every_generated_edit() {
    let mut totals = LookupCounts::default();
    let mut checked = 0_usize;
    for seed in 1..=50_u64 {
        let every_story = seed <= 5;
        let local = empty_session(41);
        let peer = empty_session(42);
        let mut rng = SplitMix(seed);
        grouped_deck(&local);
        let mut seen = BTreeSet::new();
        let mut stories = Vec::new();
        discover(&local, &mut seen, &mut stories);
        for round in 0..200 {
            if step(&mut rng, &local, &peer, &stories) {
                discover(&local, &mut seen, &mut stories);
            }
            let sample: Vec<String> = if every_story {
                stories.clone()
            } else {
                (0..2).filter_map(|_| rng.pick(&stories)).collect()
            };
            for story in sample {
                checked += 1;
                let owned = walk_finds_owner(&local, &story);
                for lookup in ["first", "second"] {
                    assert_eq!(
                        resolves(&local, &story),
                        owned,
                        "seed {seed} round {round} story {story} {lookup} lookup"
                    );
                }
            }
        }
        let counts = local.story_owners.counts.get();
        totals.hits += counts.hits;
        totals.rejected += counts.rejected;
        totals.none += counts.none;
    }
    println!("\nstories checked: {checked}, lookups: {totals:?}");
    assert!(totals.hits > 0, "{totals:?}");
    assert!(totals.rejected > 0, "{totals:?}");
    assert!(totals.none > 0, "{totals:?}");
}

#[test]
fn the_lookup_search_finds_the_strict_walk_owner_on_undamaged_generated_decks() {
    let mut compared = 0_usize;
    let mut owned = 0_usize;
    for seed in 1..=50_u64 {
        let local = empty_session(41);
        let peer = empty_session(42);
        let mut rng = SplitMix(seed);
        grouped_deck(&local);
        let mut seen = BTreeSet::new();
        let mut stories = Vec::new();
        discover(&local, &mut seen, &mut stories);
        for round in 1..=200 {
            if step(&mut rng, &local, &peer, &stories) {
                discover(&local, &mut seen, &mut stories);
            }
            if seed > 5 && round % 20 != 0 {
                continue;
            }
            let txn = local.doc.transact();
            for story in &stories {
                let walked = story_owner(&txn, story).unwrap();
                let searched = find_path(&txn, story).and_then(|path| path.owner());
                assert_eq!(searched, walked, "seed {seed} round {round} story {story}");
                compared += 1;
                owned += usize::from(walked.is_some());
            }
        }
    }
    println!("\nstories compared: {compared}, with an owner: {owned}");
    assert!(owned > 0 && owned < compared);
}
