use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::{Arc, Mutex};

use pptx_edit::{DeckSession, EditCtx, ShapeDraft, ShapeRect, TextStyle};
use yrs::branch::Branch;

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

/// Zeroes freed yrs branches and never reuses them, so a read after free fails the same way on
/// every run.
struct QuarantinedBranches;

unsafe impl GlobalAlloc for QuarantinedBranches {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        if layout == Layout::new::<Branch>() {
            unsafe { ptr.write_bytes(0, layout.size()) };
        } else {
            unsafe { System.dealloc(ptr, layout) };
        }
    }
}

#[global_allocator]
static ALLOCATOR: QuarantinedBranches = QuarantinedBranches;

struct Seed {
    update: Vec<u8>,
    slide_id: String,
    shape_id: String,
    story_id: String,
}

impl Seed {
    fn new() -> Self {
        let source = DeckSession::open(FIXTURE, 101).unwrap();
        let slide_id = source.snapshot().unwrap().slides[0].id.clone();
        let shape_id = source
            .add_text_box(&EditCtx::local("source"), &slide_id, &draft("Note"))
            .unwrap()
            .shape_id;
        let story_id = source.snapshot().unwrap().slides[0]
            .shapes
            .iter()
            .find(|shape| shape.id == shape_id)
            .unwrap()
            .text_stories[0]
            .id
            .clone();
        Self {
            update: source.encode_state_as_update_v1(),
            slide_id,
            shape_id,
            story_id,
        }
    }

    fn open(&self, client_id: u64) -> DeckSession {
        DeckSession::open_from_update(&self.update, client_id).unwrap()
    }

    fn remove_shape(&self, session: &DeckSession) {
        session
            .remove_shape(&EditCtx::local("peer"), &self.slide_id, &self.shape_id)
            .unwrap();
    }

    /// Moves, undoes and moves again, so clearing the redo stack unkeeps the shape that the
    /// restored position, still kept, points into.
    fn move_undo_move(&self, session: &DeckSession) {
        let context = EditCtx::local("local");
        session
            .move_shape(&context, &self.slide_id, &self.shape_id, 400_000, 300_000)
            .unwrap();
        session.add_undo_barrier();
        assert!(session.undo());
        session
            .move_shape(&context, &self.slide_id, &self.shape_id, 800_000, 300_000)
            .unwrap();
        session.add_undo_barrier();
    }
}

fn draft(text: &str) -> ShapeDraft {
    ShapeDraft {
        name: text.to_owned(),
        rect: ShapeRect {
            x: 100_000,
            y: 200_000,
            width: 3_000_000,
            height: 800_000,
        },
        text: text.to_owned(),
        style: TextStyle::default(),
    }
}

fn sync(from: &DeckSession, to: &DeckSession) {
    let diff = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.apply_update_v1(&diff).unwrap();
}

fn assert_converged(replicas: &[&DeckSession]) {
    for from in replicas {
        for to in replicas {
            if !std::ptr::eq(*from, *to) {
                sync(from, to);
            }
        }
    }
    let expected = replicas[0].snapshot().unwrap();
    for replica in &replicas[1..] {
        assert_eq!(replica.snapshot().unwrap(), expected);
    }
}

#[test]
fn a_peer_removing_a_shape_this_replica_edited_undid_and_edited_again_leaves_its_state_readable() {
    let seed = Seed::new();
    let local = seed.open(202);
    let peer = seed.open(303);
    seed.move_undo_move(&local);

    let sent = Arc::new(Mutex::new(Vec::new()));
    let sink = Arc::clone(&sent);
    let _subscription = peer
        .observe_update_v1(move |event| sink.lock().unwrap().push(event.update))
        .unwrap();
    seed.remove_shape(&peer);
    for update in sent.lock().unwrap().iter() {
        local.apply_update_v1(update).unwrap();
    }

    let state = local.encode_state_as_update_v1();
    peer.apply_update_v1(&state).unwrap();
    assert_eq!(local.snapshot().unwrap(), peer.snapshot().unwrap());
}

#[test]
fn a_replica_without_history_collects_what_a_peer_removes() {
    let seed = Seed::new();
    let peer = seed.open(202);
    let passive = seed.open(303);
    let before = passive.encode_state_as_update_v1().len();

    seed.remove_shape(&peer);
    sync(&peer, &passive);

    assert_eq!(passive.snapshot().unwrap(), peer.snapshot().unwrap());
    let after = passive.encode_state_as_update_v1().len();
    assert!(after < before, "state grew from {before} to {after} bytes");
}

#[test]
fn a_replica_that_collected_a_removal_still_edits_undoes_and_redoes() {
    let seed = Seed::new();
    let peer = seed.open(202);
    let passive = seed.open(303);
    seed.remove_shape(&peer);
    sync(&peer, &passive);

    let context = EditCtx::local("passive");
    let added = passive
        .add_text_box(&context, &seed.slide_id, &draft("Later"))
        .unwrap()
        .shape_id;
    passive.add_undo_barrier();
    passive
        .move_shape(&context, &seed.slide_id, &added, 500_000, 500_000)
        .unwrap();
    passive.add_undo_barrier();
    assert!(passive.undo());
    assert!(passive.redo());

    assert_converged(&[&passive, &peer]);
}

#[test]
fn a_replica_with_undo_history_survives_undo_and_redo_after_a_peer_removes_the_shape() {
    let seed = Seed::new();
    let local = seed.open(202);
    let peer = seed.open(303);
    local
        .delete_text(&EditCtx::local("local"), &seed.story_id, 0, 2)
        .unwrap();
    local.add_undo_barrier();
    seed.move_undo_move(&local);

    seed.remove_shape(&peer);
    sync(&peer, &local);
    while local.undo() {}
    while local.redo() {}

    assert_converged(&[&local, &peer]);
}

#[test]
fn a_replica_whose_undo_history_emptied_keeps_what_it_kept_after_a_peer_removes_the_shape() {
    let seed = Seed::new();
    let local = seed.open(202);
    let peer = seed.open(303);
    seed.move_undo_move(&local);

    seed.remove_shape(&peer);
    sync(&peer, &local);
    assert!(!local.undo());
    assert!(!local.can_redo());

    assert_converged(&[&local, &peer]);
}

#[test]
fn a_replica_holding_only_redo_history_keeps_what_a_peer_removes() {
    let seed = Seed::new();
    let local = seed.open(202);
    let peer = seed.open(303);
    let other = peer
        .add_text_box(&EditCtx::local("peer"), &seed.slide_id, &draft("Other"))
        .unwrap()
        .shape_id;
    sync(&peer, &local);
    let context = EditCtx::local("local");
    local
        .move_shape(&context, &seed.slide_id, &seed.shape_id, 400_000, 300_000)
        .unwrap();
    local.add_undo_barrier();
    assert!(local.undo());
    let before = local.encode_state_as_update_v1().len();

    peer.remove_shape(&EditCtx::local("peer"), &seed.slide_id, &other)
        .unwrap();
    sync(&peer, &local);

    let after = local.encode_state_as_update_v1().len();
    assert!(
        after >= before,
        "state shrank from {before} to {after} bytes"
    );
    assert!(local.redo());
    assert_converged(&[&local, &peer]);
}

#[test]
fn a_joiner_synced_from_a_collected_replica_converges_when_the_removal_is_undone() {
    let seed = Seed::new();
    let peer = seed.open(202);
    let passive = seed.open(303);
    seed.remove_shape(&peer);
    peer.add_undo_barrier();
    sync(&peer, &passive);
    let joiner = DeckSession::open_from_update(&passive.encode_state_as_update_v1(), 404).unwrap();

    assert!(peer.undo());
    assert_converged(&[&peer, &passive, &joiner]);
    let restored = &peer.snapshot().unwrap().slides[0].shapes;
    assert!(restored.iter().any(|shape| shape.id == seed.shape_id));
}
