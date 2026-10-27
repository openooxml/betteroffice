use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::{Arc, Mutex};

use docx_edit::{
    EditCtx, EditingDoc, EngineSession, FormatPolicy, Position, StoryRange, UndoSession,
};
use yrs::branch::Branch;

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

/// The seed's `body` and `note` stories on `local`, and a peer holding the same seed.
fn seeded_peer(local: &EditingDoc) -> EditingDoc {
    let source = EditingDoc::new(101);
    source
        .create_story("body", "Body", "Normal", "left")
        .unwrap();
    source
        .create_story("note", "Note", "Normal", "left")
        .unwrap();
    let seed = source.encode_state_as_update_v1();
    local.apply_update_v1(&seed).unwrap();
    let peer = EditingDoc::new(303);
    peer.apply_update_v1(&seed).unwrap();
    peer
}

/// Edits `note` and undoes it, types in `body`, then integrates the peer deleting `note`.
fn lose_an_undone_story(local: &EditingDoc, peer: &EditingDoc, mut undo: impl FnMut() -> bool) {
    let context = EditCtx::local("local", "2026-10-16T00:00:00Z");
    local
        .replace_range(&context, StoryRange::new("note", 0, 4), "Memo")
        .unwrap();
    assert!(undo());
    local
        .insert_text(
            &context,
            Position::new("body", 4),
            " text",
            FormatPolicy::Inherit,
        )
        .unwrap();

    let sent = Arc::new(Mutex::new(Vec::new()));
    let sink = Arc::clone(&sent);
    let _subscription = peer
        .observe_update_v1(move |update, _| sink.lock().unwrap().push(update.to_vec()))
        .unwrap();
    peer.delete_story("note").unwrap();
    for update in sent.lock().unwrap().iter() {
        local.apply_update_v1(update).unwrap();
    }
}

fn assert_readable(local: &EditingDoc, peer: &EditingDoc) {
    let state = local.encode_state_as_update_v1();
    peer.apply_update_v1(&state).unwrap();
    assert_eq!(local.list_stories().stories, peer.list_stories().stories);
}

#[test]
fn a_peer_deleting_a_story_edited_and_undone_before_typing_elsewhere_leaves_state_readable() {
    let local = EngineSession::new(202);
    let peer = seeded_peer(local.doc());
    let history = UndoSession::new();
    history.track(local.doc());
    lose_an_undone_story(local.doc(), &peer, || {
        history.add_undo_barrier();
        history.undo()
    });
    assert_readable(local.doc(), &peer);
}

#[test]
fn histories_attached_after_a_peer_deletes_an_undone_story_leave_state_readable() {
    let local = EngineSession::new(202);
    let peer = seeded_peer(local.doc());
    let history = UndoSession::new();
    history.track(local.doc());
    lose_an_undone_story(local.doc(), &peer, || {
        history.add_undo_barrier();
        history.undo()
    });
    let _manager = local.doc().undo_manager();
    UndoSession::new().track(local.doc());
    assert_readable(local.doc(), &peer);
    assert!(history.undo());
    assert_readable(local.doc(), &peer);
}

#[test]
fn every_way_of_attaching_undo_to_a_new_replica_keeps_what_the_history_points_into() {
    for session in [false, true] {
        let local = EditingDoc::new(202);
        let peer = seeded_peer(&local);
        if session {
            let history = UndoSession::new();
            history.track(&local);
            lose_an_undone_story(&local, &peer, || {
                history.add_undo_barrier();
                history.undo()
            });
        } else {
            let mut manager = local.undo_manager();
            lose_an_undone_story(&local, &peer, || {
                manager.add_undo_barrier();
                manager.undo()
            });
        }
        assert_readable(&local, &peer);
    }
}

fn raw_undo_restores_a_deletion(keep_deleted: bool) -> bool {
    let doc = EditingDoc::new(404);
    doc.create_story("body", "Body", "Normal", "left").unwrap();
    let context = EditCtx::local("local", "2026-10-16T00:00:00Z");
    doc.insert_text(
        &context,
        Position::new("body", 0),
        "Hello world",
        FormatPolicy::Inherit,
    )
    .unwrap();
    let before = doc.story_len("body").unwrap();
    if keep_deleted {
        doc.keep_deleted();
    }
    let mut options = yrs::undo::Options::default();
    options
        .tracked_origins
        .insert(yrs::Origin::from(doc.client_id()));
    let mut undo = yrs::undo::UndoManager::<()>::with_options(options);
    let stories = doc.yrs_doc().get_or_insert_map("stories");
    undo.expand_scope(doc.yrs_doc(), &stories);
    doc.delete_range(&context, StoryRange::new("body", 0, 5))
        .unwrap();
    undo.undo_blocking();
    doc.story_len("body").unwrap() == before
}

#[test]
fn raw_undo_managers_restore_deletions_once_the_replica_keeps_deleted_content() {
    assert!((0..32).all(|_| raw_undo_restores_a_deletion(true)));
}
