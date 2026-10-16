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

#[test]
fn a_peer_deleting_a_story_edited_and_undone_before_typing_elsewhere_leaves_state_readable() {
    let source = EditingDoc::new(101);
    source
        .create_story("body", "Body", "Normal", "left")
        .unwrap();
    source
        .create_story("note", "Note", "Normal", "left")
        .unwrap();
    let seed = source.encode_state_as_update_v1();
    let local = EngineSession::new(202);
    local.doc().apply_update_v1(&seed).unwrap();
    let peer = EditingDoc::new(303);
    peer.apply_update_v1(&seed).unwrap();

    let history = UndoSession::new();
    history.track(local.doc());
    let context = EditCtx::local("local", "2026-10-16T00:00:00Z");
    local
        .doc()
        .replace_range(&context, StoryRange::new("note", 0, 4), "Memo")
        .unwrap();
    history.add_undo_barrier();
    assert!(history.undo());
    local
        .doc()
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
        local.doc().apply_update_v1(update).unwrap();
    }

    let state = local.doc().encode_state_as_update_v1();
    peer.apply_update_v1(&state).unwrap();
    assert_eq!(
        local.doc().list_stories().stories,
        peer.list_stories().stories
    );
}
