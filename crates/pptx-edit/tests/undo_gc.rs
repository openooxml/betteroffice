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

#[test]
fn a_peer_removing_a_shape_this_replica_edited_undid_and_edited_again_leaves_its_state_readable() {
    let source = DeckSession::open(FIXTURE, 101).unwrap();
    let slide_id = source.snapshot().unwrap().slides[0].id.clone();
    let shape_id = source
        .add_text_box(
            &EditCtx::local("source"),
            &slide_id,
            &ShapeDraft {
                name: "Note".to_owned(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 200_000,
                    width: 3_000_000,
                    height: 800_000,
                },
                text: "Note".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap()
        .shape_id;
    let seed = source.encode_state_as_update_v1();
    let local = DeckSession::open_from_update(&seed, 202).unwrap();
    let peer = DeckSession::open_from_update(&seed, 303).unwrap();

    let context = EditCtx::local("local");
    local
        .move_shape(&context, &slide_id, &shape_id, 400_000, 300_000)
        .unwrap();
    local.add_undo_barrier();
    assert!(local.undo());
    local
        .move_shape(&context, &slide_id, &shape_id, 800_000, 300_000)
        .unwrap();

    let sent = Arc::new(Mutex::new(Vec::new()));
    let sink = Arc::clone(&sent);
    let _subscription = peer
        .observe_update_v1(move |event| sink.lock().unwrap().push(event.update))
        .unwrap();
    peer.remove_shape(&EditCtx::local("peer"), &slide_id, &shape_id)
        .unwrap();
    for update in sent.lock().unwrap().iter() {
        local.apply_update_v1(update).unwrap();
    }

    let state = local.encode_state_as_update_v1();
    peer.apply_update_v1(&state).unwrap();
    assert_eq!(local.snapshot().unwrap(), peer.snapshot().unwrap());
}
