use std::alloc::{GlobalAlloc, Layout, System};
use std::collections::BTreeMap;
use std::sync::Mutex;
use std::sync::atomic::{AtomicUsize, Ordering};

use docx_edit::{
    ColorPatch, EditCtx, EditingDoc, FontFamilyPatch, InlineFormatDelta, Patch, StoryRange,
    checksum, project_story, story_checksum,
};
use yrs::Any;

/// Tracks live heap bytes and their high-water mark.
struct Peak;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static HIGH: AtomicUsize = AtomicUsize::new(0);
static SERIAL: Mutex<()> = Mutex::new(());

unsafe impl GlobalAlloc for Peak {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            let live = LIVE.fetch_add(layout.size(), Ordering::Relaxed) + layout.size();
            HIGH.fetch_max(live, Ordering::Relaxed);
        }
        ptr
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) };
        LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
    }
}

#[global_allocator]
static ALLOCATOR: Peak = Peak;

/// One paragraph of `chars` characters in a single run carrying common run
/// formatting, covered by one comment.
fn formatted_story(chars: usize) -> EditingDoc {
    let text: String = "lorem ipsum ".chars().cycle().take(chars).collect();
    let doc = EditingDoc::new(7);
    doc.create_story("body", &text, "Normal", "left").unwrap();
    let ctx = EditCtx::local("Ada", "2026-10-31T00:00:00Z");
    let len = chars as u32;
    doc.format_range(
        &ctx,
        StoryRange::new("body", 0, len),
        &InlineFormatDelta {
            bold: Patch::Set(true),
            italic: Patch::Set(true),
            color: Patch::Set(ColorPatch::Rgb("1F3864".to_owned())),
            highlight: Patch::Set("yellow".to_owned()),
            font_size: Patch::Set(11.0),
            font_family: Patch::Set(FontFamilyPatch {
                ascii: "Calibri".to_owned(),
                h_ansi: Some("Calibri".to_owned()),
            }),
            other: BTreeMap::from([
                ("lang".to_owned(), Some(Any::from("en-GB"))),
                ("kern".to_owned(), Some(Any::Number(2.0))),
            ]),
            ..InlineFormatDelta::default()
        },
    )
    .unwrap();
    doc.add_comment(
        &[StoryRange::new("body", 0, len)],
        "Ada",
        "2026-10-31T00:00:00Z",
        Any::from("note"),
    )
    .unwrap();
    doc
}

#[test]
fn story_checksum_matches_the_projected_stream_on_long_formatted_commented_runs() {
    let _serial = SERIAL.lock().unwrap();
    let doc = formatted_story(5_000);
    assert_eq!(
        story_checksum(&doc, "body").unwrap(),
        checksum(&project_story(&doc, "body").unwrap())
    );
}

#[test]
fn story_checksum_memory_stays_flat_in_story_length() {
    let _serial = SERIAL.lock().unwrap();
    let doc = formatted_story(300_000);
    let base = LIVE.load(Ordering::Relaxed);
    HIGH.store(base, Ordering::Relaxed);
    story_checksum(&doc, "body").unwrap();
    let extra = HIGH.load(Ordering::Relaxed).saturating_sub(base);
    assert!(
        extra < 32 << 20,
        "story_checksum allocated {} MiB above the open document",
        extra >> 20
    );
}
