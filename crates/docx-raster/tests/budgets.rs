//! Per-page render budget probes, measured by allocation rather than by clock.

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

use docx_layout::display_list::DisplayList;
use docx_raster::{
    FontChains, ImageMap, MAX_PAGE_DIM, MAX_PAGE_PIXELS, RenderResources, render_png,
};
use ooxml_text::FontStore;
use serde_json::{Value, json};

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

thread_local! {
    static ALLOCATED: Cell<u64> = const { Cell::new(0) };
}

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let _ = ALLOCATED.try_with(|counter| counter.set(counter.get() + layout.size() as u64));
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) }
    }
}

/// Bytes this thread allocated while `body` ran. Tests share the process, so
/// the counter is per thread.
fn allocated_by<T>(body: impl FnOnce() -> T) -> (T, u64) {
    let before = ALLOCATED.with(Cell::get);
    let value = body();
    (value, ALLOCATED.with(Cell::get) - before)
}

const MIB: u64 = 1 << 20;

fn list(width: f64, height: f64, primitives: Vec<Value>) -> DisplayList {
    serde_json::from_value(json!({
        "pages": [{
            "pageIndex": 0,
            "width": width,
            "height": height,
            "primitives": primitives
        }]
    }))
    .expect("display list")
}

/// The published entry point is what a direct consumer of this crate calls, so
/// the surface cap has to live inside it rather than in a facade above it.
#[test]
fn the_published_entry_point_caps_the_page_surface_it_allocates() {
    const { assert!(4096_u64 * 4097 > MAX_PAGE_PIXELS) };
    let (fonts, chains, images) = (FontStore::new(), FontChains::new(), ImageMap::new());
    let resources = RenderResources::new(&fonts, &chains, &images);
    for (width, height, message) in [
        (
            f64::from(MAX_PAGE_DIM) + 1.0,
            1.0,
            "page is 16385x1px, past the 16384px per-side cap",
        ),
        (
            4096.0,
            4097.0,
            "page is 4096x4097px, past the 16777216-pixel surface cap",
        ),
    ] {
        let scene = list(width, height, Vec::new());
        let (result, allocated) = allocated_by(|| render_png(&scene, 0, &resources));
        assert_eq!(result.unwrap_err(), message);
        assert!(
            allocated < MIB,
            "{width}x{height} allocated {allocated} bytes before the cap refused it"
        );
    }
}
