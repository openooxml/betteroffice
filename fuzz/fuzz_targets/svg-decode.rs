#![no_main]

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};

use libfuzzer_sys::fuzz_target;
use pptx_raster::SVG_MEMORY_ENVELOPE;
use pptx_raster::fuzzing::decode;

/// Image budget the slide has left: a 4 MiB raster and its layers, outside
/// the envelope, which the decoder supersamples within and refuses past, so
/// the envelope and libFuzzer's in-memory corpus fit a 150 MB RSS cap.
const PIXELS: u64 = 1 << 20;
/// The thread stack the sandbox's nesting bound is measured against.
const STACK: usize = 512 * 1024;
/// A decode's peak heap: the envelope plus the raster it outputs.
const PEAK: usize = SVG_MEMORY_ENVELOPE as usize + PIXELS as usize * 4;

/// Counts live heap bytes, a malloc header's worth over each request, so the
/// target measures the decode rather than libFuzzer's corpus. An allocation
/// that would take a decode past [`PEAK`] fails, which aborts the run there
/// and then, before the machine ever holds more.
struct Counting;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static HIGH: AtomicUsize = AtomicUsize::new(0);
static BASE: AtomicUsize = AtomicUsize::new(0);

fn charged(layout: Layout) -> usize {
    layout.size().next_multiple_of(16) + 16
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let live = LIVE.fetch_add(charged(layout), Ordering::Relaxed) + charged(layout);
        if live.saturating_sub(BASE.load(Ordering::Relaxed)) > PEAK {
            LIVE.fetch_sub(charged(layout), Ordering::Relaxed);
            return std::ptr::null_mut();
        }
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            LIVE.fetch_sub(charged(layout), Ordering::Relaxed);
        } else {
            HIGH.fetch_max(live, Ordering::Relaxed);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
        LIVE.fetch_sub(charged(layout), Ordering::Relaxed);
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

fuzz_target!(|data: &[u8]| {
    let data = data.to_vec();
    let before = LIVE.load(Ordering::Relaxed);
    BASE.store(before, Ordering::Relaxed);
    HIGH.store(before, Ordering::Relaxed);
    std::thread::Builder::new()
        .stack_size(STACK)
        .spawn(move || decode(&data, PIXELS))
        .expect("spawn")
        .join()
        .expect("the decode panicked past its guard");
    let used = HIGH.load(Ordering::Relaxed) - before;
    assert!(used <= PEAK, "the decode's heap peaked at {used} bytes");
});
