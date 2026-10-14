//! Heap accounting for the wasm build: live and peak bytes, the size of an
//! allocation that could not be satisfied, and an optional limit on live bytes.
//! [`CountingAllocator`] is the global allocator only in a wasm32 build with the
//! `heap-stats` feature.

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};
use wasm_bindgen::prelude::*;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);
static FAILED: AtomicUsize = AtomicUsize::new(0);
static LIMIT: AtomicUsize = AtomicUsize::new(usize::MAX);

/// [`System`] with byte counters.
pub struct CountingAllocator;

fn grew(bytes: usize) {
    let live = LIVE.fetch_add(bytes, Ordering::Relaxed) + bytes;
    PEAK.fetch_max(live, Ordering::Relaxed);
}

fn shrank(bytes: usize) {
    LIVE.fetch_sub(bytes, Ordering::Relaxed);
}

/// Whether `bytes` more would take the live bytes past the limit. The failed
/// size is recorded as for an allocation the memory refused.
fn over_limit(bytes: usize) -> bool {
    let over = LIVE
        .load(Ordering::Relaxed)
        .checked_add(bytes)
        .is_none_or(|live| live > LIMIT.load(Ordering::Relaxed));
    if over {
        FAILED.store(bytes, Ordering::Relaxed);
    }
    over
}

fn counted(pointer: *mut u8, bytes: usize) -> *mut u8 {
    if pointer.is_null() {
        FAILED.store(bytes, Ordering::Relaxed);
    } else {
        grew(bytes);
    }
    pointer
}

// SAFETY: every call forwards to `System` unchanged, or returns null without
// allocating once the limit is reached, which `GlobalAlloc` permits.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        if over_limit(layout.size()) {
            return std::ptr::null_mut();
        }
        counted(unsafe { System.alloc(layout) }, layout.size())
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        if over_limit(layout.size()) {
            return std::ptr::null_mut();
        }
        counted(unsafe { System.alloc_zeroed(layout) }, layout.size())
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
        shrank(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        if new_size > layout.size() && over_limit(new_size - layout.size()) {
            FAILED.store(new_size, Ordering::Relaxed);
            return std::ptr::null_mut();
        }
        let moved = unsafe { System.realloc(pointer, layout, new_size) };
        if moved.is_null() {
            FAILED.store(new_size, Ordering::Relaxed);
        } else if new_size >= layout.size() {
            grew(new_size - layout.size());
        } else {
            shrank(layout.size() - new_size);
        }
        moved
    }
}

#[cfg(all(target_arch = "wasm32", feature = "heap-stats"))]
#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

/// Whether the counters below see every allocation of this module.
#[wasm_bindgen]
pub fn wasm_heap_counted() -> bool {
    cfg!(all(target_arch = "wasm32", feature = "heap-stats"))
}

/// Bytes currently allocated on the Rust heap.
#[wasm_bindgen]
pub fn wasm_live_bytes() -> f64 {
    LIVE.load(Ordering::Relaxed) as f64
}

/// The most bytes allocated at once since the module started or the last reset.
#[wasm_bindgen]
pub fn wasm_peak_bytes() -> f64 {
    PEAK.load(Ordering::Relaxed) as f64
}

/// Starts a new peak window at the current live bytes.
#[wasm_bindgen]
pub fn reset_wasm_peak_bytes() {
    PEAK.store(LIVE.load(Ordering::Relaxed), Ordering::Relaxed);
}

/// Size of the last allocation that failed, or 0. An allocation fails when the
/// linear memory cannot grow or the limit is reached, and the module then aborts.
#[wasm_bindgen]
pub fn wasm_failed_allocation_bytes() -> f64 {
    FAILED.load(Ordering::Relaxed) as f64
}

/// Caps the bytes allocated at once; an allocation past the cap fails. A
/// non-finite or negative value removes the cap.
#[wasm_bindgen]
pub fn set_wasm_heap_limit(bytes: f64) {
    let limit = if bytes.is_finite() && bytes >= 0.0 {
        bytes as usize
    } else {
        usize::MAX
    };
    LIMIT.store(limit, Ordering::Relaxed);
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Mutex;

    static COUNTERS: Mutex<()> = Mutex::new(());

    #[test]
    fn counts_allocations_reallocations_and_frees() {
        let _counters = COUNTERS.lock().unwrap();
        let before = LIVE.load(Ordering::Relaxed);
        let layout = Layout::from_size_align(1024, 8).unwrap();
        unsafe {
            let pointer = CountingAllocator.alloc(layout);
            assert!(!pointer.is_null());
            assert_eq!(LIVE.load(Ordering::Relaxed), before + 1024);
            let grown = CountingAllocator.realloc(pointer, layout, 4096);
            assert_eq!(LIVE.load(Ordering::Relaxed), before + 4096);
            assert!(PEAK.load(Ordering::Relaxed) >= before + 4096);
            let grown_layout = Layout::from_size_align(4096, 8).unwrap();
            let shrunk = CountingAllocator.realloc(grown, grown_layout, 512);
            assert_eq!(LIVE.load(Ordering::Relaxed), before + 512);
            CountingAllocator.dealloc(shrunk, Layout::from_size_align(512, 8).unwrap());
        }
        assert_eq!(LIVE.load(Ordering::Relaxed), before);
        reset_wasm_peak_bytes();
        assert_eq!(PEAK.load(Ordering::Relaxed), before);
    }

    #[test]
    fn counts_zeroed_allocations() {
        let _counters = COUNTERS.lock().unwrap();
        let before = LIVE.load(Ordering::Relaxed);
        let layout = Layout::from_size_align(256, 8).unwrap();
        unsafe {
            let pointer = CountingAllocator.alloc_zeroed(layout);
            assert!(
                std::slice::from_raw_parts(pointer, 256)
                    .iter()
                    .all(|byte| *byte == 0)
            );
            assert_eq!(LIVE.load(Ordering::Relaxed), before + 256);
            CountingAllocator.dealloc(pointer, layout);
        }
        assert_eq!(LIVE.load(Ordering::Relaxed), before);
    }

    #[test]
    fn a_failed_reallocation_keeps_the_block_and_records_its_size() {
        let _counters = COUNTERS.lock().unwrap();
        let before = LIVE.load(Ordering::Relaxed);
        let layout = Layout::from_size_align(64, 8).unwrap();
        let huge = isize::MAX as usize - 4096;
        unsafe {
            let pointer = CountingAllocator.alloc(layout);
            assert!(CountingAllocator.realloc(pointer, layout, huge).is_null());
            assert_eq!(LIVE.load(Ordering::Relaxed), before + 64);
            assert_eq!(wasm_failed_allocation_bytes(), huge as f64);
            CountingAllocator.dealloc(pointer, layout);
        }
        assert_eq!(LIVE.load(Ordering::Relaxed), before);
    }

    #[test]
    fn records_the_size_of_a_failed_allocation() {
        let _counters = COUNTERS.lock().unwrap();
        assert!(counted(std::ptr::null_mut(), 77).is_null());
        assert_eq!(wasm_failed_allocation_bytes(), 77.0);
    }

    #[test]
    fn an_allocation_past_the_limit_fails_without_allocating() {
        let _counters = COUNTERS.lock().unwrap();
        let before = LIVE.load(Ordering::Relaxed);
        let layout = Layout::from_size_align(64, 8).unwrap();
        set_wasm_heap_limit((before + 100) as f64);
        unsafe {
            let pointer = CountingAllocator.alloc(layout);
            assert!(!pointer.is_null());
            assert!(CountingAllocator.alloc(layout).is_null());
            assert!(CountingAllocator.alloc_zeroed(layout).is_null());
            assert_eq!(wasm_failed_allocation_bytes(), 64.0);
            assert!(CountingAllocator.realloc(pointer, layout, 128).is_null());
            assert_eq!(wasm_failed_allocation_bytes(), 128.0);
            assert_eq!(LIVE.load(Ordering::Relaxed), before + 64);
            let shrunk = CountingAllocator.realloc(pointer, layout, 32);
            assert!(!shrunk.is_null());
            set_wasm_heap_limit(f64::INFINITY);
            let grown =
                CountingAllocator.realloc(shrunk, Layout::from_size_align(32, 8).unwrap(), 256);
            assert!(!grown.is_null());
            CountingAllocator.dealloc(grown, Layout::from_size_align(256, 8).unwrap());
        }
        assert_eq!(LIVE.load(Ordering::Relaxed), before);
        FAILED.store(0, Ordering::Relaxed);
    }

    #[test]
    fn native_builds_do_not_count_the_heap() {
        assert!(!wasm_heap_counted());
    }
}
