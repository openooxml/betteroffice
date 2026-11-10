//! Per-thread heap accounting for the native engine: the live and peak bytes the calling thread
//! holds, as each WebAssembly instance counts its own heap. A block freed on a thread other than
//! the one that allocated it is subtracted on the freeing thread, saturating at zero, so the
//! allocating thread keeps counting it as live.

use std::alloc::{GlobalAlloc, Layout};
use std::cell::Cell;

#[cfg(feature = "mimalloc")]
const INNER: mimalloc::MiMalloc = mimalloc::MiMalloc;
#[cfg(not(feature = "mimalloc"))]
const INNER: std::alloc::System = std::alloc::System;

thread_local! {
    static LIVE: Cell<usize> = const { Cell::new(0) };
    static PEAK: Cell<usize> = const { Cell::new(0) };
    static FAILED: Cell<usize> = const { Cell::new(0) };
}

pub(crate) struct CountingAllocator;

fn grew(bytes: usize) {
    let _ = LIVE.try_with(|live| {
        let next = live.get().saturating_add(bytes);
        live.set(next);
        let _ = PEAK.try_with(|peak| peak.set(peak.get().max(next)));
    });
}

fn shrank(bytes: usize) {
    let _ = LIVE.try_with(|live| live.set(live.get().saturating_sub(bytes)));
}

fn failed(bytes: usize) {
    let _ = FAILED.try_with(|failed| failed.set(bytes));
}

fn counted(pointer: *mut u8, bytes: usize) -> *mut u8 {
    if pointer.is_null() {
        failed(bytes);
    } else {
        grew(bytes);
    }
    pointer
}

/// The calling thread's live bytes.
pub(crate) fn live() -> usize {
    LIVE.try_with(Cell::get).unwrap_or(0)
}

/// The most bytes the calling thread held at once since it started or its last reset.
pub(crate) fn peak() -> usize {
    PEAK.try_with(Cell::get).unwrap_or(0)
}

/// Size of the calling thread's last allocation that failed, or 0.
pub(crate) fn failed_bytes() -> usize {
    FAILED.try_with(Cell::get).unwrap_or(0)
}

/// Starts a new peak window at the calling thread's live bytes.
pub(crate) fn reset_peak() {
    let live = live();
    let _ = PEAK.try_with(|peak| peak.set(live));
}

// SAFETY: every call forwards to the inner allocator unchanged; the counters only observe it.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        counted(unsafe { INNER.alloc(layout) }, layout.size())
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        counted(unsafe { INNER.alloc_zeroed(layout) }, layout.size())
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { INNER.dealloc(pointer, layout) };
        shrank(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let moved = unsafe { INNER.realloc(pointer, layout, new_size) };
        if moved.is_null() {
            failed(new_size);
        } else if new_size >= layout.size() {
            grew(new_size - layout.size());
        } else {
            shrank(layout.size() - new_size);
        }
        moved
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn counts_allocations_reallocations_and_frees() {
        let before = live();
        let layout = Layout::from_size_align(1024, 8).unwrap();
        unsafe {
            let pointer = CountingAllocator.alloc(layout);
            assert!(!pointer.is_null());
            assert_eq!(live(), before + 1024);
            let grown = CountingAllocator.realloc(pointer, layout, 4096);
            assert_eq!(live(), before + 4096);
            assert!(peak() >= before + 4096);
            let shrunk =
                CountingAllocator.realloc(grown, Layout::from_size_align(4096, 8).unwrap(), 512);
            assert_eq!(live(), before + 512);
            CountingAllocator.dealloc(shrunk, Layout::from_size_align(512, 8).unwrap());
        }
        assert_eq!(live(), before);
        reset_peak();
        assert_eq!(peak(), before);
    }

    #[test]
    fn counts_zeroed_allocations() {
        let before = live();
        let layout = Layout::from_size_align(256, 8).unwrap();
        unsafe {
            let pointer = CountingAllocator.alloc_zeroed(layout);
            assert!(
                std::slice::from_raw_parts(pointer, 256)
                    .iter()
                    .all(|byte| *byte == 0)
            );
            assert_eq!(live(), before + 256);
            CountingAllocator.dealloc(pointer, layout);
        }
        assert_eq!(live(), before);
    }

    #[test]
    fn a_failed_reallocation_keeps_the_block_and_records_its_size() {
        let before = live();
        let layout = Layout::from_size_align(64, 8).unwrap();
        let huge = isize::MAX as usize - 4096;
        unsafe {
            let pointer = CountingAllocator.alloc(layout);
            assert!(CountingAllocator.realloc(pointer, layout, huge).is_null());
            assert_eq!(live(), before + 64);
            assert_eq!(failed_bytes(), huge);
            CountingAllocator.dealloc(pointer, layout);
        }
        assert_eq!(live(), before);
    }

    #[test]
    fn counters_belong_to_the_calling_thread() {
        let block = vec![1u8; 1 << 20];
        let (other_live, other_peak) = std::thread::spawn(|| (live(), peak()))
            .join()
            .unwrap();
        assert!(other_live < 1 << 20);
        assert!(other_peak < 1 << 20);
        assert!(live() >= 1 << 20);
        drop(block);
    }

    #[test]
    fn a_free_on_another_thread_subtracts_there_saturating_at_zero() {
        let block = vec![1u8; 1 << 20];
        let allocated = live();
        assert!(allocated >= 1 << 20);
        let freed_there = std::thread::spawn(move || {
            drop(block);
            live()
        })
        .join()
        .unwrap();
        assert_eq!(freed_there, 0);
        assert!(live() + 4096 >= allocated);
    }
}
