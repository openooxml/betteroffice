//! Per-thread live-byte accounting for the native engine. Allocations are never refused: past the
//! budget the thread's over-budget flag goes up, and the engine's progress checkpoints unwind the
//! call so the session fails with an out-of-memory error, as a WebAssembly engine at its memory
//! limit does.

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

/// The largest memory a wasm32 engine can address.
pub(crate) const DEFAULT_BUDGET: usize = 4 << 30;

/// The message of the error a call unwinds with once its thread is over budget.
pub(crate) const EXCEEDED: &str = "BetterOffice engine memory budget exceeded";

struct Counters {
    live: Cell<usize>,
    budget: Cell<usize>,
    over: Cell<bool>,
    failed: Cell<usize>,
}

thread_local! {
    static COUNTERS: Counters = const {
        Counters {
            live: Cell::new(0),
            budget: Cell::new(DEFAULT_BUDGET),
            over: Cell::new(false),
            failed: Cell::new(0),
        }
    };
}

struct CountingAllocator;

fn grew(bytes: usize) {
    let _ = COUNTERS.try_with(|counters| {
        let next = counters.live.get().saturating_add(bytes);
        counters.live.set(next);
        if next > counters.budget.get() && !counters.over.replace(true) {
            counters.failed.set(bytes.max(1));
        }
    });
}

fn shrank(bytes: usize) {
    let _ = COUNTERS.try_with(|counters| {
        let next = counters.live.get().saturating_sub(bytes);
        counters.live.set(next);
        if counters.over.get() && next <= counters.budget.get() {
            counters.over.set(false);
        }
    });
}

/// Whether the calling thread's engine memory is past its budget.
pub(crate) fn over_budget() -> bool {
    COUNTERS.try_with(|counters| counters.over.get()).unwrap_or(false)
}

/// Sets the calling thread's budget in bytes; 0 restores the default.
pub(crate) fn set_budget(bytes: usize) {
    let budget = if bytes == 0 { DEFAULT_BUDGET } else { bytes };
    let _ = COUNTERS.try_with(|counters| {
        counters.budget.set(budget);
        if counters.live.get() <= budget {
            counters.over.set(false);
        }
    });
}

/// The size of the allocation that took the calling thread past its budget, once.
pub(crate) fn take_failed() -> usize {
    COUNTERS
        .try_with(|counters| counters.failed.replace(0))
        .unwrap_or(0)
}

// SAFETY: every call forwards to `System` unchanged; the counters only observe it.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            grew(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if !pointer.is_null() {
            grew(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
        shrank(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let moved = unsafe { System.realloc(pointer, layout, new_size) };
        if !moved.is_null() {
            if new_size >= layout.size() {
                grew(new_size - layout.size());
            } else {
                shrank(layout.size() - new_size);
            }
        }
        moved
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;
