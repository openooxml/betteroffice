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

thread_local! {
    static LIVE: Cell<usize> = const { Cell::new(0) };
    static BUDGET: Cell<usize> = const { Cell::new(DEFAULT_BUDGET) };
    static OVER: Cell<bool> = const { Cell::new(false) };
    static FAILED: Cell<usize> = const { Cell::new(0) };
}

struct CountingAllocator;

fn grew(bytes: usize) {
    let _ = LIVE.try_with(|live| {
        let next = live.get().saturating_add(bytes);
        live.set(next);
        if next > BUDGET.try_with(Cell::get).unwrap_or(DEFAULT_BUDGET)
            && !OVER.try_with(|over| over.replace(true)).unwrap_or(true)
        {
            let _ = FAILED.try_with(|failed| failed.set(bytes.max(1)));
        }
    });
}

fn shrank(bytes: usize) {
    let _ = LIVE.try_with(|live| {
        let next = live.get().saturating_sub(bytes);
        live.set(next);
        if next <= BUDGET.try_with(Cell::get).unwrap_or(DEFAULT_BUDGET) {
            let _ = OVER.try_with(|over| over.set(false));
        }
    });
}

/// Whether the calling thread's engine memory is past its budget.
pub(crate) fn over_budget() -> bool {
    OVER.try_with(Cell::get).unwrap_or(false)
}

/// Sets the calling thread's budget in bytes; 0 restores the default.
pub(crate) fn set_budget(bytes: usize) {
    let budget = if bytes == 0 { DEFAULT_BUDGET } else { bytes };
    let _ = BUDGET.try_with(|slot| slot.set(budget));
    if LIVE.try_with(Cell::get).unwrap_or(0) <= budget {
        let _ = OVER.try_with(|over| over.set(false));
    }
}

/// The size of the allocation that took the calling thread past its budget, once.
pub(crate) fn take_failed() -> usize {
    FAILED.try_with(|failed| failed.replace(0)).unwrap_or(0)
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
