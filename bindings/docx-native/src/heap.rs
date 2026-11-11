//! Per-thread heap accounting for the native engine: the live and peak bytes the calling thread
//! holds, as each WebAssembly instance counts its own heap. Counting is off until [`enable`] is
//! called, after which it stays on for the life of the process, or while any thread has a heap
//! limit. While it is off an allocation does what the inner allocator does plus one relaxed load.
//! A block freed on a thread other than the one that allocated it, or allocated before counting
//! started, is subtracted on the freeing thread, saturating at zero.
//!
//! A limit never refuses an allocation: the one that takes its thread past the limit raises the
//! thread's over-limit flag, and the engine call fails at its next checkpoint.

use std::alloc::{GlobalAlloc, Layout};
use std::cell::Cell;
use std::sync::Mutex;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

#[cfg(feature = "mimalloc")]
const INNER: mimalloc::MiMalloc = mimalloc::MiMalloc;
#[cfg(not(feature = "mimalloc"))]
const INNER: std::alloc::System = std::alloc::System;

/// The start of the message of an engine call that ran past its thread's limit.
pub(crate) const EXCEEDED: &str = "BetterOffice engine memory budget exceeded";

const NO_LIMIT: usize = usize::MAX;

static COUNTING: AtomicBool = AtomicBool::new(false);

/// How many times counting has switched on from off; changed only under [`SWITCH`].
static COUNT_EPOCH: AtomicU64 = AtomicU64::new(0);

/// Whether a host asked for counting, and how many threads hold a limit.
static SWITCH: Mutex<(bool, usize)> = Mutex::new((false, 0));

thread_local! {
    static LIVE: Cell<usize> = const { Cell::new(0) };
    static PEAK: Cell<usize> = const { Cell::new(0) };
    static FAILED: Cell<usize> = const { Cell::new(0) };
    static LIMIT: Cell<usize> = const { Cell::new(NO_LIMIT) };
    static OVER: Cell<bool> = const { Cell::new(false) };
    static REPORTED: Cell<bool> = const { Cell::new(false) };
    /// The counting epoch the thread's live and peak bytes belong to.
    static LIVE_EPOCH: Cell<u64> = const { Cell::new(0) };
    static SHARE: LimitShare = const { LimitShare };
}

/// Gives up an exiting thread's share in keeping counting on, if it still holds a limit.
struct LimitShare;

impl Drop for LimitShare {
    fn drop(&mut self) {
        if limit().is_some() {
            set_limit(None);
        }
    }
}

pub(crate) struct CountingAllocator;

/// Starts counting on every thread. Idempotent; counting cannot be stopped again.
pub(crate) fn enable() -> bool {
    let mut switch = SWITCH.lock().unwrap_or_else(|error| error.into_inner());
    switch.0 = true;
    switch_counting(true);
    true
}

/// Stores whether counting is on, starting a new epoch when it switches on from off. Called with
/// [`SWITCH`] held.
fn switch_counting(counting: bool) {
    if counting && !COUNTING.load(Ordering::Relaxed) {
        COUNT_EPOCH.fetch_add(1, Ordering::Relaxed);
    }
    COUNTING.store(counting, Ordering::Relaxed);
}

/// Makes the calling thread's live and peak bytes belong to the current epoch. Bytes from an
/// earlier epoch missed the frees made while counting was off, so they restart from zero; bytes
/// already counted in the current epoch are kept.
fn reconcile() {
    let current = COUNT_EPOCH.load(Ordering::Relaxed);
    let _ = LIVE_EPOCH.try_with(|epoch| {
        if epoch.replace(current) != current {
            let _ = LIVE.try_with(|live| live.set(0));
            let _ = PEAK.try_with(|peak| peak.set(0));
        }
    });
}

/// Whether counting is on.
pub(crate) fn enabled() -> bool {
    COUNTING.load(Ordering::Relaxed)
}

/// Sets the calling thread's limit on live bytes; `None` removes it. Counting is on while any
/// thread has a limit, and a thread that exits with one gives it up. A thread taking a limit holds
/// the bytes counted since counting last switched on.
pub(crate) fn set_limit(bytes: Option<usize>) {
    let limit = bytes.unwrap_or(NO_LIMIT);
    let has = limit != NO_LIMIT;
    if has {
        let _ = SHARE.try_with(|_| ());
    }
    let had = LIMIT.try_with(|slot| slot.replace(limit)).unwrap_or(NO_LIMIT) != NO_LIMIT;
    let _ = OVER.try_with(|over| over.set(false));
    let _ = REPORTED.try_with(|reported| reported.set(false));
    let mut switch = SWITCH.lock().unwrap_or_else(|error| error.into_inner());
    match (had, has) {
        (false, true) => switch.1 += 1,
        (true, false) => switch.1 = switch.1.saturating_sub(1),
        _ => {}
    }
    switch_counting(switch.0 || switch.1 > 0);
    if has && !had {
        reconcile();
    }
}

/// The calling thread's limit.
pub(crate) fn limit() -> Option<usize> {
    Some(LIMIT.try_with(Cell::get).unwrap_or(NO_LIMIT)).filter(|limit| *limit != NO_LIMIT)
}

/// Whether an allocation took the calling thread past its limit since its current engine call
/// began; later frees do not lower it.
pub(crate) fn over_limit() -> bool {
    OVER.try_with(Cell::get).unwrap_or(false)
}

/// Starts a limit-checked engine call; false when it must fail at once. An outermost call that
/// starts within the limit forgets a crossing an earlier call already failed with; a crossing no
/// call reported, such as one made converting a return value after its call's last check, fails
/// the next call. A nested call inherits any crossing.
pub(crate) fn begin_call(outermost: bool) -> bool {
    if !over_limit() {
        return true;
    }
    if !outermost
        || !REPORTED.try_with(Cell::get).unwrap_or(false)
        || limit().is_some_and(|limit| live() > limit)
    {
        return false;
    }
    let _ = OVER.try_with(|over| over.set(false));
    let _ = REPORTED.try_with(|reported| reported.set(false));
    true
}

/// Forgets a crossing made by a call that started within the limit and is refusing without doing
/// engine work, so the refusal is not reported as running out of memory.
pub(crate) fn forget_crossing() {
    let _ = OVER.try_with(|over| over.set(false));
}

/// The error message of a call that ran past the limit, with the allocation that crossed it. The
/// crossing counts as reported.
pub(crate) fn exceeded() -> String {
    let _ = REPORTED.try_with(|reported| reported.set(true));
    let bytes = failed_bytes().max(1);
    let limit = LIMIT.try_with(Cell::get).unwrap_or(NO_LIMIT);
    format!("{EXCEEDED} ({bytes} bytes over a {limit}-byte budget)")
}

#[cold]
#[inline(never)]
fn grew(bytes: usize) {
    reconcile();
    let _ = LIVE.try_with(|live| {
        let next = live.get().saturating_add(bytes);
        live.set(next);
        let _ = PEAK.try_with(|peak| peak.set(peak.get().max(next)));
        if next > LIMIT.try_with(Cell::get).unwrap_or(NO_LIMIT)
            && !OVER.try_with(|over| over.replace(true)).unwrap_or(true)
        {
            let _ = REPORTED.try_with(|reported| reported.set(false));
            failed(bytes.max(1));
        }
    });
}

#[cold]
#[inline(never)]
fn shrank(bytes: usize) {
    reconcile();
    let _ = LIVE.try_with(|live| live.set(live.get().saturating_sub(bytes)));
}

#[cold]
#[inline(never)]
fn failed(bytes: usize) {
    let _ = FAILED.try_with(|failed| failed.set(bytes));
}

#[cold]
#[inline(never)]
fn counted(pointer: *mut u8, bytes: usize) {
    if pointer.is_null() {
        failed(bytes);
    } else {
        grew(bytes);
    }
}

#[cold]
#[inline(never)]
fn resized(moved: *mut u8, old: usize, new: usize) {
    if moved.is_null() {
        failed(new);
    } else if new >= old {
        grew(new - old);
    } else {
        shrank(old - new);
    }
}

/// The calling thread's live bytes.
pub(crate) fn live() -> usize {
    reconcile();
    LIVE.try_with(Cell::get).unwrap_or(0)
}

/// The most bytes the calling thread held at once since it started or its last reset.
pub(crate) fn peak() -> usize {
    reconcile();
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
    #[inline]
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { INNER.alloc(layout) };
        if COUNTING.load(Ordering::Relaxed) {
            counted(pointer, layout.size());
        }
        pointer
    }

    #[inline]
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { INNER.alloc_zeroed(layout) };
        if COUNTING.load(Ordering::Relaxed) {
            counted(pointer, layout.size());
        }
        pointer
    }

    #[inline]
    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { INNER.dealloc(pointer, layout) };
        if COUNTING.load(Ordering::Relaxed) {
            shrank(layout.size());
        }
    }

    #[inline]
    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let moved = unsafe { INNER.realloc(pointer, layout, new_size) };
        if COUNTING.load(Ordering::Relaxed) {
            resized(moved, layout.size(), new_size);
        }
        moved
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::MutexGuard;

    static LIMITS: std::sync::Mutex<()> = std::sync::Mutex::new(());

    /// Serialises every test that reads the counters or changes counting, limits or the epoch.
    fn limits() -> MutexGuard<'static, ()> {
        LIMITS.lock().unwrap_or_else(|error| error.into_inner())
    }

    #[test]
    fn counts_allocations_reallocations_and_frees() {
        let _limits = limits();
        enable();
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
        let _limits = limits();
        enable();
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
        let _limits = limits();
        enable();
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
        let _limits = limits();
        enable();
        let block = vec![1u8; 1 << 20];
        let (other_live, other_peak) = std::thread::spawn(|| (live(), peak())).join().unwrap();
        assert!(other_live < 1 << 20);
        assert!(other_peak < 1 << 20);
        assert!(live() >= 1 << 20);
        drop(block);
    }

    #[test]
    fn a_free_on_another_thread_subtracts_there_saturating_at_zero() {
        let _limits = limits();
        enable();
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

    #[test]
    fn a_block_from_before_counting_started_frees_saturating_at_zero() {
        let _limits = limits();
        let layout = Layout::from_size_align(1 << 20, 8).unwrap();
        let uncounted = unsafe { INNER.alloc(layout) } as usize;
        assert_ne!(uncounted, 0);
        let live_after_free = std::thread::spawn(move || {
            enable();
            let before = live();
            unsafe { CountingAllocator.dealloc(uncounted as *mut u8, layout) };
            (before, live())
        })
        .join()
        .unwrap();
        assert!(live_after_free.0 < 1 << 20);
        assert_eq!(live_after_free.1, 0);
    }

    #[test]
    fn enabling_from_any_thread_turns_counting_on_everywhere_and_stays_on() {
        let _limits = limits();
        assert!(std::thread::spawn(enable).join().unwrap());
        assert!(enabled());
        assert!(enable());
        assert!(enabled());
        let before = live();
        let block = vec![1u8; 4096];
        assert!(live() >= before + 4096);
        drop(block);
    }

    #[test]
    fn crossing_the_limit_flags_the_thread_and_keeps_the_allocation() {
        let _limits = limits();
        set_limit(Some(1 << 40));
        let cap = live() + 1024;
        set_limit(Some(cap));
        let layout = Layout::from_size_align(4096, 8).unwrap();
        unsafe {
            let pointer = CountingAllocator.alloc(layout);
            assert!(!pointer.is_null());
            assert!(over_limit());
            assert_eq!(failed_bytes(), 4096);
            CountingAllocator.dealloc(pointer, layout);
        }
        assert!(over_limit());
        assert_eq!(
            exceeded(),
            format!("{EXCEEDED} (4096 bytes over a {cap}-byte budget)")
        );
        assert!(!begin_call(false));
        assert!(begin_call(true));
        assert!(!over_limit());
        set_limit(None);
        assert_eq!(limit(), None);
    }

    #[test]
    fn an_outermost_call_still_over_the_limit_is_refused() {
        let _limits = limits();
        set_limit(Some(1 << 40));
        let block = vec![1u8; 4096];
        set_limit(Some(live() - 1024));
        let more = vec![1u8; 64];
        assert!(over_limit());
        let _ = exceeded();
        assert!(!begin_call(true));
        drop(block);
        drop(more);
        assert!(begin_call(true));
        assert!(!over_limit());
        set_limit(None);
    }

    #[test]
    fn a_crossing_no_call_reported_fails_the_next_call() {
        let _limits = limits();
        set_limit(Some(1 << 40));
        set_limit(Some(live() + 1024));
        let block = vec![1u8; 4096];
        assert!(over_limit());
        drop(block);
        assert!(!begin_call(true));
        let _ = exceeded();
        assert!(begin_call(true));
        assert!(!over_limit());
        set_limit(None);
    }

    #[test]
    fn clearing_the_limit_forgets_an_unreported_crossing() {
        let _limits = limits();
        set_limit(Some(0));
        let block = vec![1u8; 16];
        assert!(over_limit());
        set_limit(None);
        assert!(begin_call(true));
        drop(block);
    }

    /// Starts a new counting epoch, as counting switching on from off does. Other tests may have
    /// turned counting on for good, so the switch itself cannot always happen here.
    fn restart_counting() {
        let _switch = SWITCH.lock().unwrap_or_else(|error| error.into_inner());
        COUNT_EPOCH.fetch_add(1, Ordering::Relaxed);
    }

    /// Leaves the calling thread with bytes from an earlier epoch that missed their frees.
    fn stale_bytes() {
        set_limit(Some(1 << 40));
        set_limit(None);
        let _ = LIVE.try_with(|live| live.set(live.get() + (1 << 30)));
    }

    /// Keeps a block allocated in the current epoch, then takes a limit below it.
    fn block_then_limit() -> (usize, bool) {
        let block = vec![1u8; 1 << 16];
        set_limit(Some(1 << 15));
        let held = live();
        let more = vec![1u8; 16];
        let over = over_limit();
        drop(more);
        drop(block);
        set_limit(None);
        (held, over)
    }

    #[test]
    fn bytes_kept_after_counting_restarts_count_toward_a_later_limit() {
        let _limits = limits();
        let (held, over) = std::thread::spawn(|| {
            stale_bytes();
            enable();
            restart_counting();
            block_then_limit()
        })
        .join()
        .unwrap();
        assert!(held >= 1 << 16);
        assert!(held < 1 << 30);
        assert!(over);
    }

    #[test]
    fn bytes_kept_after_another_thread_restarts_counting_count_toward_a_later_limit() {
        use std::sync::mpsc::channel;
        let _limits = limits();
        let (stale_tx, stale_rx) = channel::<()>();
        let (restarted_tx, restarted_rx) = channel::<()>();
        let first = std::thread::spawn(move || {
            stale_bytes();
            stale_tx.send(()).unwrap();
            restarted_rx.recv().unwrap();
            block_then_limit()
        });
        stale_rx.recv().unwrap();
        set_limit(Some(1 << 40));
        restart_counting();
        restarted_tx.send(()).unwrap();
        let (held, over) = first.join().unwrap();
        set_limit(None);
        assert!(held >= 1 << 16);
        assert!(held < 1 << 30);
        assert!(over);
    }

    #[test]
    fn a_reallocation_past_the_limit_flags_the_thread_and_moves_the_block() {
        let _limits = limits();
        set_limit(Some(1 << 40));
        set_limit(Some(live() + 1024));
        let layout = Layout::from_size_align(64, 8).unwrap();
        unsafe {
            let pointer = CountingAllocator.alloc(layout);
            assert!(!over_limit());
            let grown = CountingAllocator.realloc(pointer, layout, 8192);
            assert!(!grown.is_null());
            assert!(over_limit());
            assert_eq!(failed_bytes(), 8192 - 64);
            CountingAllocator.dealloc(grown, Layout::from_size_align(8192, 8).unwrap());
        }
        set_limit(None);
    }

    #[test]
    fn clearing_the_limit_clears_the_flag() {
        let _limits = limits();
        set_limit(Some(0));
        let block = vec![1u8; 16];
        assert!(over_limit());
        set_limit(None);
        assert!(!over_limit());
        assert_eq!(limit(), None);
        drop(block);
    }

    #[test]
    fn without_a_limit_counting_never_flags_the_thread() {
        let _limits = limits();
        enable();
        assert_eq!(limit(), None);
        let block = vec![1u8; 1 << 20];
        assert!(!over_limit());
        assert!(begin_call(true));
        assert_eq!(limit(), None);
        drop(block);
    }

    #[test]
    fn a_limit_starts_counting_and_clearing_it_keeps_counting_a_host_asked_for() {
        let _limits = limits();
        set_limit(Some(1 << 40));
        assert!(enabled());
        enable();
        set_limit(None);
        assert!(enabled());
    }

    #[test]
    fn a_limit_belongs_to_the_thread_that_set_it() {
        let _limits = limits();
        set_limit(Some(0));
        let (other_limit, other_over) = std::thread::spawn(|| {
            let block = vec![1u8; 4096];
            let seen = (limit(), over_limit());
            drop(block);
            seen
        })
        .join()
        .unwrap();
        assert_eq!(other_limit, None);
        assert!(!other_over);
        set_limit(None);
    }

    #[test]
    fn a_thread_that_exits_with_a_limit_gives_up_its_share() {
        let _limits = limits();
        let before = SWITCH.lock().unwrap_or_else(|error| error.into_inner()).1;
        let counted_there = std::thread::spawn(|| {
            set_limit(Some(1 << 40));
            enabled()
        })
        .join()
        .unwrap();
        assert!(counted_there);
        let switch = SWITCH.lock().unwrap_or_else(|error| error.into_inner());
        assert_eq!(switch.1, before);
        assert_eq!(enabled(), switch.0 || switch.1 > 0);
        if !switch.0 && before == 0 {
            assert!(!enabled());
        }
    }

    #[test]
    fn a_limit_set_while_counting_is_on_keeps_the_bytes_already_counted() {
        let _limits = limits();
        enable();
        set_limit(Some(1 << 40));
        set_limit(None);
        let block = vec![1u8; 1 << 16];
        let counted = live();
        assert!(counted >= 1 << 16);
        set_limit(Some(1 << 40));
        assert_eq!(live(), counted);
        set_limit(None);
        drop(block);
    }
}
