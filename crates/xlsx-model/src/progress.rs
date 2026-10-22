//! Liveness reporting for long synchronous workbook calls.
//!
//! Parsing and recalculation call [`tick`] once per unit of work (a cell read,
//! a formula evaluated); every [`EVERY`] ticks the registered hook runs, so a
//! host blocked in one call can report that it is still alive.

use std::cell::Cell;

/// Ticks between two hook calls.
pub const EVERY: u32 = 4096;

thread_local! {
    static HOOK: Cell<Option<fn()>> = const { Cell::new(None) };
    static TICKS: Cell<u32> = const { Cell::new(0) };
}

/// Registers the function [`tick`] and [`pulse`] call; `None` removes it.
/// The hook must not call back into the document: it can run while a
/// transaction is held.
pub fn set_hook(hook: Option<fn()>) {
    HOOK.with(|slot| slot.set(hook));
}

/// Counts one unit of work and calls the hook every [`EVERY`] units.
pub fn tick() {
    let due = TICKS.with(|count| {
        let next = count.get().wrapping_add(1);
        count.set(next);
        next % EVERY == 0
    });
    if due {
        pulse();
    }
}

/// Counts `n` units of work at once, calling the hook once for each [`EVERY`]
/// boundary the count crosses.
pub fn tick_n(n: u64) {
    let crossed = TICKS.with(|count| {
        let before = u64::from(count.get());
        let after = before.saturating_add(n);
        count.set(after as u32);
        after / u64::from(EVERY) - before / u64::from(EVERY)
    });
    for _ in 0..crossed {
        pulse();
    }
}

/// Calls the hook now, for coarse steps that may each take a while.
pub fn pulse() {
    if let Some(hook) = HOOK.with(Cell::get) {
        hook();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    thread_local! {
        static CALLS: Cell<u32> = const { Cell::new(0) };
    }

    fn count() {
        CALLS.with(|calls| calls.set(calls.get() + 1));
    }

    #[test]
    fn the_hook_runs_once_per_every_ticks_and_on_each_pulse() {
        set_hook(Some(count));
        for _ in 0..EVERY * 3 {
            tick();
        }
        pulse();
        set_hook(None);
        pulse();
        assert_eq!(CALLS.with(Cell::get), 4);
    }

    #[test]
    fn counting_many_units_at_once_runs_the_hook_per_every_crossed() {
        set_hook(Some(count));
        tick_n(u64::from(EVERY) - 1);
        assert_eq!(CALLS.with(Cell::get), 0);
        tick();
        assert_eq!(CALLS.with(Cell::get), 1);
        tick_n(u64::from(EVERY) * 3 + 10);
        assert_eq!(CALLS.with(Cell::get), 4);
        tick_n(0);
        set_hook(None);
        assert_eq!(CALLS.with(Cell::get), 4);
    }
}
