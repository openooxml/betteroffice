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
}
