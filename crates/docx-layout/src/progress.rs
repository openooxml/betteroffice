//! Liveness reporting for long layout calls, as `ooxml_opc::progress`.
//!
//! Every [`EVERY`] ticks the hook runs. In a wasm32 build with no hook
//! registered, the hook is `globalThis.__betterofficeEngineProgress` when the
//! host installed one.

use std::cell::Cell;

/// Ticks between two hook calls.
pub const EVERY: u32 = 128;

thread_local! {
    static HOOK: Cell<Option<fn()>> = const { Cell::new(None) };
    static TICKS: Cell<u32> = const { Cell::new(0) };
}

/// Registers the function [`tick`] and [`pulse`] call; `None` restores the default.
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

/// Calls the hook now, for coarse units that may each take a while.
pub fn pulse() {
    match HOOK.with(Cell::get) {
        Some(hook) => hook(),
        None => host_progress(),
    }
}

#[cfg(target_arch = "wasm32")]
fn host_progress() {
    use wasm_bindgen::JsCast;
    let hook = js_sys::Reflect::get(&js_sys::global(), &"__betterofficeEngineProgress".into());
    if let Ok(hook) = hook.and_then(|hook| hook.dyn_into::<js_sys::Function>()) {
        let _ = hook.call0(&wasm_bindgen::JsValue::UNDEFINED);
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn host_progress() {}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::{AtomicUsize, Ordering};

    static CALLS: AtomicUsize = AtomicUsize::new(0);

    fn count() {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }

    #[test]
    fn hook_fires_once_per_every_ticks_until_cleared() {
        set_hook(Some(count));
        for _ in 0..EVERY * 3 {
            tick();
        }
        assert_eq!(CALLS.load(Ordering::Relaxed), 3);
        pulse();
        assert_eq!(CALLS.load(Ordering::Relaxed), 4);
        set_hook(None);
        for _ in 0..EVERY * 2 {
            tick();
        }
        pulse();
        assert_eq!(CALLS.load(Ordering::Relaxed), 4);
    }
}
