//! Per-thread counts of the layout work that should scale with an edit, not
//! with the document: values converted for the display compiler and measured
//! lines resolved into run segments.

use std::cell::Cell;

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct WorkCounts {
    pub transcoded: usize,
    pub resolved_lines: usize,
}

std::thread_local! {
    static COUNTS: Cell<WorkCounts> = const {
        Cell::new(WorkCounts { transcoded: 0, resolved_lines: 0 })
    };
}

/// Returns the counts since the previous call and resets them.
pub fn take() -> WorkCounts {
    COUNTS.with(|counts| counts.replace(WorkCounts::default()))
}

fn bump(update: impl FnOnce(&mut WorkCounts)) {
    COUNTS.with(|counts| {
        let mut value = counts.get();
        update(&mut value);
        counts.set(value);
    });
}

pub(crate) fn transcoded() {
    bump(|counts| counts.transcoded += 1);
}

pub(crate) fn resolved_line() {
    bump(|counts| counts.resolved_lines += 1);
}
