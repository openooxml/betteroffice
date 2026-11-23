use std::cell::Cell;

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct SpliceWork {
    pub dp_cells: usize,
    pub equal_calls: usize,
    pub hashed_nodes: usize,
    pub fragment_elements: usize,
    pub budget_work: usize,
    pub budget_peak_milli: usize,
    pub budget_exhausted: usize,
}

impl SpliceWork {
    pub(crate) fn record_budget_peak(&mut self, used: usize, weight: usize) {
        let milli = (used as u128) * 1000 / (weight.max(1) as u128);
        let milli = usize::try_from(milli).unwrap_or(usize::MAX);
        self.budget_peak_milli = self.budget_peak_milli.max(milli);
    }
}

thread_local! {
    static WORK: Cell<SpliceWork> = Cell::new(SpliceWork::default());
}

/// The work counted on this thread since the last call, which resets it.
pub fn take() -> SpliceWork {
    WORK.with(|work| work.take())
}

pub(crate) fn bump(count: impl FnOnce(&mut SpliceWork)) {
    WORK.with(|work| {
        let mut current = work.get();
        count(&mut current);
        work.set(current);
    });
}
