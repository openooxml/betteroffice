use std::cell::Cell;

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct SpliceWork {
    pub dp_cells: usize,
    pub equal_calls: usize,
    pub hashed_nodes: usize,
    pub fragment_elements: usize,
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
