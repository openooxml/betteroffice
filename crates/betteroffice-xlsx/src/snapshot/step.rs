use std::cell::Cell;

#[derive(Clone, Copy, Default, Debug)]
pub(crate) struct StepWork {
    pub(crate) records: usize,
    pub(crate) bytes: usize,
    pub(crate) initialized_bytes: usize,
}

thread_local! {
    static WORK: Cell<StepWork> = Cell::new(StepWork::default());
    static DELETED_CLOCKS: Cell<usize> = const { Cell::new(0) };
}

pub(crate) fn reset() {
    WORK.set(StepWork::default());
    DELETED_CLOCKS.set(0);
}

#[doc(hidden)]
pub(crate) fn delete(clocks: usize) {
    DELETED_CLOCKS.set(DELETED_CLOCKS.get() + clocks);
}

#[doc(hidden)]
pub(crate) fn deleted_clocks() -> usize {
    DELETED_CLOCKS.get()
}

pub(crate) fn record(records: usize, bytes: usize) {
    let mut work = WORK.get();
    work.records += records;
    work.bytes += bytes;
    WORK.set(work);
}

pub(crate) fn initialize(bytes: usize) {
    record(usize::from(bytes != 0), bytes);
    let mut work = WORK.get();
    work.initialized_bytes += bytes;
    WORK.set(work);
}

pub(crate) fn bytes_at_least(bytes: usize) {
    let mut work = WORK.get();
    work.bytes = work.bytes.max(bytes);
    WORK.set(work);
}

pub(crate) fn current() -> StepWork {
    WORK.get()
}
