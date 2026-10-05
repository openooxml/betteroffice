use std::cell::Cell;

#[derive(Clone, Copy, Default, Debug)]
pub(crate) struct StepWork {
    pub(crate) records: usize,
    pub(crate) bytes: usize,
    pub(crate) initialized_bytes: usize,
    pub(crate) scanned_bytes: usize,
    pub(crate) allocated_bytes: usize,
}

thread_local! {
    static WORK: Cell<StepWork> = Cell::new(StepWork::default());
    static DELETED_CLOCKS: Cell<usize> = const { Cell::new(0) };
    static MIGRATED_ENTRIES: Cell<usize> = const { Cell::new(0) };
    static DRAINED_ENTRIES: Cell<usize> = const { Cell::new(0) };
}

pub(crate) fn reset() {
    WORK.set(StepWork::default());
    DELETED_CLOCKS.set(0);
    DRAINED_ENTRIES.set(0);
    MIGRATED_ENTRIES.set(0);
}

pub(crate) fn migrate(entries: usize) {
    MIGRATED_ENTRIES.set(MIGRATED_ENTRIES.get() + entries);
}

pub(crate) fn migrated_entries() -> usize {
    MIGRATED_ENTRIES.get()
}

pub(crate) fn drain(entries: usize) {
    DRAINED_ENTRIES.set(DRAINED_ENTRIES.get() + entries);
}

pub(crate) fn drained_entries() -> usize {
    DRAINED_ENTRIES.get()
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

pub(crate) fn current() -> StepWork {
    WORK.get()
}

pub(crate) fn scan(bytes: usize) {
    let mut work = WORK.get();
    work.scanned_bytes += bytes;
    WORK.set(work);
}

pub(crate) fn allocate(bytes: usize) {
    let mut work = WORK.get();
    work.allocated_bytes += bytes;
    WORK.set(work);
}
