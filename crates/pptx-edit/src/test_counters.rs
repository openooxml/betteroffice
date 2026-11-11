use std::cell::Cell;

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct AttachWork {
    pub story_reads: usize,
    pub replayed_reads: usize,
    pub visible_builds: usize,
    pub replay_seeds: usize,
    pub seeded_snapshots: usize,
    pub live_snapshots: usize,
}

thread_local! {
    static WORK: Cell<AttachWork> = Cell::new(AttachWork::default());
}

/// The work counted on this thread since the last call, which resets it.
pub fn take() -> AttachWork {
    WORK.with(|work| work.take())
}

pub(crate) fn bump(count: impl FnOnce(&mut AttachWork)) {
    WORK.with(|work| {
        let mut current = work.get();
        count(&mut current);
        work.set(current);
    });
}
