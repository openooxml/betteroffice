use std::cell::RefCell;
use std::collections::BTreeMap;
use std::time::{Duration, Instant};

#[derive(Default)]
pub struct Sample {
    pub phases: BTreeMap<&'static str, Duration>,
    pub counts: BTreeMap<&'static str, u64>,
}

thread_local! {
    static SAMPLE: RefCell<Option<Sample>> = const { RefCell::new(None) };
}

pub fn begin() {
    SAMPLE.with(|sample| *sample.borrow_mut() = Some(Sample::default()));
}

pub fn finish() -> Sample {
    SAMPLE.with(|sample| sample.borrow_mut().take().unwrap())
}

pub fn count(name: &'static str, value: u64) {
    SAMPLE.with(|sample| {
        if let Some(sample) = sample.borrow_mut().as_mut() {
            *sample.counts.entry(name).or_default() += value;
        }
    });
}

pub fn span(name: &'static str) -> Option<Span> {
    SAMPLE.with(|sample| {
        sample.borrow().as_ref().map(|_| Span {
            name,
            started: Instant::now(),
        })
    })
}

pub struct Span {
    name: &'static str,
    started: Instant,
}

impl Drop for Span {
    fn drop(&mut self) {
        let elapsed = self.started.elapsed();
        SAMPLE.with(|sample| {
            if let Some(sample) = sample.borrow_mut().as_mut() {
                *sample.phases.entry(self.name).or_default() += elapsed;
            }
        });
    }
}
