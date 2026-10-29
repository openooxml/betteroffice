use std::sync::Arc;

#[derive(Clone)]
pub struct Host {
    clock: Arc<dyn Fn() -> u64 + Send + Sync>,
    entropy: Arc<dyn Fn() -> u64 + Send + Sync>,
    timer: Arc<dyn Fn() -> f64 + Send + Sync>,
}

impl Host {
    pub fn new(
        clock: Arc<dyn Fn() -> u64 + Send + Sync>,
        entropy: Arc<dyn Fn() -> u64 + Send + Sync>,
        timer: Arc<dyn Fn() -> f64 + Send + Sync>,
    ) -> Self {
        Self {
            clock,
            entropy,
            timer,
        }
    }

    pub fn clock_ms(&self) -> u64 {
        (self.clock)()
    }

    pub fn entropy(&self) -> u64 {
        (self.entropy)()
    }

    pub fn timer_ms(&self) -> f64 {
        (self.timer)()
    }

    #[cfg(test)]
    pub(crate) fn deterministic() -> Self {
        use std::sync::atomic::{AtomicU64, Ordering};
        let entropy = Arc::new(AtomicU64::new(0x9e37_79b9_7f4a_7c15));
        Self::new(
            Arc::new(|| 0),
            Arc::new(move || entropy.fetch_add(0x9e37_79b9_7f4a_7c15, Ordering::Relaxed)),
            Arc::new(|| 0.0),
        )
    }

    #[cfg(test)]
    pub(crate) fn with_entropy(entropy: Arc<dyn Fn() -> u64 + Send + Sync>) -> Self {
        Self::new(Arc::new(|| 0), entropy, Arc::new(|| 0.0))
    }
}
