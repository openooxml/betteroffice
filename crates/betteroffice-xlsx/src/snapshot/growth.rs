use std::mem::size_of;

use super::{SnapshotBudget, SnapshotError, SnapshotResult};

pub(crate) struct Growth<S> {
    pending: Option<Box<dyn Migration<S> + Send + Sync>>,
}

impl<S> Default for Growth<S> {
    fn default() -> Self {
        Self { pending: None }
    }
}

trait Migration<S> {
    fn advance(&mut self, state: &mut S, budget: SnapshotBudget) -> SnapshotResult<bool>;
}

type Target<S, T> = fn(&mut S) -> SnapshotResult<&mut Vec<T>>;

struct Move<S, T> {
    records: std::vec::IntoIter<T>,
    storage: Vec<T>,
    target: Target<S, T>,
}

impl<S, T> Migration<S> for Move<S, T> {
    fn advance(&mut self, state: &mut S, budget: SnapshotBudget) -> SnapshotResult<bool> {
        if size_of::<T>() > budget.max_bytes() {
            return Err(SnapshotError::new(
                "snapshot storage exceeds advance byte budget",
            ));
        }
        let count = self.records.len().min(
            budget
                .max_records()
                .min(budget.max_bytes() / size_of::<T>().max(1)),
        );
        self.storage.extend(self.records.by_ref().take(count));
        super::step::record(count, count * size_of::<T>());
        if self.records.as_slice().is_empty() {
            *(self.target)(state)? = std::mem::take(&mut self.storage);
            Ok(true)
        } else {
            Ok(false)
        }
    }
}

impl<S: 'static> Growth<S> {
    pub(crate) fn ensure<T: Send + Sync + 'static>(
        &mut self,
        state: &mut S,
        target: Target<S, T>,
        budget: SnapshotBudget,
    ) -> SnapshotResult<bool> {
        if let Some(pending) = &mut self.pending {
            if pending.advance(state, budget)? {
                self.pending = None;
            }
            return Ok(false);
        }
        let values = target(state)?;
        if values.len() < values.capacity() {
            return Ok(true);
        }
        if !values.is_empty() && size_of::<T>() > budget.max_bytes() {
            return Err(SnapshotError::new(
                "snapshot storage exceeds advance byte budget",
            ));
        }
        let capacity = values.len().saturating_mul(2).max(1);
        let mut storage = Vec::new();
        storage
            .try_reserve_exact(capacity)
            .map_err(|_| SnapshotError::new("cannot allocate snapshot records"))?;
        if values.is_empty() {
            *values = storage;
            return Ok(true);
        }
        let mut pending = Move {
            records: std::mem::take(values).into_iter(),
            storage,
            target,
        };
        let complete = pending.advance(state, budget);
        self.pending = Some(Box::new(pending));
        if complete? {
            self.pending = None;
        }
        Ok(false)
    }

    pub(crate) fn is_pending(&self) -> bool {
        self.pending.is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn storage_migration_respects_record_and_byte_limits() {
        let mut records: Vec<[u8; 64]> = (0..8).map(|value| [value; 64]).collect();
        records.shrink_to_fit();
        let budget = SnapshotBudget::new(3, 128).unwrap();
        let mut growth = Growth::default();
        let mut steps = 0;
        loop {
            crate::snapshot::step::reset();
            let ready = growth
                .ensure(&mut records, |records| Ok(records), budget)
                .unwrap();
            let work = crate::snapshot::step::current();
            assert!(work.records <= budget.max_records());
            assert!(work.bytes <= budget.max_bytes());
            if ready {
                break;
            }
            assert_eq!(work.records, 2);
            assert_eq!(work.bytes, 128);
            steps += 1;
        }
        assert_eq!(steps, 4);
        assert_eq!(records, (0..8).map(|value| [value; 64]).collect::<Vec<_>>());
    }

    #[test]
    fn storage_migration_requires_a_budget_that_fits_one_record() {
        let mut records: Vec<[u8; 64]> = (0..4).map(|value| [value; 64]).collect();
        records.shrink_to_fit();
        let mut growth = Growth::default();
        let too_small = SnapshotBudget::new(3, 16).unwrap();
        assert!(
            growth
                .ensure(&mut records, |records| Ok(records), too_small)
                .is_err()
        );
        let budget = SnapshotBudget::new(1, 64).unwrap();
        let mut steps = 0;
        loop {
            crate::snapshot::step::reset();
            let ready = growth
                .ensure(&mut records, |records| Ok(records), budget)
                .unwrap();
            if ready {
                break;
            }
            assert_eq!(crate::snapshot::step::current().records, 1);
            assert!(crate::snapshot::step::current().bytes <= budget.max_bytes());
            steps += 1;
        }
        assert_eq!(steps, 4);
        assert_eq!(records, (0..4).map(|value| [value; 64]).collect::<Vec<_>>());
    }
}
