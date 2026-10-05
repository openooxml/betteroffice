pub(crate) mod cells;
pub(crate) mod header;
pub(crate) mod model;
pub(crate) mod package;
pub(crate) mod preserved;
#[cfg(test)]
pub(crate) mod step;
pub(crate) mod wire;
pub(crate) mod yrs_split;

pub use crate::workbook::snapshot_assembly::{
    HydratedWorkbook, WorkbookSnapshotBuilder, WorkbookSnapshotEncoder,
};

use std::fmt;

pub type SnapshotResult<T> = std::result::Result<T, SnapshotError>;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SnapshotError(String);

impl SnapshotError {
    pub(crate) fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }
}

impl fmt::Display for SnapshotError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for SnapshotError {}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SnapshotBudget {
    max_records: usize,
    max_bytes: usize,
}

impl SnapshotBudget {
    pub fn new(max_records: usize, max_bytes: usize) -> SnapshotResult<Self> {
        if max_records == 0 || max_bytes == 0 {
            return Err(SnapshotError::new("snapshot budget must be positive"));
        }
        Ok(Self {
            max_records,
            max_bytes,
        })
    }

    pub(crate) fn max_records(self) -> usize {
        self.max_records
    }

    pub(crate) fn max_bytes(self) -> usize {
        self.max_bytes
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SnapshotProgress {
    ready: bool,
}

impl SnapshotProgress {
    pub(crate) fn pending() -> Self {
        Self { ready: false }
    }

    pub(crate) fn ready() -> Self {
        Self { ready: true }
    }

    pub fn is_ready(&self) -> bool {
        self.ready
    }
}
