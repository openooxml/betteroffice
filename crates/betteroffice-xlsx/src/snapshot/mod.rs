pub(crate) mod cells;
pub(crate) mod growth;
pub(crate) mod header;
pub(crate) mod model;
pub(crate) mod package;
pub(crate) mod preserved;
pub(crate) mod step;
pub(crate) mod wire;
pub(crate) mod yrs_split;

pub use crate::workbook::snapshot_assembly::{
    HydratedWorkbook, WorkbookSnapshotBuilder, WorkbookSnapshotEncoder,
};

use std::fmt;

#[doc(hidden)]
pub type SnapshotResult<T> = std::result::Result<T, SnapshotError>;

#[derive(Debug, Clone, PartialEq, Eq)]
#[doc(hidden)]
pub struct SnapshotError(String);

impl SnapshotError {
    pub(crate) fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }

    pub(crate) fn is_budget_refusal(&self) -> bool {
        matches!(
            self.0.as_str(),
            "snapshot storage exceeds advance byte budget"
                | "snapshot logical chunk exceeds advance byte budget"
                | "snapshot calculation exceeds advance byte budget"
                | "snapshot record exceeds advance byte budget"
                | "snapshot stream exceeds advance record or byte budget"
                | "snapshot model exceeds advance byte budget"
                | "snapshot preservation exceeds advance byte budget"
                | "snapshot decoding byte budget is too small"
                | "snapshot decoding exceeds advance byte budget"
                | "authority base decoding byte budget is too small"
                | "snapshot authority exceeds advance byte budget"
                | "Yrs snapshot record exceeds advance byte budget"
                | "snapshot validation exceeds advance byte budget"
                | "snapshot authority validation exceeds advance byte budget"
                | "snapshot authority retirement exceeds advance byte budget"
                | "snapshot chart exceeds advance byte budget"
                | "snapshot graph record exceeds advance byte budget"
                | "facts storage exceeds advance budget"
        )
    }
}

impl fmt::Display for SnapshotError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for SnapshotError {}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[doc(hidden)]
pub struct SnapshotBudget {
    max_records: usize,
    max_bytes: usize,
    partial: bool,
}

impl SnapshotBudget {
    pub fn new(max_records: usize, max_bytes: usize) -> SnapshotResult<Self> {
        if max_records == 0 || max_bytes == 0 {
            return Err(SnapshotError::new("snapshot budget must be positive"));
        }
        Ok(Self {
            max_records,
            max_bytes,
            partial: false,
        })
    }

    pub(crate) fn remaining(self, records: usize, bytes: usize) -> Option<Self> {
        let max_records = self.max_records.checked_sub(records)?;
        let max_bytes = self.max_bytes.checked_sub(bytes)?;
        if max_records == 0 || max_bytes == 0 {
            return None;
        }
        Some(Self {
            max_records,
            max_bytes,
            partial: records != 0 || bytes != 0,
        })
    }

    pub(crate) fn is_partial(self) -> bool {
        self.partial
    }

    pub(crate) fn max_records(self) -> usize {
        self.max_records
    }

    pub(crate) fn max_bytes(self) -> usize {
        self.max_bytes
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[doc(hidden)]
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
