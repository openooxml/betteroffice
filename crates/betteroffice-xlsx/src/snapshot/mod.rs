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
                | "snapshot formula exceeds advance byte budget"
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

pub(crate) fn restore_calculation_settings(
    source: Option<&[u8]>,
    full_precision: Option<bool>,
    iterate: Option<bool>,
) -> SnapshotResult<(bool, bool)> {
    let Some(source) = source.filter(|_| full_precision.is_none() || iterate.is_none()) else {
        return Ok((full_precision.unwrap_or(true), iterate.unwrap_or(false)));
    };
    let parsed =
        ooxml_opc::unzip_parts_where(source, ooxml_opc::MAX_TOTAL_UNCOMPRESSED_BYTES, |name| {
            name.trim_start_matches('/') == "xl/workbook.xml"
        })
        .map_err(SnapshotError::new)
        .and_then(|parts| {
            let Some((_, workbook_xml)) = parts.first() else {
                return Ok((true, false));
            };
            xlsx_parse::workbook_calculation_settings(workbook_xml)
                .map_err(|failure| SnapshotError::new(failure.to_string()))
        });
    let (source_precision, source_iterate) = match parsed {
        Ok(settings) => settings,
        Err(error) if full_precision.is_none() => return Err(error),
        Err(_) => (true, false),
    };
    Ok((
        full_precision.unwrap_or(source_precision),
        iterate.unwrap_or(source_iterate),
    ))
}
