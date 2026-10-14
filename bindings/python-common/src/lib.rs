//! PyO3 glue shared by every BetterOffice Python binding.

use std::path::Path;

use pyo3::exceptions::{PyOSError, PyRuntimeError};
use pyo3::prelude::*;

/// The 3-argument form makes `OSError.__new__` pick the errno subclass.
pub fn map_io_error(error: &std::io::Error, path: &Path) -> PyErr {
    PyOSError::new_err((
        error.raw_os_error(),
        error.to_string(),
        path.display().to_string(),
    ))
}

/// A random collaboration client ID, masked to the engine's maximum.
pub fn generated_client_id(max: u64) -> PyResult<u64> {
    let mut bytes = [0_u8; size_of::<u64>()];
    getrandom::fill(&mut bytes).map_err(|error| {
        PyRuntimeError::new_err(format!(
            "could not generate a collaboration client ID: {error}"
        ))
    })?;
    Ok(u64::from_le_bytes(bytes) & max)
}
