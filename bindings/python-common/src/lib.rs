//! PyO3 glue shared by every BetterOffice Python binding.

use std::path::Path;

use pyo3::exceptions::{PyOSError, PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyBytes;

/// Decodes a font pack face, returning font bytes and an optional transformed payload.
#[pyfunction]
#[pyo3(signature = (data, codec, *, dict=None, max_output=None))]
pub fn decode_font_face<'py>(
    py: Python<'py>,
    data: &[u8],
    codec: &str,
    dict: Option<&[u8]>,
    max_output: Option<Bound<'py, PyAny>>,
) -> PyResult<(Bound<'py, PyBytes>, Option<Bound<'py, PyBytes>>)> {
    let max_output = match max_output {
        Some(value)
            if codec.parse::<betteroffice_fonts::Codec>().is_ok()
                && betteroffice_fonts::is_zstd(data) =>
        {
            Some(
                value
                    .extract::<u64>()
                    .ok()
                    .and_then(|value| usize::try_from(value).ok())
                    .ok_or_else(|| {
                        PyValueError::new_err(
                            betteroffice_fonts::FaceError::InvalidMaxOutput.to_string(),
                        )
                    })?,
            )
        }
        _ => None,
    };
    let data = data.to_vec();
    let dict = dict.map(<[u8]>::to_vec);
    let decoded = py
        .detach(|| betteroffice_fonts::decode_font_face(&data, codec, dict.as_deref(), max_output))
        .map_err(|error| PyValueError::new_err(error.to_string()))?;
    Ok((
        PyBytes::new(py, &decoded.font),
        decoded.payload.map(|payload| PyBytes::new(py, &payload)),
    ))
}

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
