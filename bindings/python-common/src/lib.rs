//! PyO3 glue shared by every BetterOffice Python binding.

use std::path::{Path, PathBuf};
use std::sync::OnceLock;

use betteroffice_fonts::catalog::catalog;
use betteroffice_fonts::{BundledFontOptions, FontPack, FontSource, set_default_font_source};
use pyo3::exceptions::{
    PyModuleNotFoundError, PyOSError, PyRuntimeError, PyRuntimeWarning, PyValueError,
};
use pyo3::prelude::*;
use pyo3::types::PyBytes;

const FONT_PACK_MODULE: &str = "betteroffice_fonts_packs";
const FONT_PACKS: [FontPack; 3] = [FontPack::Latin, FontPack::Intl, FontPack::Cjk];

static INSTALLED_FONT_PACKS: OnceLock<Option<BundledFontOptions>> = OnceLock::new();

/// The installed `betteroffice-fonts-*` packs under one site directory, or None
/// when no pack of the catalog's version is installed. Packs split across
/// several directories keep the directory with the most of them.
pub fn font_pack_options(roots: &[PathBuf]) -> Option<BundledFontOptions> {
    let installed = |root: &Path, pack: FontPack| {
        catalog()
            .faces
            .iter()
            .find(|face| face.pack == pack)
            .is_some_and(|face| {
                root.join(pack.as_str())
                    .join(&face.version)
                    .join("manifest.json")
                    .is_file()
            })
    };
    let mut best: Option<BundledFontOptions> = None;
    for root in roots {
        let packs: Vec<FontPack> = FONT_PACKS
            .into_iter()
            .filter(|pack| installed(root, *pack))
            .collect();
        if !packs.is_empty()
            && best
                .as_ref()
                .is_none_or(|best| packs.len() > best.packs.len())
        {
            best = Some(BundledFontOptions {
                packs,
                base_url: Some(root.clone()),
            });
        }
    }
    best
}

fn font_pack_roots(py: Python<'_>) -> PyResult<Vec<PathBuf>> {
    let module = match py.import(FONT_PACK_MODULE) {
        Ok(module) => module,
        Err(error) if error.is_instance_of::<PyModuleNotFoundError>(py) => return Ok(Vec::new()),
        Err(error) => return Err(error),
    };
    module
        .getattr("__path__")?
        .try_iter()?
        .map(|root| root?.extract::<PathBuf>())
        .collect()
}

/// Makes the installed font pack wheels the default font source of every
/// document this process opens. Called from each binding's module init; the
/// first binding imported decides, and later calls change nothing.
pub fn install_font_packs(py: Python<'_>) -> PyResult<()> {
    if INSTALLED_FONT_PACKS.get().is_some() {
        return Ok(());
    }
    let roots = font_pack_roots(py)?;
    let options = INSTALLED_FONT_PACKS.get_or_init(|| font_pack_options(&roots));
    let Some(options) = options else {
        return Ok(());
    };
    let _ = set_default_font_source(FontSource::Bundled(options.clone()));
    let elsewhere: Vec<&str> = FONT_PACKS
        .into_iter()
        .filter(|pack| !options.packs.contains(pack))
        .filter(|pack| {
            roots.iter().any(|root| {
                font_pack_options(std::slice::from_ref(root))
                    .is_some_and(|other| other.packs.contains(pack))
            })
        })
        .map(FontPack::as_str)
        .collect();
    if !elsewhere.is_empty() {
        let message = std::ffi::CString::new(format!(
            "betteroffice font packs {} are installed outside {}; install every betteroffice-fonts-* package into one environment to use them",
            elsewhere.join(", "),
            options.base_url.as_deref().unwrap_or(Path::new("")).display()
        ))
        .unwrap_or_default();
        PyErr::warn(py, py.get_type::<PyRuntimeWarning>().as_any(), &message, 1)?;
    }
    Ok(())
}

/// The font pack directory and packs this process renders with, or None.
#[pyfunction(name = "_font_packs")]
pub fn font_packs() -> Option<(PathBuf, Vec<&'static str>)> {
    INSTALLED_FONT_PACKS.get()?.as_ref().map(|options| {
        (
            options.base_url.clone().unwrap_or_default(),
            options.packs.iter().map(|pack| pack.as_str()).collect(),
        )
    })
}

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
