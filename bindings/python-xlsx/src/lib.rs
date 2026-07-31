//! PyO3 bindings over the `betteroffice-xlsx` facade.

use std::fs;
use std::path::PathBuf;

use pyo3::create_exception;
use pyo3::exceptions::{PyException, PyIndexError, PyKeyError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyBytes, PyInt};
use python_common::{generated_client_id, map_io_error};

use betteroffice_xlsx::{
    CalculationOptions, CellAddress, CellInput, CellRange, CellRef, CellValue, Error as CoreError,
    HorizontalAlignment, MAX_COLLABORATION_BYTES, MAX_COLLABORATION_CLIENT_ID, MutationResult,
    NumberFormatMutation, Proposal, ProposalEditInput, ProposalRequest, RenderOptions, SheetId,
    StylePatch, TextWrapping, VerticalAlignment, Workbook as CoreWorkbook,
};

create_exception!(
    _betteroffice_xlsx,
    XlsxError,
    PyException,
    "Base class for every error raised by the engine."
);
create_exception!(
    _betteroffice_xlsx,
    ParseError,
    XlsxError,
    "The workbook could not be read."
);
create_exception!(
    _betteroffice_xlsx,
    RangeError,
    XlsxError,
    "A sheet, cell, or range was out of bounds or too large."
);
create_exception!(
    _betteroffice_xlsx,
    RenderError,
    XlsxError,
    "Rendering failed or exceeded a size limit."
);
create_exception!(
    _betteroffice_xlsx,
    InvalidUpdateError,
    XlsxError,
    "A peer sent an invalid collaboration update."
);
create_exception!(
    _betteroffice_xlsx,
    CollaborativeStateError,
    XlsxError,
    "The local collaborative state is invalid."
);
create_exception!(
    _betteroffice_xlsx,
    StaleProposalError,
    XlsxError,
    "A proposal's target cells changed before acceptance."
);
create_exception!(
    _betteroffice_xlsx,
    NotCollaborativeError,
    XlsxError,
    "The operation requires a collaborative workbook."
);

fn stale_proposal_error(message: String, cells: Vec<CellAddress>) -> PyErr {
    let cells = cells
        .into_iter()
        .map(|address| address.cell.to_a1())
        .collect::<Vec<_>>();
    Python::attach(|py| {
        let error = StaleProposalError::new_err(message);
        match error.value(py).setattr("cells", cells) {
            Ok(()) => error,
            Err(attribute_error) => attribute_error,
        }
    })
}

fn map_error(error: CoreError) -> PyErr {
    let message = error.to_string();
    match error {
        CoreError::InvalidUpdate(_) => InvalidUpdateError::new_err(message),
        CoreError::CollaborativeState(_) => CollaborativeStateError::new_err(message),
        CoreError::StaleProposal(cells) => stale_proposal_error(message, cells),
        CoreError::ProposalNotFound(_) => PyKeyError::new_err(message),
        CoreError::NotCollaborative => NotCollaborativeError::new_err(message),
        CoreError::Package(_)
        | CoreError::Spreadsheet(_)
        | CoreError::DuplicatePart(_)
        | CoreError::NoSheets => ParseError::new_err(message),
        CoreError::SheetOutOfRange(_)
        | CoreError::CellOutOfRange(_)
        | CoreError::RangeTooLarge { .. }
        | CoreError::DisplayTooLarge { .. }
        | CoreError::InvalidViewport => RangeError::new_err(message),
        CoreError::InvalidScale(_)
        | CoreError::RenderTooLarge { .. }
        | CoreError::RenderAreaTooLarge { .. }
        | CoreError::Raster(_) => RenderError::new_err(message),
        _ => XlsxError::new_err(message),
    }
}

fn map_style_error(error: CoreError) -> PyErr {
    if matches!(error, CoreError::Operation(_)) {
        PyValueError::new_err(error.to_string())
    } else {
        map_error(error)
    }
}

fn parse_cell(address: &str) -> PyResult<CellRef> {
    CellRef::parse_a1(&address.to_ascii_uppercase())
        .map_err(|error| RangeError::new_err(format!("invalid cell {address:?}: {error}")))
}

fn parse_range(address: &str) -> PyResult<CellRange> {
    CellRange::parse_a1(&address.to_ascii_uppercase())
        .map_err(|error| RangeError::new_err(format!("invalid range {address:?}: {error}")))
}

/// An Excel error value, distinct from a cell holding that text.
#[pyclass(module = "betteroffice_xlsx", name = "CellError", frozen)]
pub struct PyCellError {
    #[pyo3(get)]
    code: String,
}

#[pymethods]
impl PyCellError {
    fn __str__(&self) -> &str {
        &self.code
    }

    fn __repr__(&self) -> String {
        format!("CellError({:?})", self.code)
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        if let Ok(other) = other.extract::<PyRef<'_, Self>>() {
            return self.code == other.code;
        }
        other
            .extract::<String>()
            .is_ok_and(|text| text == self.code)
    }

    /// Must match `str`'s hash, since `__eq__` accepts one.
    fn __hash__(&self, py: Python<'_>) -> PyResult<isize> {
        self.code.as_str().into_pyobject(py)?.hash()
    }
}

#[pyclass(module = "betteroffice_xlsx", name = "Png", frozen)]
pub struct PyPng {
    data: Vec<u8>,
    #[pyo3(get)]
    width: u32,
    #[pyo3(get)]
    height: u32,
}

#[pymethods]
impl PyPng {
    #[getter]
    fn bytes<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.data)
    }

    fn write(&self, py: Python<'_>, path: PathBuf) -> PyResult<()> {
        py.detach(|| fs::write(&path, &self.data))
            .map_err(|error| map_io_error(&error, &path))
    }

    fn __len__(&self) -> usize {
        self.data.len()
    }

    fn __repr__(&self) -> String {
        format!(
            "Png(width={}, height={}, bytes={})",
            self.width,
            self.height,
            self.data.len()
        )
    }
}

#[pyclass(module = "betteroffice_xlsx", name = "Calculation", frozen)]
pub struct PyCalculation {
    #[pyo3(get)]
    changed: usize,
    #[pyo3(get)]
    cycles: usize,
    #[pyo3(get)]
    limited: usize,
}

#[pymethods]
impl PyCalculation {
    fn __repr__(&self) -> String {
        format!(
            "Calculation(changed={}, cycles={}, limited={})",
            self.changed, self.cycles, self.limited
        )
    }
}

/// What a mutating call changed.
#[pyclass(module = "betteroffice_xlsx", name = "Mutation", frozen)]
pub struct PyMutation {
    #[pyo3(get)]
    applied: bool,
    #[pyo3(get)]
    changed: Vec<String>,
    #[pyo3(get)]
    cycles: Vec<String>,
    #[pyo3(get)]
    limited: Vec<String>,
}

impl PyMutation {
    fn from_core(workbook: &CoreWorkbook, result: &MutationResult) -> Self {
        let names = |cells: &[CellAddress]| -> Vec<String> {
            cells
                .iter()
                .map(|at| workbook.format_address(*at))
                .collect()
        };
        Self {
            applied: result.applied,
            changed: names(&result.changed),
            cycles: names(&result.cycle_cells),
            limited: names(&result.limited_cells),
        }
    }
}

#[pymethods]
impl PyMutation {
    fn __bool__(&self) -> bool {
        self.applied
    }

    fn __repr__(&self) -> String {
        format!(
            "Mutation(applied={}, changed={})",
            if self.applied { "True" } else { "False" },
            self.changed.len()
        )
    }
}

#[pyclass(module = "betteroffice_xlsx", name = "History", frozen)]
pub struct PyHistory {
    #[pyo3(get)]
    can_undo: bool,
    #[pyo3(get)]
    can_redo: bool,
    #[pyo3(get)]
    undo_depth: usize,
    #[pyo3(get)]
    redo_depth: usize,
}

#[pymethods]
impl PyHistory {
    fn __repr__(&self) -> String {
        format!(
            "History(undo_depth={}, redo_depth={})",
            self.undo_depth, self.redo_depth
        )
    }
}

/// A staged edit set awaiting a human decision.
#[pyclass(module = "betteroffice_xlsx", name = "Proposal", frozen)]
pub struct PyProposal {
    #[pyo3(get)]
    id: String,
    #[pyo3(get)]
    agent_id: String,
    #[pyo3(get)]
    note: Option<String>,
    #[pyo3(get)]
    edits: Vec<PyProposedEdit>,
}

impl PyProposal {
    fn from_core(proposal: &Proposal) -> Self {
        Self {
            id: proposal.id.clone(),
            agent_id: proposal.agent_id.clone(),
            note: proposal.note.clone(),
            edits: proposal
                .edits
                .iter()
                .map(|edit| PyProposedEdit {
                    sheet: edit.sheet as usize,
                    address: CellRef::new(edit.row, edit.col).to_a1(),
                    input: edit.input.clone(),
                    before: edit.old_text.clone(),
                    after: edit.new_text.clone(),
                })
                .collect(),
        }
    }
}

#[pymethods]
impl PyProposal {
    fn __repr__(&self) -> String {
        format!(
            "Proposal(id={:?}, agent_id={:?}, edits={})",
            self.id,
            self.agent_id,
            self.edits.len()
        )
    }
}

/// One cell inside a proposal, with the text a reviewer would compare.
#[pyclass(
    module = "betteroffice_xlsx",
    name = "ProposedEdit",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyProposedEdit {
    #[pyo3(get)]
    sheet: usize,
    #[pyo3(get)]
    address: String,
    #[pyo3(get)]
    input: String,
    #[pyo3(get)]
    before: String,
    #[pyo3(get)]
    after: String,
}

#[pymethods]
impl PyProposedEdit {
    fn __repr__(&self) -> String {
        format!(
            "ProposedEdit({} {:?} -> {:?})",
            self.address, self.before, self.after
        )
    }
}

#[pyclass(module = "betteroffice_xlsx", name = "Workbook")]
pub struct PyWorkbook {
    inner: CoreWorkbook,
}

impl PyWorkbook {
    /// Copies first: a borrow of a Python buffer cannot cross `detach`.
    fn open_bytes(
        py: Python<'_>,
        data: &[u8],
        options: Option<CalculationOptions>,
    ) -> PyResult<Self> {
        let data = data.to_vec();
        py.detach(|| match options {
            Some(options) => CoreWorkbook::open_recalculated(&data, options),
            None => CoreWorkbook::open(&data),
        })
        .map(|inner| Self { inner })
        .map_err(map_error)
    }

    fn resolve_sheet(&self, key: &Bound<'_, PyAny>) -> PyResult<SheetId> {
        // bool is an int subclass, so True would otherwise select sheet 1.
        if key.is_instance_of::<PyBool>() {
            return Err(PyTypeError::new_err(
                "sheet must be a name (str) or an index (int), not bool",
            ));
        }
        if let Ok(name) = key.extract::<String>() {
            return self
                .inner
                .sheet_id(&name)
                .ok_or_else(|| PyKeyError::new_err(format!("no sheet named {name:?}")));
        }
        if key.is_instance_of::<PyInt>() {
            let count = self.inner.sheet_count();
            let index = key.extract::<usize>().map_err(|_| {
                PyIndexError::new_err(format!(
                    "sheet index {key} out of range for {count} sheet(s)"
                ))
            })?;
            if index >= count {
                return Err(PyIndexError::new_err(format!(
                    "sheet index {index} out of range for {count} sheet(s)"
                )));
            }
            return Ok(SheetId(index as u32));
        }
        Err(PyTypeError::new_err(
            "sheet must be a name (str) or an index (int)",
        ))
    }
}

#[pymethods]
impl PyWorkbook {
    /// Keeps the values already cached in the file.
    #[staticmethod]
    fn open(py: Python<'_>, data: &[u8]) -> PyResult<Self> {
        Self::open_bytes(py, data, None)
    }

    #[staticmethod]
    fn open_path(py: Python<'_>, path: PathBuf) -> PyResult<Self> {
        let data = py
            .detach(|| fs::read(&path))
            .map_err(|error| map_io_error(&error, &path))?;
        Self::open_bytes(py, &data, None)
    }

    #[staticmethod]
    #[pyo3(signature = (data, *, now_serial = None))]
    fn open_recalculated(py: Python<'_>, data: &[u8], now_serial: Option<f64>) -> PyResult<Self> {
        Self::open_bytes(py, data, Some(CalculationOptions { now_serial }))
    }

    #[pyo3(signature = (*, now_serial = None))]
    fn recalculate(&mut self, py: Python<'_>, now_serial: Option<f64>) -> PyCalculation {
        let result = py.detach(|| {
            self.inner
                .recalculate_all(CalculationOptions { now_serial })
        });
        PyCalculation {
            changed: result.changed.len(),
            cycles: result.cycle_cells.len(),
            limited: result.limited_cells.len(),
        }
    }

    #[getter]
    fn sheet_count(&self) -> usize {
        self.inner.sheet_count()
    }

    fn sheet_index(&self, sheet: &Bound<'_, PyAny>) -> PyResult<usize> {
        self.resolve_sheet(sheet).map(|sheet| sheet.0 as usize)
    }

    #[getter]
    fn sheet_names(&self) -> PyResult<Vec<String>> {
        (0..self.inner.sheet_count())
            .map(|index| {
                self.inner
                    .sheet(SheetId(index as u32))
                    .map(|sheet| sheet.name.clone())
                    .map_err(map_error)
            })
            .collect()
    }

    fn value(
        &self,
        py: Python<'_>,
        sheet: &Bound<'_, PyAny>,
        address: &str,
    ) -> PyResult<Py<PyAny>> {
        let sheet = self.resolve_sheet(sheet)?;
        let cell = parse_cell(address)?;
        let sheet = self.inner.sheet(sheet).map_err(map_error)?;
        let Some(found) = sheet.cell(cell) else {
            return Ok(py.None());
        };
        match &found.value {
            CellValue::Empty => Ok(py.None()),
            CellValue::Number { value } => Ok(value.into_pyobject(py)?.unbind().into_any()),
            CellValue::Text { value } => Ok(value.into_pyobject(py)?.unbind().into_any()),
            CellValue::Bool { value } => Ok(value
                .into_pyobject(py)
                .map(|bound| bound.to_owned())?
                .unbind()
                .into_any()),
            CellValue::Error { value } => Ok(Py::new(
                py,
                PyCellError {
                    code: value.as_str().to_string(),
                },
            )?
            .into_any()),
        }
    }

    /// Source formula without the leading `=`.
    fn formula(&self, sheet: &Bound<'_, PyAny>, address: &str) -> PyResult<Option<String>> {
        let sheet = self.resolve_sheet(sheet)?;
        let cell = parse_cell(address)?;
        let sheet = self.inner.sheet(sheet).map_err(map_error)?;
        Ok(sheet.cell(cell).and_then(|found| found.formula.clone()))
    }

    /// Takes what a user would type; returns whether anything changed.
    #[pyo3(signature = (sheet, address, value, *, now_serial = None))]
    fn set(
        &mut self,
        sheet: &Bound<'_, PyAny>,
        address: &str,
        value: &str,
        now_serial: Option<f64>,
    ) -> PyResult<PyMutation> {
        let sheet = self.resolve_sheet(sheet)?;
        let cell = parse_cell(address)?;
        let result = self
            .inner
            .edit_cell(sheet, cell, value, CalculationOptions { now_serial })
            .map_err(map_error)?;
        Ok(PyMutation::from_core(&self.inner, &result))
    }

    #[pyo3(signature = (sheet, *, scale = 1.0, range = None, max_width = None, max_height = None))]
    fn render_png(
        &self,
        py: Python<'_>,
        sheet: &Bound<'_, PyAny>,
        scale: f32,
        range: Option<&str>,
        max_width: Option<u32>,
        max_height: Option<u32>,
    ) -> PyResult<PyPng> {
        let sheet = self.resolve_sheet(sheet)?;
        let range = range.map(parse_range).transpose()?;
        let options = RenderOptions {
            range,
            scale,
            max_width,
            max_height,
        };
        let rendered = py
            .detach(|| self.inner.render_sheet(sheet, &options))
            .map_err(map_error)?;
        Ok(PyPng {
            data: rendered.bytes,
            width: rendered.width,
            height: rendered.height,
        })
    }

    fn save<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let bytes = py.detach(|| self.inner.save()).map_err(map_error)?;
        Ok(PyBytes::new(py, &bytes))
    }

    fn save_path(&self, py: Python<'_>, path: PathBuf) -> PyResult<()> {
        enum Failure {
            Engine(CoreError),
            Io(std::io::Error),
        }

        py.detach(|| {
            let bytes = self.inner.save().map_err(Failure::Engine)?;
            fs::write(&path, bytes).map_err(Failure::Io)
        })
        .map_err(|failure| match failure {
            Failure::Engine(error) => map_error(error),
            Failure::Io(error) => map_io_error(&error, &path),
        })
    }

    /// Open a replica with a generated ID unless `client_id` is supplied.
    #[staticmethod]
    #[pyo3(signature = (data, *, client_id = None, recalculate = false, now_serial = None))]
    fn open_collaborative(
        py: Python<'_>,
        data: &[u8],
        client_id: Option<u64>,
        recalculate: bool,
        now_serial: Option<f64>,
    ) -> PyResult<Self> {
        let client_id =
            client_id.map_or_else(|| generated_client_id(MAX_COLLABORATION_CLIENT_ID), Ok)?;
        let data = data.to_vec();
        py.detach(|| {
            if recalculate {
                CoreWorkbook::open_collaborative_recalculated(
                    &data,
                    client_id,
                    CalculationOptions { now_serial },
                )
            } else {
                CoreWorkbook::open_collaborative(&data, client_id)
            }
        })
        .map(|inner| Self { inner })
        .map_err(map_error)
    }

    #[getter]
    fn client_id(&self) -> u64 {
        self.inner.client_id()
    }

    #[getter]
    fn is_collaborative(&self) -> bool {
        self.inner.is_collaborative()
    }

    /// This replica's state vector, to hand a peer so it can compute a diff.
    fn state_vector<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.inner.encode_state_vector_v1())
    }

    /// The whole document as one update, for a peer joining from nothing.
    fn state_as_update<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.inner.encode_state_as_update_v1())
    }

    /// The update carrying everything the peer's state vector is missing.
    fn diff<'py>(&self, py: Python<'py>, state_vector: &[u8]) -> PyResult<Bound<'py, PyBytes>> {
        let update = self.inner.encode_diff_v1(state_vector).map_err(map_error)?;
        Ok(PyBytes::new(py, &update))
    }

    #[pyo3(signature = (update, *, now_serial = None))]
    fn apply_update(
        &mut self,
        py: Python<'_>,
        update: &[u8],
        now_serial: Option<f64>,
    ) -> PyResult<PyMutation> {
        if update.len() > MAX_COLLABORATION_BYTES {
            return Err(map_error(CoreError::CollaborationDataTooLarge {
                bytes: update.len(),
                max: MAX_COLLABORATION_BYTES,
            }));
        }
        let update = update.to_vec();
        let result = py
            .detach(|| {
                self.inner
                    .apply_update_v1(&update, CalculationOptions { now_serial })
            })
            .map_err(map_error)?;
        Ok(PyMutation::from_core(&self.inner, &result))
    }

    #[getter]
    fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    #[getter]
    fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }

    fn history(&self) -> PyHistory {
        let state = self.inner.history_state();
        PyHistory {
            can_undo: state.can_undo,
            can_redo: state.can_redo,
            undo_depth: state.undo_depth,
            redo_depth: state.redo_depth,
        }
    }

    #[pyo3(signature = (*, now_serial = None))]
    fn undo(&mut self, now_serial: Option<f64>) -> PyResult<PyMutation> {
        let result = self
            .inner
            .undo(CalculationOptions { now_serial })
            .map_err(map_error)?;
        Ok(PyMutation::from_core(&self.inner, &result))
    }

    #[pyo3(signature = (*, now_serial = None))]
    fn redo(&mut self, now_serial: Option<f64>) -> PyResult<PyMutation> {
        let result = self
            .inner
            .redo(CalculationOptions { now_serial })
            .map_err(map_error)?;
        Ok(PyMutation::from_core(&self.inner, &result))
    }

    /// Write many cells as one undo step.
    #[pyo3(signature = (sheet, edits, *, now_serial = None))]
    fn set_many(
        &mut self,
        sheet: &Bound<'_, PyAny>,
        edits: Vec<(String, String)>,
        now_serial: Option<f64>,
    ) -> PyResult<PyMutation> {
        let sheet = self.resolve_sheet(sheet)?;
        let inputs = edits
            .into_iter()
            .map(|(address, input)| parse_cell(&address).map(|cell| CellInput { cell, input }))
            .collect::<PyResult<Vec<_>>>()?;
        let result = self
            .inner
            .edit_cells(sheet, &inputs, CalculationOptions { now_serial })
            .map_err(map_error)?;
        Ok(PyMutation::from_core(&self.inner, &result))
    }

    /// Stage edits for a human to accept or reject instead of applying them.
    #[pyo3(signature = (agent_id, edits, *, note = None, now_serial = None))]
    fn propose(
        &mut self,
        agent_id: String,
        edits: Vec<(Bound<'_, PyAny>, String, String)>,
        note: Option<String>,
        now_serial: Option<f64>,
    ) -> PyResult<PyProposal> {
        let edits = edits
            .into_iter()
            .map(|(sheet, address, input)| {
                Ok(ProposalEditInput {
                    sheet: self.resolve_sheet(&sheet)?,
                    cell: parse_cell(&address)?,
                    input,
                    number_format: None,
                })
            })
            .collect::<PyResult<Vec<_>>>()?;

        let proposal = self
            .inner
            .propose(
                ProposalRequest {
                    agent_id,
                    note,
                    edits,
                },
                CalculationOptions { now_serial },
            )
            .map_err(map_error)?;
        Ok(PyProposal::from_core(&proposal))
    }

    fn proposals(&self) -> Vec<PyProposal> {
        self.inner
            .proposals()
            .iter()
            .map(PyProposal::from_core)
            .collect()
    }

    #[pyo3(signature = (proposal_id, *, force = false, now_serial = None))]
    fn accept_proposal(
        &mut self,
        proposal_id: &str,
        force: bool,
        now_serial: Option<f64>,
    ) -> PyResult<PyMutation> {
        let acceptance = self
            .inner
            .accept_proposal(proposal_id, force, CalculationOptions { now_serial })
            .map_err(map_error)?;
        Ok(PyMutation::from_core(&self.inner, &acceptance.mutation))
    }

    fn reject_proposal(&mut self, proposal_id: &str) -> bool {
        self.inner.reject_proposal(proposal_id)
    }

    #[getter]
    fn active_sheet(&self) -> usize {
        self.inner.active_sheet().0 as usize
    }

    /// Select a sheet and persist it as the workbook's active tab.
    fn set_active_sheet(&mut self, sheet: &Bound<'_, PyAny>) -> PyResult<()> {
        let sheet = self.resolve_sheet(sheet)?;
        self.inner.set_active_sheet(sheet).map_err(map_error)
    }

    fn merged_ranges(&self, sheet: &Bound<'_, PyAny>, range: &str) -> PyResult<Vec<String>> {
        let sheet = self.resolve_sheet(sheet)?;
        let range = parse_range(range)?;
        Ok(self
            .inner
            .merged_ranges(sheet, range)
            .map_err(map_error)?
            .into_iter()
            .map(|found| found.to_a1())
            .collect())
    }

    fn last_calculation(&self) -> PyCalculation {
        let result = self.inner.last_calculation();
        PyCalculation {
            changed: result.changed.len(),
            cycles: result.cycle_cells.len(),
            limited: result.limited_cells.len(),
        }
    }

    /// Set a number format over a range. `format` is one of automatic, text,
    /// number, percent, scientific, currency, date, time, or a custom pattern.
    #[pyo3(signature = (sheet, range, format, *, now_serial = None))]
    fn set_number_format(
        &mut self,
        sheet: &Bound<'_, PyAny>,
        range: &str,
        format: &str,
        now_serial: Option<f64>,
    ) -> PyResult<PyMutation> {
        let sheet = self.resolve_sheet(sheet)?;
        let range = parse_range(range)?;
        let format = match format.to_ascii_lowercase().as_str() {
            "automatic" => NumberFormatMutation::Automatic,
            "text" => NumberFormatMutation::PlainText,
            "number" => NumberFormatMutation::Number,
            "percent" => NumberFormatMutation::Percent,
            "scientific" => NumberFormatMutation::Scientific,
            "currency" => NumberFormatMutation::Currency,
            "date" => NumberFormatMutation::Date,
            "time" => NumberFormatMutation::Time,
            _ => NumberFormatMutation::Custom {
                pattern: format.to_string(),
            },
        };
        let result = self
            .inner
            .set_range_number_format(sheet, range, format, CalculationOptions { now_serial })
            .map_err(map_error)?;
        Ok(PyMutation::from_core(&self.inner, &result))
    }

    /// Patch styling over a range. Every argument is optional; omitted ones are
    /// left as they are.
    #[pyo3(signature = (
        sheet, range, *,
        bold = None, italic = None, strikethrough = None,
        font_family = None, font_size = None,
        text_color = None, fill_color = None,
        horizontal_alignment = None, vertical_alignment = None, text_wrapping = None,
        now_serial = None
    ))]
    #[allow(clippy::too_many_arguments)]
    fn set_style(
        &mut self,
        sheet: &Bound<'_, PyAny>,
        range: &str,
        bold: Option<bool>,
        italic: Option<bool>,
        strikethrough: Option<bool>,
        font_family: Option<String>,
        font_size: Option<f64>,
        text_color: Option<String>,
        fill_color: Option<String>,
        horizontal_alignment: Option<String>,
        vertical_alignment: Option<String>,
        text_wrapping: Option<String>,
        now_serial: Option<f64>,
    ) -> PyResult<PyMutation> {
        let sheet = self.resolve_sheet(sheet)?;
        let range = parse_range(range)?;

        let horizontal = horizontal_alignment
            .as_deref()
            .map(|value| match value.to_ascii_lowercase().as_str() {
                "left" => Ok(HorizontalAlignment::Left),
                "center" => Ok(HorizontalAlignment::Center),
                "right" => Ok(HorizontalAlignment::Right),
                other => Err(PyValueError::new_err(format!(
                    "horizontal_alignment must be left, center, or right, not {other:?}"
                ))),
            })
            .transpose()?;
        let vertical = vertical_alignment
            .as_deref()
            .map(|value| match value.to_ascii_lowercase().as_str() {
                "top" => Ok(VerticalAlignment::Top),
                "middle" => Ok(VerticalAlignment::Middle),
                "bottom" => Ok(VerticalAlignment::Bottom),
                other => Err(PyValueError::new_err(format!(
                    "vertical_alignment must be top, middle, or bottom, not {other:?}"
                ))),
            })
            .transpose()?;
        let wrapping = text_wrapping
            .as_deref()
            .map(|value| match value.to_ascii_lowercase().as_str() {
                "overflow" => Ok(TextWrapping::Overflow),
                "wrap" => Ok(TextWrapping::Wrap),
                "clip" => Ok(TextWrapping::Clip),
                other => Err(PyValueError::new_err(format!(
                    "text_wrapping must be overflow, wrap, or clip, not {other:?}"
                ))),
            })
            .transpose()?;

        let patch = StylePatch {
            bold,
            italic,
            strikethrough,
            font_family,
            font_size,
            text_color,
            fill_color,
            border: None,
            horizontal_alignment: horizontal,
            vertical_alignment: vertical,
            text_wrapping: wrapping,
            clear: Vec::new(),
        };
        let result = self
            .inner
            .patch_range_style(sheet, range, patch, CalculationOptions { now_serial })
            .map_err(map_style_error)?;
        Ok(PyMutation::from_core(&self.inner, &result))
    }

    fn __len__(&self) -> usize {
        self.inner.sheet_count()
    }

    fn __repr__(&self) -> String {
        format!("Workbook(sheets={})", self.inner.sheet_count())
    }
}

#[pymodule]
fn _betteroffice_xlsx(module: &Bound<'_, PyModule>) -> PyResult<()> {
    let py = module.py();
    module.add("__version__", env!("CARGO_PKG_VERSION"))?;
    module.add_class::<PyWorkbook>()?;
    module.add_class::<PyCellError>()?;
    module.add_class::<PyPng>()?;
    module.add_class::<PyCalculation>()?;
    module.add_class::<PyMutation>()?;
    module.add_class::<PyHistory>()?;
    module.add_class::<PyProposal>()?;
    module.add_class::<PyProposedEdit>()?;
    module.add("XlsxError", py.get_type::<XlsxError>())?;
    module.add("ParseError", py.get_type::<ParseError>())?;
    module.add("RangeError", py.get_type::<RangeError>())?;
    module.add("RenderError", py.get_type::<RenderError>())?;
    module.add("InvalidUpdateError", py.get_type::<InvalidUpdateError>())?;
    module.add(
        "CollaborativeStateError",
        py.get_type::<CollaborativeStateError>(),
    )?;
    module.add("StaleProposalError", py.get_type::<StaleProposalError>())?;
    module.add(
        "NotCollaborativeError",
        py.get_type::<NotCollaborativeError>(),
    )?;
    module.add("MAX_COLLABORATION_BYTES", MAX_COLLABORATION_BYTES)?;
    Ok(())
}
