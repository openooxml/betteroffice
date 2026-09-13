use std::fs;
use std::path::PathBuf;

use betteroffice_vsdx::{Diagram as CoreDiagram, Error as CoreError};
use pyo3::create_exception;
use pyo3::exceptions::PyException;
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyMemoryView};

create_exception!(betteroffice_vsdx, VsdxError, PyException);
create_exception!(betteroffice_vsdx, ParseError, VsdxError);
create_exception!(betteroffice_vsdx, RangeError, VsdxError);
create_exception!(betteroffice_vsdx, RenderError, VsdxError);

fn map_error(error: CoreError) -> PyErr {
    match error {
        CoreError::Parse(error) => ParseError::new_err(error.to_string()),
        CoreError::Resolve(error) => RenderError::new_err(error.to_string()),
        CoreError::Policy(error) => RangeError::new_err(error),
    }
}

#[pyclass(name = "Cell", frozen, skip_from_py_object)]
#[derive(Clone)]
struct PyCell {
    #[pyo3(get)]
    name: String,
    #[pyo3(get)]
    formula: Option<String>,
    #[pyo3(get)]
    value: Option<String>,
    #[pyo3(get)]
    unit: Option<String>,
}

#[pyclass(name = "Connect", frozen, skip_from_py_object)]
#[derive(Clone)]
struct PyConnect {
    #[pyo3(get)]
    from_sheet: u32,
    #[pyo3(get)]
    from_cell: Option<String>,
    #[pyo3(get)]
    from_part: Option<i32>,
    #[pyo3(get)]
    to_sheet: u32,
    #[pyo3(get)]
    to_cell: Option<String>,
    #[pyo3(get)]
    to_part: Option<i32>,
}

#[pyclass(name = "Shape", frozen, skip_from_py_object)]
#[derive(Clone)]
struct PyShape {
    #[pyo3(get)]
    id: u32,
    #[pyo3(get)]
    name: Option<String>,
    #[pyo3(get)]
    text: Option<String>,
    #[pyo3(get)]
    pin_x: Option<f64>,
    #[pyo3(get)]
    pin_y: Option<f64>,
    #[pyo3(get)]
    width: Option<f64>,
    #[pyo3(get)]
    height: Option<f64>,
    #[pyo3(get)]
    cells: Vec<PyCell>,
    #[pyo3(get)]
    children: Vec<PyShape>,
}

#[pyclass(name = "Page", frozen, skip_from_py_object)]
#[derive(Clone)]
struct PyPage {
    #[pyo3(get)]
    id: u32,
    #[pyo3(get)]
    name: Option<String>,
    #[pyo3(get)]
    source_part_path: String,
    #[pyo3(get)]
    shapes: Vec<PyShape>,
    #[pyo3(get)]
    connects: Vec<PyConnect>,
}

#[pyclass(name = "Diagram", unsendable)]
struct PyDiagram {
    diagram: CoreDiagram,
}

impl PyCell {
    fn from_core(cell: &vsdx_parse::Cell) -> Self {
        Self {
            name: cell.name.clone(),
            formula: cell.formula.clone(),
            value: cell.value.clone(),
            unit: cell.unit.clone(),
        }
    }
}

impl PyShape {
    fn from_core(shape: &vsdx_parse::Shape) -> Self {
        let value = |name: &str| {
            shape
                .cells()
                .find(|cell| cell.name == name)
                .and_then(|cell| cell.value.as_deref())
                .and_then(|value| value.parse().ok())
        };
        Self {
            id: shape.id,
            name: shape.name.clone().or_else(|| shape.name_u.clone()),
            text: shape.text().map(|tokens| {
                tokens
                    .iter()
                    .map(|token| match token {
                        vsdx_parse::TextToken::Literal(value) => value.as_str(),
                        vsdx_parse::TextToken::ParagraphRun(_) => "\n",
                        vsdx_parse::TextToken::Tab(_) => "\t",
                        vsdx_parse::TextToken::CharacterRun(_)
                        | vsdx_parse::TextToken::Field(_) => "",
                    })
                    .collect()
            }),
            pin_x: value("PinX"),
            pin_y: value("PinY"),
            width: value("Width"),
            height: value("Height"),
            cells: shape.cells().map(PyCell::from_core).collect(),
            children: shape.shapes().map(PyShape::from_core).collect(),
        }
    }
}

impl PyPage {
    fn from_core(id: u32, path: &str, sheet: &vsdx_parse::Sheet, name: Option<String>) -> Self {
        Self {
            id,
            name,
            source_part_path: path.to_owned(),
            shapes: sheet.shapes().map(PyShape::from_core).collect(),
            connects: sheet
                .connects()
                .map(|connect| PyConnect {
                    from_sheet: connect.from_sheet,
                    from_cell: connect.from_cell.clone(),
                    from_part: connect.from_part,
                    to_sheet: connect.to_sheet,
                    to_cell: connect.to_cell.clone(),
                    to_part: connect.to_part,
                })
                .collect(),
        }
    }
}

#[pymethods]
impl PyDiagram {
    #[staticmethod]
    fn open(py: Python<'_>, data: &Bound<'_, PyAny>) -> PyResult<Self> {
        let bytes = if let Ok(bytes) = data.cast::<PyBytes>() {
            bytes.clone()
        } else {
            PyMemoryView::from(data)?
                .call_method0("tobytes")?
                .cast_into::<PyBytes>()?
        };
        let data = bytes.as_bytes();
        Ok(Self {
            diagram: py.detach(|| CoreDiagram::open(data)).map_err(map_error)?,
        })
    }

    #[staticmethod]
    fn open_path(py: Python<'_>, path: PathBuf) -> PyResult<Self> {
        let data = py
            .detach(|| fs::read(&path))
            .map_err(|error| python_common::map_io_error(&error, &path))?;
        Ok(Self {
            diagram: py.detach(|| CoreDiagram::open(&data)).map_err(map_error)?,
        })
    }

    #[getter]
    fn pages(&self) -> Vec<PyPage> {
        let package = self.diagram.package();
        package
            .page_part_paths
            .iter()
            .filter_map(|path| package.page_contents.get(path).map(|sheet| (path, sheet)))
            .map(|(path, sheet)| {
                let id = *package.page_part_ids.get(path).unwrap_or(&0);
                PyPage::from_core(id, path, sheet, package.page_names.get(&id).cloned())
            })
            .collect()
    }

    fn __len__(&self) -> usize {
        self.diagram.package().page_contents.len()
    }

    fn __repr__(&self) -> String {
        format!("Diagram(pages={})", self.__len__())
    }
}

#[pymodule]
fn _betteroffice_vsdx(module: &Bound<'_, PyModule>) -> PyResult<()> {
    let py = module.py();
    module.add("__version__", env!("CARGO_PKG_VERSION"))?;
    module.add_class::<PyDiagram>()?;
    module.add_class::<PyPage>()?;
    module.add_class::<PyShape>()?;
    module.add_class::<PyCell>()?;
    module.add_class::<PyConnect>()?;
    module.add("VsdxError", py.get_type::<VsdxError>())?;
    module.add("ParseError", py.get_type::<ParseError>())?;
    module.add("RangeError", py.get_type::<RangeError>())?;
    module.add("RenderError", py.get_type::<RenderError>())?;
    Ok(())
}
