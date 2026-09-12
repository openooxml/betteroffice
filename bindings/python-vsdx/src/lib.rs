use std::fs;
use std::path::PathBuf;

use betteroffice_vsdx::{Diagram as CoreDiagram, Error as CoreError};
use pyo3::create_exception;
use pyo3::exceptions::PyException;
use pyo3::prelude::*;

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
    fn open(data: &[u8]) -> PyResult<Self> {
        Ok(Self {
            diagram: CoreDiagram::open(data).map_err(map_error)?,
        })
    }

    #[staticmethod]
    fn open_path(py: Python<'_>, path: PathBuf) -> PyResult<Self> {
        let data = py
            .detach(|| fs::read(&path))
            .map_err(|error| VsdxError::new_err(format!("{}: {error}", path.display())))?;
        Self::open(&data)
    }

    #[getter]
    fn pages(&self) -> Vec<PyPage> {
        let package = self.diagram.package();
        package
            .page_contents
            .iter()
            .map(|(path, sheet)| {
                let id = *package.page_part_ids.get(path).unwrap_or(&0);
                PyPage::from_core(id, path, sheet, catalogued_page_name(package, id))
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

fn catalogued_page_name(package: &vsdx_parse::VsdxPackage, page_id: u32) -> Option<String> {
    let pages = package.part_bytes(package.pages_part_path.as_deref()?)?;
    catalogued_page_name_in_pages(pages, page_id)
}

fn catalogued_page_name_in_pages(pages: &[u8], page_id: u32) -> Option<String> {
    let mut offset = 0;
    while let Some(index) = pages[offset..].windows(5).position(|window| window == b"<Page") {
        let start = offset + index;
        let next = pages.get(start + 5)?;
        if !next.is_ascii_whitespace() && *next != b'>' {
            offset = start + 5;
            continue;
        }
        let end = tag_end(pages, start + 5)?;
        let attributes = &pages[start + 5..end];
        if attribute(attributes, b"ID").and_then(|value| value.parse::<u32>().ok()) == Some(page_id) {
            return attribute(attributes, b"Name")
                .or_else(|| attribute(attributes, b"NameU"))
                .map(|value| unescape_xml(&value));
        }
        offset = end + 1;
    }
    None
}

fn tag_end(bytes: &[u8], start: usize) -> Option<usize> {
    let mut quote = None;
    for (index, byte) in bytes.iter().enumerate().skip(start) {
        if let Some(current) = quote {
            if *byte == current {
                quote = None;
            }
        } else if matches!(*byte, b'\'' | b'"') {
            quote = Some(*byte);
        } else if *byte == b'>' {
            return Some(index);
        }
    }
    None
}

fn attribute(bytes: &[u8], name: &[u8]) -> Option<String> {
    let mut offset = 0;
    while offset < bytes.len() {
        while bytes.get(offset).is_some_and(u8::is_ascii_whitespace) {
            offset += 1;
        }
        let key_start = offset;
        while bytes.get(offset).is_some_and(|byte| !byte.is_ascii_whitespace() && *byte != b'=') {
            offset += 1;
        }
        let key = &bytes[key_start..offset];
        while bytes.get(offset).is_some_and(u8::is_ascii_whitespace) {
            offset += 1;
        }
        if bytes.get(offset) != Some(&b'=') {
            offset += 1;
            continue;
        }
        offset += 1;
        while bytes.get(offset).is_some_and(u8::is_ascii_whitespace) {
            offset += 1;
        }
        let quote = *bytes.get(offset)?;
        if !matches!(quote, b'\'' | b'"') {
            return None;
        }
        offset += 1;
        let value_start = offset;
        while bytes.get(offset).is_some_and(|byte| *byte != quote) {
            offset += 1;
        }
        if bytes.get(offset) != Some(&quote) {
            return None;
        }
        if key == name {
            return Some(String::from_utf8_lossy(&bytes[value_start..offset]).into_owned());
        }
        offset += 1;
    }
    None
}

fn unescape_xml(value: &str) -> String {
    let mut result = String::with_capacity(value.len());
    let mut remaining = value;
    while let Some(start) = remaining.find('&') {
        result.push_str(&remaining[..start]);
        let entity = &remaining[start + 1..];
        let Some(end) = entity.find(';') else {
            result.push_str(&remaining[start..]);
            break;
        };
        let name = &entity[..end];
        if let Some(decoded) = decode_xml_entity(name) {
            result.push(decoded);
        } else {
            result.push('&');
            result.push_str(name);
            result.push(';');
        }
        remaining = &entity[end + 1..];
    }
    if !remaining.is_empty() {
        result.push_str(remaining);
    }
    result
}

fn decode_xml_entity(entity: &str) -> Option<char> {
    match entity {
        "quot" => Some('"'),
        "apos" => Some('\''),
        "lt" => Some('<'),
        "gt" => Some('>'),
        "amp" => Some('&'),
        _ => entity
            .strip_prefix("#x")
            .or_else(|| entity.strip_prefix("#X"))
            .and_then(|value| u32::from_str_radix(value, 16).ok())
            .or_else(|| entity.strip_prefix('#').and_then(|value| value.parse().ok()))
            .and_then(char::from_u32),
    }
}

#[cfg(test)]
mod tests {
    use super::catalogued_page_name_in_pages;

    #[test]
    fn page_name_uses_catalogue_name() {
        assert_eq!(
            catalogued_page_name_in_pages(
                br#"<Pages><Page ID='1' NameU='Universal' Name='Display &amp; name'/></Pages>"#,
                1,
            ),
            Some("Display & name".to_owned())
        );
    }

    #[test]
    fn page_name_falls_back_to_catalogue_name_u() {
        assert_eq!(
            catalogued_page_name_in_pages(br#"<Pages><Page ID='1' NameU='Universal'/></Pages>"#, 1),
            Some("Universal".to_owned())
        );
    }

    #[test]
    fn page_name_decodes_numeric_character_references() {
        assert_eq!(
            catalogued_page_name_in_pages(
                br#"<Pages><Page ID='1' Name='A&#x20;B'/><Page ID='2' Name='C&#32;D'/></Pages>"#,
                1,
            ),
            Some("A B".to_owned())
        );
        assert_eq!(
            catalogued_page_name_in_pages(
                br#"<Pages><Page ID='1' Name='A&#x20;B'/><Page ID='2' Name='C&#32;D'/></Pages>"#,
                2,
            ),
            Some("C D".to_owned())
        );
    }

    #[test]
    fn page_name_is_none_without_catalogue_name() {
        assert_eq!(
            catalogued_page_name_in_pages(br#"<Pages><Page ID='1'/></Pages>"#, 1),
            None
        );
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
