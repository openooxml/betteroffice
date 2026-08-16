//! PyO3 bindings over the `betteroffice-docx` facade.

use std::fs;
use std::path::PathBuf;

use pyo3::create_exception;
use pyo3::exceptions::{PyException, PyIndexError, PyKeyError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyBytes, PyDict, PyInt};
use python_common::map_io_error;

use betteroffice_docx::{
    BlockContent, DisplayList, Document as CoreDocument, EditCtx, EditOrigin, Error as CoreError,
    HeaderFooter, ImageScope, InlineNode, LayoutInput, NoteKind, Paragraph, ParagraphContent,
    ParseLimits, Receipt, Run, RunContent, SaveOptions, Section, Table, get_paragraph_text,
};

/// The engine builds and discards an editing document per call, so one client
/// ID serves every edit this binding makes.
const EDIT_CLIENT_ID: u64 = 1;

create_exception!(
    _betteroffice_docx,
    DocxError,
    PyException,
    "Base class for every error raised by the engine."
);
create_exception!(
    _betteroffice_docx,
    ParseError,
    DocxError,
    "The document could not be read."
);
create_exception!(
    _betteroffice_docx,
    EditError,
    DocxError,
    "The engine rejected an edit."
);
create_exception!(
    _betteroffice_docx,
    UnsupportedEditError,
    EditError,
    "The engine cannot rewrite this paragraph yet."
);
create_exception!(
    _betteroffice_docx,
    LayoutError,
    DocxError,
    "Pagination or display-list construction failed."
);
create_exception!(
    _betteroffice_docx,
    RenderError,
    DocxError,
    "Rasterization failed or exceeded a resource limit."
);

fn map_error(error: CoreError) -> PyErr {
    let message = error.to_string();
    match error {
        CoreError::Parse(_) => ParseError::new_err(message),
        CoreError::Edit(_) | CoreError::Operation(_) => EditError::new_err(message),
        CoreError::ParagraphNotFound(_) => PyKeyError::new_err(message),
        CoreError::UnsupportedParagraphEdit(_) => UnsupportedEditError::new_err(message),
        CoreError::Layout(_) | CoreError::DisplayList(_) => LayoutError::new_err(message),
        CoreError::Font(_) | CoreError::Image(_) => PyValueError::new_err(message),
        CoreError::ResourceLimit(_)
        | CoreError::Render(_)
        | CoreError::RenderTooLarge { .. }
        | CoreError::RenderAreaTooLarge { .. } => RenderError::new_err(message),
        _ => DocxError::new_err(message),
    }
}

fn parse_limits(limits: Option<&Bound<'_, PyDict>>) -> PyResult<ParseLimits> {
    let mut parsed = ParseLimits::default();
    let Some(limits) = limits else {
        return Ok(parsed);
    };
    for (key, value) in limits.iter() {
        let key = key
            .extract::<String>()
            .map_err(|_| PyTypeError::new_err("parse limit names must be strings"))?;
        let value = value.extract::<usize>()?;
        match key.as_str() {
            "max_xml_bytes" => parsed.max_xml_bytes = value,
            "max_xml_events" => parsed.max_xml_events = value,
            "max_xml_text_bytes" => parsed.max_xml_text_bytes = value,
            "max_xml_depth" => parsed.max_xml_depth = value,
            "max_attributes_per_element" => parsed.max_attributes_per_element = value,
            "max_attribute_bytes" => parsed.max_attribute_bytes = value,
            "max_relationships" => parsed.max_relationships = value,
            "max_leaf_values" => parsed.max_leaf_values = value,
            "max_blocks" => parsed.max_blocks = value,
            "max_paragraphs" => parsed.max_paragraphs = value,
            "max_tables" => parsed.max_tables = value,
            "max_table_rows" => parsed.max_table_rows = value,
            "max_table_cells" => parsed.max_table_cells = value,
            "max_notes" => parsed.max_notes = value,
            "max_comments" => parsed.max_comments = value,
            "max_nesting_depth" => parsed.max_nesting_depth = value,
            other => {
                return Err(PyValueError::new_err(format!(
                    "unknown parse limit {other:?}"
                )));
            }
        }
    }
    Ok(parsed)
}

fn parse_origin(origin: &str) -> PyResult<EditOrigin> {
    match origin.to_ascii_lowercase().as_str() {
        "local" => Ok(EditOrigin::Local),
        "agent" => Ok(EditOrigin::Agent),
        "remote" => Ok(EditOrigin::Remote),
        "system" => Ok(EditOrigin::System),
        other => Err(PyValueError::new_err(format!(
            "origin must be local, agent, remote, or system, not {other:?}"
        ))),
    }
}

fn origin_name(origin: EditOrigin) -> &'static str {
    match origin {
        EditOrigin::Local => "local",
        EditOrigin::Agent => "agent",
        EditOrigin::Remote => "remote",
        EditOrigin::System => "system",
    }
}

fn extract_page_index(page: &Bound<'_, PyAny>) -> PyResult<usize> {
    page.extract::<usize>()
        .map_err(|_| PyIndexError::new_err(format!("page index {page} out of range")))
}

/// Renders `Option<String>` the way Python would, not the way Rust would.
fn optional_repr(value: &Option<String>) -> String {
    value
        .as_ref()
        .map_or_else(|| "None".to_owned(), |text| format!("{text:?}"))
}

fn join_text<'a>(parts: impl Iterator<Item = &'a str>) -> String {
    parts
        .filter(|part| !part.is_empty())
        .collect::<Vec<_>>()
        .join("\n")
}

/// The text of one run. `separators` mirrors the engine's own paragraph
/// extraction, which keeps tabs and breaks only for top-level runs.
fn run_text(run: &Run, separators: bool) -> String {
    let mut text = String::new();
    for content in &run.content {
        match content {
            RunContent::Text { text: value, .. } => text.push_str(value),
            RunContent::Tab if separators => text.push('\t'),
            RunContent::Break { break_type, .. } if separators => {
                text.push(if break_type.as_deref() == Some("page") {
                    '\u{000c}'
                } else {
                    '\n'
                });
            }
            _ => {}
        }
    }
    text
}

/// Every run whose text `get_paragraph_text` would emit, paired with whether
/// it contributes tabs and breaks.
fn paragraph_runs(paragraph: &Paragraph) -> Vec<(&Run, bool)> {
    let mut runs = Vec::new();
    for content in &paragraph.content {
        let ParagraphContent::Inline(node) = content else {
            continue;
        };
        match node {
            InlineNode::Run(run) => runs.push((run, true)),
            InlineNode::Hyperlink(link) => {
                for child in &link.children {
                    if let InlineNode::Run(run) = child {
                        runs.push((run, false));
                    }
                }
            }
            InlineNode::SimpleField(field) => {
                runs.extend(field.content.iter().map(|run| (run, false)));
            }
            InlineNode::ComplexField(field) => {
                runs.extend(field.field_result.iter().map(|run| (run, false)));
            }
            _ => {}
        }
    }
    runs
}

fn collect_paragraphs<'a>(blocks: &'a [BlockContent], output: &mut Vec<&'a Paragraph>) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => output.push(paragraph),
            BlockContent::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        collect_paragraphs(&cell.content, output);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => collect_paragraphs(&sdt.content, output),
        }
    }
}

fn paragraphs_of(blocks: &[BlockContent]) -> Vec<PyParagraph> {
    let mut found = Vec::new();
    collect_paragraphs(blocks, &mut found);
    found.into_iter().map(PyParagraph::from_core).collect()
}

fn collect_tables<'a>(blocks: &'a [BlockContent], output: &mut Vec<&'a Table>) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(_) => {}
            BlockContent::Table(table) => {
                output.push(table);
                for row in &table.rows {
                    for cell in &row.cells {
                        collect_tables(&cell.content, output);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => collect_tables(&sdt.content, output),
        }
    }
}

fn tables_of(blocks: &[BlockContent]) -> Vec<PyTable> {
    let mut found = Vec::new();
    collect_tables(blocks, &mut found);
    found.into_iter().map(PyTable::from_core).collect()
}

/// One styled run of text inside a paragraph.
#[pyclass(
    module = "betteroffice_docx",
    name = "TextRun",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyTextRun {
    #[pyo3(get)]
    text: String,
    #[pyo3(get)]
    bold: Option<bool>,
    #[pyo3(get)]
    italic: Option<bool>,
    #[pyo3(get)]
    underline: Option<String>,
    #[pyo3(get)]
    strike: Option<bool>,
    #[pyo3(get)]
    font_size: Option<f64>,
    #[pyo3(get)]
    font_family: Option<String>,
    #[pyo3(get)]
    color: Option<String>,
    #[pyo3(get)]
    style_id: Option<String>,
}

impl PyTextRun {
    fn from_core(run: &Run, separators: bool) -> Self {
        let formatting = run.formatting.as_ref();
        Self {
            text: run_text(run, separators),
            bold: formatting.and_then(|format| format.bold),
            italic: formatting.and_then(|format| format.italic),
            underline: formatting
                .and_then(|format| format.underline.as_ref())
                .map(|underline| underline.style.clone()),
            strike: formatting.and_then(|format| format.strike),
            // `w:sz` counts half-points.
            font_size: formatting
                .and_then(|format| format.font_size)
                .map(|size| size / 2.0),
            font_family: formatting
                .and_then(|format| format.font_family.as_ref())
                .and_then(|family| family.ascii.clone().or_else(|| family.h_ansi.clone())),
            color: formatting
                .and_then(|format| format.color.as_ref())
                .and_then(|color| color.rgb.clone()),
            style_id: formatting.and_then(|format| format.style_id.clone()),
        }
    }
}

#[pymethods]
impl PyTextRun {
    fn __repr__(&self) -> String {
        format!("TextRun({:?})", self.text)
    }
}

#[pyclass(
    module = "betteroffice_docx",
    name = "Paragraph",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyParagraph {
    #[pyo3(get)]
    id: Option<String>,
    #[pyo3(get)]
    style: Option<String>,
    #[pyo3(get)]
    alignment: Option<String>,
    #[pyo3(get)]
    runs: Vec<PyTextRun>,
    #[pyo3(get)]
    text: String,
}

impl PyParagraph {
    fn from_core(paragraph: &Paragraph) -> Self {
        let formatting = paragraph.formatting.as_ref();
        Self {
            id: paragraph.para_id.clone(),
            style: formatting.and_then(|format| format.style_id.clone()),
            alignment: formatting.and_then(|format| format.alignment.clone()),
            runs: paragraph_runs(paragraph)
                .into_iter()
                .map(|(run, separators)| PyTextRun::from_core(run, separators))
                .collect(),
            text: get_paragraph_text(paragraph),
        }
    }
}

#[pymethods]
impl PyParagraph {
    fn __repr__(&self) -> String {
        format!(
            "Paragraph(id={}, text={:?})",
            optional_repr(&self.id),
            self.text
        )
    }
}

#[pyclass(
    module = "betteroffice_docx",
    name = "TableCell",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyTableCell {
    #[pyo3(get)]
    paragraphs: Vec<PyParagraph>,
    #[pyo3(get)]
    tables: Vec<PyTable>,
    #[pyo3(get)]
    text: String,
}

impl PyTableCell {
    fn from_core(cell: &betteroffice_docx::TableCell) -> Self {
        let paragraphs = paragraphs_of(&cell.content);
        let tables = tables_of(&cell.content);
        let text = join_text(paragraphs.iter().map(|paragraph| paragraph.text.as_str()));
        Self {
            paragraphs,
            tables,
            text,
        }
    }
}

#[pymethods]
impl PyTableCell {
    fn __repr__(&self) -> String {
        format!("TableCell(text={:?})", self.text)
    }
}

#[pyclass(
    module = "betteroffice_docx",
    name = "TableRow",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyTableRow {
    #[pyo3(get)]
    cells: Vec<PyTableCell>,
    #[pyo3(get)]
    text: String,
}

impl PyTableRow {
    fn from_core(row: &betteroffice_docx::TableRow) -> Self {
        let cells: Vec<PyTableCell> = row.cells.iter().map(PyTableCell::from_core).collect();
        let text = cells
            .iter()
            .map(|cell| cell.text.as_str())
            .collect::<Vec<_>>()
            .join("\t");
        Self { cells, text }
    }
}

#[pymethods]
impl PyTableRow {
    fn __len__(&self) -> usize {
        self.cells.len()
    }

    fn __repr__(&self) -> String {
        format!("TableRow(cells={})", self.cells.len())
    }
}

#[pyclass(
    module = "betteroffice_docx",
    name = "Table",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyTable {
    #[pyo3(get)]
    rows: Vec<PyTableRow>,
    #[pyo3(get)]
    column_widths: Option<Vec<f64>>,
    #[pyo3(get)]
    text: String,
}

impl PyTable {
    fn from_core(table: &Table) -> Self {
        let rows: Vec<PyTableRow> = table.rows.iter().map(PyTableRow::from_core).collect();
        let text = join_text(rows.iter().map(|row| row.text.as_str()));
        Self {
            rows,
            column_widths: table.column_widths.clone(),
            text,
        }
    }
}

#[pymethods]
impl PyTable {
    fn __len__(&self) -> usize {
        self.rows.len()
    }

    fn __repr__(&self) -> String {
        format!("Table(rows={})", self.rows.len())
    }
}

/// One section's page setup and content. Lengths are twips: 1440 to the inch.
#[pyclass(
    module = "betteroffice_docx",
    name = "Section",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PySection {
    #[pyo3(get)]
    id: Option<String>,
    #[pyo3(get)]
    page_width: Option<f64>,
    #[pyo3(get)]
    page_height: Option<f64>,
    #[pyo3(get)]
    orientation: Option<String>,
    #[pyo3(get)]
    margin_top: Option<f64>,
    #[pyo3(get)]
    margin_bottom: Option<f64>,
    #[pyo3(get)]
    margin_left: Option<f64>,
    #[pyo3(get)]
    margin_right: Option<f64>,
    #[pyo3(get)]
    paragraphs: Vec<PyParagraph>,
    #[pyo3(get)]
    text: String,
}

impl PySection {
    fn from_core(section: &Section) -> Self {
        let properties = &section.properties;
        let paragraphs = paragraphs_of(&section.content);
        let text = join_text(paragraphs.iter().map(|paragraph| paragraph.text.as_str()));
        Self {
            id: section.id.clone().or_else(|| properties.section_id.clone()),
            page_width: properties.page_width,
            page_height: properties.page_height,
            orientation: properties.orientation.clone(),
            margin_top: properties.margin_top,
            margin_bottom: properties.margin_bottom,
            margin_left: properties.margin_left,
            margin_right: properties.margin_right,
            paragraphs,
            text,
        }
    }
}

#[pymethods]
impl PySection {
    fn __repr__(&self) -> String {
        format!(
            "Section(id={}, paragraphs={})",
            optional_repr(&self.id),
            self.paragraphs.len()
        )
    }
}

#[pyclass(
    module = "betteroffice_docx",
    name = "HeaderFooter",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyHeaderFooter {
    #[pyo3(get)]
    rel_id: String,
    #[pyo3(get)]
    kind: String,
    #[pyo3(get)]
    paragraphs: Vec<PyParagraph>,
    #[pyo3(get)]
    text: String,
}

impl PyHeaderFooter {
    fn from_core(entry: &(String, HeaderFooter)) -> Self {
        let (rel_id, story) = entry;
        let paragraphs = paragraphs_of(&story.content);
        let text = join_text(paragraphs.iter().map(|paragraph| paragraph.text.as_str()));
        Self {
            rel_id: rel_id.clone(),
            kind: story.hdr_ftr_type.clone(),
            paragraphs,
            text,
        }
    }
}

#[pymethods]
impl PyHeaderFooter {
    fn __repr__(&self) -> String {
        format!(
            "HeaderFooter(rel_id={:?}, kind={:?})",
            self.rel_id, self.kind
        )
    }
}

/// How much of each kind the document holds.
#[pyclass(module = "betteroffice_docx", name = "Structure", frozen)]
pub struct PyStructure {
    #[pyo3(get)]
    body_paragraphs: usize,
    #[pyo3(get)]
    body_tables: usize,
    #[pyo3(get)]
    sections: usize,
    #[pyo3(get)]
    headers: usize,
    #[pyo3(get)]
    footers: usize,
    #[pyo3(get)]
    footnotes: usize,
    #[pyo3(get)]
    endnotes: usize,
}

#[pymethods]
impl PyStructure {
    fn __repr__(&self) -> String {
        format!(
            "Structure(body_paragraphs={}, body_tables={}, sections={})",
            self.body_paragraphs, self.body_tables, self.sections
        )
    }
}

/// What an edit touched. Offsets are UTF-16 code units inside the paragraph.
#[pyclass(module = "betteroffice_docx", name = "Edit", frozen)]
pub struct PyEdit {
    #[pyo3(get)]
    para_id: Option<String>,
    #[pyo3(get)]
    story: Option<String>,
    #[pyo3(get)]
    start: Option<u32>,
    #[pyo3(get)]
    end: Option<u32>,
    #[pyo3(get)]
    new_para_ids: Vec<String>,
    #[pyo3(get)]
    revision_ids: Vec<String>,
}

impl PyEdit {
    fn from_core(receipt: Receipt) -> Self {
        let range = receipt.range;
        Self {
            para_id: range.as_ref().map(|range| range.start.para.clone()),
            story: range.as_ref().map(|range| range.start.story.clone()),
            start: range.as_ref().map(|range| range.start.offset),
            end: range.as_ref().map(|range| range.end.offset),
            new_para_ids: receipt.new_para_ids,
            revision_ids: receipt.revision_ids,
        }
    }
}

#[pymethods]
impl PyEdit {
    fn __repr__(&self) -> String {
        format!(
            "Edit(para_id={}, start={}, end={})",
            optional_repr(&self.para_id),
            offset_repr(self.start),
            offset_repr(self.end)
        )
    }
}

fn offset_repr(offset: Option<u32>) -> String {
    offset.map_or_else(|| "None".to_owned(), |offset| offset.to_string())
}

/// A paginated document as the renderer's drawing contract.
#[pyclass(
    module = "betteroffice_docx",
    name = "DisplayList",
    frozen,
    from_py_object
)]
#[derive(Clone)]
pub struct PyDisplayList {
    inner: DisplayList,
    payload: String,
    #[pyo3(get)]
    contract_version: Option<u32>,
}

impl PyDisplayList {
    fn new(inner: DisplayList) -> PyResult<Self> {
        let payload = serde_json::to_string(&inner)
            .map_err(|error| LayoutError::new_err(error.to_string()))?;
        Ok(Self {
            contract_version: inner.contract_version,
            inner,
            payload,
        })
    }
}

#[pymethods]
impl PyDisplayList {
    /// Adopt a display list produced elsewhere, so it can be rasterized here.
    #[staticmethod]
    fn from_json(text: &str) -> PyResult<Self> {
        let inner: DisplayList = serde_json::from_str(text)
            .map_err(|error| PyValueError::new_err(format!("invalid display list: {error}")))?;
        Self::new(inner)
    }

    #[getter]
    fn json(&self) -> &str {
        &self.payload
    }

    #[getter]
    fn primitives(&self) -> usize {
        self.inner
            .pages
            .iter()
            .map(|page| page.primitives.len())
            .sum()
    }

    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        py.import("json")?
            .call_method1("loads", (self.payload.as_str(),))
    }

    fn write(&self, py: Python<'_>, path: PathBuf) -> PyResult<()> {
        py.detach(|| fs::write(&path, self.payload.as_bytes()))
            .map_err(|error| map_io_error(&error, &path))
    }

    fn __len__(&self) -> usize {
        self.inner.pages.len()
    }

    fn __repr__(&self) -> String {
        format!(
            "DisplayList(pages={}, primitives={})",
            self.inner.pages.len(),
            self.primitives()
        )
    }
}

/// A paginated document: the page boxes plus the display list that paints them.
#[pyclass(module = "betteroffice_docx", name = "Layout", frozen)]
pub struct PyLayout {
    payload: String,
    #[pyo3(get)]
    pages: usize,
    #[pyo3(get)]
    display_list: PyDisplayList,
}

#[pymethods]
impl PyLayout {
    #[getter]
    fn json(&self) -> &str {
        &self.payload
    }

    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        py.import("json")?
            .call_method1("loads", (self.payload.as_str(),))
    }

    fn write(&self, py: Python<'_>, path: PathBuf) -> PyResult<()> {
        py.detach(|| fs::write(&path, self.payload.as_bytes()))
            .map_err(|error| map_io_error(&error, &path))
    }

    fn __len__(&self) -> usize {
        self.pages
    }

    fn __repr__(&self) -> String {
        format!("Layout(pages={})", self.pages)
    }
}

/// One rasterized page. `skipped_images` counts the image references the
/// backend could not resolve and left undrawn.
#[pyclass(module = "betteroffice_docx", name = "Png", frozen)]
pub struct PyPng {
    data: Vec<u8>,
    #[pyo3(get)]
    skipped_images: usize,
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
            "Png(bytes={}, skipped_images={})",
            self.data.len(),
            self.skipped_images
        )
    }
}

#[pyclass(module = "betteroffice_docx", name = "Document")]
pub struct PyDocument {
    inner: CoreDocument,
    author: String,
    origin: EditOrigin,
    timestamp: String,
}

impl PyDocument {
    fn wrap(inner: CoreDocument) -> Self {
        Self {
            inner,
            author: "python".to_owned(),
            origin: EditOrigin::Local,
            timestamp: SaveOptions::default().now,
        }
    }

    fn open_owned(py: Python<'_>, data: Vec<u8>, limits: ParseLimits) -> PyResult<Self> {
        py.detach(|| CoreDocument::open_with_limits(&data, &limits))
            .map(Self::wrap)
            .map_err(map_error)
    }

    /// Copies first: a borrow of a Python buffer cannot cross `detach`.
    fn open_bytes(
        py: Python<'_>,
        data: &[u8],
        limits: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<Self> {
        let limits = parse_limits(limits)?;
        let data = data.to_vec();
        Self::open_owned(py, data, limits)
    }

    fn edit_ctx(&self) -> EditCtx {
        EditCtx {
            author: self.author.clone(),
            origin: self.origin,
            suggesting: None,
            now_iso: self.timestamp.clone(),
        }
    }

    fn save_options(
        &self,
        now: Option<String>,
        update_modified_date: bool,
        modified_by: Option<String>,
    ) -> SaveOptions {
        SaveOptions {
            now: now.unwrap_or_else(|| self.timestamp.clone()),
            update_modified_date,
            modified_by,
        }
    }

    fn resolve_paragraph(&self, key: &Bound<'_, PyAny>) -> PyResult<PyParagraph> {
        // bool is an int subclass, so True would otherwise select paragraph 1.
        if key.is_instance_of::<PyBool>() {
            return Err(PyTypeError::new_err(
                "paragraph must be an ID (str) or an index (int), not bool",
            ));
        }
        let paragraphs = self.inner.paragraphs();
        if let Ok(id) = key.extract::<String>() {
            return paragraphs
                .into_iter()
                .find(|paragraph| paragraph.para_id.as_deref() == Some(id.as_str()))
                .map(PyParagraph::from_core)
                .ok_or_else(|| PyKeyError::new_err(format!("no paragraph with ID {id:?}")));
        }
        if key.is_instance_of::<PyInt>() {
            return key
                .extract::<usize>()
                .ok()
                .and_then(|index| {
                    paragraphs
                        .get(index)
                        .map(|found| PyParagraph::from_core(found))
                })
                .ok_or_else(|| {
                    PyIndexError::new_err(format!(
                        "paragraph index {key} out of range for {} paragraph(s)",
                        paragraphs.len()
                    ))
                });
        }
        Err(PyTypeError::new_err(
            "paragraph must be an ID (str) or an index (int)",
        ))
    }

    fn image_scope<'a>(scope: &str, part: Option<&'a str>) -> PyResult<ImageScope<'a>> {
        match scope.to_ascii_lowercase().as_str() {
            "body" => Ok(ImageScope::Body),
            "header_footer" => part.map(ImageScope::HeaderFooter).ok_or_else(|| {
                PyValueError::new_err("the header_footer scope needs the part's relationship id")
            }),
            "footnotes" => Ok(ImageScope::Notes(NoteKind::Footnote)),
            "endnotes" => Ok(ImageScope::Notes(NoteKind::Endnote)),
            other => Err(PyValueError::new_err(format!(
                "scope must be body, header_footer, footnotes, or endnotes, not {other:?}"
            ))),
        }
    }
}

#[pymethods]
impl PyDocument {
    #[staticmethod]
    #[pyo3(signature = (data, *, limits = None))]
    fn open(py: Python<'_>, data: &[u8], limits: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        Self::open_bytes(py, data, limits)
    }

    #[staticmethod]
    #[pyo3(signature = (path, *, limits = None))]
    fn open_path(
        py: Python<'_>,
        path: PathBuf,
        limits: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<Self> {
        let data = py
            .detach(|| fs::read(&path))
            .map_err(|error| map_io_error(&error, &path))?;
        let limits = parse_limits(limits)?;
        Self::open_owned(py, data, limits)
    }

    /// Stamped on every edit this document makes.
    #[getter]
    fn author(&self) -> String {
        self.author.clone()
    }

    #[setter]
    fn set_author(&mut self, author: String) {
        self.author = author;
    }

    /// One of local, agent, remote, or system.
    #[getter]
    fn origin(&self) -> &'static str {
        origin_name(self.origin)
    }

    #[setter]
    fn set_origin(&mut self, origin: &str) -> PyResult<()> {
        self.origin = parse_origin(origin)?;
        Ok(())
    }

    /// The ISO-8601 clock reading edits and saves record. The engine has no
    /// clock of its own, so it defaults to the epoch and output stays
    /// reproducible until you set one.
    #[getter]
    fn timestamp(&self) -> String {
        self.timestamp.clone()
    }

    #[setter]
    fn set_timestamp(&mut self, timestamp: String) {
        self.timestamp = timestamp;
    }

    /// Body paragraph IDs in document order. A paragraph Word never stamped
    /// with a `w14:paraId` reads as `None` and cannot be edited by ID.
    #[getter]
    fn paragraph_ids(&self) -> Vec<Option<String>> {
        self.inner
            .paragraphs()
            .into_iter()
            .map(|paragraph| paragraph.para_id.clone())
            .collect()
    }

    /// What the parser could not represent. Parsing still succeeded.
    #[getter]
    fn warnings(&self) -> Vec<String> {
        self.inner.model().warnings.clone()
    }

    /// `{{name}}` placeholders the parser found in body text.
    #[getter]
    fn template_variables(&self) -> Vec<String> {
        self.inner.model().template_variables.clone()
    }

    #[getter]
    fn text(&self) -> String {
        let paragraphs: Vec<String> = self
            .inner
            .paragraphs()
            .into_iter()
            .map(get_paragraph_text)
            .collect();
        join_text(paragraphs.iter().map(String::as_str))
    }

    fn structure(&self) -> PyStructure {
        let structure = self.inner.structure();
        PyStructure {
            body_paragraphs: structure.body_paragraphs,
            body_tables: structure.body_tables,
            sections: structure.sections,
            headers: structure.headers,
            footers: structure.footers,
            footnotes: structure.footnotes,
            endnotes: structure.endnotes,
        }
    }

    /// Every paragraph in the body, table cells and content controls included.
    fn paragraphs(&self) -> Vec<PyParagraph> {
        self.inner
            .paragraphs()
            .into_iter()
            .map(PyParagraph::from_core)
            .collect()
    }

    fn tables(&self) -> Vec<PyTable> {
        self.inner
            .tables()
            .into_iter()
            .map(PyTable::from_core)
            .collect()
    }

    fn sections(&self) -> Vec<PySection> {
        self.inner
            .sections()
            .iter()
            .map(PySection::from_core)
            .collect()
    }

    fn headers(&self) -> Vec<PyHeaderFooter> {
        self.inner
            .headers()
            .iter()
            .map(PyHeaderFooter::from_core)
            .collect()
    }

    fn footers(&self) -> Vec<PyHeaderFooter> {
        self.inner
            .footers()
            .iter()
            .map(PyHeaderFooter::from_core)
            .collect()
    }

    /// One paragraph by ID or by index into the body.
    fn paragraph(&self, paragraph: &Bound<'_, PyAny>) -> PyResult<PyParagraph> {
        self.resolve_paragraph(paragraph)
    }

    /// Rewrite one paragraph's text, keeping its style and run formatting.
    ///
    /// Raises `KeyError` for an unknown ID and `UnsupportedEditError` for a
    /// paragraph the engine cannot rebuild from a single run.
    fn replace_text(&mut self, para_id: &str, text: &str) -> PyResult<PyEdit> {
        let context = self.edit_ctx();
        let receipt = self
            .inner
            .replace_paragraph_text_with(para_id, text, EDIT_CLIENT_ID, &context)
            .map_err(map_error)?;
        Ok(PyEdit::from_core(receipt))
    }

    /// Paginate a `{"measured": [...], "options": {...}}` envelope.
    ///
    /// The engine paginates blocks that were already measured; it does not
    /// measure them, so the envelope carries the metrics.
    fn layout(&self, py: Python<'_>, input: &Bound<'_, PyAny>) -> PyResult<PyLayout> {
        let text = match input.extract::<String>() {
            Ok(text) => text,
            Err(_) => py
                .import("json")?
                .call_method1("dumps", (input,))?
                .extract::<String>()?,
        };
        let input: LayoutInput = serde_json::from_str(&text)
            .map_err(|error| PyValueError::new_err(format!("invalid layout input: {error}")))?;
        let result = py.detach(|| self.inner.layout(input)).map_err(map_error)?;
        let payload = serde_json::to_string(&result.layout)
            .map_err(|error| LayoutError::new_err(error.to_string()))?;
        Ok(PyLayout {
            payload,
            pages: result.layout.pages.len(),
            display_list: PyDisplayList::new(result.display_list)?,
        })
    }

    /// Register a face for rasterization. No face is embedded in the wheel, so
    /// text paints only in families registered here.
    #[pyo3(signature = (family, data, *, bold = false, italic = false))]
    fn register_font(
        &mut self,
        family: &str,
        data: &[u8],
        bold: bool,
        italic: bool,
    ) -> PyResult<u32> {
        self.inner
            .register_font(family, bold, italic, data)
            .map_err(map_error)
    }

    /// Supply the bytes behind one relationship id the display list left
    /// unresolved, scoped to the part that owns the id.
    #[pyo3(signature = (rel_id, data, *, scope = "body", part = None))]
    fn register_image(
        &mut self,
        rel_id: &str,
        data: &[u8],
        scope: &str,
        part: Option<&str>,
    ) -> PyResult<()> {
        let scope = Self::image_scope(scope, part)?;
        self.inner
            .register_image(scope, rel_id, data)
            .map_err(map_error)
    }

    /// Rasterize one display-list page to deterministic PNG bytes.
    #[pyo3(signature = (display_list, page = 0))]
    fn render_png(
        &self,
        py: Python<'_>,
        display_list: &PyDisplayList,
        #[pyo3(from_py_with = extract_page_index)] page: usize,
    ) -> PyResult<PyPng> {
        let list = &display_list.inner;
        let rendered = py
            .detach(|| self.inner.render_png(list, page))
            .map_err(map_error)?;
        Ok(PyPng {
            data: rendered.bytes,
            skipped_images: rendered.skipped_images,
        })
    }

    #[pyo3(signature = (*, now = None, update_modified_date = false, modified_by = None))]
    fn save<'py>(
        &self,
        py: Python<'py>,
        now: Option<String>,
        update_modified_date: bool,
        modified_by: Option<String>,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let options = self.save_options(now, update_modified_date, modified_by);
        let bytes = py
            .detach(|| self.inner.save_with_options(options))
            .map_err(map_error)?;
        Ok(PyBytes::new(py, &bytes))
    }

    #[pyo3(signature = (path, *, now = None, update_modified_date = false, modified_by = None))]
    fn save_path(
        &self,
        py: Python<'_>,
        path: PathBuf,
        now: Option<String>,
        update_modified_date: bool,
        modified_by: Option<String>,
    ) -> PyResult<()> {
        enum Failure {
            Engine(CoreError),
            Io(std::io::Error),
        }

        let options = self.save_options(now, update_modified_date, modified_by);
        py.detach(|| {
            let bytes = self
                .inner
                .save_with_options(options)
                .map_err(Failure::Engine)?;
            fs::write(&path, bytes).map_err(Failure::Io)
        })
        .map_err(|failure| match failure {
            Failure::Engine(error) => map_error(error),
            Failure::Io(error) => map_io_error(&error, &path),
        })
    }

    fn __len__(&self) -> usize {
        self.inner.paragraphs().len()
    }

    fn __getitem__(&self, paragraph: &Bound<'_, PyAny>) -> PyResult<PyParagraph> {
        self.resolve_paragraph(paragraph)
    }

    fn __repr__(&self) -> String {
        let structure = self.inner.structure();
        format!(
            "Document(paragraphs={}, tables={}, sections={})",
            structure.body_paragraphs, structure.body_tables, structure.sections
        )
    }
}

#[pymodule]
fn _betteroffice_docx(module: &Bound<'_, PyModule>) -> PyResult<()> {
    let py = module.py();
    module.add("__version__", env!("CARGO_PKG_VERSION"))?;
    module.add_class::<PyDocument>()?;
    module.add_class::<PyParagraph>()?;
    module.add_class::<PyTextRun>()?;
    module.add_class::<PyTable>()?;
    module.add_class::<PyTableRow>()?;
    module.add_class::<PyTableCell>()?;
    module.add_class::<PySection>()?;
    module.add_class::<PyHeaderFooter>()?;
    module.add_class::<PyStructure>()?;
    module.add_class::<PyEdit>()?;
    module.add_class::<PyLayout>()?;
    module.add_class::<PyDisplayList>()?;
    module.add_class::<PyPng>()?;
    module.add("DocxError", py.get_type::<DocxError>())?;
    module.add("ParseError", py.get_type::<ParseError>())?;
    module.add("EditError", py.get_type::<EditError>())?;
    module.add(
        "UnsupportedEditError",
        py.get_type::<UnsupportedEditError>(),
    )?;
    module.add("LayoutError", py.get_type::<LayoutError>())?;
    module.add("RenderError", py.get_type::<RenderError>())?;
    module.add("MAX_PIXMAP_DIM", betteroffice_docx::MAX_PIXMAP_DIM)?;
    module.add("MAX_PIXMAP_PIXELS", betteroffice_docx::MAX_PIXMAP_PIXELS)?;
    Ok(())
}
