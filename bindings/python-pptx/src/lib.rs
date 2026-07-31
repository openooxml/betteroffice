//! PyO3 bindings over the `betteroffice-pptx` facade.

use std::cell::Cell;
use std::collections::BTreeMap;
use std::fs;
use std::path::PathBuf;

use pyo3::create_exception;
use pyo3::exceptions::{PyException, PyIndexError, PyKeyError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyBytes, PyDict, PyInt};
use python_common::{generated_client_id, map_io_error};

use betteroffice_pptx::{
    DeckSnapshot, EditCtx, EditError, EditOrigin, Error as CoreError, MAX_COLLABORATION_BYTES,
    MAX_COLLABORATION_CLIENT_ID, ParagraphSnapshot, ParseLimits, Presentation as CorePresentation,
    PresetShapeDraft, ShapeAdjustReceipt, ShapeDraft, ShapeFillReceipt, ShapeKind, ShapeReceipt,
    ShapeRect, ShapeSnapshot, ShapeStroke, ShapeStrokeReceipt, SlideReceipt, SlideSnapshot,
    StorySnapshot, TextReceipt, TextRunSnapshot, TextStyle, TextStylePatch, TransformReceipt,
};

create_exception!(
    _betteroffice_pptx,
    PptxError,
    PyException,
    "Base class for every error raised by the engine."
);
create_exception!(
    _betteroffice_pptx,
    ParseError,
    PptxError,
    "The presentation could not be read."
);
create_exception!(
    _betteroffice_pptx,
    RangeError,
    PptxError,
    "A slide, shape, or text index was out of bounds."
);
create_exception!(
    _betteroffice_pptx,
    RenderError,
    PptxError,
    "Layout failed or exceeded a resource limit."
);
create_exception!(
    _betteroffice_pptx,
    InvalidUpdateError,
    PptxError,
    "A peer sent an invalid collaboration update."
);
create_exception!(
    _betteroffice_pptx,
    CollaborativeStateError,
    PptxError,
    "The local collaborative state is invalid."
);
create_exception!(
    _betteroffice_pptx,
    NotCollaborativeError,
    PptxError,
    "The operation requires a collaborative presentation."
);
create_exception!(
    _betteroffice_pptx,
    UnsupportedWriteError,
    PptxError,
    "The engine cannot write this change back to PPTX yet."
);

fn map_edit_error(error: EditError, message: String) -> PyErr {
    match error {
        EditError::Parse(_) => ParseError::new_err(message),
        EditError::InvalidState(_) => CollaborativeStateError::new_err(message),
        EditError::InvalidUpdate(_) | EditError::InvalidStateVector(_) => {
            InvalidUpdateError::new_err(message)
        }
        EditError::InvalidClientId(_)
        | EditError::InvalidGeometry(_)
        | EditError::InvalidAdjustment(_) => PyValueError::new_err(message),
        EditError::SlideNotFound(_) | EditError::ShapeNotFound(_) | EditError::StoryNotFound(_) => {
            PyKeyError::new_err(message)
        }
        EditError::OutOfBounds { .. } | EditError::ParagraphBoundary { .. } => {
            RangeError::new_err(message)
        }
        _ => PptxError::new_err(message),
    }
}

fn map_error(error: CoreError) -> PyErr {
    let message = error.to_string();
    match error {
        CoreError::Parse(_) => ParseError::new_err(message),
        CoreError::Render(_) => RenderError::new_err(message),
        CoreError::Edit(edit) => map_edit_error(edit, message),
        _ => PptxError::new_err(message),
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
            "max_shapes" => parsed.max_shapes = value,
            "max_paragraphs" => parsed.max_paragraphs = value,
            "max_runs" => parsed.max_runs = value,
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

fn kind_name(kind: ShapeKind) -> &'static str {
    match kind {
        ShapeKind::Shape => "shape",
        ShapeKind::Picture => "picture",
        ShapeKind::GraphicFrame => "graphicFrame",
        ShapeKind::Group => "group",
    }
}

fn text_style(
    bold: Option<bool>,
    italic: Option<bool>,
    font_size: Option<f64>,
    color: Option<String>,
    font_family: Option<String>,
    underline: Option<String>,
) -> TextStyle {
    TextStyle {
        bold,
        italic,
        font_size_pt: font_size,
        color,
        font_family,
        underline,
    }
}

fn text_style_patch(style: TextStyle) -> TextStylePatch {
    TextStylePatch {
        bold: style.bold,
        italic: style.italic,
        font_size_pt: style.font_size_pt,
        color: style.color,
        font_family: style.font_family,
        underline: style.underline,
    }
}

/// Renders `Option<String>` the way Python would, not the way Rust would.
fn optional_repr(value: &Option<String>) -> String {
    value
        .as_ref()
        .map_or_else(|| "None".to_owned(), |text| format!("{text:?}"))
}

fn index_repr(index: Option<u32>) -> String {
    index.map_or_else(|| "None".to_owned(), |index| index.to_string())
}

fn rect(x: i64, y: i64, width: i64, height: i64) -> ShapeRect {
    ShapeRect {
        x,
        y,
        width,
        height,
    }
}

/// A shape's box in English Metric Units.
#[pyclass(
    module = "betteroffice_pptx",
    name = "Rect",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyRect {
    #[pyo3(get)]
    x: i64,
    #[pyo3(get)]
    y: i64,
    #[pyo3(get)]
    width: i64,
    #[pyo3(get)]
    height: i64,
}

impl PyRect {
    fn from_core(rect: ShapeRect) -> Self {
        Self {
            x: rect.x,
            y: rect.y,
            width: rect.width,
            height: rect.height,
        }
    }
}

#[pymethods]
impl PyRect {
    fn __repr__(&self) -> String {
        format!(
            "Rect(x={}, y={}, width={}, height={})",
            self.x, self.y, self.width, self.height
        )
    }
}

#[pyclass(
    module = "betteroffice_pptx",
    name = "Stroke",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyStroke {
    #[pyo3(get)]
    color: Option<String>,
    #[pyo3(get)]
    width_pt: Option<f64>,
}

impl PyStroke {
    fn from_core(stroke: ShapeStroke) -> Self {
        Self {
            color: stroke.color,
            width_pt: stroke.width_pt,
        }
    }
}

#[pymethods]
impl PyStroke {
    fn __repr__(&self) -> String {
        let width = self
            .width_pt
            .map_or_else(|| "None".to_owned(), |width| width.to_string());
        format!(
            "Stroke(color={}, width_pt={})",
            optional_repr(&self.color),
            width
        )
    }
}

/// One styled run of text inside a paragraph.
#[pyclass(
    module = "betteroffice_pptx",
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
    font_size: Option<f64>,
    #[pyo3(get)]
    color: Option<String>,
    #[pyo3(get)]
    font_family: Option<String>,
    #[pyo3(get)]
    underline: Option<String>,
}

impl PyTextRun {
    fn from_core(run: &TextRunSnapshot) -> Self {
        Self {
            text: run.text.clone(),
            bold: run.style.bold,
            italic: run.style.italic,
            font_size: run.style.font_size_pt,
            color: run.style.color.clone(),
            font_family: run.style.font_family.clone(),
            underline: run.style.underline.clone(),
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
    module = "betteroffice_pptx",
    name = "Paragraph",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyParagraph {
    #[pyo3(get)]
    id: String,
    #[pyo3(get)]
    alignment: Option<String>,
    #[pyo3(get)]
    level: u32,
    #[pyo3(get)]
    runs: Vec<PyTextRun>,
    #[pyo3(get)]
    text: String,
}

impl PyParagraph {
    fn from_core(paragraph: &ParagraphSnapshot) -> Self {
        Self {
            id: paragraph.id.clone(),
            alignment: paragraph.alignment.clone(),
            level: paragraph.level,
            runs: paragraph.runs.iter().map(PyTextRun::from_core).collect(),
            text: paragraph.runs.iter().map(|run| run.text.as_str()).collect(),
        }
    }
}

#[pymethods]
impl PyParagraph {
    fn __repr__(&self) -> String {
        format!("Paragraph(level={}, text={:?})", self.level, self.text)
    }
}

/// One editable text flow. Text offsets are UTF-16 code units, and every
/// paragraph ends with a pilcrow that occupies one of them.
#[pyclass(
    module = "betteroffice_pptx",
    name = "Story",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyStory {
    #[pyo3(get)]
    id: String,
    #[pyo3(get)]
    length: u32,
    #[pyo3(get)]
    paragraphs: Vec<PyParagraph>,
    #[pyo3(get)]
    text: String,
}

impl PyStory {
    fn from_core(story: &StorySnapshot) -> Self {
        Self {
            id: story.id.clone(),
            length: story.length,
            paragraphs: story
                .paragraphs
                .iter()
                .map(PyParagraph::from_core)
                .collect(),
            text: story.plain_text(),
        }
    }
}

#[pymethods]
impl PyStory {
    fn __repr__(&self) -> String {
        format!("Story(id={:?}, length={})", self.id, self.length)
    }
}

#[pyclass(
    module = "betteroffice_pptx",
    name = "Shape",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PyShape {
    #[pyo3(get)]
    id: String,
    #[pyo3(get)]
    name: String,
    #[pyo3(get)]
    kind: &'static str,
    #[pyo3(get)]
    x: i64,
    #[pyo3(get)]
    y: i64,
    #[pyo3(get)]
    width: i64,
    #[pyo3(get)]
    height: i64,
    #[pyo3(get)]
    rotation: f64,
    #[pyo3(get)]
    flip_h: bool,
    #[pyo3(get)]
    flip_v: bool,
    #[pyo3(get)]
    geometry: String,
    #[pyo3(get)]
    adjustments: BTreeMap<String, f64>,
    #[pyo3(get)]
    fill_color: Option<String>,
    #[pyo3(get)]
    outline_color: Option<String>,
    #[pyo3(get)]
    placeholder: Option<String>,
    #[pyo3(get)]
    media_path: Option<String>,
    #[pyo3(get)]
    stories: Vec<PyStory>,
    #[pyo3(get)]
    children: Vec<PyShape>,
    #[pyo3(get)]
    text: String,
}

impl PyShape {
    fn from_core(shape: &ShapeSnapshot) -> Self {
        let stories: Vec<PyStory> = shape.text_stories.iter().map(PyStory::from_core).collect();
        let children: Vec<PyShape> = shape.children.iter().map(PyShape::from_core).collect();
        let text = join_text(
            stories
                .iter()
                .map(|story| story.text.as_str())
                .chain(children.iter().map(|child| child.text.as_str())),
        );
        Self {
            id: shape.id.clone(),
            name: shape.name.clone(),
            kind: kind_name(shape.kind),
            x: shape.x,
            y: shape.y,
            width: shape.width,
            height: shape.height,
            rotation: shape.rotation_deg,
            flip_h: shape.flip_h,
            flip_v: shape.flip_v,
            geometry: shape.geometry.clone(),
            adjustments: shape.adjust_values.clone(),
            fill_color: shape.resolved_fill_color.clone(),
            outline_color: shape.resolved_outline_color.clone(),
            placeholder: shape
                .placeholder
                .as_ref()
                .and_then(|placeholder| placeholder.placeholder_type.clone()),
            media_path: shape.media_part_path.clone(),
            stories,
            children,
            text,
        }
    }
}

#[pymethods]
impl PyShape {
    fn __repr__(&self) -> String {
        format!(
            "Shape(id={:?}, kind={:?}, name={:?})",
            self.id, self.kind, self.name
        )
    }
}

#[pyclass(
    module = "betteroffice_pptx",
    name = "Slide",
    frozen,
    skip_from_py_object
)]
#[derive(Clone)]
pub struct PySlide {
    #[pyo3(get)]
    id: String,
    #[pyo3(get)]
    index: usize,
    #[pyo3(get)]
    name: Option<String>,
    #[pyo3(get)]
    layout: Option<String>,
    #[pyo3(get)]
    part_path: Option<String>,
    #[pyo3(get)]
    shapes: Vec<PyShape>,
    #[pyo3(get)]
    text: String,
}

impl PySlide {
    fn from_core(slide: &SlideSnapshot, index: usize) -> Self {
        let shapes: Vec<PyShape> = slide.shapes.iter().map(PyShape::from_core).collect();
        let text = join_text(shapes.iter().map(|shape| shape.text.as_str()));
        Self {
            id: slide.id.clone(),
            index,
            name: slide.name.clone(),
            layout: slide.layout_part_path.clone(),
            part_path: slide.source_part_path.clone(),
            shapes,
            text,
        }
    }
}

#[pymethods]
impl PySlide {
    fn __len__(&self) -> usize {
        self.shapes.len()
    }

    fn __repr__(&self) -> String {
        format!(
            "Slide(index={}, id={:?}, shapes={})",
            self.index,
            self.id,
            self.shapes.len()
        )
    }
}

/// The whole deck as plain data, as of one moment.
#[pyclass(module = "betteroffice_pptx", name = "Deck", frozen)]
pub struct PyDeck {
    #[pyo3(get)]
    width_emu: i64,
    #[pyo3(get)]
    height_emu: i64,
    #[pyo3(get)]
    slides: Vec<PySlide>,
}

impl PyDeck {
    fn from_core(snapshot: &DeckSnapshot) -> Self {
        Self {
            width_emu: snapshot.width_emu,
            height_emu: snapshot.height_emu,
            slides: snapshot
                .slides
                .iter()
                .enumerate()
                .map(|(index, slide)| PySlide::from_core(slide, index))
                .collect(),
        }
    }
}

#[pymethods]
impl PyDeck {
    fn __len__(&self) -> usize {
        self.slides.len()
    }

    fn __repr__(&self) -> String {
        format!(
            "Deck(width_emu={}, height_emu={}, slides={})",
            self.width_emu,
            self.height_emu,
            self.slides.len()
        )
    }
}

/// One embedded binary part, such as a slide's image.
#[pyclass(module = "betteroffice_pptx", name = "Media", frozen)]
pub struct PyMedia {
    data: Vec<u8>,
    #[pyo3(get)]
    path: String,
    #[pyo3(get)]
    content_type: String,
}

#[pymethods]
impl PyMedia {
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
            "Media(path={:?}, content_type={:?}, bytes={})",
            self.path,
            self.content_type,
            self.data.len()
        )
    }
}

/// A laid-out slide as the renderer's display list. There is no PPTX
/// rasterizer yet, so this is the drawing contract rather than pixels.
#[pyclass(module = "betteroffice_pptx", name = "DisplayList", frozen)]
pub struct PyDisplayList {
    payload: String,
    #[pyo3(get)]
    width: f32,
    #[pyo3(get)]
    height: f32,
    #[pyo3(get)]
    contract_version: u32,
    primitives: usize,
}

#[pymethods]
impl PyDisplayList {
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
        self.primitives
    }

    fn __repr__(&self) -> String {
        format!(
            "DisplayList(width={}, height={}, primitives={})",
            self.width, self.height, self.primitives
        )
    }
}

/// Where a slide moved. A missing index means the slide did not exist on that
/// side of the edit.
#[pyclass(module = "betteroffice_pptx", name = "SlideEdit", frozen)]
pub struct PySlideEdit {
    #[pyo3(get)]
    slide_id: String,
    #[pyo3(get)]
    from_index: Option<u32>,
    #[pyo3(get)]
    to_index: Option<u32>,
}

impl PySlideEdit {
    fn from_core(receipt: SlideReceipt) -> Self {
        Self {
            slide_id: receipt.slide_id,
            from_index: receipt.from_index,
            to_index: receipt.to_index,
        }
    }
}

#[pymethods]
impl PySlideEdit {
    fn __repr__(&self) -> String {
        format!(
            "SlideEdit(slide_id={:?}, from_index={}, to_index={})",
            self.slide_id,
            index_repr(self.from_index),
            index_repr(self.to_index)
        )
    }
}

#[pyclass(module = "betteroffice_pptx", name = "ShapeEdit", frozen)]
pub struct PyShapeEdit {
    #[pyo3(get)]
    slide_id: String,
    #[pyo3(get)]
    shape_id: String,
    #[pyo3(get)]
    index: u32,
}

impl PyShapeEdit {
    fn from_core(receipt: ShapeReceipt) -> Self {
        Self {
            slide_id: receipt.slide_id,
            shape_id: receipt.shape_id,
            index: receipt.index,
        }
    }
}

#[pymethods]
impl PyShapeEdit {
    fn __repr__(&self) -> String {
        format!(
            "ShapeEdit(shape_id={:?}, index={})",
            self.shape_id, self.index
        )
    }
}

#[pyclass(module = "betteroffice_pptx", name = "TransformEdit", frozen)]
pub struct PyTransformEdit {
    #[pyo3(get)]
    slide_id: String,
    #[pyo3(get)]
    shape_id: String,
    #[pyo3(get)]
    before: PyRect,
    #[pyo3(get)]
    after: PyRect,
}

impl PyTransformEdit {
    fn from_core(receipt: TransformReceipt) -> Self {
        Self {
            slide_id: receipt.slide_id,
            shape_id: receipt.shape_id,
            before: PyRect::from_core(receipt.before),
            after: PyRect::from_core(receipt.after),
        }
    }
}

#[pymethods]
impl PyTransformEdit {
    fn __repr__(&self) -> String {
        format!(
            "TransformEdit(shape_id={:?}, after={})",
            self.shape_id,
            self.after.__repr__()
        )
    }
}

#[pyclass(module = "betteroffice_pptx", name = "FillEdit", frozen)]
pub struct PyFillEdit {
    #[pyo3(get)]
    slide_id: String,
    #[pyo3(get)]
    shape_id: String,
    #[pyo3(get)]
    before: Option<String>,
    #[pyo3(get)]
    after: Option<String>,
}

impl PyFillEdit {
    fn from_core(receipt: ShapeFillReceipt) -> Self {
        Self {
            slide_id: receipt.slide_id,
            shape_id: receipt.shape_id,
            before: receipt.before,
            after: receipt.after,
        }
    }
}

#[pymethods]
impl PyFillEdit {
    fn __repr__(&self) -> String {
        format!(
            "FillEdit(shape_id={:?}, before={}, after={})",
            self.shape_id,
            optional_repr(&self.before),
            optional_repr(&self.after)
        )
    }
}

#[pyclass(module = "betteroffice_pptx", name = "StrokeEdit", frozen)]
pub struct PyStrokeEdit {
    #[pyo3(get)]
    slide_id: String,
    #[pyo3(get)]
    shape_id: String,
    #[pyo3(get)]
    before: Option<PyStroke>,
    #[pyo3(get)]
    after: Option<PyStroke>,
}

impl PyStrokeEdit {
    fn from_core(receipt: ShapeStrokeReceipt) -> Self {
        Self {
            slide_id: receipt.slide_id,
            shape_id: receipt.shape_id,
            before: receipt.before.map(PyStroke::from_core),
            after: receipt.after.map(PyStroke::from_core),
        }
    }
}

#[pymethods]
impl PyStrokeEdit {
    fn __repr__(&self) -> String {
        let render = |stroke: &Option<PyStroke>| {
            stroke
                .as_ref()
                .map_or_else(|| "None".to_owned(), PyStroke::__repr__)
        };
        format!(
            "StrokeEdit(shape_id={:?}, before={}, after={})",
            self.shape_id,
            render(&self.before),
            render(&self.after)
        )
    }
}

/// Adjustments the engine clamped into their guide's legal range.
#[pyclass(module = "betteroffice_pptx", name = "AdjustEdit", frozen)]
pub struct PyAdjustEdit {
    #[pyo3(get)]
    slide_id: String,
    #[pyo3(get)]
    shape_id: String,
    #[pyo3(get)]
    before: BTreeMap<String, f64>,
    #[pyo3(get)]
    after: BTreeMap<String, f64>,
}

impl PyAdjustEdit {
    fn from_core(receipt: ShapeAdjustReceipt) -> Self {
        Self {
            slide_id: receipt.slide_id,
            shape_id: receipt.shape_id,
            before: receipt.before,
            after: receipt.after,
        }
    }
}

#[pymethods]
impl PyAdjustEdit {
    fn __repr__(&self) -> String {
        format!(
            "AdjustEdit(shape_id={:?}, after={:?})",
            self.shape_id, self.after
        )
    }
}

/// The affected range, in UTF-16 code units, and the text that occupied it.
#[pyclass(module = "betteroffice_pptx", name = "TextEdit", frozen)]
pub struct PyTextEdit {
    #[pyo3(get)]
    story_id: String,
    #[pyo3(get)]
    start: u32,
    #[pyo3(get)]
    end: u32,
    #[pyo3(get)]
    text: String,
}

impl PyTextEdit {
    fn from_core(receipt: TextReceipt) -> Self {
        Self {
            story_id: receipt.story_id,
            start: receipt.start,
            end: receipt.end,
            text: receipt.text,
        }
    }
}

#[pymethods]
impl PyTextEdit {
    fn __len__(&self) -> usize {
        (self.end - self.start) as usize
    }

    fn __repr__(&self) -> String {
        format!(
            "TextEdit(story_id={:?}, start={}, end={})",
            self.story_id, self.start, self.end
        )
    }
}

fn join_text<'a>(parts: impl Iterator<Item = &'a str>) -> String {
    parts
        .filter(|part| !part.is_empty())
        .collect::<Vec<_>>()
        .join("\n")
}

#[pyclass(module = "betteroffice_pptx", name = "Presentation", unsendable)]
pub struct PyPresentation {
    presentation: CorePresentation,
    author: String,
    origin: EditOrigin,
    edited: Cell<bool>,
    collaborative: bool,
}

impl PyPresentation {
    fn edit_ctx(&self) -> EditCtx {
        EditCtx {
            origin: self.origin,
            author: self.author.clone(),
        }
    }

    /// Marks the deck dirty only once the engine has accepted the edit, so a
    /// raised edit leaves `save` willing to serialize.
    fn committed<T>(&self, edit: Result<T, CoreError>) -> PyResult<T> {
        let receipt = edit.map_err(map_error)?;
        self.edited.set(true);
        Ok(receipt)
    }

    fn all_slide_ids(&self) -> PyResult<Vec<String>> {
        Ok(self
            .presentation
            .snapshot()
            .map_err(map_error)?
            .slides
            .into_iter()
            .map(|slide| slide.id)
            .collect())
    }

    /// bool is an int subclass, so True would otherwise select slide 1.
    fn reject_bool(key: &Bound<'_, PyAny>) -> PyResult<()> {
        if key.is_instance_of::<PyBool>() {
            return Err(PyTypeError::new_err(
                "slide must be an ID (str) or an index (int), not bool",
            ));
        }
        Ok(())
    }

    fn resolve_slide_index(&self, key: &Bound<'_, PyAny>) -> PyResult<usize> {
        Self::reject_bool(key)?;
        let ids = self.all_slide_ids()?;
        if let Ok(id) = key.extract::<String>() {
            return ids
                .iter()
                .position(|candidate| *candidate == id)
                .ok_or_else(|| PyKeyError::new_err(format!("no slide with ID {id:?}")));
        }
        if key.is_instance_of::<PyInt>() {
            let index = key
                .extract::<usize>()
                .ok()
                .filter(|index| *index < ids.len());
            return index.ok_or_else(|| {
                PyIndexError::new_err(format!(
                    "slide index {key} out of range for {} slide(s)",
                    ids.len()
                ))
            });
        }
        Err(PyTypeError::new_err(
            "slide must be an ID (str) or an index (int)",
        ))
    }

    fn resolve_slide_id(&self, key: &Bound<'_, PyAny>) -> PyResult<String> {
        Self::reject_bool(key)?;
        if let Ok(id) = key.extract::<String>() {
            return Ok(id);
        }
        let index = self.resolve_slide_index(key)?;
        self.all_slide_ids()?
            .get(index)
            .cloned()
            .ok_or_else(|| PyIndexError::new_err(format!("slide index {index} out of range")))
    }

    fn wrap(presentation: CorePresentation, collaborative: bool) -> Self {
        Self {
            presentation,
            author: "python".to_owned(),
            origin: EditOrigin::Local,
            edited: Cell::new(false),
            collaborative,
        }
    }

    /// Every non-collaborative deck shares one client ID, so two of them would
    /// author under the same identity and never converge.
    fn require_collaborative(&self) -> PyResult<()> {
        if self.collaborative {
            return Ok(());
        }
        Err(NotCollaborativeError::new_err(
            "this presentation was not opened with open_collaborative, so it \
             has no unique client ID; exchanging updates between standalone \
             decks would silently diverge",
        ))
    }

    /// The engine writes the parsed package, not the edited model, so an
    /// edited deck refuses to serialize rather than dropping the edits.
    fn saved_bytes(&self) -> PyResult<Vec<u8>> {
        if self.edited.get() {
            return Err(UnsupportedWriteError::new_err(
                "this presentation has been edited, and the engine cannot write \
                 model edits back to PPTX yet; saving would drop them",
            ));
        }
        self.presentation.save().map_err(map_error)
    }

    fn open_bytes(
        data: &[u8],
        limits: Option<&Bound<'_, PyDict>>,
        client_id: Option<u64>,
    ) -> PyResult<Self> {
        let limits = parse_limits(limits)?;
        match client_id {
            Some(client_id) => {
                CorePresentation::open_collaborative_with_limits(data, client_id, &limits)
            }
            None => CorePresentation::open_with_limits(data, &limits),
        }
        .map(|presentation| Self::wrap(presentation, client_id.is_some()))
        .map_err(map_error)
    }
}

#[pymethods]
impl PyPresentation {
    #[staticmethod]
    #[pyo3(signature = (data, *, limits = None))]
    fn open(data: &[u8], limits: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        Self::open_bytes(data, limits, None)
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
        Self::open_bytes(&data, limits, None)
    }

    /// Open a replica with a generated ID unless `client_id` is supplied.
    #[staticmethod]
    #[pyo3(signature = (data, *, client_id = None, limits = None))]
    fn open_collaborative(
        data: &[u8],
        client_id: Option<u64>,
        limits: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<Self> {
        // The engine rejects 0, so a masked draw of 0 becomes 1.
        let client_id = match client_id {
            Some(client_id) => client_id,
            None => generated_client_id(MAX_COLLABORATION_CLIENT_ID)?.max(1),
        };
        Self::open_bytes(data, limits, Some(client_id))
    }

    #[getter]
    fn client_id(&self) -> u64 {
        self.presentation.client_id()
    }

    /// Whether this deck was opened with `open_collaborative`. Only then does
    /// it carry a client ID peers can converge against.
    #[getter]
    fn is_collaborative(&self) -> bool {
        self.collaborative
    }

    /// Whether an edit the engine accepted has made `save` refuse.
    #[getter]
    fn is_edited(&self) -> bool {
        self.edited.get()
    }

    #[getter]
    fn author(&self) -> String {
        self.author.clone()
    }

    #[setter]
    fn set_author(&mut self, author: String) {
        self.author = author;
    }

    /// One of local, agent, remote, or system. Only local edits enter the
    /// local undo stack.
    #[getter]
    fn origin(&self) -> &'static str {
        origin_name(self.origin)
    }

    #[setter]
    fn set_origin(&mut self, origin: &str) -> PyResult<()> {
        self.origin = parse_origin(origin)?;
        Ok(())
    }

    #[getter]
    fn slide_count(&self) -> PyResult<usize> {
        Ok(self.all_slide_ids()?.len())
    }

    #[getter]
    fn slide_ids(&self) -> PyResult<Vec<String>> {
        self.all_slide_ids()
    }

    #[getter]
    fn width_emu(&self) -> PyResult<i64> {
        Ok(self.presentation.snapshot().map_err(map_error)?.width_emu)
    }

    #[getter]
    fn height_emu(&self) -> PyResult<i64> {
        Ok(self.presentation.snapshot().map_err(map_error)?.height_emu)
    }

    /// Layout part paths, as `insert_slide` expects them.
    #[getter]
    fn layouts(&self) -> Vec<String> {
        self.presentation
            .layouts()
            .iter()
            .map(|layout| layout.part_path.clone())
            .collect()
    }

    fn snapshot(&self) -> PyResult<PyDeck> {
        let snapshot = self.presentation.snapshot().map_err(map_error)?;
        Ok(PyDeck::from_core(&snapshot))
    }

    fn slide(&self, slide: &Bound<'_, PyAny>) -> PyResult<PySlide> {
        let index = self.resolve_slide_index(slide)?;
        let snapshot = self.presentation.snapshot().map_err(map_error)?;
        let found = snapshot
            .slides
            .get(index)
            .ok_or_else(|| PyIndexError::new_err(format!("slide index {index} out of range")))?;
        Ok(PySlide::from_core(found, index))
    }

    fn story(&self, story_id: &str) -> PyResult<PyStory> {
        let story = self.presentation.story(story_id).map_err(map_error)?;
        Ok(PyStory::from_core(&story))
    }

    fn media(&self) -> Vec<PyMedia> {
        self.presentation
            .media()
            .iter()
            .map(|part| PyMedia {
                data: part.bytes.clone(),
                path: part.part_path.clone(),
                content_type: part.content_type.clone(),
            })
            .collect()
    }

    #[pyo3(signature = (index, *, layout = None))]
    fn insert_slide(&self, index: u32, layout: Option<&str>) -> PyResult<PySlideEdit> {
        let receipt = self.committed(self.presentation.insert_slide(
            &self.edit_ctx(),
            index,
            layout,
        ))?;
        Ok(PySlideEdit::from_core(receipt))
    }

    fn delete_slide(&self, slide: &Bound<'_, PyAny>) -> PyResult<PySlideEdit> {
        let slide_id = self.resolve_slide_id(slide)?;
        let receipt =
            self.committed(self.presentation.delete_slide(&self.edit_ctx(), &slide_id))?;
        Ok(PySlideEdit::from_core(receipt))
    }

    fn move_slide(&self, slide: &Bound<'_, PyAny>, index: u32) -> PyResult<PySlideEdit> {
        let slide_id = self.resolve_slide_id(slide)?;
        let receipt = self.committed(self.presentation.move_slide(
            &self.edit_ctx(),
            &slide_id,
            index,
        ))?;
        Ok(PySlideEdit::from_core(receipt))
    }

    #[pyo3(signature = (
        slide, *, x, y, width, height, text = String::new(), name = None,
        bold = None, italic = None, font_size = None,
        color = None, font_family = None, underline = None
    ))]
    #[allow(clippy::too_many_arguments)]
    fn add_text_box(
        &self,
        slide: &Bound<'_, PyAny>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        text: String,
        name: Option<String>,
        bold: Option<bool>,
        italic: Option<bool>,
        font_size: Option<f64>,
        color: Option<String>,
        font_family: Option<String>,
        underline: Option<String>,
    ) -> PyResult<PyShapeEdit> {
        let slide_id = self.resolve_slide_id(slide)?;
        let draft = ShapeDraft {
            name: name.unwrap_or_else(|| "TextBox".to_owned()),
            rect: rect(x, y, width, height),
            text,
            style: text_style(bold, italic, font_size, color, font_family, underline),
        };
        let receipt = self.committed(self.presentation.add_text_box(
            &self.edit_ctx(),
            &slide_id,
            &draft,
        ))?;
        Ok(PyShapeEdit::from_core(receipt))
    }

    #[pyo3(signature = (slide, geometry, *, x, y, width, height, fill = None, name = None))]
    #[allow(clippy::too_many_arguments)]
    fn add_shape(
        &self,
        slide: &Bound<'_, PyAny>,
        geometry: String,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill: Option<String>,
        name: Option<String>,
    ) -> PyResult<PyShapeEdit> {
        let slide_id = self.resolve_slide_id(slide)?;
        let draft = PresetShapeDraft {
            name: name.unwrap_or_else(|| geometry.clone()),
            geometry,
            rect: rect(x, y, width, height),
            fill,
        };
        let receipt = self.committed(self.presentation.add_shape(
            &self.edit_ctx(),
            &slide_id,
            &draft,
        ))?;
        Ok(PyShapeEdit::from_core(receipt))
    }

    /// `None` clears the fill.
    fn set_shape_fill(
        &self,
        slide: &Bound<'_, PyAny>,
        shape_id: &str,
        color: Option<&str>,
    ) -> PyResult<PyFillEdit> {
        let slide_id = self.resolve_slide_id(slide)?;
        let receipt = self.committed(self.presentation.set_shape_fill(
            &self.edit_ctx(),
            &slide_id,
            shape_id,
            color,
        ))?;
        Ok(PyFillEdit::from_core(receipt))
    }

    /// Omitting both arguments clears the outline.
    #[pyo3(signature = (slide, shape_id, *, color = None, width_pt = None))]
    fn set_shape_stroke(
        &self,
        slide: &Bound<'_, PyAny>,
        shape_id: &str,
        color: Option<String>,
        width_pt: Option<f64>,
    ) -> PyResult<PyStrokeEdit> {
        let slide_id = self.resolve_slide_id(slide)?;
        let stroke = ShapeStroke { color, width_pt };
        let receipt = self.committed(self.presentation.set_shape_stroke(
            &self.edit_ctx(),
            &slide_id,
            shape_id,
            &stroke,
        ))?;
        Ok(PyStrokeEdit::from_core(receipt))
    }

    fn set_shape_adjust(
        &self,
        slide: &Bound<'_, PyAny>,
        shape_id: &str,
        adjustments: BTreeMap<String, f64>,
    ) -> PyResult<PyAdjustEdit> {
        let slide_id = self.resolve_slide_id(slide)?;
        let receipt = self.committed(self.presentation.set_shape_adjust(
            &self.edit_ctx(),
            &slide_id,
            shape_id,
            &adjustments,
        ))?;
        Ok(PyAdjustEdit::from_core(receipt))
    }

    fn remove_shape(&self, slide: &Bound<'_, PyAny>, shape_id: &str) -> PyResult<PyShapeEdit> {
        let slide_id = self.resolve_slide_id(slide)?;
        let receipt = self.committed(self.presentation.remove_shape(
            &self.edit_ctx(),
            &slide_id,
            shape_id,
        ))?;
        Ok(PyShapeEdit::from_core(receipt))
    }

    fn move_shape(
        &self,
        slide: &Bound<'_, PyAny>,
        shape_id: &str,
        x: i64,
        y: i64,
    ) -> PyResult<PyTransformEdit> {
        let slide_id = self.resolve_slide_id(slide)?;
        let receipt = self.committed(self.presentation.move_shape(
            &self.edit_ctx(),
            &slide_id,
            shape_id,
            x,
            y,
        ))?;
        Ok(PyTransformEdit::from_core(receipt))
    }

    fn resize_shape(
        &self,
        slide: &Bound<'_, PyAny>,
        shape_id: &str,
        width: i64,
        height: i64,
    ) -> PyResult<PyTransformEdit> {
        let slide_id = self.resolve_slide_id(slide)?;
        let receipt = self.committed(self.presentation.resize_shape(
            &self.edit_ctx(),
            &slide_id,
            shape_id,
            width,
            height,
        ))?;
        Ok(PyTransformEdit::from_core(receipt))
    }

    /// `index` is a UTF-16 offset into the story.
    #[pyo3(signature = (
        story_id, index, text, *,
        bold = None, italic = None, font_size = None,
        color = None, font_family = None, underline = None
    ))]
    #[allow(clippy::too_many_arguments)]
    fn insert_text(
        &self,
        story_id: &str,
        index: u32,
        text: &str,
        bold: Option<bool>,
        italic: Option<bool>,
        font_size: Option<f64>,
        color: Option<String>,
        font_family: Option<String>,
        underline: Option<String>,
    ) -> PyResult<PyTextEdit> {
        let style = text_style(bold, italic, font_size, color, font_family, underline);
        let receipt = self.committed(self.presentation.insert_text(
            &self.edit_ctx(),
            story_id,
            index,
            text,
            &style,
        ))?;
        Ok(PyTextEdit::from_core(receipt))
    }

    fn delete_text(&self, story_id: &str, start: u32, end: u32) -> PyResult<PyTextEdit> {
        let receipt =
            self.committed(
                self.presentation
                    .delete_text(&self.edit_ctx(), story_id, start, end),
            )?;
        Ok(PyTextEdit::from_core(receipt))
    }

    #[pyo3(signature = (
        story_id, start, end, *,
        bold = None, italic = None, font_size = None,
        color = None, font_family = None, underline = None
    ))]
    #[allow(clippy::too_many_arguments)]
    fn format_text(
        &self,
        story_id: &str,
        start: u32,
        end: u32,
        bold: Option<bool>,
        italic: Option<bool>,
        font_size: Option<f64>,
        color: Option<String>,
        font_family: Option<String>,
        underline: Option<String>,
    ) -> PyResult<PyTextEdit> {
        let patch = text_style_patch(text_style(
            bold,
            italic,
            font_size,
            color,
            font_family,
            underline,
        ));
        let receipt = self.committed(self.presentation.format_text(
            &self.edit_ctx(),
            story_id,
            start,
            end,
            &patch,
        ))?;
        Ok(PyTextEdit::from_core(receipt))
    }

    fn insert_paragraph_break(&self, story_id: &str, index: u32) -> PyResult<PyTextEdit> {
        let receipt = self.committed(self.presentation.insert_paragraph_break(
            &self.edit_ctx(),
            story_id,
            index,
        ))?;
        Ok(PyTextEdit::from_core(receipt))
    }

    /// Register a face for layout. Nothing is embedded in the wheel, so text
    /// only measures against families registered here.
    #[pyo3(signature = (family, data, *, bold = false, italic = false))]
    fn register_font(
        &mut self,
        family: &str,
        data: &[u8],
        bold: bool,
        italic: bool,
    ) -> PyResult<u32> {
        self.presentation
            .register_font(family, bold, italic, data)
            .map_err(map_error)
    }

    fn render_slide(&self, slide: &Bound<'_, PyAny>) -> PyResult<PyDisplayList> {
        let index = self.resolve_slide_index(slide)?;
        let rendered = self.presentation.render_slide(index).map_err(map_error)?;
        let payload = serde_json::to_string(&rendered.display_list)
            .map_err(|error| RenderError::new_err(error.to_string()))?;
        Ok(PyDisplayList {
            payload,
            width: rendered.display_list.width,
            height: rendered.display_list.height,
            contract_version: rendered.display_list.contract_version,
            primitives: rendered.display_list.primitives.len(),
        })
    }

    fn save<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let bytes = self.saved_bytes()?;
        Ok(PyBytes::new(py, &bytes))
    }

    fn save_path(&self, py: Python<'_>, path: PathBuf) -> PyResult<()> {
        let bytes = self.saved_bytes()?;
        py.detach(|| fs::write(&path, bytes))
            .map_err(|error| map_io_error(&error, &path))
    }

    /// This replica's state vector, to hand a peer so it can compute a diff.
    fn state_vector<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        self.require_collaborative()?;
        let bytes = self.presentation.encode_state_vector_v1();
        Ok(PyBytes::new(py, &bytes))
    }

    /// The whole document as one update, for a peer joining from nothing.
    fn state_as_update<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        self.require_collaborative()?;
        let bytes = self.presentation.encode_state_as_update_v1();
        Ok(PyBytes::new(py, &bytes))
    }

    /// The update carrying everything the peer's state vector is missing.
    fn diff<'py>(&self, py: Python<'py>, state_vector: &[u8]) -> PyResult<Bound<'py, PyBytes>> {
        self.require_collaborative()?;
        let update = self
            .presentation
            .encode_diff_v1(state_vector)
            .map_err(map_error)?;
        Ok(PyBytes::new(py, &update))
    }

    fn apply_update(&self, update: &[u8]) -> PyResult<PyDeck> {
        self.require_collaborative()?;
        if update.len() > MAX_COLLABORATION_BYTES {
            return Err(InvalidUpdateError::new_err(format!(
                "collaboration payload is {} bytes, exceeds the {MAX_COLLABORATION_BYTES}-byte limit",
                update.len()
            )));
        }
        let snapshot = self
            .presentation
            .apply_update_v1(update)
            .map_err(map_error)?;
        self.edited.set(true);
        Ok(PyDeck::from_core(&snapshot))
    }

    #[getter]
    fn can_undo(&self) -> bool {
        self.presentation.can_undo()
    }

    #[getter]
    fn can_redo(&self) -> bool {
        self.presentation.can_redo()
    }

    fn undo(&self) -> bool {
        let undone = self.presentation.undo();
        self.edited.set(self.edited.get() || undone);
        undone
    }

    fn redo(&self) -> bool {
        let redone = self.presentation.redo();
        self.edited.set(self.edited.get() || redone);
        redone
    }

    /// Close the current undo step so the next edit starts a new one.
    fn add_undo_barrier(&self) {
        self.presentation.add_undo_barrier();
    }

    fn __len__(&self) -> PyResult<usize> {
        self.slide_count()
    }

    fn __getitem__(&self, slide: &Bound<'_, PyAny>) -> PyResult<PySlide> {
        self.slide(slide)
    }

    fn __repr__(&self) -> PyResult<String> {
        Ok(format!("Presentation(slides={})", self.slide_count()?))
    }
}

#[pymodule]
fn _betteroffice_pptx(module: &Bound<'_, PyModule>) -> PyResult<()> {
    let py = module.py();
    module.add("__version__", env!("CARGO_PKG_VERSION"))?;
    module.add_class::<PyPresentation>()?;
    module.add_class::<PyDeck>()?;
    module.add_class::<PySlide>()?;
    module.add_class::<PyShape>()?;
    module.add_class::<PyStory>()?;
    module.add_class::<PyParagraph>()?;
    module.add_class::<PyTextRun>()?;
    module.add_class::<PyRect>()?;
    module.add_class::<PyStroke>()?;
    module.add_class::<PyMedia>()?;
    module.add_class::<PyDisplayList>()?;
    module.add_class::<PySlideEdit>()?;
    module.add_class::<PyShapeEdit>()?;
    module.add_class::<PyTransformEdit>()?;
    module.add_class::<PyFillEdit>()?;
    module.add_class::<PyStrokeEdit>()?;
    module.add_class::<PyAdjustEdit>()?;
    module.add_class::<PyTextEdit>()?;
    module.add("PptxError", py.get_type::<PptxError>())?;
    module.add("ParseError", py.get_type::<ParseError>())?;
    module.add("RangeError", py.get_type::<RangeError>())?;
    module.add("RenderError", py.get_type::<RenderError>())?;
    module.add("InvalidUpdateError", py.get_type::<InvalidUpdateError>())?;
    module.add(
        "CollaborativeStateError",
        py.get_type::<CollaborativeStateError>(),
    )?;
    module.add(
        "NotCollaborativeError",
        py.get_type::<NotCollaborativeError>(),
    )?;
    module.add(
        "UnsupportedWriteError",
        py.get_type::<UnsupportedWriteError>(),
    )?;
    module.add("MAX_COLLABORATION_BYTES", MAX_COLLABORATION_BYTES)?;
    Ok(())
}
