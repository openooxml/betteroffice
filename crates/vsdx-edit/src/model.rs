use serde::{Deserialize, Deserializer, Serialize, Serializer};
use thiserror::Error;
use vsdx_parse::{CellLocator, CellRow, CellSheet, MutationGesture};

#[derive(Clone, Copy, Debug, Default, Eq, Hash, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum EditOrigin {
    #[default]
    Local,
    Agent,
    Remote,
    System,
}

#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct EditCtx {
    pub origin: EditOrigin,
    pub author: String,
}

impl EditCtx {
    pub fn local(author: impl Into<String>) -> Self {
        Self {
            origin: EditOrigin::Local,
            author: author.into(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellSnapshot {
    #[serde(
        serialize_with = "serialize_cell_locator",
        deserialize_with = "deserialize_cell_locator"
    )]
    pub locator: CellLocator,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_type: Option<String>,
    pub name: String,
    pub formula: Option<String>,
    #[serde(default)]
    pub value: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeSnapshot {
    pub id: String,
    pub source_id: u32,
    pub name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub master: Option<u32>,
    pub cells: Vec<CellSnapshot>,
    pub children: Vec<ShapeSnapshot>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub copy_source_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub copy_source_page_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub copy_refusal: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageSnapshot {
    pub id: String,
    pub source_part_path: String,
    pub name: Option<String>,
    pub shapes: Vec<ShapeSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DiagramSnapshot {
    pub pages: Vec<PageSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellFormulaReceipt {
    pub page_id: String,
    pub shape_id: String,
    pub cell_name: String,
    pub before: Option<String>,
    pub after: String,
}

/// One shape's pin move in a batch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ShapeMove {
    pub page_id: String,
    pub shape_id: String,
    pub x: String,
    pub y: String,
}

/// One shape's deletion in a batch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ShapeDelete {
    pub page_id: String,
    pub shape_id: String,
}

/// One cell write in a batch that may span shapes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellFormulaWrite {
    pub page_id: String,
    pub shape_id: String,
    pub cell_name: String,
    pub formula: String,
}

/// One shape-data row write in a batch against a single shape.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ShapeDataWrite {
    pub row: CellRow,
    pub section_index: Option<u32>,
    pub formula: String,
}

/// Per-row decision; `after` is present only when the batch succeeds.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeDataReceipt {
    pub page_id: String,
    pub shape_id: String,
    pub row_name: Option<String>,
    pub row_index: Option<u32>,
    pub section_index: Option<u32>,
    pub before: Option<String>,
    pub after: Option<String>,
    pub refusal: Option<String>,
}

impl ShapeDataReceipt {
    pub fn refused(&self) -> bool {
        self.refusal.is_some()
    }
}

/// One cell to probe, with the gesture to probe it as; the cell name picks one when absent.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellWriteQuery {
    pub locator: CellLocator,
    pub gesture: Option<MutationGesture>,
}

/// What the mutation policy would do with a write to one cell, without writing it.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellWriteProbe {
    pub cell_name: String,
    pub allowed: bool,
    /// Cell the write would land on after SETATREF redirection; absent when refused.
    pub target_cell_name: Option<String>,
    /// `guard`, `lock` or `unsupported`; absent when the write is allowed.
    pub refusal: Option<String>,
    pub reason: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeReceipt {
    pub page_id: String,
    pub shape_id: String,
    pub from_index: Option<u32>,
    pub to_index: Option<u32>,
}

/// Typed receipt for a committed shape-text edit.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextReceipt {
    pub page_id: String,
    pub shape_id: String,
    pub before: String,
    pub after: String,
}

/// Paired receipts for a shape inserted with its connector in one transaction.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConnectedShapeReceipt {
    pub shape: ShapeReceipt,
    pub connector: ShapeReceipt,
}

/// A hand-routed connector path written in scene-space points.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConnectorRouteReceipt {
    pub page_id: String,
    pub shape_id: String,
    pub points: u32,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeDraft {
    pub name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub master: Option<u32>,
    pub cells: Vec<CellSnapshot>,
}

/// A pasted group subtree; every node names its live copy source.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeTreeDraft {
    pub name: Option<String>,
    pub cells: Vec<CellSnapshot>,
    #[serde(default)]
    pub text: String,
    #[serde(default)]
    pub copy_source_id: Option<u32>,
    #[serde(default)]
    pub copy_source_page_id: Option<u32>,
    #[serde(default)]
    pub source_shape_id: Option<String>,
    #[serde(default)]
    pub source_id: Option<u32>,
    #[serde(default)]
    pub copy_refusal: Option<String>,
    #[serde(default)]
    pub glue: Vec<ShapeTreeGlue>,
    #[serde(default)]
    pub children: Vec<ShapeTreeDraft>,
}

/// One internal glue record of a copied subtree, addressed by copy sources.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeTreeGlue {
    pub connector_source: String,
    pub endpoint: String,
    pub target_source: String,
    pub to_cell: String,
}

/// One glued connector endpoint; `to_cell` defaults to `PinX` when absent.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConnectorGlue {
    pub shape_id: String,
    pub to_cell: Option<String>,
}

#[derive(Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
enum SnapshotCellSheet {
    Document,
    Page(u32),
    Master(u32),
}

#[derive(Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
enum SnapshotCellRow {
    Index(u32),
    Name(String),
}

#[derive(Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
struct SnapshotCellLocator {
    sheet: SnapshotCellSheet,
    shape_id: Option<u32>,
    section: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    section_index: Option<u32>,
    row: Option<SnapshotCellRow>,
    cell_name: String,
}

impl From<&CellLocator> for SnapshotCellLocator {
    fn from(locator: &CellLocator) -> Self {
        Self {
            sheet: match locator.sheet {
                CellSheet::Document => SnapshotCellSheet::Document,
                CellSheet::Page(id) => SnapshotCellSheet::Page(id),
                CellSheet::Master(id) => SnapshotCellSheet::Master(id),
            },
            shape_id: locator.shape_id,
            section: locator.section.clone(),
            section_index: locator.section_index,
            row: locator.row.as_ref().map(|row| match row {
                CellRow::Index(id) => SnapshotCellRow::Index(*id),
                CellRow::Name(name) => SnapshotCellRow::Name(name.clone()),
            }),
            cell_name: locator.cell_name.clone(),
        }
    }
}

impl From<SnapshotCellLocator> for CellLocator {
    fn from(locator: SnapshotCellLocator) -> Self {
        Self {
            sheet: match locator.sheet {
                SnapshotCellSheet::Document => CellSheet::Document,
                SnapshotCellSheet::Page(id) => CellSheet::Page(id),
                SnapshotCellSheet::Master(id) => CellSheet::Master(id),
            },
            shape_id: locator.shape_id,
            section: locator.section,
            section_index: locator.section_index,
            row: locator.row.map(|row| match row {
                SnapshotCellRow::Index(id) => CellRow::Index(id),
                SnapshotCellRow::Name(name) => CellRow::Name(name),
            }),
            cell_name: locator.cell_name,
        }
    }
}

fn serialize_cell_locator<S>(locator: &CellLocator, serializer: S) -> Result<S::Ok, S::Error>
where
    S: Serializer,
{
    SnapshotCellLocator::from(locator).serialize(serializer)
}

fn deserialize_cell_locator<'de, D>(deserializer: D) -> Result<CellLocator, D::Error>
where
    D: Deserializer<'de>,
{
    SnapshotCellLocator::deserialize(deserializer).map(Into::into)
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum UpdateOrigin {
    Local,
    Remote,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct UpdateEvent {
    pub update: Vec<u8>,
    pub origin: UpdateOrigin,
}

#[derive(Debug, Error)]
pub enum EditError {
    #[error("invalid client ID {0}")]
    InvalidClientId(u64),
    #[error("could not parse VSDX: {0}")]
    Parse(String),
    #[error("invalid diagram state: {0}")]
    InvalidState(String),
    #[error("invalid yrs update: {0}")]
    InvalidUpdate(String),
    #[error("invalid yrs state vector: {0}")]
    InvalidStateVector(String),
    #[error("page {0:?} was not found")]
    PageNotFound(String),
    #[error("shape {0:?} was not found")]
    ShapeNotFound(String),
    #[error("cell {0:?} was not found")]
    CellNotFound(String),
    #[error("index {index} is outside length {length}")]
    OutOfBounds { index: u32, length: u32 },
    #[error("update observer failed: {0}")]
    Observer(String),
    #[error("JSON boundary error: {0}")]
    Json(String),
}

pub type EditResult<T> = Result<T, EditError>;
