use serde::{Deserialize, Deserializer, Serialize, Serializer};
use thiserror::Error;
use vsdx_parse::{CellLocator, CellRow, CellSheet, SemanticCellEdit};

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
    pub name: String,
    pub formula: Option<String>,
    pub value: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeSnapshot {
    pub id: String,
    pub source_id: u32,
    pub name: Option<String>,
    pub cells: Vec<CellSnapshot>,
    pub children: Vec<ShapeSnapshot>,
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

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeReceipt {
    pub page_id: String,
    pub shape_id: String,
    pub from_index: Option<u32>,
    pub to_index: Option<u32>,
}

/// A shape that exists only in the session, so a save must insert it into the package.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AddedShape {
    pub page_id: String,
    pub source_page_id: u32,
    pub shape_id: String,
    pub name: Option<String>,
    pub cells: Vec<CellSnapshot>,
}

/// The whole of a session's divergence from its package, partitioned by how a save applies it.
#[derive(Clone, Debug, PartialEq)]
pub struct SessionExport {
    pub cell_edits: Vec<SemanticCellEdit>,
    pub added_shapes: Vec<AddedShape>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeDraft {
    pub source_id: u32,
    pub name: Option<String>,
    pub cells: Vec<CellSnapshot>,
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
