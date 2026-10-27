//! Native table-structure operations over the yrs structural-table embed.
//!
//! A table is one map-backed unit in its parent story. Its `tblPr`, `grid`, and
//! `rows` fields are JSON-shaped [`Any`] values; every authored cell points at
//! an independent flat story. Since the three structural fields are atomic in
//! the persisted schema, each operation snapshots them, validates the grid,
//! rewrites them together in one transaction, and creates/removes cell stories
//! in that same transaction.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use serde::{Deserialize, Serialize};
use serde_json::Value;
use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{
    Any, Assoc, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextPrelim, TextRef, Transact,
    TransactionMut,
};

use crate::identity::IdAllocator;
use crate::op::{OpError, OpResult};
use crate::ops::paragraph_content_before;
use crate::seed::border_side_sources;
use crate::sticky_positions::{sticky_index, sticky_offset};
use crate::{
    DEL, EditCtx, EditingDoc, HYPERLINK, INS, KIND_KEY, PARA_ID, PILCROW_KIND, ParagraphIdOrigin,
    Position, STORIES, check_position, insertion_attrs, is_identity_key, is_pilcrow, map_string,
    out_len, revision_value, story_ref,
};

const TR_INS: &str = "trIns";
const TR_DEL: &str = "trDel";

/// Story-local locator for one structural table embed.
///
/// `table_index` is the zero-based table ordinal in the parent story (the same
/// `t{n}` ordinal used when the initial cell-story IDs are minted).
#[derive(Clone, Debug, Eq, PartialEq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableLocator {
    pub story: String,
    pub table_index: u32,
}

impl TableLocator {
    pub fn new(story: impl Into<String>, table_index: u32) -> Self {
        Self {
            story: story.into(),
            table_index,
        }
    }
}

/// One cell addressed in the table's resolved rectangular grid.
#[derive(Clone, Debug, Eq, PartialEq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellLoc {
    pub story: String,
    pub table_index: u32,
    pub row: u32,
    pub column: u32,
}

impl CellLoc {
    pub fn new(story: impl Into<String>, table_index: u32, row: u32, column: u32) -> Self {
        Self {
            story: story.into(),
            table_index,
            row,
            column,
        }
    }

    pub fn table(&self) -> TableLocator {
        TableLocator::new(self.story.clone(), self.table_index)
    }
}

/// Anchor-cell to head-cell rectangular selection.
#[derive(Clone, Debug, Eq, PartialEq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRange {
    pub anchor: CellLoc,
    pub head: CellLoc,
}

impl TableRange {
    pub fn new(anchor: CellLoc, head: CellLoc) -> Self {
        Self { anchor, head }
    }

    pub fn cell(at: CellLoc) -> Self {
        Self {
            anchor: at.clone(),
            head: at,
        }
    }
}

/// Common receipt returned by every native table mutation.
#[derive(Clone, Debug, Eq, PartialEq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableReceipt {
    pub table: TableLocator,
    pub rows: u32,
    pub columns: u32,
    pub created_story_ids: Vec<String>,
    pub deleted_story_ids: Vec<String>,
    /// Existing stories whose content changed: the table's story, plus a merge's surviving cell.
    #[serde(default)]
    pub changed_story_ids: Vec<String>,
    pub new_para_ids: Vec<String>,
    pub deleted_table: bool,
    #[serde(default)]
    pub revision_ids: Vec<String>,
}

#[derive(Clone, Debug, Eq, PartialEq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableFragmentReceipt {
    pub table: TableLocator,
    pub rows: u32,
    pub columns: u32,
    pub created_story_ids: Vec<String>,
    pub deleted_story_ids: Vec<String>,
    #[serde(default)]
    pub changed_story_ids: Vec<String>,
    pub new_para_ids: Vec<String>,
    pub deleted_table: bool,
    #[serde(default)]
    pub revision_ids: Vec<String>,
    pub after: FragmentCaret,
}

/// Caret position right after an inserted table fragment's embed.
#[derive(Clone, Debug, Eq, PartialEq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FragmentCaret {
    pub story: String,
    pub para_id: String,
    pub offset: u32,
}

#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub(crate) enum TableRowChangeKind {
    Insertion,
    Deletion,
    TableInsertion,
    TableDeletion,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct TableRowChange {
    pub revision_id: String,
    pub kind: TableRowChangeKind,
    pub author: String,
    pub date: String,
    pub start: u32,
}

#[derive(Clone, Debug)]
struct CellData {
    tc_pr: HashMap<String, Any>,
    story: String,
}

#[derive(Clone, Debug)]
struct RowData {
    tr_pr: HashMap<String, Any>,
    cells: Vec<CellData>,
}

#[derive(Clone, Debug)]
struct TableData {
    tbl_pr: HashMap<String, Any>,
    grid: Vec<Any>,
    rows: Vec<RowData>,
}

#[derive(Clone, Debug)]
struct CellAnchor {
    row: usize,
    column: usize,
    rowspan: usize,
    colspan: usize,
    cell_index: usize,
    cell: CellData,
}

#[derive(Clone, Copy, Debug)]
struct Rect {
    top: usize,
    bottom: usize,
    left: usize,
    right: usize,
}

#[derive(Clone)]
enum StoryInsert {
    Text(String),
    Embed(Vec<(String, Any)>),
}

#[derive(Clone)]
struct StoryChunk {
    insert: StoryInsert,
    attrs: Attrs,
}

fn invalid(message: impl Into<String>) -> OpError {
    OpError::InvalidTable(message.into())
}

fn any_map(value: &Any, detail: &str) -> OpResult<HashMap<String, Any>> {
    match value {
        Any::Map(map) => Ok(map.as_ref().clone()),
        _ => Err(invalid(format!("{detail} must be a map"))),
    }
}

fn any_array(value: &Any, detail: &str) -> OpResult<Vec<Any>> {
    match value {
        Any::Array(values) => Ok(values.to_vec()),
        _ => Err(invalid(format!("{detail} must be an array"))),
    }
}

fn any_string(value: Option<&Any>, detail: &str) -> OpResult<String> {
    match value {
        Some(Any::String(value)) => Ok(value.to_string()),
        _ => Err(invalid(format!("{detail} must be a string"))),
    }
}

fn any_usize(value: Option<&Any>, default: usize, detail: &str) -> OpResult<usize> {
    let number = match value {
        None | Some(Any::Null | Any::Undefined) => return Ok(default),
        Some(Any::Number(value)) => *value,
        Some(Any::BigInt(value)) => *value as f64,
        _ => return Err(invalid(format!("{detail} must be a positive integer"))),
    };
    if !number.is_finite() || number < 1.0 || number.fract() != 0.0 {
        return Err(invalid(format!("{detail} must be a positive integer")));
    }
    Ok(number as usize)
}

fn any_number(value: Option<&Any>) -> Option<f64> {
    match value {
        Some(Any::Number(value)) if value.is_finite() => Some(*value),
        Some(Any::BigInt(value)) => Some(*value as f64),
        _ => None,
    }
}

fn row_revision_parts(value: &Any) -> Option<(String, String, String)> {
    let Any::Map(map) = value else {
        return None;
    };
    let id = match map.get("id").or_else(|| map.get("revisionId")) {
        Some(Any::String(value)) => value.to_string(),
        Some(Any::Number(value)) if value.is_finite() => value.to_string(),
        Some(Any::BigInt(value)) => value.to_string(),
        _ => return None,
    };
    let string = |key: &str| match map.get(key) {
        Some(Any::String(value)) => value.to_string(),
        _ => String::new(),
    };
    Some((id, string("author"), string("date")))
}

fn matching_row_revision(value: Option<&Any>, filter: Option<&str>) -> Option<Any> {
    let value = value?.clone();
    if value == Any::Null {
        return None;
    }
    match filter {
        None => Some(value),
        Some(filter) => row_revision_parts(&value)
            .filter(|(id, ..)| id == filter)
            .map(|_| value),
    }
}

fn record_row_revision(resolved: &mut Vec<String>, value: Option<&Any>) {
    let Some((id, ..)) = value.and_then(row_revision_parts) else {
        return;
    };
    if !resolved.contains(&id) {
        resolved.push(id);
    }
}

fn shared_any<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> OpResult<Any> {
    match map.get(txn, key) {
        Some(Out::Any(value)) => Ok(value),
        _ => Err(invalid(format!("table is missing {key}"))),
    }
}

fn read_table<T: ReadTxn>(map: &MapRef, txn: &T) -> OpResult<TableData> {
    let tbl_pr = any_map(&shared_any(map, txn, "tblPr")?, "tblPr")?;
    let grid = any_array(&shared_any(map, txn, "grid")?, "grid")?;
    let row_values = any_array(&shared_any(map, txn, "rows")?, "rows")?;
    if row_values.is_empty() {
        return Err(invalid("table must contain at least one row"));
    }
    let mut rows = Vec::with_capacity(row_values.len());
    for (row_index, row_value) in row_values.iter().enumerate() {
        let row = any_map(row_value, &format!("row {row_index}"))?;
        let tr_pr = any_map(
            row.get("trPr")
                .ok_or_else(|| invalid(format!("row {row_index} is missing trPr")))?,
            &format!("row {row_index} trPr"),
        )?;
        let cells = any_array(
            row.get("cells")
                .ok_or_else(|| invalid(format!("row {row_index} is missing cells")))?,
            &format!("row {row_index} cells"),
        )?
        .into_iter()
        .enumerate()
        .map(|(cell_index, cell_value)| {
            let cell = any_map(&cell_value, &format!("row {row_index} cell {cell_index}"))?;
            let tc_pr = any_map(
                cell.get("tcPr").ok_or_else(|| {
                    invalid(format!("row {row_index} cell {cell_index} is missing tcPr"))
                })?,
                &format!("row {row_index} cell {cell_index} tcPr"),
            )?;
            let story = any_string(
                cell.get("story"),
                &format!("row {row_index} cell {cell_index} story"),
            )?;
            Ok(CellData { tc_pr, story })
        })
        .collect::<OpResult<Vec<_>>>()?;
        rows.push(RowData { tr_pr, cells });
    }
    let data = TableData { tbl_pr, grid, rows };
    anchors(&data)?;
    Ok(data)
}

fn to_any_map(map: HashMap<String, Any>) -> Any {
    Any::Map(Arc::new(map))
}

fn to_any_array(values: Vec<Any>) -> Any {
    Any::Array(Arc::from(values))
}

/// Merges physical border edges into a cell's `tcPr.borders`. A `None` value
/// drops that edge; an emptied object drops the whole key.
fn merge_cell_borders<'a>(
    cell: &mut CellData,
    edges: impl IntoIterator<Item = (&'a str, &'a Option<Any>)>,
) {
    let mut merged = cell
        .tc_pr
        .get("borders")
        .and_then(|value| match value {
            Any::Map(map) => Some(map.as_ref().clone()),
            _ => None,
        })
        .unwrap_or_default();
    for (edge, value) in edges {
        match value {
            Some(value) => merged.insert(edge.to_owned(), value.clone()),
            None => merged.remove(edge),
        };
    }
    if merged.is_empty() {
        cell.tc_pr.remove("borders");
    } else {
        cell.tc_pr.insert("borders".to_owned(), to_any_map(merged));
    }
}

/// An authored border; a `none` or `nil` side keeps only its style, since
/// line size, spacing and color mean nothing on a side that draws no line.
fn authored_border(value: &Any) -> Any {
    match value {
        Any::Map(map) => match map.get("style") {
            Some(Any::String(style)) if matches!(style.as_ref(), "none" | "nil") => to_any_map(
                HashMap::from([("style".to_owned(), Any::from(style.as_ref()))]),
            ),
            _ => value.clone(),
        },
        _ => value.clone(),
    }
}

/// `tcPr.borders` for a freshly inserted table: Word's Table Grid rules, a
/// single 0.5pt (`w:sz="4"`) black line on every edge.
fn default_cell_borders() -> Any {
    let edge = to_any_map(HashMap::from([
        ("style".to_owned(), Any::from("single")),
        ("size".to_owned(), Any::Number(4.0)),
        (
            "color".to_owned(),
            to_any_map(HashMap::from([("rgb".to_owned(), Any::from("000000"))])),
        ),
    ]));
    to_any_map(HashMap::from([
        ("top".to_owned(), edge.clone()),
        ("bottom".to_owned(), edge.clone()),
        ("left".to_owned(), edge.clone()),
        ("right".to_owned(), edge),
    ]))
}

fn write_table(txn: &mut TransactionMut<'_>, map: &MapRef, data: &TableData) {
    let rows = data
        .rows
        .iter()
        .map(|row| {
            let cells = row
                .cells
                .iter()
                .map(|cell| {
                    to_any_map(HashMap::from([
                        ("tcPr".to_owned(), to_any_map(cell.tc_pr.clone())),
                        ("story".to_owned(), Any::from(cell.story.as_str())),
                    ]))
                })
                .collect();
            to_any_map(HashMap::from([
                ("trPr".to_owned(), to_any_map(row.tr_pr.clone())),
                ("cells".to_owned(), to_any_array(cells)),
            ]))
        })
        .collect();
    map.insert(txn, "tblPr", to_any_map(data.tbl_pr.clone()));
    map.insert(txn, "grid", to_any_array(data.grid.clone()));
    map.insert(txn, "rows", to_any_array(rows));
}

fn span(cell: &CellData) -> OpResult<(usize, usize)> {
    Ok((
        any_usize(cell.tc_pr.get("rowspan"), 1, "rowspan")?,
        any_usize(cell.tc_pr.get("colspan"), 1, "colspan")?,
    ))
}

fn set_spans(cell: &mut CellData, rowspan: usize, colspan: usize) {
    cell.tc_pr
        .insert("rowspan".to_owned(), Any::Number(rowspan as f64));
    cell.tc_pr
        .insert("colspan".to_owned(), Any::Number(colspan as f64));
    // `gridSpan`/`vMerge` are OOXML serialization details. The canonical yrs
    // model stores resolved `colspan`/`rowspan`; yrsToDocument reconstructs
    // the OOXML continuation cells on save.
    cell.tc_pr.remove("gridSpan");
    cell.tc_pr.remove("vMerge");
}

const MAX_FRAGMENT_ROWS: usize = 1000;
const MAX_FRAGMENT_COLUMNS: usize = 1000;
const MAX_FRAGMENT_CELLS: usize = 10_000;
const MAX_FRAGMENT_DEPTH: usize = 8;

struct FragmentText {
    text: String,
    attrs: Attrs,
}

struct FragmentTabBreak {
    kind: &'static str,
    payload: Vec<(String, Any)>,
    attrs: Attrs,
}

struct FragmentPilcrow {
    props: Vec<(String, Any)>,
    attrs: Attrs,
}

struct FragmentNested {
    table: FragmentTable,
    attrs: Attrs,
}

enum FragmentSegment {
    Text(FragmentText),
    TabBreak(FragmentTabBreak),
    Pilcrow(FragmentPilcrow),
    Table(FragmentNested),
}

struct FragmentCell {
    tc_pr: HashMap<String, Any>,
    content: Vec<FragmentSegment>,
    shading: Option<String>,
}

struct FragmentRow {
    tr_pr: HashMap<String, Any>,
    cells: Vec<FragmentCell>,
}

struct FragmentTable {
    tbl_pr: HashMap<String, Any>,
    grid: Vec<Any>,
    rows: Vec<FragmentRow>,
}

#[derive(Default)]
struct FragmentBuild {
    created_story_ids: Vec<String>,
    new_para_ids: Vec<String>,
}

fn json_shaped(value: &Value) -> bool {
    match value {
        Value::Null | Value::Bool(_) | Value::String(_) => true,
        Value::Number(number) => number.as_f64().is_some_and(f64::is_finite),
        Value::Array(items) => items.iter().all(json_shaped),
        Value::Object(map) => map.values().all(json_shaped),
    }
}

fn fragment_any(value: &Value, detail: &str) -> OpResult<Any> {
    if !json_shaped(value) {
        return Err(invalid(format!("{detail} is not JSON-shaped")));
    }
    Any::from_json(&value.to_string())
        .map_err(|error| invalid(format!("{detail} is not JSON-shaped: {error:?}")))
}

fn fragment_object<'a>(
    value: &'a Value,
    detail: &str,
) -> OpResult<&'a serde_json::Map<String, Value>> {
    value
        .as_object()
        .ok_or_else(|| invalid(format!("{detail} must be an object")))
}

fn sanitize_fragment_hyperlink(value: &Value, detail: &str) -> OpResult<Any> {
    match value {
        Value::String(href) => Ok(Any::Map(Arc::new(HashMap::from([(
            "href".to_owned(),
            Any::from(href.as_str()),
        )])))),
        Value::Object(map) => {
            let href = map
                .get("href")
                .and_then(Value::as_str)
                .ok_or_else(|| invalid(format!("{detail} hyperlink requires a string href")))?;
            let mut link = HashMap::from([("href".to_owned(), Any::from(href))]);
            if let Some(Value::String(tooltip)) = map.get("tooltip") {
                link.insert("tooltip".to_owned(), Any::from(tooltip.as_str()));
            }
            if let Some(Value::Bool(history)) = map.get("history") {
                link.insert("history".to_owned(), Any::Bool(*history));
            }
            Ok(Any::Map(Arc::new(link)))
        }
        _ => Err(invalid(format!("{detail} hyperlink must be an object"))),
    }
}

fn sanitize_fragment_attrs(value: Option<&Value>, detail: &str) -> OpResult<Attrs> {
    let Some(entries) = value else {
        return Ok(Attrs::new());
    };
    let map = entries
        .as_object()
        .ok_or_else(|| invalid(format!("{detail} attributes must be an object")))?;
    let mut attrs = Attrs::new();
    for (key, entry) in map {
        if key == INS || key == DEL {
            continue;
        }
        if key == HYPERLINK {
            attrs.insert(
                Arc::from(HYPERLINK),
                sanitize_fragment_hyperlink(entry, detail)?,
            );
            continue;
        }
        attrs.insert(
            Arc::from(key.as_str()),
            fragment_any(entry, &format!("{detail} attribute {key:?}"))?,
        );
    }
    Ok(attrs)
}

fn fragment_deleted(attributes: Option<&Value>) -> bool {
    attributes
        .and_then(Value::as_object)
        .and_then(|map| map.get("del"))
        .is_some_and(|stamp| !stamp.is_null())
}

fn validate_fragment_table(
    value: &Value,
    depth: usize,
    remaining: &mut usize,
) -> OpResult<(FragmentTable, usize)> {
    if depth > MAX_FRAGMENT_DEPTH {
        return Err(invalid("table nesting exceeds eight levels"));
    }
    let object = fragment_object(value, "table")?;
    let tbl_pr = match object.get("tblPr") {
        Some(Value::Object(map)) => map
            .iter()
            .map(|(key, entry)| Ok((key.clone(), fragment_any(entry, "tblPr entry")?)))
            .collect::<OpResult<HashMap<String, Any>>>()?,
        _ => return Err(invalid("table tblPr must be an object")),
    };
    let grid = match object.get("grid") {
        Some(Value::Array(items)) => items
            .iter()
            .map(|entry| fragment_any(entry, "grid entry"))
            .collect::<OpResult<Vec<Any>>>()?,
        _ => return Err(invalid("table grid must be an array")),
    };
    let rows = match object.get("rows") {
        Some(Value::Array(rows)) if !rows.is_empty() => rows,
        _ => return Err(invalid("table must contain at least one row")),
    };
    if rows.len() > MAX_FRAGMENT_ROWS {
        return Err(invalid("table holds more than 1000 rows"));
    }
    let mut validated = Vec::with_capacity(rows.len());
    for (index, row) in rows.iter().enumerate() {
        validated.push(validate_fragment_row(
            row,
            depth,
            remaining,
            &format!("row {index}"),
        )?);
    }
    let valid = FragmentTable {
        tbl_pr,
        grid,
        rows: validated,
    };
    let columns = validation_anchors(&valid, depth)?;
    Ok((valid, columns))
}

fn validate_fragment_row(
    value: &Value,
    depth: usize,
    remaining: &mut usize,
    detail: &str,
) -> OpResult<FragmentRow> {
    let object = fragment_object(value, detail)?;
    let tr_pr = match object.get("trPr") {
        Some(Value::Object(map)) => map
            .iter()
            .map(|(key, entry)| {
                Ok((
                    key.clone(),
                    fragment_any(entry, &format!("{detail} trPr entry"))?,
                ))
            })
            .collect::<OpResult<HashMap<String, Any>>>()?,
        _ => return Err(invalid(format!("{detail} trPr must be an object"))),
    };
    let cells = match object.get("cells") {
        Some(Value::Array(cells)) => cells,
        _ => return Err(invalid(format!("{detail} cells must be an array"))),
    };
    if cells.len() > MAX_FRAGMENT_COLUMNS {
        return Err(invalid(format!("{detail} holds more than 1000 cells")));
    }
    let mut validated = Vec::with_capacity(cells.len());
    for (index, cell) in cells.iter().enumerate() {
        validated.push(validate_fragment_cell(
            cell,
            depth,
            remaining,
            &format!("{detail} cell {index}"),
        )?);
    }
    Ok(FragmentRow {
        tr_pr,
        cells: validated,
    })
}

fn validate_fragment_cell(
    value: &Value,
    depth: usize,
    remaining: &mut usize,
    detail: &str,
) -> OpResult<FragmentCell> {
    let object = fragment_object(value, detail)?;
    let tc_pr = match object.get("tcPr") {
        Some(Value::Object(map)) => map
            .iter()
            .map(|(key, entry)| {
                Ok((
                    key.clone(),
                    fragment_any(entry, &format!("{detail} tcPr entry"))?,
                ))
            })
            .collect::<OpResult<HashMap<String, Any>>>()?,
        _ => return Err(invalid(format!("{detail} tcPr must be an object"))),
    };
    for (key, cap) in [
        ("rowspan", MAX_FRAGMENT_ROWS),
        ("colspan", MAX_FRAGMENT_COLUMNS),
    ] {
        let span = any_usize(tc_pr.get(key), 1, key)?;
        if span > cap {
            return Err(invalid(format!("{detail} {key} exceeds {cap}")));
        }
    }
    if *remaining == 0 {
        return Err(invalid("fragment holds more than 10000 cells"));
    }
    *remaining -= 1;
    let shading = match object.get("shading") {
        None => None,
        Some(Value::String(color)) => {
            let plain = color.trim_start_matches('#');
            if plain.len() != 6 || !plain.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return Err(invalid(format!(
                    "{detail} shading must be a six-digit hex color"
                )));
            }
            Some(plain.to_owned())
        }
        Some(_) => {
            return Err(invalid(format!(
                "{detail} shading must be a six-digit hex color"
            )));
        }
    };
    let content = match object.get("content") {
        Some(Value::Array(items)) if !items.is_empty() => items,
        _ => {
            return Err(invalid(format!(
                "{detail} content must be a non-empty array"
            )));
        }
    };
    let mut validated = Vec::with_capacity(content.len());
    for (index, segment) in content.iter().enumerate() {
        if let Some(segment) = validate_fragment_segment(
            segment,
            depth,
            remaining,
            &format!("{detail} segment {index}"),
        )? {
            validated.push(segment);
        }
    }
    if validated.is_empty() {
        return Err(invalid(format!(
            "{detail} content must be a non-empty array"
        )));
    }
    if !matches!(validated.last(), Some(FragmentSegment::Pilcrow(_))) {
        return Err(invalid(format!("{detail} content must end with a pilcrow")));
    }
    Ok(FragmentCell {
        tc_pr,
        content: validated,
        shading,
    })
}

fn validate_fragment_segment(
    value: &Value,
    depth: usize,
    remaining: &mut usize,
    detail: &str,
) -> OpResult<Option<FragmentSegment>> {
    let object = fragment_object(value, detail)?;
    let kind = object
        .get("kind")
        .and_then(Value::as_str)
        .ok_or_else(|| invalid(format!("{detail} requires a string kind")))?;
    let attrs = sanitize_fragment_attrs(object.get("attributes"), detail)?;
    match kind {
        "text" => {
            if fragment_deleted(object.get("attributes")) {
                return Ok(None);
            }
            let text = object
                .get("text")
                .and_then(Value::as_str)
                .ok_or_else(|| invalid(format!("{detail} requires text")))?;
            if text.is_empty() || text.contains(['\n', '\r', '\u{000B}', '\u{2028}', '\u{2029}']) {
                return Err(invalid(format!(
                    "{detail} text must be non-empty without line breaks"
                )));
            }
            let mut merged = attrs;
            if let Some(format) = object.get("format") {
                let map = format
                    .as_object()
                    .ok_or_else(|| invalid(format!("{detail} format must be an object")))?;
                let delta = crate::format::parse_inline_format_delta(map)
                    .map_err(|error| invalid(format!("{detail} format is invalid: {error}")))?;
                let lowered = delta
                    .to_attrs()
                    .map_err(|error| invalid(format!("{detail} format is invalid: {error}")))?;
                for (key, entry) in lowered {
                    if entry == Any::Null {
                        merged.remove(&key);
                    } else {
                        merged.insert(key, entry);
                    }
                }
            }
            Ok(Some(FragmentSegment::Text(FragmentText {
                text: text.to_owned(),
                attrs: merged,
            })))
        }
        "tab" | "break" => {
            let mut payload = Vec::new();
            if let Some(entries) = object.get("payload") {
                let map = entries
                    .as_object()
                    .ok_or_else(|| invalid(format!("{detail} payload must be an object")))?;
                for (key, entry) in map {
                    if key == KIND_KEY {
                        continue;
                    }
                    payload.push((
                        key.clone(),
                        fragment_any(entry, &format!("{detail} payload entry"))?,
                    ));
                }
            }
            Ok(Some(FragmentSegment::TabBreak(FragmentTabBreak {
                kind: if kind == "tab" { "tab" } else { "break" },
                payload,
                attrs,
            })))
        }
        "pilcrow" => {
            let mut props = Vec::new();
            if let Some(entries) = object.get("properties") {
                let map = entries
                    .as_object()
                    .ok_or_else(|| invalid(format!("{detail} properties must be an object")))?;
                for (key, entry) in map {
                    if is_identity_key(key) {
                        continue;
                    }
                    props.push((
                        key.clone(),
                        fragment_any(entry, &format!("{detail} property {key:?}"))?,
                    ));
                }
            }
            Ok(Some(FragmentSegment::Pilcrow(FragmentPilcrow {
                props,
                attrs,
            })))
        }
        "table" => {
            let nested = object
                .get("table")
                .ok_or_else(|| invalid(format!("{detail} requires a table")))?;
            let (table, _) = validate_fragment_table(nested, depth + 1, remaining)?;
            Ok(Some(FragmentSegment::Table(FragmentNested {
                table,
                attrs,
            })))
        }
        _ => Err(invalid(format!(
            "{detail} has an unknown segment kind {kind:?}"
        ))),
    }
}

fn validation_anchors(valid: &FragmentTable, depth: usize) -> OpResult<usize> {
    let mut rows = Vec::with_capacity(valid.rows.len());
    for (row_index, row) in valid.rows.iter().enumerate() {
        let mut cells = Vec::with_capacity(row.cells.len());
        for (cell_index, cell) in row.cells.iter().enumerate() {
            cells.push(CellData {
                tc_pr: cell.tc_pr.clone(),
                story: format!("frag:{depth}:{row_index}:{cell_index}"),
            });
        }
        rows.push(RowData {
            tr_pr: row.tr_pr.clone(),
            cells,
        });
    }
    let data = TableData {
        tbl_pr: valid.tbl_pr.clone(),
        grid: valid.grid.clone(),
        rows,
    };
    let (cell_anchors, columns) = anchors(&data)?;
    if columns > MAX_FRAGMENT_COLUMNS {
        return Err(invalid("table holds more than 1000 columns"));
    }
    if cell_anchors.is_empty() {
        return Err(invalid("table must contain at least one cell"));
    }
    for (row_index, row) in data.rows.iter().enumerate() {
        if row.cells.is_empty()
            && (0..columns).any(|column| covering(&cell_anchors, row_index, column).is_none())
        {
            return Err(invalid(format!(
                "row {row_index} is not covered by a merged cell"
            )));
        }
    }
    Ok(columns)
}

fn caret_after_insert(
    story: &TextRef,
    txn: &TransactionMut<'_>,
    story_id: &str,
    insert_index: u32,
) -> OpResult<FragmentCaret> {
    let mut offset = 0u32;
    let mut para_start = 0u32;
    let mut fallback: Option<(String, u32)> = None;
    for diff in story.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        if let Out::YMap(map) = &diff.insert
            && is_pilcrow(map, txn)
        {
            let para_id = map_string(map, txn, PARA_ID).unwrap_or_default();
            fallback = Some((para_id.clone(), para_start));
            if offset > insert_index {
                return Ok(FragmentCaret {
                    story: story_id.to_owned(),
                    para_id,
                    offset: insert_index + 1 - para_start,
                });
            }
            para_start = offset + 1;
        }
        offset += len;
    }
    let (para_id, start) = fallback.unwrap_or_default();
    Ok(FragmentCaret {
        story: story_id.to_owned(),
        para_id,
        offset: insert_index + 1 - start,
    })
}

fn omitted_columns(row: &RowData, key: &str) -> OpResult<usize> {
    let value = row.tr_pr.get(key).or_else(|| {
        let Any::Map(original) = row.tr_pr.get("_originalFormatting")? else {
            return None;
        };
        original.get(key)
    });
    if matches!(value, None | Some(Any::Null | Any::Undefined)) {
        return Ok(0);
    }
    let count = any_number(value).ok_or_else(|| invalid(format!("invalid {key}")))?;
    if count < 0.0 || count.fract() != 0.0 || count > MAX_FRAGMENT_COLUMNS as f64 {
        return Err(invalid(format!("invalid {key}")));
    }
    Ok(count as usize)
}

fn anchors(data: &TableData) -> OpResult<(Vec<CellAnchor>, usize)> {
    let mut occupied: Vec<Vec<bool>> = vec![Vec::new(); data.rows.len()];
    let mut result = Vec::new();
    let mut total_columns = data.grid.len();
    let mut seen_stories = HashSet::new();
    for (row_index, row) in data.rows.iter().enumerate() {
        let mut column = omitted_columns(row, "gridBefore")?;
        for (cell_index, cell) in row.cells.iter().enumerate() {
            while occupied[row_index].get(column).copied().unwrap_or(false) {
                column += 1;
            }
            let (rowspan, colspan) = span(cell)?;
            if colspan > MAX_FRAGMENT_COLUMNS {
                return Err(invalid(format!(
                    "cell {} exceeds the column limit",
                    cell.story
                )));
            }
            let row_end = row_index.checked_add(rowspan).ok_or_else(|| {
                invalid(format!("cell {} spans beyond the final row", cell.story))
            })?;
            if row_end > data.rows.len() {
                return Err(invalid(format!(
                    "cell {} spans beyond the final row",
                    cell.story
                )));
            }
            let column_end = column
                .checked_add(colspan)
                .ok_or_else(|| invalid(format!("cell {} exceeds the column limit", cell.story)))?;
            if column_end > MAX_FRAGMENT_COLUMNS {
                return Err(invalid(format!(
                    "cell {} exceeds the column limit",
                    cell.story
                )));
            }
            if !seen_stories.insert(cell.story.clone()) {
                return Err(invalid(format!(
                    "cell story {:?} is referenced more than once",
                    cell.story
                )));
            }
            for row_slot in occupied.iter_mut().skip(row_index).take(rowspan) {
                if row_slot.len() < column + colspan {
                    row_slot.resize(column + colspan, false);
                }
                if row_slot[column..column + colspan].iter().any(|used| *used) {
                    return Err(invalid(format!(
                        "cell {} overlaps another cell",
                        cell.story
                    )));
                }
                row_slot[column..column + colspan].fill(true);
            }
            result.push(CellAnchor {
                row: row_index,
                column,
                rowspan,
                colspan,
                cell_index,
                cell: cell.clone(),
            });
            column += colspan;
            total_columns = total_columns.max(column);
        }
        let row_end = column.max(occupied[row_index].len()) + omitted_columns(row, "gridAfter")?;
        if row_end > MAX_FRAGMENT_COLUMNS {
            return Err(invalid("row exceeds the column limit"));
        }
        total_columns = total_columns.max(row_end);
    }
    if total_columns == 0 {
        return Err(invalid("table must contain at least one column"));
    }
    Ok((result, total_columns))
}

fn covering<'a>(anchors: &'a [CellAnchor], row: usize, column: usize) -> Option<&'a CellAnchor> {
    anchors.iter().find(|anchor| {
        anchor.row <= row
            && row < anchor.row + anchor.rowspan
            && anchor.column <= column
            && column < anchor.column + anchor.colspan
    })
}

fn table_at<T: ReadTxn>(txn: &T, locator: &TableLocator) -> OpResult<(TextRef, MapRef, u32)> {
    let story = story_ref(txn, &locator.story).map_err(OpError::from)?;
    let mut offset = 0u32;
    let mut table_index = 0u32;
    for diff in story.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        if let Out::YMap(map) = diff.insert
            && map_string(&map, txn, KIND_KEY).as_deref() == Some("table")
        {
            if table_index == locator.table_index {
                return Ok((story, map, offset));
            }
            table_index += 1;
        }
        offset += len;
    }
    Err(OpError::UnknownTable {
        story: locator.story.clone(),
        table_index: locator.table_index,
    })
}

fn checked_rect(data: &TableData, range: &TableRange) -> OpResult<Rect> {
    if range.anchor.story != range.head.story || range.anchor.table_index != range.head.table_index
    {
        return Err(invalid("a cell range must stay inside one table"));
    }
    let locator = range.anchor.table();
    if locator.story != range.anchor.story {
        return Err(invalid("invalid table range locator"));
    }
    let (cell_anchors, columns) = anchors(data)?;
    let rows = data.rows.len();
    let anchor_row = range.anchor.row as usize;
    let anchor_column = range.anchor.column as usize;
    let head_row = range.head.row as usize;
    let head_column = range.head.column as usize;
    if anchor_row >= rows || head_row >= rows || anchor_column >= columns || head_column >= columns
    {
        return Err(invalid(format!(
            "cell range is outside the {rows}x{columns} table"
        )));
    }
    if covering(&cell_anchors, anchor_row, anchor_column).is_none()
        || covering(&cell_anchors, head_row, head_column).is_none()
    {
        return Err(invalid("cell range points at an unoccupied grid slot"));
    }
    let mut rect = Rect {
        top: anchor_row.min(head_row),
        bottom: anchor_row.max(head_row) + 1,
        left: anchor_column.min(head_column),
        right: anchor_column.max(head_column) + 1,
    };
    // A cell selection expands across a merged cell rather than selecting a
    // fraction of it. Repeat until every intersecting span is fully enclosed.
    loop {
        let before = (rect.top, rect.bottom, rect.left, rect.right);
        for anchor in &cell_anchors {
            let intersects = anchor.row < rect.bottom
                && anchor.row + anchor.rowspan > rect.top
                && anchor.column < rect.right
                && anchor.column + anchor.colspan > rect.left;
            if intersects {
                rect.top = rect.top.min(anchor.row);
                rect.bottom = rect.bottom.max(anchor.row + anchor.rowspan);
                rect.left = rect.left.min(anchor.column);
                rect.right = rect.right.max(anchor.column + anchor.colspan);
            }
        }
        if before == (rect.top, rect.bottom, rect.left, rect.right) {
            return Ok(rect);
        }
    }
}

fn selected_anchors(data: &TableData, range: &TableRange) -> OpResult<Vec<CellAnchor>> {
    let rect = checked_rect(data, range)?;
    let (cell_anchors, _) = anchors(data)?;
    Ok(cell_anchors
        .into_iter()
        .filter(|anchor| {
            anchor.row >= rect.top
                && anchor.row + anchor.rowspan <= rect.bottom
                && anchor.column >= rect.left
                && anchor.column + anchor.colspan <= rect.right
        })
        .collect())
}

fn reset_cell_spans(mut cell: CellData) -> CellData {
    let source_story = Any::from(cell.story.as_str());
    cell.tc_pr
        .entry("_sourceCellStory".to_owned())
        .or_insert(source_story);
    set_spans(&mut cell, 1, 1);
    cell.tc_pr.remove("colwidth");
    cell.tc_pr.remove("cellMarker");
    cell.tc_pr.remove("_originalFormatting");
    cell
}

/// `tr_pr` for a row synthesized from it, without the source row's own
/// `w:tr` attributes (`w14:paraId`, `w:rsid*`).
pub(crate) fn fresh_row_properties(mut tr_pr: HashMap<String, Any>) -> HashMap<String, Any> {
    tr_pr.remove("_extraAttributes");
    tr_pr
}

fn reconstruct_rows(
    rows: Vec<RowData>,
    mut cell_anchors: Vec<CellAnchor>,
) -> OpResult<Vec<RowData>> {
    cell_anchors.sort_by_key(|anchor| (anchor.row, anchor.column));
    let mut result: Vec<RowData> = rows
        .into_iter()
        .map(|row| RowData {
            tr_pr: row.tr_pr,
            cells: Vec::new(),
        })
        .collect();
    for anchor in cell_anchors {
        if anchor.row >= result.len() {
            return Err(invalid("cell anchor lies beyond the final row"));
        }
        let mut cell = anchor.cell;
        set_spans(&mut cell, anchor.rowspan, anchor.colspan);
        result[anchor.row].cells.push(cell);
    }
    Ok(result)
}

fn table_base(locator: &TableLocator) -> String {
    format!("{}:t{}", locator.story, locator.table_index)
}

/// Story ids without the `~{client}` suffix of an edit-born cell, so a slot is fresh only when
/// no replica's cell holds it.
fn story_ids<T: ReadTxn>(txn: &T) -> HashSet<String> {
    let Some(stories) = txn.get_map(STORIES) else {
        return HashSet::new();
    };
    stories
        .iter(txn)
        .map(|(id, _)| bare_story_id(id).to_owned())
        .collect()
}

fn bare_story_id(id: &str) -> &str {
    let Some((bare, client)) = id.rsplit_once('~') else {
        return id;
    };
    let digits = !client.is_empty() && client.bytes().all(|byte| byte.is_ascii_digit());
    if digits { bare } else { id }
}

fn fresh_row_slot(used: &HashSet<String>, base: &str, start: usize, columns: usize) -> usize {
    (start..)
        .find(|row| (0..columns).all(|column| !used.contains(&format!("{base}:r{row}c{column}"))))
        .expect("usize address space cannot be exhausted by story ids")
}

fn fresh_table_slot(used: &HashSet<String>, story: &str, start: usize) -> usize {
    (start..)
        .find(|slot| {
            let prefix = format!("{story}:t{slot}:");
            !used.iter().any(|id| id.starts_with(&prefix))
        })
        .expect("usize address space cannot be exhausted by story ids")
}

fn fresh_column_slot(used: &HashSet<String>, base: &str, rows: usize, start: usize) -> usize {
    (start..)
        .find(|column| (0..rows).all(|row| !used.contains(&format!("{base}:r{row}c{column}"))))
        .expect("usize address space cannot be exhausted by story ids")
}

/// A cell story id minted by an edit. The `~{client}` suffix keeps it apart from the cell a
/// concurrent replica mints for the same slot; ids read from a document carry none.
fn fresh_cell_story(
    used: &mut HashSet<String>,
    base: &str,
    row: usize,
    mut column: usize,
    client: u64,
) -> String {
    loop {
        let candidate = format!("{base}:r{row}c{column}");
        if used.insert(candidate.clone()) {
            return format!("{candidate}~{client}");
        }
        column += 1;
    }
}

fn create_cell_story(
    doc: &EditingDoc,
    txn: &mut TransactionMut<'_>,
    ids: &mut IdAllocator,
    story_id: &str,
) -> OpResult<String> {
    let stories = txn
        .get_map(STORIES)
        .expect("stories root is declared by EditingDoc::new");
    if stories.contains_key(txn, story_id) {
        return Err(OpError::StoryExists(story_id.to_owned()));
    }
    let para_id = ids.session_key(doc);
    let story = stories.insert(txn, story_id, TextPrelim::new(""));
    let pilcrow = story.insert_embed_with_attributes(
        txn,
        0,
        MapPrelim::default(),
        insertion_attrs(None, None),
    );
    pilcrow.insert(txn, KIND_KEY, crate::PILCROW_KIND);
    pilcrow.insert(txn, crate::PARA_ID, para_id.as_str());
    ids.bind(txn, &pilcrow, &para_id, ParagraphIdOrigin::Authored);
    Ok(para_id)
}

fn remove_story_exact(txn: &mut TransactionMut<'_>, story_id: &str, deleted: &mut Vec<String>) {
    if story_ref(txn, story_id).is_ok() {
        crate::mark_story_removed(txn, story_id);
        deleted.push(story_id.to_owned());
    }
}

fn remove_story_tree(txn: &mut TransactionMut<'_>, story_id: &str, deleted: &mut Vec<String>) {
    let stories = txn
        .get_map(STORIES)
        .expect("stories root is declared by EditingDoc::new");
    let prefix = format!("{story_id}:");
    let mut ids: Vec<String> = crate::live_stories(&stories, txn)
        .map(|(id, _)| id.to_string())
        .filter(|id| id == story_id || id.starts_with(&prefix))
        .collect();
    // Children first is easier to inspect in receipts and avoids temporarily
    // leaving a descendant whose parent has already vanished.
    ids.sort_by_key(|id| std::cmp::Reverse(id.len()));
    for id in ids {
        remove_story_exact(txn, &id, deleted);
    }
}

fn snapshot_story<T: ReadTxn>(story: &TextRef, txn: &T) -> OpResult<Vec<StoryChunk>> {
    story
        .diff(txn, YChange::identity)
        .into_iter()
        .map(|diff| {
            let attrs = diff.attributes.as_deref().cloned().unwrap_or_default();
            let insert = match diff.insert {
                Out::Any(Any::String(text)) => StoryInsert::Text(text.to_string()),
                Out::YMap(map) => {
                    let entries = map
                        .iter(txn)
                        .map(|(key, value)| match value {
                            Out::Any(value) => Ok((key.to_string(), value)),
                            _ => Err(invalid("a story embed property must be JSON-shaped")),
                        })
                        .collect::<OpResult<Vec<_>>>()?;
                    StoryInsert::Embed(entries)
                }
                _ => return Err(invalid("cell story contains an unsupported shared value")),
            };
            Ok(StoryChunk { insert, attrs })
        })
        .collect()
}

struct MergedCommentRange {
    index: usize,
    start: u32,
    end: u32,
    start_assoc: Assoc,
    end_assoc: Assoc,
}

struct MergedComment {
    comment: MapRef,
    anchors: Vec<Any>,
    ranges: Vec<MergedCommentRange>,
}

fn merged_comments<T: ReadTxn>(
    txn: &T,
    offsets: &HashMap<String, u32>,
) -> OpResult<Vec<MergedComment>> {
    let mut moved = Vec::new();
    let Some(comments) = txn.get_map(crate::COMMENTS) else {
        return Ok(moved);
    };
    for (_, value) in comments.iter(txn) {
        let Out::YMap(comment) = value else {
            continue;
        };
        let Some(Out::Any(Any::Array(anchors))) = comment.get(txn, "anchors") else {
            continue;
        };
        let mut ranges = Vec::new();
        for (index, encoded) in anchors.iter().enumerate() {
            let Any::Map(fields) = encoded else {
                continue;
            };
            let Some(Any::String(story)) = fields.get("story") else {
                continue;
            };
            let Some(offset) = offsets.get(story.as_ref()) else {
                continue;
            };
            let anchor = crate::decode_anchor(encoded).map_err(OpError::from)?;
            let source = story_ref(txn, &anchor.story).map_err(OpError::from)?;
            let start = sticky_offset(&anchor.start, txn)
                .ok_or_else(|| OpError::InvalidComment("start anchor no longer resolves".into()))?;
            let end = sticky_offset(&anchor.end, txn)
                .ok_or_else(|| OpError::InvalidComment("end anchor no longer resolves".into()))?;
            if start.index > source.len(txn) || end.index > source.len(txn) {
                return Err(OpError::InvalidComment("anchor exceeds its story".into()));
            }
            ranges.push(MergedCommentRange {
                index,
                start: offset + start.index,
                end: offset + end.index,
                start_assoc: anchor.start.assoc,
                end_assoc: anchor.end.assoc,
            });
        }
        if !ranges.is_empty() {
            moved.push(MergedComment {
                comment,
                anchors: anchors.to_vec(),
                ranges,
            });
        }
    }
    Ok(moved)
}

fn move_merged_comments(
    txn: &mut TransactionMut<'_>,
    destination: &TextRef,
    story_id: &str,
    comments: Vec<MergedComment>,
) {
    for mut comment in comments {
        for range in comment.ranges {
            let start = sticky_index(destination, txn, range.start, range.start_assoc)
                .expect("merged comment start lies in appended content");
            let end = sticky_index(destination, txn, range.end, range.end_assoc)
                .expect("merged comment end lies in appended content");
            comment.anchors[range.index] = crate::anchor_value(story_id, &start, &end);
        }
        comment
            .comment
            .insert(txn, "anchors", to_any_array(comment.anchors));
    }
}

fn append_story(txn: &mut TransactionMut<'_>, destination: &TextRef, chunks: &[StoryChunk]) {
    let mut at = destination.len(txn);
    for chunk in chunks {
        match &chunk.insert {
            StoryInsert::Text(text) => {
                destination.insert_with_attributes(txn, at, text, chunk.attrs.clone());
                at += text.encode_utf16().count() as u32;
            }
            StoryInsert::Embed(entries) => {
                let map = destination.insert_embed_with_attributes(
                    txn,
                    at,
                    MapPrelim::default(),
                    chunk.attrs.clone(),
                );
                for (key, value) in entries {
                    map.insert(txn, key.clone(), value.clone());
                }
                at += 1;
            }
        }
    }
}

fn ensure_grid(data: &mut TableData, columns: usize) {
    if data.grid.len() == columns
        && data
            .grid
            .iter()
            .all(|width| any_number(Some(width)).is_some_and(|width| width > 0.0))
    {
        return;
    }
    let table_width = any_number(data.tbl_pr.get("width"))
        .filter(|width| *width > 0.0)
        .or_else(|| {
            let widths: Vec<f64> = data
                .grid
                .iter()
                .filter_map(|width| any_number(Some(width)))
                .collect();
            (!widths.is_empty()).then(|| widths.iter().sum())
        })
        .unwrap_or(9360.0);
    let width = (table_width / columns.max(1) as f64).floor().max(1.0);
    data.grid = vec![Any::Number(width); columns];
}

fn receipt(
    locator: TableLocator,
    data: Option<&TableData>,
    created_story_ids: Vec<String>,
    deleted_story_ids: Vec<String>,
    new_para_ids: Vec<String>,
) -> OpResult<TableReceipt> {
    let (rows, columns, deleted_table) = if let Some(data) = data {
        let (_, columns) = anchors(data)?;
        (data.rows.len() as u32, columns as u32, false)
    } else {
        (0, 0, true)
    };
    Ok(TableReceipt {
        changed_story_ids: vec![locator.story.clone()],
        table: locator,
        rows,
        columns,
        created_story_ids,
        deleted_story_ids,
        new_para_ids,
        deleted_table,
        revision_ids: Vec::new(),
    })
}

fn remove_row_at(
    txn: &mut TransactionMut<'_>,
    data: &mut TableData,
    row: usize,
    deleted: &mut Vec<String>,
) -> OpResult<()> {
    let (cell_anchors, _) = anchors(data)?;
    let mut kept = Vec::new();
    for mut anchor in cell_anchors {
        let end = anchor.row + anchor.rowspan;
        if anchor.row > row {
            anchor.row -= 1;
            kept.push(anchor);
        } else if anchor.row == row && anchor.rowspan == 1 {
            remove_story_tree(txn, &anchor.cell.story, deleted);
        } else if anchor.row <= row && row < end {
            anchor.rowspan -= 1;
            if anchor.row == row {
                anchor.row = row;
            }
            kept.push(anchor);
        } else {
            kept.push(anchor);
        }
    }
    data.rows.remove(row);
    data.rows = reconstruct_rows(std::mem::take(&mut data.rows), kept)?;
    Ok(())
}

pub(crate) fn table_revisions<T: ReadTxn>(table: &MapRef, txn: &T) -> [Option<Any>; 2] {
    let Ok(data) = read_table(table, txn) else {
        return [None, None];
    };
    [TR_INS, TR_DEL].map(|key| {
        let stamp = data.rows.first()?.tr_pr.get(key)?;
        data.rows
            .iter()
            .all(|row| row.tr_pr.get(key).and_then(row_revision_parts).is_some())
            .then(|| stamp.clone())
    })
}

pub(crate) fn table_revision_stamps<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
) -> Vec<(String, String, String)> {
    if map_string(map, txn, KIND_KEY).as_deref() != Some("table") {
        return Vec::new();
    }
    let Ok(data) = read_table(map, txn) else {
        return Vec::new();
    };
    data.rows
        .iter()
        .flat_map(|row| {
            [TR_INS, TR_DEL]
                .into_iter()
                .filter_map(move |key| row.tr_pr.get(key).and_then(row_revision_parts))
        })
        .collect()
}

/// Collects structural row revisions at their containing table embed. Multiple
/// rows from one table insertion intentionally collapse to one entry when they
/// share a revision id, matching the sidebar's "Inserted table" grouping.
pub(crate) fn table_row_changes<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<TableRowChange> {
    let mut changes = Vec::new();
    let mut offset = 0u32;
    for diff in story.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        if let Out::YMap(map) = diff.insert {
            push_table_row_changes(&map, txn, offset, &mut changes);
        }
        offset += len;
    }
    changes
}

/// [`table_row_changes`] of the one embed `map` at story index `offset`, if it is a table.
pub(crate) fn push_table_row_changes<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    offset: u32,
    changes: &mut Vec<TableRowChange>,
) {
    if map_string(map, txn, KIND_KEY).as_deref() != Some("table") {
        return;
    }
    let Ok(data) = read_table(map, txn) else {
        return;
    };
    for (key, row_kind, table_kind) in [
        (
            TR_INS,
            TableRowChangeKind::Insertion,
            TableRowChangeKind::TableInsertion,
        ),
        (
            TR_DEL,
            TableRowChangeKind::Deletion,
            TableRowChangeKind::TableDeletion,
        ),
    ] {
        let revisions: Vec<(String, String, String)> = data
            .rows
            .iter()
            .filter_map(|row| row.tr_pr.get(key).and_then(row_revision_parts))
            .collect();
        let whole_table = revisions.len() == data.rows.len()
            && revisions
                .first()
                .is_some_and(|first| revisions.iter().all(|revision| revision == first));
        if whole_table {
            let (revision_id, author, date) = revisions[0].clone();
            changes.push(TableRowChange {
                revision_id,
                kind: table_kind,
                author,
                date,
                start: offset,
            });
            continue;
        }
        for (revision_id, author, date) in revisions {
            if changes.iter().any(|change: &TableRowChange| {
                change.start == offset
                    && change.kind == row_kind
                    && change.revision_id == revision_id
            }) {
                continue;
            }
            changes.push(TableRowChange {
                revision_id,
                kind: row_kind,
                author,
                date,
                start: offset,
            });
        }
    }
}

/// Resolves every matching `trIns`/`trDel` in one story. Row removals also
/// delete the row's now-unreachable cell-story trees; removing every row
/// removes the structural table embed itself.
pub(crate) fn resolve_table_row_revisions(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    story_id: &str,
    accept: bool,
    span: Option<(u32, u32)>,
    filter: Option<&str>,
    resolved: &mut Vec<String>,
) -> OpResult<u32> {
    let (span_start, span_end) = span.unwrap_or((0, u32::MAX));
    let mut tables = Vec::new();
    let mut offset = 0u32;
    let mut table_index = 0u32;
    for diff in story.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        if let Out::YMap(map) = diff.insert
            && map_string(&map, txn, KIND_KEY).as_deref() == Some("table")
        {
            if offset >= span_start && offset < span_end {
                tables.push((offset, table_index, map));
            }
            table_index += 1;
        }
        offset += len;
    }

    let mut removed_tables = 0;
    for (table_offset, table_index, table) in tables.into_iter().rev() {
        let mut data = read_table(&table, txn)?;
        let mut remove = Vec::new();
        let mut changed = false;
        for (row_index, row) in data.rows.iter_mut().enumerate() {
            let ins = matching_row_revision(row.tr_pr.get(TR_INS), filter);
            let del = matching_row_revision(row.tr_pr.get(TR_DEL), filter);
            let remove_row = if accept { del.is_some() } else { ins.is_some() };
            if remove_row {
                record_row_revision(resolved, if accept { del.as_ref() } else { ins.as_ref() });
                remove.push(row_index);
                changed = true;
            } else if accept && ins.is_some() {
                record_row_revision(resolved, ins.as_ref());
                row.tr_pr.remove(TR_INS);
                changed = true;
            } else if !accept && del.is_some() {
                record_row_revision(resolved, del.as_ref());
                row.tr_pr.remove(TR_DEL);
                changed = true;
            }
        }
        if !changed {
            continue;
        }
        if remove.len() == data.rows.len() {
            let locator = TableLocator::new(story_id, table_index);
            delete_table_in_txn(txn, &locator, story, table_offset, &data)?;
            removed_tables += 1;
            continue;
        }
        let mut deleted = Vec::new();
        for row in remove.into_iter().rev() {
            remove_row_at(txn, &mut data, row, &mut deleted)?;
        }
        write_table(txn, &table, &data);
    }
    Ok(removed_tables)
}

fn delete_table_in_txn(
    txn: &mut TransactionMut<'_>,
    locator: &TableLocator,
    parent_story: &TextRef,
    table_index: u32,
    data: &TableData,
) -> OpResult<TableReceipt> {
    let mut deleted = Vec::new();
    for row in &data.rows {
        for cell in &row.cells {
            remove_story_tree(txn, &cell.story, &mut deleted);
        }
    }
    parent_story.remove_range(txn, table_index, 1);
    receipt(locator.clone(), None, Vec::new(), deleted, Vec::new())
}

impl EditingDoc {
    /// The payload of `locator`'s table embed, as [`EditingDoc::story_segments`]
    /// gives it, without reading the rest of the story's segments.
    pub fn table_payload(&self, locator: &TableLocator) -> OpResult<BTreeMap<String, Any>> {
        let txn = self.yrs_doc().transact();
        let (_, table, _) = table_at(&txn, locator)?;
        Ok(crate::embed_payload(&table, &txn))
    }

    /// Story-global position of the table embed. Used by wasm awareness state
    /// to make a cell selection sticky without putting it in the document.
    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn table_embed_index(&self, locator: &TableLocator) -> OpResult<u32> {
        let txn = self.yrs_doc().transact();
        table_at(&txn, locator).map(|(_, _, index)| index)
    }

    /// Resolves the table ordinal currently sitting at a sticky story index.
    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn table_locator_at_index(
        &self,
        story_id: &str,
        target_index: u32,
    ) -> OpResult<TableLocator> {
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, story_id).map_err(OpError::from)?;
        let mut offset = 0u32;
        let mut ordinal = 0u32;
        for diff in story.diff(&txn, YChange::identity) {
            let len = out_len(&diff.insert);
            if let Out::YMap(map) = diff.insert
                && map_string(&map, &txn, KIND_KEY).as_deref() == Some("table")
            {
                if offset == target_index {
                    return Ok(TableLocator::new(story_id, ordinal));
                }
                ordinal += 1;
            }
            offset += len;
        }
        Err(OpError::UnknownTable {
            story: story_id.to_owned(),
            table_index: ordinal,
        })
    }

    /// Canonicalizes a grid location to the covering cell's anchor and story.
    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn resolve_cell_identity(&self, loc: &CellLoc) -> OpResult<(CellLoc, String)> {
        let txn = self.yrs_doc().transact();
        let (_, table, _) = table_at(&txn, &loc.table())?;
        let data = read_table(&table, &txn)?;
        let (cell_anchors, columns) = anchors(&data)?;
        let row = loc.row as usize;
        let column = loc.column as usize;
        if row >= data.rows.len() || column >= columns {
            return Err(invalid(format!(
                "cell {},{} is outside the {}x{} table",
                loc.row,
                loc.column,
                data.rows.len(),
                columns
            )));
        }
        let anchor = covering(&cell_anchors, row, column)
            .ok_or_else(|| invalid("cell location points at an unoccupied grid slot"))?;
        Ok((
            CellLoc::new(
                loc.story.clone(),
                loc.table_index,
                anchor.row as u32,
                anchor.column as u32,
            ),
            anchor.cell.story.clone(),
        ))
    }

    /// Finds a cell by stable cell-story identity, falling back to a clamped
    /// grid coordinate if the selected cell was removed by a local edit.
    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn cell_loc_for_story(
        &self,
        locator: &TableLocator,
        cell_story: &str,
        fallback_row: u32,
        fallback_column: u32,
    ) -> OpResult<CellLoc> {
        let txn = self.yrs_doc().transact();
        let (_, table, _) = table_at(&txn, locator)?;
        let data = read_table(&table, &txn)?;
        let (cell_anchors, columns) = anchors(&data)?;
        if let Some(anchor) = cell_anchors
            .iter()
            .find(|anchor| anchor.cell.story == cell_story)
        {
            return Ok(CellLoc::new(
                locator.story.clone(),
                locator.table_index,
                anchor.row as u32,
                anchor.column as u32,
            ));
        }
        let row = (fallback_row as usize).min(data.rows.len() - 1);
        let column = (fallback_column as usize).min(columns - 1);
        let anchor = covering(&cell_anchors, row, column)
            .or_else(|| cell_anchors.last())
            .ok_or_else(|| invalid("table has no resolvable cells"))?;
        Ok(CellLoc::new(
            locator.story.clone(),
            locator.table_index,
            anchor.row as u32,
            anchor.column as u32,
        ))
    }

    /// Inserts a rectangular structural table at a story position. Each cell
    /// receives an independent one-paragraph story. In suggesting mode every
    /// row carries the same `trIns` revision, so one resolve action accepts or
    /// rejects the complete table insertion.
    pub fn insert_table(
        &self,
        ctx: &EditCtx,
        mut at: Position,
        rows: u32,
        columns: u32,
    ) -> OpResult<TableReceipt> {
        if rows == 0 || columns == 0 {
            return Err(invalid("table dimensions must be positive"));
        }
        let revision_id = ctx.is_suggesting().then(|| self.next_id());
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        check_position(&story, &txn, at.index)?;
        at.index = crate::ops::code_point_range(&story, &txn, at.index, at.index).0;
        let split_para_id = if paragraph_content_before(&story, &txn, at.index) {
            let split = self.split_paragraph_in_transaction(
                ctx,
                &mut txn,
                at.clone(),
                None,
                revision_id.clone(),
            )?;
            at.index += 1;
            Some(split.second_para_id)
        } else {
            None
        };

        let mut ordinal = 0u32;
        let mut offset = 0u32;
        for diff in story.diff(&txn, YChange::identity) {
            if offset >= at.index {
                break;
            }
            if let Out::YMap(map) = &diff.insert
                && map_string(map, &txn, KIND_KEY).as_deref() == Some("table")
            {
                ordinal += 1;
            }
            offset += out_len(&diff.insert);
        }
        let locator = TableLocator::new(at.story.clone(), ordinal);
        let mut used = story_ids(&txn);
        let mut ids = IdAllocator::new(self, &txn);
        let table_slot = fresh_table_slot(&used, &at.story, ordinal as usize);
        let base = format!("{}:t{table_slot}", at.story);
        let mut created_story_ids = Vec::with_capacity((rows * columns) as usize);
        let mut new_para_ids = Vec::with_capacity((rows * columns) as usize + 1);
        new_para_ids.extend(split_para_id);
        let revision = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        let borders = default_cell_borders();
        let mut table_rows = Vec::with_capacity(rows as usize);
        for row in 0..rows as usize {
            let mut cells = Vec::with_capacity(columns as usize);
            for column in 0..columns as usize {
                let story_id = fresh_cell_story(&mut used, &base, row, column, self.client_id());
                let para_id = create_cell_story(self, &mut txn, &mut ids, &story_id)?;
                cells.push(CellData {
                    tc_pr: HashMap::from([
                        ("rowspan".to_owned(), Any::Number(1.0)),
                        ("colspan".to_owned(), Any::Number(1.0)),
                        ("borders".to_owned(), borders.clone()),
                    ]),
                    story: story_id.clone(),
                });
                created_story_ids.push(story_id);
                new_para_ids.push(para_id);
            }
            let mut tr_pr = HashMap::new();
            if let Some(revision) = revision.as_ref() {
                tr_pr.insert(TR_INS.to_owned(), revision.clone());
            }
            table_rows.push(RowData { tr_pr, cells });
        }
        let width = (9360.0 / columns as f64).floor().max(1.0);
        let data = TableData {
            tbl_pr: HashMap::from([
                ("width".to_owned(), Any::Number(9360.0)),
                ("widthType".to_owned(), Any::from("dxa")),
            ]),
            grid: vec![Any::Number(width); columns as usize],
            rows: table_rows,
        };
        let table = story.insert_embed_with_attributes(
            &mut txn,
            at.index,
            MapPrelim::default(),
            insertion_attrs(None, None),
        );
        table.insert(&mut txn, KIND_KEY, "table");
        write_table(&mut txn, &table, &data);
        let mut receipt = receipt(
            locator,
            Some(&data),
            created_story_ids,
            Vec::new(),
            new_para_ids,
        )?;
        receipt.revision_ids = revision_id.into_iter().collect();
        Ok(receipt)
    }

    pub fn validate_table_fragment(fragment: &Value) -> OpResult<()> {
        let mut remaining = MAX_FRAGMENT_CELLS;
        validate_fragment_table(fragment, 1, &mut remaining)?;
        Ok(())
    }

    /// Inserts a clipboard table fragment at a story position, creating one
    /// story per cell with the fragment's content. Mirrors [`EditingDoc::insert_table`].
    pub fn insert_table_fragment(
        &self,
        ctx: &EditCtx,
        mut at: Position,
        fragment: &Value,
    ) -> OpResult<TableFragmentReceipt> {
        let mut remaining = MAX_FRAGMENT_CELLS;
        let (valid, columns) = validate_fragment_table(fragment, 1, &mut remaining)?;
        let revision_id = ctx.is_suggesting().then(|| self.next_id());
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        check_position(&story, &txn, at.index)?;
        at.index = crate::ops::code_point_range(&story, &txn, at.index, at.index).0;
        let split_para_id = if paragraph_content_before(&story, &txn, at.index) {
            let split = self.split_paragraph_in_transaction(
                ctx,
                &mut txn,
                at.clone(),
                None,
                revision_id.clone(),
            )?;
            at.index += 1;
            Some(split.second_para_id)
        } else {
            None
        };
        let mut ordinal = 0u32;
        let mut offset = 0u32;
        for diff in story.diff(&txn, YChange::identity) {
            if offset >= at.index {
                break;
            }
            if let Out::YMap(map) = &diff.insert
                && map_string(map, &txn, KIND_KEY).as_deref() == Some("table")
            {
                ordinal += 1;
            }
            offset += out_len(&diff.insert);
        }
        let locator = TableLocator::new(at.story.clone(), ordinal);
        let mut used = story_ids(&txn);
        let mut ids = IdAllocator::new(self, &txn);
        let slot = fresh_table_slot(&used, &at.story, ordinal as usize);
        let revision = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        let mut build = FragmentBuild::default();
        build.new_para_ids.extend(split_para_id);
        let data = self.insert_fragment_table_data(
            &mut txn,
            &mut ids,
            &mut used,
            &story,
            &at.story,
            at.index,
            slot,
            &valid,
            revision.as_ref(),
            insertion_attrs(None, None),
            &mut build,
        )?;
        let after = caret_after_insert(&story, &txn, &at.story, at.index)?;
        debug_assert_eq!(data.rows.len() as u32, valid.rows.len() as u32);
        Ok(TableFragmentReceipt {
            table: locator,
            rows: data.rows.len() as u32,
            columns: columns as u32,
            created_story_ids: build.created_story_ids,
            deleted_story_ids: Vec::new(),
            changed_story_ids: vec![at.story.clone()],
            new_para_ids: build.new_para_ids,
            deleted_table: false,
            revision_ids: revision_id.into_iter().collect(),
            after,
        })
    }

    /// Builds one validated table's rows, creating each cell story with the
    /// fragment's content, then inserts the table embed. Nested tables
    /// recurse under their cell story.
    #[allow(clippy::too_many_arguments)]
    fn insert_fragment_table_data(
        &self,
        txn: &mut TransactionMut<'_>,
        ids: &mut IdAllocator,
        used: &mut HashSet<String>,
        parent: &TextRef,
        parent_story_id: &str,
        index: u32,
        slot: usize,
        valid: &FragmentTable,
        revision: Option<&Any>,
        embed_attrs: Attrs,
        build: &mut FragmentBuild,
    ) -> OpResult<TableData> {
        let base = format!("{parent_story_id}:t{slot}");
        let mut table_rows = Vec::with_capacity(valid.rows.len());
        for (row_index, row) in valid.rows.iter().enumerate() {
            let mut cells = Vec::with_capacity(row.cells.len());
            for (cell_index, cell) in row.cells.iter().enumerate() {
                let story_id =
                    fresh_cell_story(used, &base, row_index, cell_index, self.client_id());
                let stories = txn
                    .get_map(STORIES)
                    .expect("stories root is declared by EditingDoc::new");
                if stories.contains_key(txn, &story_id) {
                    return Err(OpError::StoryExists(story_id));
                }
                let cell_story = stories.insert(txn, story_id.as_str(), TextPrelim::new(""));
                build.created_story_ids.push(story_id.clone());
                self.write_fragment_content(
                    txn,
                    ids,
                    used,
                    &cell_story,
                    &story_id,
                    &cell.content,
                    revision,
                    build,
                )?;
                let rowspan = any_usize(cell.tc_pr.get("rowspan"), 1, "rowspan")?;
                let colspan = any_usize(cell.tc_pr.get("colspan"), 1, "colspan")?;
                let mut data = CellData {
                    tc_pr: cell.tc_pr.clone(),
                    story: story_id,
                };
                set_spans(&mut data, rowspan, colspan);
                if let Some(shading) = cell.shading.as_ref() {
                    data.tc_pr.insert(
                        "backgroundColor".to_owned(),
                        Any::from(shading.trim_start_matches('#')),
                    );
                }
                cells.push(data);
            }
            let mut tr_pr = fresh_row_properties(row.tr_pr.clone());
            tr_pr.remove(TR_INS);
            tr_pr.remove(TR_DEL);
            if let Some(stamp) = revision {
                tr_pr.insert(TR_INS.to_owned(), stamp.clone());
            }
            table_rows.push(RowData { tr_pr, cells });
        }
        let mut tbl_pr = valid.tbl_pr.clone();
        if tbl_pr.is_empty() {
            tbl_pr.insert("width".to_owned(), Any::Number(9360.0));
            tbl_pr.insert("widthType".to_owned(), Any::from("dxa"));
        }
        let mut data = TableData {
            tbl_pr,
            grid: valid.grid.clone(),
            rows: table_rows,
        };
        let (_, columns) = anchors(&data)?;
        ensure_grid(&mut data, columns);
        let table =
            parent.insert_embed_with_attributes(txn, index, MapPrelim::default(), embed_attrs);
        table.insert(txn, KIND_KEY, "table");
        write_table(txn, &table, &data);
        Ok(data)
    }

    /// Writes validated cell content into a fresh cell story. Every pilcrow
    /// is minted a fresh paragraph id; nested tables recurse.
    #[allow(clippy::too_many_arguments)]
    fn write_fragment_content(
        &self,
        txn: &mut TransactionMut<'_>,
        ids: &mut IdAllocator,
        used: &mut HashSet<String>,
        story: &TextRef,
        story_id: &str,
        content: &[FragmentSegment],
        revision: Option<&Any>,
        build: &mut FragmentBuild,
    ) -> OpResult<()> {
        let mut at = 0u32;
        let mut nested = 0usize;
        for segment in content {
            match segment {
                FragmentSegment::Text(part) => {
                    story.insert_with_attributes(txn, at, part.text.as_str(), part.attrs.clone());
                    at += part.text.encode_utf16().count() as u32;
                }
                FragmentSegment::TabBreak(part) => {
                    let map = story.insert_embed_with_attributes(
                        txn,
                        at,
                        MapPrelim::default(),
                        part.attrs.clone(),
                    );
                    map.insert(txn, KIND_KEY, part.kind);
                    for (key, value) in &part.payload {
                        map.insert(txn, key.clone(), value.clone());
                    }
                    at += 1;
                }
                FragmentSegment::Pilcrow(part) => {
                    let key = ids.session_key(self);
                    let map = story.insert_embed_with_attributes(
                        txn,
                        at,
                        MapPrelim::default(),
                        part.attrs.clone(),
                    );
                    map.insert(txn, KIND_KEY, PILCROW_KIND);
                    map.insert(txn, PARA_ID, key.as_str());
                    for (prop_key, prop_value) in &part.props {
                        map.insert(txn, prop_key.clone(), prop_value.clone());
                    }
                    ids.bind(txn, &map, &key, ParagraphIdOrigin::Authored);
                    build.new_para_ids.push(key);
                    at += 1;
                }
                FragmentSegment::Table(part) => {
                    let slot = fresh_table_slot(used, story_id, nested);
                    nested += 1;
                    let mut attrs = insertion_attrs(None, None);
                    for (key, value) in part.attrs.iter() {
                        attrs.insert(key.clone(), value.clone());
                    }
                    self.insert_fragment_table_data(
                        txn,
                        ids,
                        used,
                        story,
                        story_id,
                        at,
                        slot,
                        &part.table,
                        revision,
                        attrs,
                        build,
                    )?;
                    at += 1;
                }
            }
        }
        Ok(())
    }

    /// Inserts a row above (`after = false`) or below (`after = true`) the
    /// cell's covering row. Spans crossing the insertion boundary grow by one.
    pub fn insert_row(&self, ctx: &EditCtx, at: &CellLoc, after: bool) -> OpResult<TableReceipt> {
        let locator = at.table();
        let revision_id = ctx.is_suggesting().then(|| self.next_id());
        let mut txn = self.transact_for(ctx);
        let (_, table, _) = table_at(&txn, &locator)?;
        let mut data = read_table(&table, &txn)?;
        let (mut cell_anchors, columns) = anchors(&data)?;
        let original_anchors = cell_anchors.clone();
        let target = covering(&cell_anchors, at.row as usize, at.column as usize)
            .ok_or_else(|| invalid("insert-row target is outside the table"))?;
        let boundary = if after {
            target.row + target.rowspan
        } else {
            target.row
        };
        let template_row = target.row.min(data.rows.len() - 1);
        let mut tr_pr = fresh_row_properties(data.rows[template_row].tr_pr.clone());
        tr_pr.remove(TR_INS);
        tr_pr.remove(TR_DEL);
        if let Some(revision_id) = revision_id.as_ref() {
            tr_pr.insert(
                TR_INS.to_owned(),
                revision_value(revision_id, &ctx.revision_author()),
            );
        }
        let used = story_ids(&txn);
        let mut ids = IdAllocator::new(self, &txn);
        let base = table_base(&locator);
        let row_slot = fresh_row_slot(&used, &base, data.rows.len(), columns);
        let mut used = used;
        let mut created = Vec::new();
        let mut new_para_ids = Vec::new();
        let mut inserted_anchors = Vec::new();

        for anchor in &mut cell_anchors {
            if anchor.row < boundary && boundary < anchor.row + anchor.rowspan {
                anchor.rowspan += 1;
            } else if anchor.row >= boundary {
                anchor.row += 1;
            }
        }

        for column in 0..columns {
            let covered_by_crossing_span = cell_anchors.iter().any(|anchor| {
                anchor.row < boundary
                    && boundary < anchor.row + anchor.rowspan
                    && anchor.column <= column
                    && column < anchor.column + anchor.colspan
            });
            if covered_by_crossing_span {
                continue;
            }
            let template = covering(&original_anchors, template_row, column)
                .map(|anchor| anchor.cell.clone())
                .unwrap_or(CellData {
                    tc_pr: HashMap::new(),
                    story: String::new(),
                });
            let story_id = fresh_cell_story(&mut used, &base, row_slot, column, self.client_id());
            let para_id = create_cell_story(self, &mut txn, &mut ids, &story_id)?;
            let mut cell = reset_cell_spans(template);
            cell.story = story_id.clone();
            inserted_anchors.push(CellAnchor {
                row: boundary,
                column,
                rowspan: 1,
                colspan: 1,
                cell_index: 0,
                cell,
            });
            created.push(story_id);
            new_para_ids.push(para_id);
        }

        data.rows.insert(
            boundary,
            RowData {
                tr_pr,
                cells: Vec::new(),
            },
        );
        cell_anchors.extend(inserted_anchors);
        data.rows = reconstruct_rows(data.rows, cell_anchors)?;
        ensure_grid(&mut data, columns);
        write_table(&mut txn, &table, &data);
        let mut receipt = receipt(locator, Some(&data), created, Vec::new(), new_para_ids)?;
        receipt.revision_ids = revision_id.into_iter().collect();
        Ok(receipt)
    }

    /// Inserts a column left (`after = false`) or right (`after = true`) of
    /// the target cell. A colspan crossing the boundary grows by one.
    pub fn insert_column(
        &self,
        ctx: &EditCtx,
        at: &CellLoc,
        after: bool,
    ) -> OpResult<TableReceipt> {
        let locator = at.table();
        let mut txn = self.transact_for(ctx);
        let (_, table, _) = table_at(&txn, &locator)?;
        let mut data = read_table(&table, &txn)?;
        let (mut cell_anchors, columns) = anchors(&data)?;
        let original_anchors = cell_anchors.clone();
        let target = covering(&cell_anchors, at.row as usize, at.column as usize)
            .ok_or_else(|| invalid("insert-column target is outside the table"))?;
        let boundary = if after {
            target.column + target.colspan
        } else {
            target.column
        };
        let used = story_ids(&txn);
        let mut ids = IdAllocator::new(self, &txn);
        let base = table_base(&locator);
        let column_slot = fresh_column_slot(&used, &base, data.rows.len(), columns);
        let mut used = used;
        let mut created = Vec::new();
        let mut new_para_ids = Vec::new();

        for anchor in &mut cell_anchors {
            if anchor.column < boundary && boundary < anchor.column + anchor.colspan {
                anchor.colspan += 1;
            } else if anchor.column >= boundary {
                anchor.column += 1;
            }
        }

        let mut inserted = Vec::new();
        for row in 0..data.rows.len() {
            let covered_by_crossing_span = cell_anchors.iter().any(|anchor| {
                anchor.column < boundary
                    && boundary < anchor.column + anchor.colspan
                    && anchor.row <= row
                    && row < anchor.row + anchor.rowspan
            });
            if covered_by_crossing_span {
                continue;
            }
            let template_column = if boundary == 0 {
                0
            } else {
                (boundary - 1).min(columns - 1)
            };
            let template = covering(&original_anchors, row, template_column)
                .map(|anchor| anchor.cell.clone())
                .unwrap_or(CellData {
                    tc_pr: HashMap::new(),
                    story: String::new(),
                });
            let story_id = fresh_cell_story(&mut used, &base, row, column_slot, self.client_id());
            let para_id = create_cell_story(self, &mut txn, &mut ids, &story_id)?;
            let mut cell = reset_cell_spans(template);
            cell.story = story_id.clone();
            inserted.push(CellAnchor {
                row,
                column: boundary,
                rowspan: 1,
                colspan: 1,
                cell_index: 0,
                cell,
            });
            created.push(story_id);
            new_para_ids.push(para_id);
        }
        cell_anchors.extend(inserted);
        data.rows = reconstruct_rows(data.rows, cell_anchors)?;
        ensure_grid(&mut data, columns);
        let inserted_width = if boundary > 0 {
            data.grid[boundary - 1].clone()
        } else {
            data.grid[0].clone()
        };
        data.grid.insert(boundary, inserted_width);
        write_table(&mut txn, &table, &data);
        receipt(locator, Some(&data), created, Vec::new(), new_para_ids)
    }

    /// Deletes every row touched by `range`. Cells crossing the deletion band
    /// shrink; a cell anchored in a deleted row but extending below it moves to
    /// the first surviving row. Deleting every row removes the whole table.
    pub fn delete_row(&self, ctx: &EditCtx, range: &TableRange) -> OpResult<TableReceipt> {
        let locator = range.anchor.table();
        let mut txn = self.transact_for(ctx);
        let (parent_story, table, table_story_index) = table_at(&txn, &locator)?;
        let mut data = read_table(&table, &txn)?;
        let rect = checked_rect(&data, range)?;
        if ctx.is_suggesting() {
            let revision_id = self.next_id();
            let revision = revision_value(&revision_id, &ctx.revision_author());
            for row in &mut data.rows[rect.top..rect.bottom] {
                row.tr_pr.insert(TR_DEL.to_owned(), revision.clone());
            }
            write_table(&mut txn, &table, &data);
            let mut receipt = receipt(locator, Some(&data), Vec::new(), Vec::new(), Vec::new())?;
            receipt.revision_ids.push(revision_id);
            return Ok(receipt);
        }
        if rect.top == 0 && rect.bottom == data.rows.len() {
            return delete_table_in_txn(
                &mut txn,
                &locator,
                &parent_story,
                table_story_index,
                &data,
            );
        }
        let removed_rows = rect.bottom - rect.top;
        let (cell_anchors, _) = anchors(&data)?;
        let mut kept = Vec::new();
        let mut deleted = Vec::new();
        for mut anchor in cell_anchors {
            let start = anchor.row;
            let end = anchor.row + anchor.rowspan;
            let overlap = end.min(rect.bottom).saturating_sub(start.max(rect.top));
            if overlap == 0 {
                if anchor.row >= rect.bottom {
                    anchor.row -= removed_rows;
                }
                kept.push(anchor);
            } else if overlap == anchor.rowspan {
                remove_story_tree(&mut txn, &anchor.cell.story, &mut deleted);
            } else {
                anchor.rowspan -= overlap;
                if start >= rect.top {
                    anchor.row = rect.top;
                }
                kept.push(anchor);
            }
        }
        data.rows.drain(rect.top..rect.bottom);
        data.rows = reconstruct_rows(data.rows, kept)?;
        write_table(&mut txn, &table, &data);
        receipt(locator, Some(&data), Vec::new(), deleted, Vec::new())
    }

    /// Deletes every grid column touched by `range`. Colspans crossing the
    /// deletion band shrink. Deleting all columns removes the whole table.
    pub fn delete_column(&self, ctx: &EditCtx, range: &TableRange) -> OpResult<TableReceipt> {
        let locator = range.anchor.table();
        let mut txn = self.transact_for(ctx);
        let (parent_story, table, table_story_index) = table_at(&txn, &locator)?;
        let mut data = read_table(&table, &txn)?;
        let rect = checked_rect(&data, range)?;
        let (_, columns) = anchors(&data)?;
        if rect.left == 0 && rect.right == columns {
            return delete_table_in_txn(
                &mut txn,
                &locator,
                &parent_story,
                table_story_index,
                &data,
            );
        }
        let removed_columns = rect.right - rect.left;
        let (cell_anchors, _) = anchors(&data)?;
        let mut kept = Vec::new();
        let mut deleted = Vec::new();
        for mut anchor in cell_anchors {
            let start = anchor.column;
            let end = anchor.column + anchor.colspan;
            let overlap = end.min(rect.right).saturating_sub(start.max(rect.left));
            if overlap == 0 {
                if anchor.column >= rect.right {
                    anchor.column -= removed_columns;
                }
                kept.push(anchor);
            } else if overlap == anchor.colspan {
                remove_story_tree(&mut txn, &anchor.cell.story, &mut deleted);
            } else {
                anchor.colspan -= overlap;
                if start >= rect.left {
                    anchor.column = rect.left;
                }
                kept.push(anchor);
            }
        }
        data.rows = reconstruct_rows(data.rows, kept)?;
        ensure_grid(&mut data, columns);
        data.grid.drain(rect.left..rect.right);
        write_table(&mut txn, &table, &data);
        receipt(locator, Some(&data), Vec::new(), deleted, Vec::new())
    }

    /// Removes the structural embed and every cell story reachable beneath it.
    pub fn delete_table(&self, ctx: &EditCtx, locator: &TableLocator) -> OpResult<TableReceipt> {
        let mut txn = self.transact_for(ctx);
        let (parent_story, table, table_story_index) = table_at(&txn, locator)?;
        let data = read_table(&table, &txn)?;
        delete_table_in_txn(&mut txn, locator, &parent_story, table_story_index, &data)
    }

    /// Merges a rectangular range into its top-left cell. Source cell stories
    /// are appended (paragraph-for-paragraph, preserving attrs and embeds) to
    /// the survivor before being marked removed.
    pub fn merge_cells(&self, ctx: &EditCtx, range: &TableRange) -> OpResult<TableReceipt> {
        let locator = range.anchor.table();
        let mut txn = self.transact_for(ctx);
        let (_, table, _) = table_at(&txn, &locator)?;
        let mut data = read_table(&table, &txn)?;
        let rect = checked_rect(&data, range)?;
        let (cell_anchors, _) = anchors(&data)?;
        let mut selected: Vec<CellAnchor> = cell_anchors
            .iter()
            .filter(|anchor| {
                anchor.row >= rect.top
                    && anchor.row + anchor.rowspan <= rect.bottom
                    && anchor.column >= rect.left
                    && anchor.column + anchor.colspan <= rect.right
            })
            .cloned()
            .collect();
        selected.sort_by_key(|anchor| (anchor.row, anchor.column));
        if selected.len() < 2 {
            return Err(invalid("merge requires at least two cells"));
        }
        let target = selected
            .first()
            .cloned()
            .ok_or_else(|| invalid("merge range contains no cells"))?;
        if target.row != rect.top || target.column != rect.left {
            return Err(invalid("merge range does not have a top-left anchor cell"));
        }
        for row in rect.top..rect.bottom {
            for column in rect.left..rect.right {
                if covering(&cell_anchors, row, column).is_none() {
                    return Err(invalid("merge range contains an unoccupied grid slot"));
                }
            }
        }
        let destination = story_ref(&txn, &target.cell.story).map_err(OpError::from)?;
        let mut deleted = Vec::new();
        let source_ids: HashSet<String> = selected
            .iter()
            .skip(1)
            .map(|anchor| anchor.cell.story.clone())
            .collect();
        let mut sources = Vec::new();
        let mut offsets = HashMap::new();
        let mut offset = destination.len(&txn);
        for source in selected.iter().skip(1) {
            let source_story = story_ref(&txn, &source.cell.story).map_err(OpError::from)?;
            sources.push(snapshot_story(&source_story, &txn)?);
            offsets.insert(source.cell.story.clone(), offset);
            offset = offset
                .checked_add(source_story.len(&txn))
                .ok_or_else(|| invalid("merged story exceeds the length limit"))?;
        }
        let comments = merged_comments(&txn, &offsets)?;
        for chunks in sources {
            append_story(&mut txn, &destination, &chunks);
        }
        move_merged_comments(&mut txn, &destination, &target.cell.story, comments);
        for source in selected.iter().skip(1) {
            remove_story_exact(&mut txn, &source.cell.story, &mut deleted);
        }
        let mut kept: Vec<CellAnchor> = cell_anchors
            .into_iter()
            .filter(|anchor| !source_ids.contains(&anchor.cell.story))
            .collect();
        let survivor = kept
            .iter_mut()
            .find(|anchor| anchor.cell.story == target.cell.story)
            .ok_or_else(|| invalid("merge survivor disappeared"))?;
        survivor.row = rect.top;
        survivor.column = rect.left;
        survivor.rowspan = rect.bottom - rect.top;
        survivor.colspan = rect.right - rect.left;
        data.rows = reconstruct_rows(data.rows, kept)?;
        write_table(&mut txn, &table, &data);
        let mut receipt = receipt(locator, Some(&data), Vec::new(), deleted, Vec::new())?;
        receipt.changed_story_ids.push(target.cell.story);
        Ok(receipt)
    }

    /// Splits the merged cell covering `at` back into one cell per grid slot.
    /// The original story remains in the top-left slot; all other slots get a
    /// fresh one-paragraph cell story.
    pub fn split_cell(&self, ctx: &EditCtx, at: &CellLoc) -> OpResult<TableReceipt> {
        self.split_cell_grid(ctx, at, None, None)
    }

    /// Splits the cell covering `at` into the requested rectangular grid.
    /// Omitting the dimensions preserves `split_cell`'s unmerge behavior.
    pub fn split_cell_grid(
        &self,
        ctx: &EditCtx,
        at: &CellLoc,
        rows: Option<u32>,
        columns: Option<u32>,
    ) -> OpResult<TableReceipt> {
        let locator = at.table();
        let mut txn = self.transact_for(ctx);
        let (_, table, _) = table_at(&txn, &locator)?;
        let mut data = read_table(&table, &txn)?;
        let (cell_anchors, existing_columns) = anchors(&data)?;
        let target_index = cell_anchors
            .iter()
            .position(|anchor| {
                anchor.row <= at.row as usize
                    && (at.row as usize) < anchor.row + anchor.rowspan
                    && anchor.column <= at.column as usize
                    && (at.column as usize) < anchor.column + anchor.colspan
            })
            .ok_or_else(|| invalid("split target is outside the table"))?;
        let target = cell_anchors[target_index].clone();
        let requested_rows = rows.map(|value| value as usize).unwrap_or(target.rowspan);
        let requested_columns = columns
            .map(|value| value as usize)
            .unwrap_or(target.colspan);
        if requested_rows < target.rowspan || requested_columns < target.colspan {
            return Err(invalid("split dimensions cannot shrink the covered grid"));
        }
        if requested_rows == 0 || requested_columns == 0 {
            return Err(invalid("split dimensions must be positive"));
        }
        if requested_rows == 1 && requested_columns == 1 {
            return Err(invalid("cell is not merged"));
        }

        let delta_rows = requested_rows - target.rowspan;
        let delta_columns = requested_columns - target.colspan;
        let target_row_end = target.row + target.rowspan;
        let target_column_end = target.column + target.colspan;
        let mut next_anchors =
            Vec::with_capacity(cell_anchors.len() - 1 + requested_rows * requested_columns);
        for (index, mut anchor) in cell_anchors.into_iter().enumerate() {
            if index == target_index {
                continue;
            }
            let row_end = anchor.row + anchor.rowspan;
            let column_end = anchor.column + anchor.colspan;
            let row_intersects = anchor.row < target_row_end && row_end > target.row;
            let column_intersects = anchor.column < target_column_end && column_end > target.column;
            if anchor.row >= target_row_end {
                anchor.row += delta_rows;
            }
            if anchor.column >= target_column_end {
                anchor.column += delta_columns;
            }
            if delta_rows > 0 && row_intersects && !column_intersects {
                anchor.rowspan += delta_rows;
            }
            if delta_columns > 0 && column_intersects && !row_intersects {
                anchor.colspan += delta_columns;
            }
            next_anchors.push(anchor);
        }

        let base = table_base(&locator);
        let mut used = story_ids(&txn);
        let mut ids = IdAllocator::new(self, &txn);
        let mut created = Vec::new();
        let mut new_para_ids = Vec::new();
        for row in target.row..target.row + requested_rows {
            for column in target.column..target.column + requested_columns {
                if row == target.row && column == target.column {
                    let cell = reset_cell_spans(target.cell.clone());
                    next_anchors.push(CellAnchor {
                        row,
                        column,
                        rowspan: 1,
                        colspan: 1,
                        cell_index: 0,
                        cell,
                    });
                    continue;
                }
                let story_id = fresh_cell_story(&mut used, &base, row, column, self.client_id());
                let para_id = create_cell_story(self, &mut txn, &mut ids, &story_id)?;
                let mut cell = reset_cell_spans(target.cell.clone());
                cell.story = story_id.clone();
                next_anchors.push(CellAnchor {
                    row,
                    column,
                    rowspan: 1,
                    colspan: 1,
                    cell_index: 0,
                    cell,
                });
                created.push(story_id);
                new_para_ids.push(para_id);
            }
        }

        let original_rows = data.rows.clone();
        data.rows = (0..original_rows.len() + delta_rows)
            .map(|row| {
                let tr_pr = if row < target_row_end {
                    original_rows[row].tr_pr.clone()
                } else if row < target.row + requested_rows {
                    fresh_row_properties(original_rows[target_row_end - 1].tr_pr.clone())
                } else {
                    original_rows[row - delta_rows].tr_pr.clone()
                };
                RowData {
                    tr_pr,
                    cells: Vec::new(),
                }
            })
            .collect();

        ensure_grid(&mut data, existing_columns);
        let split_width: f64 = data.grid[target.column..target_column_end]
            .iter()
            .filter_map(|width| any_number(Some(width)))
            .sum();
        let segment_width = (split_width / requested_columns as f64).floor();
        let remainder = split_width - segment_width * requested_columns as f64;
        let replacement = (0..requested_columns)
            .map(|index| {
                Any::Number(segment_width + if (index as f64) < remainder { 1.0 } else { 0.0 })
            })
            .collect::<Vec<_>>();
        data.grid
            .splice(target.column..target_column_end, replacement);
        data.rows = reconstruct_rows(data.rows, next_anchors)?;
        write_table(&mut txn, &table, &data);
        receipt(locator, Some(&data), created, Vec::new(), new_para_ids)
    }

    /// Sets or clears `tcPr.backgroundColor` for every cell in `range`.
    pub fn set_cell_shading(
        &self,
        ctx: &EditCtx,
        range: &TableRange,
        color: Option<&str>,
    ) -> OpResult<TableReceipt> {
        let patch = HashMap::from([(
            "backgroundColor".to_owned(),
            color
                .map(|value| Any::from(value.trim_start_matches('#')))
                .unwrap_or(Any::Null),
        )]);
        self.set_cell_text_format(ctx, range, &patch)
    }

    /// Generic cell-format passthrough. Keys are merged into `tcPr`; a JSON
    /// `null` removes a key. This covers vertical alignment, margins, text
    /// direction, no-wrap, and future independently-authored cell properties.
    pub fn set_cell_text_format(
        &self,
        ctx: &EditCtx,
        range: &TableRange,
        patch: &HashMap<String, Any>,
    ) -> OpResult<TableReceipt> {
        for reserved in ["rowspan", "colspan", "gridSpan", "vMerge"] {
            if patch.contains_key(reserved) {
                return Err(invalid(format!(
                    "{reserved} is managed by merge/split operations"
                )));
            }
        }
        let locator = range.anchor.table();
        let mut txn = self.transact_for(ctx);
        let (_, table, _) = table_at(&txn, &locator)?;
        let mut data = read_table(&table, &txn)?;
        let selected = selected_anchors(&data, range)?;
        for anchor in selected {
            let cell = &mut data.rows[anchor.row].cells[anchor.cell_index];
            for (key, value) in patch {
                if matches!(value, Any::Null | Any::Undefined) {
                    cell.tc_pr.remove(key);
                } else {
                    cell.tc_pr.insert(key.clone(), value.clone());
                }
            }
        }
        write_table(&mut txn, &table, &data);
        receipt(locator, Some(&data), Vec::new(), Vec::new(), Vec::new())
    }

    /// Merges the supplied border sides into every selected cell's
    /// `tcPr.borders`. `insideH`/`insideV` resolve per cell to the physical
    /// edges interior to the selection, the same mapping seeding uses to push
    /// `tblBorders` down. An omitted side is left untouched, `style: "none"`
    /// authors an explicit no-border, and JSON `null` drops the authored side.
    pub fn set_cell_borders(
        &self,
        ctx: &EditCtx,
        range: &TableRange,
        borders: &HashMap<String, Any>,
    ) -> OpResult<TableReceipt> {
        let locator = range.anchor.table();
        let mut txn = self.transact_for(ctx);
        let (_, table, _) = table_at(&txn, &locator)?;
        let mut data = read_table(&table, &txn)?;
        let rect = checked_rect(&data, range)?;
        let selected = selected_anchors(&data, range)?;
        let (all, _) = anchors(&data)?;

        let resolved: Vec<(&CellAnchor, Vec<(&'static str, Option<Any>)>)> = selected
            .iter()
            .map(|anchor| {
                let edges = border_side_sources(
                    anchor.row == rect.top,
                    anchor.row + anchor.rowspan == rect.bottom,
                    anchor.column == rect.left,
                    anchor.column + anchor.colspan == rect.right,
                )
                .into_iter()
                .filter_map(|(edge, source)| {
                    let value = borders.get(source)?;
                    let keep = !matches!(value, Any::Null | Any::Undefined);
                    Some((edge, keep.then(|| authored_border(value))))
                })
                .collect();
                (anchor, edges)
            })
            .collect();

        for (anchor, edges) in &resolved {
            merge_cell_borders(
                &mut data.rows[anchor.row].cells[anchor.cell_index],
                edges.iter().map(|(edge, value)| (*edge, value)),
            );
        }

        // Keep shared edges symmetric. A resolved edge is mirrored onto the
        // facing edge of every cell across it, including cells just outside
        // the selection.
        for (anchor, edges) in &resolved {
            for (edge, value) in edges {
                let (facing, slots): (&str, Vec<(usize, usize)>) = match *edge {
                    "top" if anchor.row > 0 => (
                        "bottom",
                        (anchor.column..anchor.column + anchor.colspan)
                            .map(|column| (anchor.row - 1, column))
                            .collect(),
                    ),
                    "bottom" => (
                        "top",
                        (anchor.column..anchor.column + anchor.colspan)
                            .map(|column| (anchor.row + anchor.rowspan, column))
                            .collect(),
                    ),
                    "left" if anchor.column > 0 => (
                        "right",
                        (anchor.row..anchor.row + anchor.rowspan)
                            .map(|row| (row, anchor.column - 1))
                            .collect(),
                    ),
                    "right" => (
                        "left",
                        (anchor.row..anchor.row + anchor.rowspan)
                            .map(|row| (row, anchor.column + anchor.colspan))
                            .collect(),
                    ),
                    _ => continue,
                };
                let mut neighbors = HashSet::new();
                for (row, column) in slots {
                    let Some(neighbor) = covering(&all, row, column) else {
                        continue;
                    };
                    if !neighbors.insert(neighbor.cell.story.clone()) {
                        continue;
                    }
                    merge_cell_borders(
                        &mut data.rows[neighbor.row].cells[neighbor.cell_index],
                        std::iter::once((facing, value)),
                    );
                }
            }
        }

        write_table(&mut txn, &table, &data);
        receipt(locator, Some(&data), Vec::new(), Vec::new(), Vec::new())
    }

    /// Sets one authored grid-column width in twips and refreshes each cell's
    /// dxa width from the sum of the grid columns it spans.
    pub fn set_column_width(
        &self,
        ctx: &EditCtx,
        at: &CellLoc,
        width_twips: f64,
    ) -> OpResult<TableReceipt> {
        if !width_twips.is_finite() || width_twips <= 0.0 {
            return Err(OpError::InvalidFormatValue(
                "column width must be a positive finite twip value".to_owned(),
            ));
        }
        let locator = at.table();
        let mut txn = self.transact_for(ctx);
        let (_, table, _) = table_at(&txn, &locator)?;
        let mut data = read_table(&table, &txn)?;
        let (mut cell_anchors, columns) = anchors(&data)?;
        let column = at.column as usize;
        if column >= columns {
            return Err(invalid(format!("column {column} is outside the table")));
        }
        ensure_grid(&mut data, columns);
        data.grid[column] = Any::Number(width_twips);
        for anchor in &mut cell_anchors {
            let width: f64 = data.grid[anchor.column..anchor.column + anchor.colspan]
                .iter()
                .filter_map(|value| any_number(Some(value)))
                .sum();
            anchor
                .cell
                .tc_pr
                .insert("width".to_owned(), Any::Number(width));
            anchor
                .cell
                .tc_pr
                .insert("widthType".to_owned(), Any::from("dxa"));
        }
        data.rows = reconstruct_rows(data.rows, cell_anchors)?;
        write_table(&mut txn, &table, &data);
        receipt(locator, Some(&data), Vec::new(), Vec::new(), Vec::new())
    }

    /// Sets the table-wide preferred width in twips (`tblPr.widthType = dxa`).
    pub fn set_table_width(
        &self,
        ctx: &EditCtx,
        locator: &TableLocator,
        width_twips: f64,
    ) -> OpResult<TableReceipt> {
        if !width_twips.is_finite() || width_twips <= 0.0 {
            return Err(OpError::InvalidFormatValue(
                "table width must be a positive finite twip value".to_owned(),
            ));
        }
        let mut txn = self.transact_for(ctx);
        let (_, table, _) = table_at(&txn, locator)?;
        let mut data = read_table(&table, &txn)?;
        data.tbl_pr
            .insert("width".to_owned(), Any::Number(width_twips));
        data.tbl_pr.insert("widthType".to_owned(), Any::from("dxa"));
        write_table(&mut txn, &table, &data);
        receipt(
            locator.clone(),
            Some(&data),
            Vec::new(),
            Vec::new(),
            Vec::new(),
        )
    }

    pub fn set_table_properties(
        &self,
        ctx: &EditCtx,
        locator: &TableLocator,
        patch: &HashMap<String, Any>,
    ) -> OpResult<TableReceipt> {
        for (key, value) in patch {
            let valid = match key.as_str() {
                "width" => {
                    matches!(value, Any::Null | Any::Undefined)
                        || any_number(Some(value))
                            .is_some_and(|width| width >= 0.0 && width.fract() == 0.0)
                }
                "widthType" => {
                    matches!(value, Any::Null | Any::Undefined)
                        || matches!(value, Any::String(kind) if ["auto", "dxa", "pct", "nil"].contains(&&**kind))
                }
                "justification" => {
                    matches!(value, Any::Null | Any::Undefined)
                        || matches!(value, Any::String(kind) if ["left", "center", "right"].contains(&&**kind))
                }
                _ => false,
            };
            if !valid {
                return Err(OpError::InvalidFormatValue(format!(
                    "invalid table property {key}"
                )));
            }
        }
        let mut txn = self.transact_for(ctx);
        let (_, table, _) = table_at(&txn, locator)?;
        let mut data = read_table(&table, &txn)?;
        for (key, value) in patch {
            if matches!(value, Any::Null | Any::Undefined) {
                data.tbl_pr.remove(key);
            } else {
                data.tbl_pr.insert(key.clone(), value.clone());
            }
        }
        if matches!(data.tbl_pr.get("widthType"), Some(Any::String(kind)) if &**kind == "auto") {
            data.tbl_pr.remove("width");
        }
        table.insert(&mut txn, "tblPr", to_any_map(data.tbl_pr.clone()));
        receipt(
            locator.clone(),
            Some(&data),
            Vec::new(),
            Vec::new(),
            Vec::new(),
        )
    }
}

#[cfg(test)]
mod tests {
    use serde_json::Value;
    use yrs::Transact;

    use super::*;
    use crate::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
    use crate::{CanonicalItem, RawOp, project_story};

    fn direct() -> EditCtx {
        EditCtx::local("Ada", "2026-07-14T10:00:00Z")
    }

    fn suggesting() -> EditCtx {
        direct().suggesting()
    }

    fn seed_table() -> EditingDoc {
        let doc = EditingDoc::new(71);
        for (story, text) in [
            ("body:t0:r0c0", "A"),
            ("body:t0:r0c1", "B"),
            ("body:t0:r1c0", "C"),
            ("body:t0:r1c1", "D"),
        ] {
            doc.create_story(story, text, "Normal", "left").unwrap();
        }
        doc.create_story("body", "", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Delete { index: 0, len: 1 },
                RawOp::InsertEmbed {
                    index: 0,
                    kind: "table".to_owned(),
                    payload: vec![
                        (
                            "tblPr".to_owned(),
                            Any::from_json(r#"{"width":4000,"widthType":"dxa"}"#).unwrap(),
                        ),
                        (
                            "grid".to_owned(),
                            Any::from_json("[2000,2000]").unwrap(),
                        ),
                        (
                            "rows".to_owned(),
                            Any::from_json(
                                r#"[
                                  {"trPr":{"height":360},"cells":[
                                    {"tcPr":{"colspan":1,"rowspan":1,"noWrap":false},"story":"body:t0:r0c0"},
                                    {"tcPr":{"colspan":1,"rowspan":1,"noWrap":false},"story":"body:t0:r0c1"}
                                  ]},
                                  {"trPr":{"height":360},"cells":[
                                    {"tcPr":{"colspan":1,"rowspan":1,"noWrap":false},"story":"body:t0:r1c0"},
                                    {"tcPr":{"colspan":1,"rowspan":1,"noWrap":false},"story":"body:t0:r1c1"}
                                  ]}
                                ]"#,
                            )
                            .unwrap(),
                        ),
                    ],
                    attrs: Attrs::new(),
                },
            ],
            &direct(),
        )
        .unwrap();
        doc
    }

    fn table_value(doc: &EditingDoc) -> (Value, Value, Value) {
        let item = project_story(doc, "body").unwrap().remove(0);
        let CanonicalItem::Table { tbl_pr, grid, rows } = item else {
            panic!("expected canonical table");
        };
        (tbl_pr, grid, rows)
    }

    fn cell(row: u32, column: u32) -> CellLoc {
        CellLoc::new("body", 0, row, column)
    }

    #[test]
    fn merge_moves_comment_anchors_with_cell_text() {
        let doc = seed_table();
        let moved = doc
            .add_comment(
                &[
                    crate::StoryRange::new("body:t0:r0c1", 0, 1),
                    crate::StoryRange::new("body:t0:r1c1", 0, 1),
                ],
                "Ada",
                "2026-10-21",
                Any::from("note"),
            )
            .unwrap();
        let retained = doc
            .add_comment(
                &[crate::StoryRange::new("body:t0:r0c0", 0, 1)],
                "Ada",
                "2026-10-21",
                Any::from("keep"),
            )
            .unwrap();
        let before = doc.resolve_comment(&retained).unwrap();
        doc.merge_cells(&direct(), &TableRange::new(cell(0, 0), cell(1, 1)))
            .unwrap();
        let expected = vec![
            crate::ResolvedCommentAnchor {
                story: "body:t0:r0c0".to_owned(),
                start: 2,
                end: 3,
            },
            crate::ResolvedCommentAnchor {
                story: "body:t0:r0c0".to_owned(),
                start: 6,
                end: 7,
            },
        ];
        assert_eq!(doc.resolve_comment(&moved).unwrap(), expected);
        assert_eq!(doc.resolve_comment(&retained).unwrap(), before);
        assert_eq!(
            doc.paragraphs("body:t0:r0c0")
                .unwrap()
                .iter()
                .map(|paragraph| paragraph.text.as_str())
                .collect::<Vec<_>>(),
            ["A", "B", "C", "D"]
        );
        let replica = EditingDoc::new(72);
        replica
            .apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(replica.resolve_comment(&moved).unwrap(), expected);
    }

    #[test]
    fn merge_uses_grid_before_and_preserves_unselected_cells() {
        let cell_xml =
            |text: &str| format!("<w:tc><w:tcPr/><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>");
        let body = format!(
            concat!(
                "<w:tbl><w:tblPr/><w:tblGrid>",
                "<w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/>",
                "<w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/>",
                "</w:tblGrid><w:tr>{}{}{}{}</w:tr>",
                "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/>",
                "<w:gridAfter w:val=\"1\"/></w:trPr>{}{}</w:tr>",
                "<w:tr>{}{}{}{}</w:tr></w:tbl>"
            ),
            cell_xml("Left"),
            cell_xml("A"),
            cell_xml("B"),
            cell_xml("Right"),
            cell_xml("C"),
            cell_xml("D"),
            cell_xml("Keep1"),
            cell_xml("Keep2"),
            cell_xml("Keep3"),
            cell_xml("Keep4"),
        );
        let bytes = crate::seed::fixture::Package::new(&body).bytes();
        let doc = EditingDoc::new(73);
        crate::seed_from_docx(&doc, &bytes).unwrap();
        let (_, _, before) = table_value(&doc);
        let receipt = doc
            .merge_cells(&direct(), &TableRange::new(cell(0, 1), cell(1, 2)))
            .unwrap();
        assert_eq!(receipt.columns, 4);
        let (_, _, after) = table_value(&doc);
        assert_eq!(after[0]["cells"][0], before[0]["cells"][0]);
        assert_eq!(after[0]["cells"][2], before[0]["cells"][3]);
        assert_eq!(after[2], before[2]);
        assert_eq!(after[1]["trPr"], before[1]["trPr"]);
        assert!(after[1]["cells"].as_array().unwrap().is_empty());
        assert_eq!(after[0]["cells"][1]["tcPr"]["rowspan"], 2);
        assert_eq!(after[0]["cells"][1]["tcPr"]["colspan"], 2);
        assert_eq!(
            doc.paragraphs("body:t0:r0c1")
                .unwrap()
                .iter()
                .map(|paragraph| paragraph.text.as_str())
                .collect::<Vec<_>>(),
            ["A", "B", "C", "D"]
        );
    }

    #[test]
    fn a_table_payload_is_the_one_its_story_segment_carries() {
        let doc = seed_table();
        doc.insert_table(&direct(), Position::new("body", 1), 1, 2)
            .unwrap();
        let tables: Vec<_> = doc
            .story_segments("body")
            .unwrap()
            .into_iter()
            .filter_map(|segment| match segment.content {
                crate::SegmentContent::OtherEmbed { kind, payload } if kind == "table" => {
                    Some(payload)
                }
                _ => None,
            })
            .collect();
        assert_eq!(tables.len(), 2);
        for (index, payload) in tables.iter().enumerate() {
            assert_eq!(
                &doc.table_payload(&TableLocator::new("body", index as u32))
                    .unwrap(),
                payload
            );
        }
        assert!(matches!(
            doc.table_payload(&TableLocator::new("body", 2)),
            Err(OpError::UnknownTable { .. })
        ));
        assert!(matches!(
            doc.table_payload(&TableLocator::new("missing", 0)),
            Err(OpError::UnknownStory(_))
        ));
    }

    #[test]
    fn insert_row_creates_stories_and_still_lowers_to_layout_table() {
        let doc = seed_table();
        let receipt = doc.insert_row(&direct(), &cell(0, 0), true).unwrap();
        assert_eq!((receipt.rows, receipt.columns), (3, 2));
        assert_eq!(receipt.created_story_ids.len(), 2);
        assert!(
            receipt
                .created_story_ids
                .contains(&"body:t0:r2c0~71".to_owned())
        );
        assert!(
            receipt
                .created_story_ids
                .contains(&"body:t0:r2c1~71".to_owned())
        );
        for story in &receipt.created_story_ids {
            assert_eq!(doc.paragraphs(story).unwrap().len(), 1);
        }

        let (_, grid, rows) = table_value(&doc);
        assert_eq!(grid.as_array().unwrap().len(), 2);
        assert_eq!(rows.as_array().unwrap().len(), 3);
        let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        assert_eq!(blocks.len(), 1);
        let value = serde_json::to_value(&blocks).unwrap();
        assert_eq!(value[0]["kind"], "table");
        assert_eq!(value[0]["rows"].as_array().unwrap().len(), 3);
    }

    fn inserted(ctx: &EditCtx, rows: u32, columns: u32) -> EditingDoc {
        let doc = EditingDoc::new(74);
        doc.create_story("body", "", "Normal", "left").unwrap();
        doc.insert_table(ctx, Position::new("body", 0), rows, columns)
            .unwrap();
        doc
    }

    #[test]
    fn insert_table_closes_its_anchor_paragraph_and_undo_restores_it() {
        for offset in [0, 2, 4] {
            let doc = EditingDoc::new(74);
            doc.create_story("body", "Seed", "Normal", "center")
                .unwrap();
            let before = project_story(&doc, "body").unwrap();
            let original_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
            let mut undo = doc.undo_manager();
            let receipt = doc
                .insert_table(&direct(), Position::new("body", offset), 2, 2)
                .unwrap();
            let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
            let value = serde_json::to_value(&blocks).unwrap();
            let table_index = usize::from(offset > 0);
            assert_eq!(value[table_index]["kind"], "table");
            let paragraphs = doc.paragraphs("body").unwrap();
            let expected = match offset {
                0 => vec!["Seed"],
                2 => vec!["Se", "ed"],
                _ => vec!["Seed", ""],
            };
            assert_eq!(
                paragraphs
                    .iter()
                    .map(|p| p.text.as_str())
                    .collect::<Vec<_>>(),
                expected
            );
            assert_eq!(paragraphs[0].para_id, original_id);
            assert_eq!(receipt.new_para_ids.len(), 4 + table_index);
            assert!(undo.undo());
            assert_eq!(project_story(&doc, "body").unwrap(), before);
            assert_eq!(doc.story_ids_for_test(), vec!["body"]);
            assert!(undo.redo());
            assert!(yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).is_ok());
        }
    }

    #[test]
    fn rejecting_a_suggested_table_also_rejoins_its_split_paragraph() {
        let doc = EditingDoc::new(74);
        doc.create_story("body", "Seed", "Normal", "left").unwrap();
        let receipt = doc
            .insert_table(&suggesting(), Position::new("body", 2), 1, 1)
            .unwrap();
        assert!(yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).is_ok());
        doc.reject_change(
            &direct(),
            &crate::ChangeTarget::Revision(receipt.revision_ids[0].clone()),
        )
        .unwrap();
        assert_eq!(doc.paragraphs("body").unwrap().len(), 1);
        assert_eq!(doc.paragraphs("body").unwrap()[0].text, "Seed");
        assert_eq!(doc.story_ids_for_test(), vec!["body"]);
    }

    #[test]
    fn revision_stamps_keep_distinct_dates_for_rows_sharing_a_revision() {
        for key in [TR_INS, TR_DEL] {
            let doc = seed_table();
            let stamp = |date: &str| {
                Any::from_json(&format!(
                    r#"{{"id":"shared","author":"Ada","date":"{date}"}}"#
                ))
                .unwrap()
            };
            {
                let mut txn = doc.yrs_doc().transact_mut();
                let (_, table, _) = table_at(&txn, &TableLocator::new("body", 0)).unwrap();
                let mut data = read_table(&table, &txn).unwrap();
                data.rows[0]
                    .tr_pr
                    .insert(key.into(), stamp("2026-07-14T10:00:00Z"));
                data.rows[1]
                    .tr_pr
                    .insert(key.into(), stamp("2026-07-14T10:01:00Z"));
                write_table(&mut txn, &table, &data);
            }
            assert_eq!(
                doc.revision_stamps(&["shared".into()]).unwrap(),
                BTreeMap::from([(
                    "shared".into(),
                    std::collections::BTreeSet::from([
                        ("Ada".into(), "2026-07-14T10:00:00Z".into()),
                        ("Ada".into(), "2026-07-14T10:01:00Z".into()),
                    ]),
                )])
            );
        }
    }

    fn assert_grid_borders(doc: &EditingDoc) {
        let (_, _, rows) = table_value(doc);
        for row in rows.as_array().unwrap() {
            for cell in row["cells"].as_array().unwrap() {
                let borders = &cell["tcPr"]["borders"];
                for side in ["top", "bottom", "left", "right"] {
                    assert_eq!(borders[side]["style"], "single", "{side} of {cell}");
                    assert_eq!(borders[side]["size"], 4.0, "{side} of {cell}");
                    assert_eq!(borders[side]["color"]["rgb"], "000000", "{side} of {cell}");
                }
            }
        }
    }

    #[test]
    fn inserted_table_carries_word_default_grid_borders() {
        assert_grid_borders(&inserted(&direct(), 3, 2));
        assert_grid_borders(&inserted(&suggesting(), 3, 2));
    }

    #[test]
    fn inserted_table_keeps_grid_borders_when_rows_and_columns_grow() {
        let doc = inserted(&direct(), 2, 2);
        doc.insert_row(&direct(), &cell(1, 0), true).unwrap();
        doc.insert_column(&direct(), &cell(0, 1), true).unwrap();
        let (_, grid, rows) = table_value(&doc);
        assert_eq!(grid.as_array().unwrap().len(), 3);
        assert_eq!(rows.as_array().unwrap().len(), 3);
        assert_grid_borders(&doc);
    }

    #[test]
    fn inserted_table_paints_a_rule_on_every_cell_edge() {
        use docx_layout::display_list::{LineRole, Primitive};
        use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
        use docx_layout::types::{Input, LayoutOptions, MeasuredBlock, PageMargins, Size};

        let doc = inserted(&direct(), 2, 2);
        let mut blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        let measures = measure_blocks(&mut blocks, 600.0, &MeasurementConfig::default()).unwrap();
        let mut input = Input {
            measured: blocks
                .into_iter()
                .zip(measures)
                .map(|(block, measure)| MeasuredBlock { block, measure })
                .collect(),
            options: LayoutOptions {
                page_size: Some(Size {
                    w: 816.0,
                    h: 1056.0,
                }),
                margins: Some(PageMargins {
                    top: 96.0,
                    right: 108.0,
                    bottom: 96.0,
                    left: 108.0,
                    header: None,
                    footer: None,
                }),
                ..LayoutOptions::default()
            },
        };
        let layout = docx_layout::compute_layout_input(&mut input).unwrap();
        let display_list = docx_layout::build_display_list(&input, &layout).unwrap();
        let rules: Vec<_> = display_list.pages[0]
            .primitives
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Line(line) if line.role == Some(LineRole::TableBorder) => Some(line),
                _ => None,
            })
            .collect();
        // A 2x2 grid draws three horizontal and three vertical rules, each
        // split into one segment per cell it borders.
        let horizontal = rules.iter().filter(|line| line.y1 == line.y2).count();
        let vertical = rules.iter().filter(|line| line.x1 == line.x2).count();
        assert_eq!((horizontal, vertical), (6, 6), "every grid edge is painted");
        for line in &rules {
            assert_eq!(line.color, "#000000");
            assert_eq!(line.stroke_width.as_f64(), Some(0.667));
        }
    }

    #[test]
    fn suggested_table_insert_shares_trins_and_resolves_as_one_revision() {
        let accepted = EditingDoc::new(72);
        accepted.create_story("body", "", "Normal", "left").unwrap();
        let inserted = accepted
            .insert_table(&suggesting(), Position::new("body", 0), 3, 2)
            .unwrap();
        assert_eq!((inserted.rows, inserted.columns), (3, 2));
        assert_eq!(inserted.created_story_ids.len(), 6);
        assert_eq!(inserted.revision_ids.len(), 1);
        let revision_id = inserted.revision_ids[0].clone();
        let (_, _, rows) = table_value(&accepted);
        for row in rows.as_array().unwrap() {
            assert_eq!(row["trPr"][TR_INS]["id"], revision_id);
        }
        let listed = accepted.list_revisions().unwrap();
        assert_eq!(
            listed
                .iter()
                .filter(|revision| revision.change.revision_id == revision_id)
                .count(),
            1
        );
        accepted
            .accept_change(
                &direct(),
                &crate::ChangeTarget::Revision(revision_id.clone()),
            )
            .unwrap();
        let (_, _, rows) = table_value(&accepted);
        assert_eq!(rows.as_array().unwrap().len(), 3);
        assert!(
            rows.as_array()
                .unwrap()
                .iter()
                .all(|row| row["trPr"].get(TR_INS).is_none())
        );

        let rejected = EditingDoc::new(73);
        rejected.create_story("body", "", "Normal", "left").unwrap();
        let inserted = rejected
            .insert_table(&suggesting(), Position::new("body", 0), 2, 2)
            .unwrap();
        rejected
            .reject_change(
                &direct(),
                &crate::ChangeTarget::Revision(inserted.revision_ids[0].clone()),
            )
            .unwrap();
        assert!(
            project_story(&rejected, "body")
                .unwrap()
                .iter()
                .all(|item| !matches!(item, CanonicalItem::Table { .. }))
        );
        for story in inserted.created_story_ids {
            assert!(matches!(
                rejected.paragraphs(&story),
                Err(crate::EditError::StoryNotFound(_))
            ));
        }
    }

    #[test]
    fn suggested_row_insertion_stamps_trins_and_accept_reject_resolve_it() {
        let accepted = seed_table();
        let inserted = accepted
            .insert_row(&suggesting(), &cell(0, 0), true)
            .unwrap();
        let revision_id = inserted.revision_ids[0].clone();
        let (_, _, rows) = table_value(&accepted);
        assert_eq!(rows.as_array().unwrap().len(), 3);
        assert_eq!(rows[1]["trPr"][TR_INS]["author"], "Ada");
        assert_eq!(rows[1]["trPr"][TR_INS]["id"], revision_id);
        accepted
            .accept_change(
                &direct(),
                &crate::ChangeTarget::Revision(revision_id.clone()),
            )
            .unwrap();
        let (_, _, rows) = table_value(&accepted);
        assert_eq!(rows.as_array().unwrap().len(), 3);
        assert!(rows[1]["trPr"].get(TR_INS).is_none());

        let rejected = seed_table();
        let inserted = rejected
            .insert_row(&suggesting(), &cell(0, 0), true)
            .unwrap();
        rejected
            .reject_change(
                &direct(),
                &crate::ChangeTarget::Revision(inserted.revision_ids[0].clone()),
            )
            .unwrap();
        let (_, _, rows) = table_value(&rejected);
        assert_eq!(rows.as_array().unwrap().len(), 2);
        for story in inserted.created_story_ids {
            assert!(matches!(
                rejected.paragraphs(&story),
                Err(crate::EditError::StoryNotFound(_))
            ));
        }
    }

    #[test]
    fn suggested_row_deletion_stamps_trdel_and_accept_reject_resolve_it() {
        let rejected = seed_table();
        let deletion = rejected
            .delete_row(&suggesting(), &TableRange::cell(cell(1, 0)))
            .unwrap();
        let revision_id = deletion.revision_ids[0].clone();
        assert_eq!(deletion.rows, 2, "suggested delete retains the row");
        let (_, _, rows) = table_value(&rejected);
        assert_eq!(rows[1]["trPr"][TR_DEL]["author"], "Ada");
        rejected
            .reject_change(&direct(), &crate::ChangeTarget::Revision(revision_id))
            .unwrap();
        let (_, _, rows) = table_value(&rejected);
        assert_eq!(rows.as_array().unwrap().len(), 2);
        assert!(rows[1]["trPr"].get(TR_DEL).is_none());

        let accepted = seed_table();
        let deletion = accepted
            .delete_row(&suggesting(), &TableRange::cell(cell(1, 0)))
            .unwrap();
        accepted
            .accept_change(
                &direct(),
                &crate::ChangeTarget::Revision(deletion.revision_ids[0].clone()),
            )
            .unwrap();
        let (_, _, rows) = table_value(&accepted);
        assert_eq!(rows.as_array().unwrap().len(), 1);
        assert!(matches!(
            accepted.paragraphs("body:t0:r1c0"),
            Err(crate::EditError::StoryNotFound(_))
        ));
    }

    #[test]
    fn delete_column_removes_cell_stories_and_shrinks_grid() {
        let doc = seed_table();
        let receipt = doc
            .delete_column(&direct(), &TableRange::cell(cell(0, 1)))
            .unwrap();
        assert_eq!((receipt.rows, receipt.columns), (2, 1));
        assert_eq!(receipt.deleted_story_ids.len(), 2);
        assert!(matches!(
            doc.paragraphs("body:t0:r0c1"),
            Err(crate::EditError::StoryNotFound(_))
        ));
        let (_, grid, rows) = table_value(&doc);
        assert_eq!(grid.as_array().unwrap(), &[Value::from(2000)]);
        assert_eq!(rows[0]["cells"].as_array().unwrap().len(), 1);
        let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        let value = serde_json::to_value(blocks).unwrap();
        assert_eq!(value[0]["rows"][0]["cells"].as_array().unwrap().len(), 1);
    }

    #[test]
    fn insert_column_and_delete_row_keep_grid_and_stories_consistent() {
        let doc = seed_table();
        let inserted = doc.insert_column(&direct(), &cell(0, 0), false).unwrap();
        assert_eq!((inserted.rows, inserted.columns), (2, 3));
        assert_eq!(inserted.created_story_ids.len(), 2);
        assert!(
            inserted
                .created_story_ids
                .contains(&"body:t0:r0c2~71".to_owned())
        );
        assert!(
            inserted
                .created_story_ids
                .contains(&"body:t0:r1c2~71".to_owned())
        );
        let (_, grid, rows) = table_value(&doc);
        assert_eq!(grid.as_array().unwrap().len(), 3);
        assert_eq!(rows[0]["cells"].as_array().unwrap().len(), 3);

        let deleted = doc
            .delete_row(&direct(), &TableRange::cell(cell(1, 0)))
            .unwrap();
        assert_eq!((deleted.rows, deleted.columns), (1, 3));
        assert_eq!(deleted.deleted_story_ids.len(), 3);
        let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        let value = serde_json::to_value(blocks).unwrap();
        assert_eq!(value[0]["rows"].as_array().unwrap().len(), 1);
        assert_eq!(value[0]["rows"][0]["cells"].as_array().unwrap().len(), 3);
    }

    #[test]
    fn merge_then_split_updates_spans_and_cell_story_lifecycle() {
        let doc = seed_table();
        let range = TableRange::new(cell(0, 0), cell(1, 1));
        let merged = doc.merge_cells(&direct(), &range).unwrap();
        assert_eq!(merged.deleted_story_ids.len(), 3);
        assert_eq!(merged.changed_story_ids, ["body", "body:t0:r0c0"]);
        let (_, _, rows) = table_value(&doc);
        assert_eq!(rows[0]["cells"].as_array().unwrap().len(), 1);
        assert_eq!(rows[1]["cells"].as_array().unwrap().len(), 0);
        assert_eq!(rows[0]["cells"][0]["tcPr"]["colspan"], 2);
        assert_eq!(rows[0]["cells"][0]["tcPr"]["rowspan"], 2);
        assert_eq!(doc.paragraphs("body:t0:r0c0").unwrap().len(), 4);

        let split = doc.split_cell(&direct(), &cell(0, 0)).unwrap();
        assert_eq!(split.created_story_ids.len(), 3);
        let (_, _, rows) = table_value(&doc);
        assert_eq!(rows[0]["cells"].as_array().unwrap().len(), 2);
        assert_eq!(rows[1]["cells"].as_array().unwrap().len(), 2);
        for row in rows.as_array().unwrap() {
            for cell in row["cells"].as_array().unwrap() {
                assert_eq!(cell["tcPr"]["colspan"], 1);
                assert_eq!(cell["tcPr"]["rowspan"], 1);
            }
        }
        yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    }

    #[test]
    fn set_table_properties_patches_width_type_and_justification() {
        let doc = seed_table();
        let at = TableLocator::new("body", 0);
        doc.set_table_properties(
            &direct(),
            &at,
            &HashMap::from([
                ("width".to_owned(), Any::Number(2500.0)),
                ("widthType".to_owned(), Any::from("pct")),
                ("justification".to_owned(), Any::from("center")),
            ]),
        )
        .unwrap();
        let (tbl_pr, _, _) = table_value(&doc);
        assert_eq!(tbl_pr["width"], 2500);
        assert_eq!(tbl_pr["widthType"], "pct");
        assert_eq!(tbl_pr["justification"], "center");
        doc.set_table_properties(
            &direct(),
            &at,
            &HashMap::from([
                ("width".to_owned(), Any::Null),
                ("widthType".to_owned(), Any::from("auto")),
                ("justification".to_owned(), Any::Null),
            ]),
        )
        .unwrap();
        let (tbl_pr, _, _) = table_value(&doc);
        assert_eq!(tbl_pr["width"], Value::Null);
        assert_eq!(tbl_pr["widthType"], "auto");
        assert_eq!(tbl_pr["justification"], Value::Null);
        assert!(
            doc.set_table_properties(
                &direct(),
                &at,
                &HashMap::from([("widthType".to_owned(), Any::from("wide")),])
            )
            .is_err()
        );
        assert!(
            doc.set_table_properties(
                &direct(),
                &at,
                &HashMap::from([("width".to_owned(), Any::Number(-1.0)),])
            )
            .is_err()
        );
        assert!(
            doc.set_table_properties(
                &direct(),
                &at,
                &HashMap::from([("width".to_owned(), Any::Number(2500.5)),])
            )
            .is_err()
        );
        let (tbl_pr, _, _) = table_value(&doc);
        assert_eq!(tbl_pr["width"], Value::Null);
        doc.set_table_properties(
            &direct(),
            &at,
            &HashMap::from([
                ("width".to_owned(), Any::Number(3000.0)),
                ("widthType".to_owned(), Any::from("dxa")),
            ]),
        )
        .unwrap();
        doc.set_table_properties(
            &direct(),
            &at,
            &HashMap::from([("widthType".to_owned(), Any::from("auto"))]),
        )
        .unwrap();
        let (tbl_pr, _, _) = table_value(&doc);
        assert_eq!(tbl_pr["width"], Value::Null);
        assert_eq!(tbl_pr["widthType"], "auto");
        assert!(
            doc.set_table_properties(
                &direct(),
                &at,
                &HashMap::from([("tableLayout".to_owned(), Any::from("fixed")),])
            )
            .is_err()
        );
    }

    #[test]
    fn set_table_properties_does_not_rewrite_rows_or_grid() {
        let baseline = seed_table();
        let state = baseline.encode_state_as_update_v1();
        let rows_peer = EditingDoc::new(72);
        rows_peer.apply_update_v1(&state).unwrap();
        let props_peer = EditingDoc::new(99);
        props_peer.apply_update_v1(&state).unwrap();
        rows_peer
            .set_cell_borders(
                &direct(),
                &TableRange::cell(cell(0, 0)),
                &HashMap::from([("top".to_owned(), line("FF0000"))]),
            )
            .unwrap();
        rows_peer
            .set_column_width(&direct(), &cell(0, 0), 2400.0)
            .unwrap();
        props_peer
            .set_table_properties(
                &direct(),
                &TableLocator::new("body", 0),
                &HashMap::from([("justification".to_owned(), Any::from("center"))]),
            )
            .unwrap();
        let from_rows = rows_peer.encode_state_as_update_v1();
        let from_props = props_peer.encode_state_as_update_v1();
        rows_peer.apply_update_v1(&from_props).unwrap();
        props_peer.apply_update_v1(&from_rows).unwrap();
        for doc in [&rows_peer, &props_peer] {
            let (tbl_pr, grid, rows) = table_value(doc);
            assert_eq!(tbl_pr["justification"], "center");
            assert_eq!(grid[0], 2400);
            assert_eq!(
                rows[0]["cells"][0]["tcPr"]["borders"]["top"]["color"]["rgb"],
                "FF0000"
            );
        }
    }

    #[test]
    fn split_cell_grid_subdivides_an_ordinary_cell_and_adjusts_neighbors() {
        let doc = seed_table();
        let split = doc
            .split_cell_grid(&direct(), &cell(0, 0), Some(1), Some(2))
            .unwrap();
        assert_eq!((split.rows, split.columns), (2, 3));
        assert_eq!(split.created_story_ids.len(), 1);

        let (_, grid, rows) = table_value(&doc);
        assert_eq!(
            grid.as_array().unwrap(),
            &[Value::from(1000), Value::from(1000), Value::from(2000)]
        );
        assert_eq!(rows[0]["cells"].as_array().unwrap().len(), 3);
        assert_eq!(rows[1]["cells"].as_array().unwrap().len(), 2);
        assert_eq!(rows[1]["cells"][0]["tcPr"]["colspan"], 2);
        yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    }

    fn row_source_attributes(doc: &EditingDoc) -> Vec<Option<Value>> {
        let payload = doc.table_payload(&TableLocator::new("body", 0)).unwrap();
        let rows = serde_json::to_value(&payload["rows"]).unwrap();
        rows.as_array()
            .unwrap()
            .iter()
            .map(|row| row["trPr"].get("_extraAttributes").cloned())
            .collect()
    }

    #[test]
    fn synthesized_rows_do_not_repeat_source_row_attributes() {
        let body = concat!(
            r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>"#,
            r#"<w:tr w14:paraId="11111111" w14:textId="77777777" w:rsidR="00AA00BB">"#,
            "<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        );
        let bytes = crate::seed::fixture::Package::new(body).bytes();
        let doc = EditingDoc::new(74);
        crate::seed_from_docx(&doc, &bytes).unwrap();
        let source = row_source_attributes(&doc).remove(0).unwrap();
        assert!(source.to_string().contains("w14:paraId"));

        doc.split_cell_grid(&direct(), &cell(0, 0), Some(3), Some(1))
            .unwrap();
        doc.insert_row(&direct(), &cell(0, 0), true).unwrap();

        let rows = row_source_attributes(&doc);
        assert_eq!(rows.len(), 4);
        assert_eq!(rows[0].as_ref(), Some(&source));
        assert!(rows[1..].iter().all(Option::is_none), "{rows:?}");
        let para_ids = rows
            .iter()
            .flatten()
            .flat_map(|attributes| attributes.as_array().unwrap())
            .filter(|attribute| attribute["name"] == "w14:paraId")
            .count();
        assert_eq!(para_ids, 1);
    }

    #[test]
    fn cell_formatting_borders_and_width_are_authored_in_tcpr_and_grid() {
        let doc = seed_table();
        let range = TableRange::cell(cell(0, 0));
        doc.set_cell_shading(&direct(), &range, Some("#AABBCC"))
            .unwrap();
        doc.set_cell_text_format(
            &direct(),
            &range,
            &HashMap::from([
                ("verticalAlign".to_owned(), Any::from("center")),
                ("textDirection".to_owned(), Any::from("tbRl")),
            ]),
        )
        .unwrap();
        doc.set_cell_borders(
            &direct(),
            &range,
            &HashMap::from([
                (
                    "top".to_owned(),
                    Any::from_json(r#"{"style":"single","size":4,"color":{"rgb":"000000"}}"#)
                        .unwrap(),
                ),
                (
                    "right".to_owned(),
                    Any::from_json(r#"{"style":"single","size":4,"color":{"rgb":"000000"}}"#)
                        .unwrap(),
                ),
            ]),
        )
        .unwrap();
        doc.set_column_width(&direct(), &cell(0, 0), 2400.0)
            .unwrap();
        doc.set_table_width(&direct(), &TableLocator::new("body", 0), 5000.0)
            .unwrap();
        let (tbl_pr, grid, rows) = table_value(&doc);
        let tc_pr = &rows[0]["cells"][0]["tcPr"];
        assert_eq!(tc_pr["backgroundColor"], "AABBCC");
        assert_eq!(tc_pr["verticalAlign"], "center");
        assert_eq!(tc_pr["textDirection"], "tbRl");
        assert_eq!(tc_pr["borders"]["top"]["style"], "single");
        assert_eq!(
            rows[0]["cells"][1]["tcPr"]["borders"]["left"]["style"],
            "single"
        );
        assert_eq!(tc_pr["width"], 2400);
        assert_eq!(grid[0], 2400);
        assert_eq!(tbl_pr["width"], 5000);
        assert_eq!(tbl_pr["widthType"], "dxa");
    }

    fn line(color: &str) -> Any {
        Any::from_json(&format!(
            r#"{{"style":"single","size":4,"color":{{"rgb":"{color}"}}}}"#
        ))
        .unwrap()
    }

    fn cleared() -> Any {
        Any::from_json(r#"{"style":"none","size":0,"color":{"rgb":"000000"}}"#).unwrap()
    }

    fn sides(names: &[&str], value: &Any) -> HashMap<String, Any> {
        names
            .iter()
            .map(|name| ((*name).to_owned(), value.clone()))
            .collect()
    }

    /// The six-, four- and two-key payloads the border toolbar actually sends.
    const ALL: &[&str] = &["top", "bottom", "left", "right", "insideH", "insideV"];
    const OUTSIDE: &[&str] = &["top", "bottom", "left", "right"];
    const INSIDE: &[&str] = &["insideH", "insideV"];

    fn whole_table() -> TableRange {
        TableRange::new(cell(0, 0), cell(1, 1))
    }

    fn grid_table() -> EditingDoc {
        let doc = seed_table();
        doc.set_cell_borders(&direct(), &whole_table(), &sides(ALL, &line("000000")))
            .unwrap();
        doc
    }

    /// Every cell's four physical edge styles, `None` where unauthored.
    fn edges(doc: &EditingDoc, row: usize, column: usize) -> [Option<String>; 4] {
        let (_, _, rows) = table_value(doc);
        let borders = &rows[row]["cells"][column]["tcPr"]["borders"];
        ["top", "bottom", "left", "right"].map(|side| {
            borders[side]["style"]
                .as_str()
                .filter(|style| *style != "none")
                .map(str::to_owned)
        })
    }

    fn authored(styles: [&str; 4]) -> [Option<String>; 4] {
        styles.map(|style| Some(style.to_owned()))
    }

    /// Painted table-border rules as `(x1, y1, x2, y2, color)` page pixels.
    fn rules(doc: &EditingDoc) -> Vec<(f64, f64, f64, f64, String)> {
        use docx_layout::display_list::{LineRole, Primitive};
        use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
        use docx_layout::types::{Input, LayoutOptions, MeasuredBlock, PageMargins, Size};

        let mut blocks = yrs_doc_to_layout_blocks(doc, "body", &RenderEnv::default()).unwrap();
        let measures = measure_blocks(&mut blocks, 600.0, &MeasurementConfig::default()).unwrap();
        let mut input = Input {
            measured: blocks
                .into_iter()
                .zip(measures)
                .map(|(block, measure)| MeasuredBlock { block, measure })
                .collect(),
            options: LayoutOptions {
                page_size: Some(Size {
                    w: 816.0,
                    h: 1056.0,
                }),
                margins: Some(PageMargins {
                    top: 96.0,
                    right: 108.0,
                    bottom: 96.0,
                    left: 108.0,
                    header: None,
                    footer: None,
                }),
                ..LayoutOptions::default()
            },
        };
        let layout = docx_layout::compute_layout_input(&mut input).unwrap();
        let display_list = docx_layout::build_display_list(&input, &layout).unwrap();
        let number = |value: &serde_json::Number| value.as_f64().unwrap();
        display_list.pages[0]
            .primitives
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Line(l) if l.role == Some(LineRole::TableBorder) => Some((
                    number(&l.x1),
                    number(&l.y1),
                    number(&l.x2),
                    number(&l.y2),
                    l.color.clone(),
                )),
                _ => None,
            })
            .collect()
    }

    /// `(horizontal, vertical)` painted rule segments.
    fn rule_counts(doc: &EditingDoc) -> (usize, usize) {
        let painted = rules(doc);
        (
            painted.iter().filter(|rule| rule.1 == rule.3).count(),
            painted.iter().filter(|rule| rule.0 == rule.2).count(),
        )
    }

    #[test]
    fn border_all_authors_four_physical_edges_and_no_inside_keys() {
        let doc = grid_table();
        let (_, _, rows) = table_value(&doc);
        for row in 0..2 {
            for column in 0..2 {
                assert_eq!(edges(&doc, row, column), authored(["single"; 4]));
                let borders = &rows[row]["cells"][column]["tcPr"]["borders"];
                assert!(borders["insideH"].is_null(), "insideH reached a cell");
                assert!(borders["insideV"].is_null(), "insideV reached a cell");
            }
        }
        assert_eq!(rule_counts(&doc), (6, 6));
    }

    #[test]
    fn a_single_side_button_keeps_the_cell_s_other_three_sides() {
        let doc = grid_table();
        assert_eq!(rule_counts(&doc), (6, 6));
        doc.set_cell_borders(
            &direct(),
            &TableRange::cell(cell(0, 0)),
            &sides(&["top"], &line("FF0000")),
        )
        .unwrap();
        assert_eq!(edges(&doc, 0, 0), authored(["single"; 4]));
        assert_eq!(rule_counts(&doc), (6, 6));
    }

    #[test]
    fn single_side_buttons_apply_to_the_selection_s_own_edge() {
        for (side, painted, authored_at) in [
            ("top", (2, 0), [(0, 0), (0, 1)]),
            ("bottom", (2, 0), [(1, 0), (1, 1)]),
            ("left", (0, 2), [(0, 0), (1, 0)]),
            ("right", (0, 2), [(0, 1), (1, 1)]),
        ] {
            let doc = seed_table();
            doc.set_cell_borders(&direct(), &whole_table(), &sides(&[side], &line("000000")))
                .unwrap();
            let index = ["top", "bottom", "left", "right"]
                .iter()
                .position(|name| *name == side)
                .unwrap();
            for row in 0..2 {
                for column in 0..2 {
                    let expected = authored_at.contains(&(row, column));
                    assert_eq!(
                        edges(&doc, row, column)[index].is_some(),
                        expected,
                        "{side} on cell ({row},{column})"
                    );
                }
            }
            assert_eq!(rule_counts(&doc), painted, "{side}");
        }
    }

    #[test]
    fn border_outside_leaves_the_selection_s_interior_alone() {
        let doc = seed_table();
        doc.set_cell_borders(&direct(), &whole_table(), &sides(OUTSIDE, &line("000000")))
            .unwrap();
        assert_eq!(
            edges(&doc, 0, 0),
            [
                Some("single".to_owned()),
                None,
                Some("single".to_owned()),
                None
            ]
        );
        assert_eq!(
            edges(&doc, 1, 1),
            [
                None,
                Some("single".to_owned()),
                None,
                Some("single".to_owned())
            ]
        );
        assert_eq!(rule_counts(&doc), (4, 4));

        let grid = grid_table();
        grid.set_cell_borders(&direct(), &whole_table(), &sides(OUTSIDE, &line("FF0000")))
            .unwrap();
        assert_eq!(rule_counts(&grid), (6, 6));
        let painted = rules(&grid);
        let interior = painted.iter().filter(|rule| rule.4 == "#000000").count();
        assert_eq!(interior, 4, "the four interior segments keep their colour");
    }

    #[test]
    fn border_inside_paints_the_interior_rules_and_no_exterior_ones() {
        let doc = seed_table();
        doc.set_cell_borders(&direct(), &whole_table(), &sides(INSIDE, &line("000000")))
            .unwrap();
        let interior = Some("single".to_owned());
        assert_eq!(
            edges(&doc, 0, 0),
            [None, interior.clone(), None, interior.clone()]
        );
        // the same two edges seen from the far side, kept symmetric
        assert_eq!(edges(&doc, 1, 1), [interior.clone(), None, interior, None]);

        let painted = rules(&doc);
        let horizontal: Vec<_> = painted.iter().filter(|rule| rule.1 == rule.3).collect();
        let vertical: Vec<_> = painted.iter().filter(|rule| rule.0 == rule.2).collect();
        assert_eq!((horizontal.len(), vertical.len()), (2, 2));
        let y = horizontal[0].1;
        let x = vertical[0].0;
        assert!(horizontal.iter().all(|rule| rule.1 == y));
        assert!(vertical.iter().all(|rule| rule.0 == x));
        let left = horizontal
            .iter()
            .map(|rule| rule.0)
            .fold(f64::MAX, f64::min);
        let right = horizontal
            .iter()
            .map(|rule| rule.2)
            .fold(f64::MIN, f64::max);
        let top = vertical.iter().map(|rule| rule.1).fold(f64::MAX, f64::min);
        let bottom = vertical.iter().map(|rule| rule.3).fold(f64::MIN, f64::max);
        assert!(left < x && x < right, "the column rule is interior");
        assert!(top < y && y < bottom, "the row rule is interior");
    }

    #[test]
    fn border_none_clears_the_selection_and_the_edges_its_neighbours_own() {
        let doc = grid_table();
        doc.set_cell_borders(&direct(), &whole_table(), &sides(ALL, &cleared()))
            .unwrap();
        assert_eq!(edges(&doc, 0, 0), [const { None }; 4]);
        assert_eq!(rule_counts(&doc), (0, 0));
        let (_, _, rows) = table_value(&doc);
        for row in rows.as_array().unwrap() {
            for cell in row["cells"].as_array().unwrap() {
                for side in ["top", "bottom", "left", "right"] {
                    assert_eq!(
                        cell["tcPr"]["borders"][side],
                        serde_json::json!({"style": "none"})
                    );
                }
            }
        }

        let corner = grid_table();
        corner
            .set_cell_borders(
                &direct(),
                &TableRange::cell(cell(1, 1)),
                &sides(ALL, &cleared()),
            )
            .unwrap();
        assert_eq!(rule_counts(&corner), (4, 4));
    }

    #[test]
    fn one_side_clears_by_style_none_or_by_null_without_touching_the_rest() {
        for value in [cleared(), Any::Null] {
            let doc = grid_table();
            doc.set_cell_borders(
                &direct(),
                &TableRange::cell(cell(0, 0)),
                &sides(&["top"], &value),
            )
            .unwrap();
            assert_eq!(
                edges(&doc, 0, 0),
                [
                    None,
                    Some("single".to_owned()),
                    Some("single".to_owned()),
                    Some("single".to_owned())
                ]
            );
            assert_eq!(rule_counts(&doc), (5, 6));
        }
    }

    fn fragment_doc() -> EditingDoc {
        let doc = EditingDoc::new(77);
        doc.create_story("body", "", "Normal", "left").unwrap();
        doc
    }

    fn fragment_cell(text: &str) -> Value {
        serde_json::json!({
            "tcPr": {},
            "content": [
                {"kind": "text", "text": text, "attributes": {}},
                {"kind": "pilcrow", "properties": {}, "attributes": {}},
            ],
        })
    }

    fn fragment_table(cells: Vec<Value>) -> Value {
        serde_json::json!({
            "tblPr": {},
            "grid": [],
            "rows": [{"trPr": {}, "cells": cells}],
        })
    }

    fn fragment_texts(doc: &EditingDoc, story: &str) -> Vec<String> {
        doc.story_segments(story)
            .unwrap()
            .into_iter()
            .filter_map(|segment| match segment.content {
                crate::SegmentContent::Text(text) => Some(text),
                _ => None,
            })
            .collect()
    }

    fn assert_fragment_refused(doc: &EditingDoc, fragment: &Value) {
        let before = project_story(doc, "body").unwrap();
        let stories = doc.story_ids_for_test();
        let result = doc.insert_table_fragment(&direct(), Position::new("body", 0), fragment);
        assert!(
            matches!(result, Err(OpError::InvalidTable(_))),
            "expected InvalidTable, got {result:?}"
        );
        assert_eq!(project_story(doc, "body").unwrap(), before);
        assert_eq!(doc.story_ids_for_test(), stories);
    }

    #[test]
    fn fragment_round_trip_preserves_content_and_ids() {
        let doc = fragment_doc();
        let fragment = serde_json::json!({
            "tblPr": {"width": 4000, "widthType": "dxa"},
            "grid": [2000, 2000],
            "rows": [
                {"trPr": {"height": 360}, "cells": [
                    {"tcPr": {"rowspan": 1, "colspan": 1}, "content": [
                        {"kind": "text", "text": "A", "attributes": {"bold": true}},
                        {"kind": "pilcrow", "properties": {"alignment": "center"}, "attributes": {}},
                    ]},
                    {"tcPr": {"rowspan": 1, "colspan": 1}, "content": [
                        {"kind": "text", "text": "B", "attributes": {}},
                        {"kind": "pilcrow", "properties": {}, "attributes": {}},
                    ]},
                ]},
                {"trPr": {"height": 360}, "cells": [
                    {"tcPr": {"rowspan": 1, "colspan": 2}, "content": [
                        {"kind": "text", "text": "C", "attributes": {}},
                        {"kind": "pilcrow", "properties": {}, "attributes": {}},
                    ]},
                ]},
            ],
        });
        let receipt = doc
            .insert_table_fragment(&direct(), Position::new("body", 0), &fragment)
            .unwrap();
        assert_eq!((receipt.rows, receipt.columns), (2, 2));
        assert_eq!(receipt.created_story_ids.len(), 3);
        assert!(!receipt.deleted_table);
        assert!(receipt.deleted_story_ids.is_empty());
        let mut para_ids = receipt.new_para_ids.clone();
        para_ids.sort();
        para_ids.dedup();
        assert_eq!(para_ids.len(), receipt.new_para_ids.len());
        assert_eq!(
            fragment_texts(&doc, &receipt.created_story_ids[0]),
            vec!["A"]
        );
        let segments = doc.story_segments(&receipt.created_story_ids[0]).unwrap();
        assert_eq!(segments[0].attributes.get("bold"), Some(&Any::Bool(true)));
        let payload = doc.table_payload(&receipt.table).unwrap();
        assert_eq!(
            payload.get("tblPr"),
            Some(&Any::from_json(r#"{"width":4000,"widthType":"dxa"}"#).unwrap())
        );
        assert_eq!(
            payload.get("grid"),
            Some(&Any::from_json("[2000,2000]").unwrap())
        );
        assert_eq!(receipt.after.story, "body");
        assert_eq!(receipt.after.offset, 1);
        let paragraphs = doc.paragraphs("body").unwrap();
        assert_eq!(receipt.after.para_id, paragraphs[0].para_id);
    }

    #[test]
    fn fragment_insertion_mid_paragraph_splits() {
        let doc = EditingDoc::new(77);
        doc.create_story("body", "Seed", "Normal", "left").unwrap();
        let receipt = doc
            .insert_table_fragment(
                &direct(),
                Position::new("body", 2),
                &fragment_table(vec![fragment_cell("x")]),
            )
            .unwrap();
        let paragraphs = doc.paragraphs("body").unwrap();
        assert_eq!(
            paragraphs
                .iter()
                .map(|p| p.text.as_str())
                .collect::<Vec<_>>(),
            vec!["Se", "ed"]
        );
        assert_eq!(receipt.after.para_id, paragraphs[1].para_id);
        assert_eq!(receipt.after.offset, 1);
        assert_eq!(receipt.new_para_ids.len(), 2);
    }

    #[test]
    fn fragment_nested_table_creates_stories_under_its_cell() {
        let doc = fragment_doc();
        let nested = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {}, "cells": [fragment_cell("deep")]}],
        });
        let fragment = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {}, "cells": [{
                "tcPr": {}, "content": [
                    {"kind": "table", "table": nested, "attributes": {}},
                    {"kind": "pilcrow", "properties": {}, "attributes": {}},
                ]}]}],
        });
        let receipt = doc
            .insert_table_fragment(&direct(), Position::new("body", 0), &fragment)
            .unwrap();
        assert_eq!(receipt.created_story_ids.len(), 2);
        let cell_story = &receipt.created_story_ids[0];
        let nested_story = &receipt.created_story_ids[1];
        assert!(nested_story.starts_with(&format!("{cell_story}:")));
        assert_eq!(fragment_texts(&doc, nested_story), vec!["deep"]);
    }

    #[test]
    fn fragment_suggesting_stamps_and_reject_removes_table() {
        let doc = EditingDoc::new(77);
        doc.create_story("body", "Seed", "Normal", "left").unwrap();
        let before = project_story(&doc, "body").unwrap();
        let receipt = doc
            .insert_table_fragment(
                &suggesting(),
                Position::new("body", 2),
                &fragment_table(vec![fragment_cell("x")]),
            )
            .unwrap();
        assert_eq!(receipt.revision_ids.len(), 1);
        let payload = doc.table_payload(&receipt.table).unwrap();
        let rows = match payload.get("rows") {
            Some(Any::Array(rows)) => rows.clone(),
            other => panic!("expected rows array, got {other:?}"),
        };
        for row in rows.iter() {
            let Any::Map(row) = row else {
                panic!("expected row map")
            };
            let tr_pr = match row.get("trPr") {
                Some(Any::Map(tr_pr)) => tr_pr,
                other => panic!("expected trPr map, got {other:?}"),
            };
            assert!(tr_pr.contains_key("trIns"), "suggesting stamps trIns");
        }
        doc.reject_change(
            &direct(),
            &crate::ChangeTarget::Revision(receipt.revision_ids[0].clone()),
        )
        .unwrap();
        assert_eq!(project_story(&doc, "body").unwrap(), before);
        assert_eq!(doc.story_ids_for_test(), vec!["body".to_owned()]);
    }

    #[test]
    fn fragment_undo_removes_the_whole_insertion() {
        let doc = EditingDoc::new(77);
        doc.create_story("body", "Seed", "Normal", "left").unwrap();
        let before = project_story(&doc, "body").unwrap();
        let mut undo = doc.undo_manager();
        doc.insert_table_fragment(
            &direct(),
            Position::new("body", 2),
            &fragment_table(vec![fragment_cell("x")]),
        )
        .unwrap();
        assert!(undo.undo());
        assert_eq!(project_story(&doc, "body").unwrap(), before);
        assert_eq!(doc.story_ids_for_test(), vec!["body".to_owned()]);
    }

    #[test]
    fn fragment_grid_mismatch_is_recomputed() {
        for grid in [
            serde_json::json!([1]),
            serde_json::json!([2000, -5]),
            serde_json::json!([2000, "wide"]),
        ] {
            let doc = fragment_doc();
            let fragment = serde_json::json!({
                "tblPr": {}, "grid": grid,
                "rows": [{"trPr": {}, "cells": [fragment_cell("a"), fragment_cell("b")]}],
            });
            let receipt = doc
                .insert_table_fragment(&direct(), Position::new("body", 0), &fragment)
                .unwrap();
            assert_eq!((receipt.rows, receipt.columns), (1, 2));
            let payload = doc.table_payload(&receipt.table).unwrap();
            assert_eq!(
                payload.get("grid"),
                Some(&Any::from_json("[4680,4680]").unwrap())
            );
        }
    }

    #[test]
    fn fragment_empty_tbl_pr_gets_default_width() {
        let doc = fragment_doc();
        let receipt = doc
            .insert_table_fragment(
                &direct(),
                Position::new("body", 0),
                &fragment_table(vec![fragment_cell("a")]),
            )
            .unwrap();
        let payload = doc.table_payload(&receipt.table).unwrap();
        assert_eq!(
            payload.get("tblPr"),
            Some(&Any::from_json(r#"{"width":9360,"widthType":"dxa"}"#).unwrap())
        );
    }

    #[test]
    fn fragment_sanitizes_revisions_and_identity_keys() {
        let doc = fragment_doc();
        let fragment = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {"trIns": {"id": "r1"}, "height": 100,
                "_extraAttributes": [{"name": "w:rsidR", "value": "00112233"}]}, "cells": [{
                "tcPr": {"gridSpan": 2, "vMerge": "rest", "rowspan": 1, "colspan": 1, "backgroundColor": "old"},
                "shading": "FF0000",
                "content": [
                    {"kind": "text", "text": "gone", "attributes": {"del": {"id": "r2"}}},
                    {"kind": "text", "text": "kept", "attributes": {
                        "ins": {"id": "r3"}, "bold": true,
                        "hyperlink": {"href": "https://example.com", "tooltip": "t", "history": false, "rId": "rId9", "extra": 1}
                    }},
                    {"kind": "pilcrow", "properties": {
                        "paraId": "p", "_kind": "pilcrow", "ooxmlParaId": "1",
                        "sourceParaId": "s", "paraOrigin": "synthetic", "alignment": "right"
                    }, "attributes": {}},
                ]}]}],
        });
        let receipt = doc
            .insert_table_fragment(&direct(), Position::new("body", 0), &fragment)
            .unwrap();
        assert_eq!(
            fragment_texts(&doc, &receipt.created_story_ids[0]),
            vec!["kept"]
        );
        let segments = doc.story_segments(&receipt.created_story_ids[0]).unwrap();
        assert_eq!(segments[0].attributes.get("bold"), Some(&Any::Bool(true)));
        assert!(!segments[0].attributes.contains_key("ins"));
        let expected_link =
            Any::from_json(r#"{"href":"https://example.com","tooltip":"t","history":false}"#)
                .unwrap();
        assert_eq!(
            segments[0].attributes.get("hyperlink"),
            Some(&expected_link)
        );
        let payload = doc.table_payload(&receipt.table).unwrap();
        let rows = match payload.get("rows") {
            Some(Any::Array(rows)) => rows.clone(),
            other => panic!("expected rows array, got {other:?}"),
        };
        let Any::Map(row) = &rows[0] else {
            panic!("expected row map")
        };
        let tr_pr = match row.get("trPr") {
            Some(Any::Map(tr_pr)) => tr_pr.clone(),
            other => panic!("expected trPr map, got {other:?}"),
        };
        assert!(!tr_pr.contains_key("trIns"));
        assert!(!tr_pr.contains_key("_extraAttributes"));
        let cells = match row.get("cells") {
            Some(Any::Array(cells)) => cells.clone(),
            other => panic!("expected cells array, got {other:?}"),
        };
        let Any::Map(cell) = &cells[0] else {
            panic!("expected cell map")
        };
        let tc_pr = match cell.get("tcPr") {
            Some(Any::Map(tc_pr)) => tc_pr.clone(),
            other => panic!("expected tcPr map, got {other:?}"),
        };
        assert!(!tc_pr.contains_key("gridSpan"));
        assert!(!tc_pr.contains_key("vMerge"));
        assert_eq!(tc_pr.get("backgroundColor"), Some(&Any::from("FF0000")));
    }

    #[test]
    fn fragment_fully_covered_row_inserts_and_round_trips() {
        let doc = fragment_doc();
        let fragment = serde_json::json!({
            "tblPr": {}, "grid": [2000],
            "rows": [
                {"trPr": {"height": 10}, "cells": [
                    {"tcPr": {"rowspan": 2, "colspan": 1},
                     "content": [
                        {"kind": "text", "text": "A", "attributes": {}},
                        {"kind": "pilcrow", "properties": {}, "attributes": {}},
                     ]},
                ]},
                {"trPr": {"height": 20}, "cells": []},
            ],
        });
        let receipt = doc
            .insert_table_fragment(&direct(), Position::new("body", 0), &fragment)
            .unwrap();
        assert_eq!((receipt.rows, receipt.columns), (2, 1));
        assert_eq!(receipt.created_story_ids.len(), 1);
        let payload = doc.table_payload(&receipt.table).unwrap();
        let rows = match payload.get("rows") {
            Some(Any::Array(rows)) => rows.clone(),
            other => panic!("expected rows array, got {other:?}"),
        };
        assert_eq!(rows.len(), 2);
        let Any::Map(second) = &rows[1] else {
            panic!("expected row map")
        };
        assert_eq!(second.get("cells"), Some(&Any::from_json("[]").unwrap()));
        assert_eq!(
            second.get("trPr"),
            Some(&Any::from_json(r#"{"height":20}"#).unwrap())
        );
    }

    #[test]
    fn fragment_uncovered_empty_row_is_rejected() {
        let uncovered = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [
                {"trPr": {}, "cells": [fragment_cell("a")]},
                {"trPr": {}, "cells": []},
            ],
        });
        assert_fragment_refused(&fragment_doc(), &uncovered);
        let empty = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {}, "cells": []}],
        });
        assert_fragment_refused(&fragment_doc(), &empty);
    }

    #[test]
    fn fragment_rowspan_occupancy_inserts() {
        let doc = fragment_doc();
        let fragment = serde_json::json!({
            "tblPr": {}, "grid": [2000, 2000],
            "rows": [
                {"trPr": {}, "cells": [
                    {"tcPr": {"rowspan": 2, "colspan": 1},
                     "content": [
                        {"kind": "text", "text": "A", "attributes": {}},
                        {"kind": "pilcrow", "properties": {}, "attributes": {}},
                     ]},
                    {"tcPr": {"rowspan": 1, "colspan": 1},
                     "content": [
                        {"kind": "text", "text": "B", "attributes": {}},
                        {"kind": "pilcrow", "properties": {}, "attributes": {}},
                     ]},
                ]},
                {"trPr": {}, "cells": [
                    {"tcPr": {"rowspan": 1, "colspan": 1},
                     "content": [
                        {"kind": "text", "text": "C", "attributes": {}},
                        {"kind": "pilcrow", "properties": {}, "attributes": {}},
                     ]},
                ]},
            ],
        });
        let receipt = doc
            .insert_table_fragment(&direct(), Position::new("body", 0), &fragment)
            .unwrap();
        assert_eq!((receipt.rows, receipt.columns), (2, 2));
        assert_eq!(receipt.created_story_ids.len(), 3);
        assert_eq!(
            fragment_texts(&doc, &receipt.created_story_ids[2]),
            vec!["C"]
        );
    }

    #[test]
    fn fragment_huge_span_is_rejected_without_panic() {
        let colspan = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {}, "cells": [{
                "tcPr": {"colspan": 1_000_000_000},
                "content": [
                    {"kind": "text", "text": "a", "attributes": {}},
                    {"kind": "pilcrow", "properties": {}, "attributes": {}},
                ]}]}],
        });
        assert_fragment_refused(&fragment_doc(), &colspan);
        let rowspan = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {}, "cells": [{
                "tcPr": {"rowspan": 1_000_000_000},
                "content": [
                    {"kind": "text", "text": "a", "attributes": {}},
                    {"kind": "pilcrow", "properties": {}, "attributes": {}},
                ]}]}],
        });
        assert_fragment_refused(&fragment_doc(), &rowspan);
    }

    #[test]
    fn fragment_refusals_leave_the_document_unchanged() {
        let overlapping = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [
                {"trPr": {}, "cells": [
                    {"tcPr": {"colspan": 2}, "content": [
                        {"kind": "text", "text": "a", "attributes": {}},
                        {"kind": "pilcrow", "properties": {}, "attributes": {}},
                    ]},
                    {"tcPr": {"rowspan": 2}, "content": [
                        {"kind": "text", "text": "b", "attributes": {}},
                        {"kind": "pilcrow", "properties": {}, "attributes": {}},
                    ]},
                ]},
                {"trPr": {}, "cells": [
                    {"tcPr": {"colspan": 3}, "content": [
                        {"kind": "text", "text": "c", "attributes": {}},
                        {"kind": "pilcrow", "properties": {}, "attributes": {}},
                    ]},
                ]},
            ],
        });
        let no_pilcrow = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {}, "cells": [{
                "tcPr": {}, "content": [{"kind": "text", "text": "a", "attributes": {}}],
            }]}],
        });
        let line_break = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {}, "cells": [{
                "tcPr": {}, "content": [
                    {"kind": "text", "text": "a\nb", "attributes": {}},
                    {"kind": "pilcrow", "properties": {}, "attributes": {}},
                ]}]}],
        });
        let span_beyond = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {}, "cells": [{
                "tcPr": {"rowspan": 5}, "content": [
                    {"kind": "text", "text": "a", "attributes": {}},
                    {"kind": "pilcrow", "properties": {}, "attributes": {}},
                ]}]}],
        });
        let wide = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {}, "cells": [{
                "tcPr": {"colspan": 1001}, "content": [
                    {"kind": "text", "text": "a", "attributes": {}},
                    {"kind": "pilcrow", "properties": {}, "attributes": {}},
                ]}]}],
        });
        let bad_shading = serde_json::json!({
            "tblPr": {}, "grid": [],
            "rows": [{"trPr": {}, "cells": [{
                "tcPr": {}, "shading": "xyz",
                "content": [
                    {"kind": "text", "text": "a", "attributes": {}},
                    {"kind": "pilcrow", "properties": {}, "attributes": {}},
                ]}]}],
        });
        let cases: Vec<Value> = vec![
            serde_json::json!({"tblPr": {}, "grid": [], "rows": []}),
            serde_json::json!({"tblPr": {}, "grid": [], "rows": [{"trPr": {}, "cells": []}]}),
            serde_json::json!({"tblPr": [], "grid": [], "rows": [{"trPr": {}, "cells": [fragment_cell("a")]}]}),
            overlapping,
            no_pilcrow,
            line_break,
            span_beyond,
            wide,
            bad_shading,
        ];
        for fragment in &cases {
            assert_fragment_refused(&fragment_doc(), fragment);
        }
        let mut many_rows = Vec::new();
        for _ in 0..1001 {
            many_rows.push(serde_json::json!({"trPr": {}, "cells": [fragment_cell("a")]}));
        }
        assert_fragment_refused(
            &fragment_doc(),
            &serde_json::json!({"tblPr": {}, "grid": [], "rows": many_rows}),
        );
        let deep = (0..9).fold(fragment_cell("x"), |content, _| {
            serde_json::json!({
                "tcPr": {}, "content": [
                    {"kind": "table", "table": {
                        "tblPr": {}, "grid": [],
                        "rows": [{"trPr": {}, "cells": [content]}],
                    }, "attributes": {}},
                    {"kind": "pilcrow", "properties": {}, "attributes": {}},
                ],
            })
        });
        assert_fragment_refused(
            &fragment_doc(),
            &serde_json::json!({"tblPr": {}, "grid": [], "rows": [{"trPr": {}, "cells": [deep]}]}),
        );
    }

    #[test]
    fn delete_table_removes_embed_and_cell_story_tree() {
        let doc = seed_table();
        let receipt = doc
            .delete_table(&direct(), &TableLocator::new("body", 0))
            .unwrap();
        assert!(receipt.deleted_table);
        assert_eq!(receipt.deleted_story_ids.len(), 4);
        assert_eq!(doc.story_len("body").unwrap(), 0);
        assert_eq!(doc.story_ids_for_test(), vec!["body".to_owned()]);
    }

    impl EditingDoc {
        fn story_ids_for_test(&self) -> Vec<String> {
            let txn = self.yrs_doc().transact();
            let stories = txn.get_map(STORIES).unwrap();
            let mut ids: Vec<String> = crate::live_stories(&stories, &txn)
                .map(|(id, _)| id.to_string())
                .collect();
            ids.sort();
            ids
        }
    }
}
