//! Native table-structure operations over the yrs structural-table embed.
//!
//! A table is one map-backed unit in its parent story. Its `tblPr`, `grid`, and
//! `rows` fields are JSON-shaped [`Any`] values; every authored cell points at
//! an independent flat story. Since the three structural fields are atomic in
//! the persisted schema, each operation snapshots them, validates the grid,
//! rewrites them together in one transaction, and creates/removes cell stories
//! in that same transaction.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use serde::{Deserialize, Serialize};
use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{
    Any, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextPrelim, TextRef, Transact, TransactionMut,
};

use crate::op::{OpError, OpResult};
use crate::{
    EditCtx, EditingDoc, KIND_KEY, Position, STORIES, check_position, insertion_attrs, map_string,
    out_len, revision_value, story_ref, write_pilcrow_properties,
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
    pub new_para_ids: Vec<String>,
    pub deleted_table: bool,
    #[serde(default)]
    pub revision_ids: Vec<String>,
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

fn anchors(data: &TableData) -> OpResult<(Vec<CellAnchor>, usize)> {
    let mut occupied: Vec<Vec<bool>> = vec![Vec::new(); data.rows.len()];
    let mut result = Vec::new();
    let mut total_columns = data.grid.len();
    let mut seen_stories = HashSet::new();
    for (row_index, row) in data.rows.iter().enumerate() {
        let mut column = 0usize;
        for (cell_index, cell) in row.cells.iter().enumerate() {
            while occupied[row_index].get(column).copied().unwrap_or(false) {
                column += 1;
            }
            let (rowspan, colspan) = span(cell)?;
            if row_index + rowspan > data.rows.len() {
                return Err(invalid(format!(
                    "cell {} spans beyond the final row",
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
    // A PM CellSelection expands across a merged cell rather than selecting a
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
    set_spans(&mut cell, 1, 1);
    cell.tc_pr.remove("colwidth");
    cell.tc_pr.remove("cellMarker");
    cell.tc_pr.remove("_originalFormatting");
    cell
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

fn story_ids<T: ReadTxn>(txn: &T) -> HashSet<String> {
    let Some(stories) = txn.get_map(STORIES) else {
        return HashSet::new();
    };
    stories.iter(txn).map(|(id, _)| id.to_string()).collect()
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

fn fresh_cell_story(
    used: &mut HashSet<String>,
    base: &str,
    row: usize,
    mut column: usize,
) -> String {
    loop {
        let candidate = format!("{base}:r{row}c{column}");
        if used.insert(candidate.clone()) {
            return candidate;
        }
        column += 1;
    }
}

fn create_cell_story(
    doc: &EditingDoc,
    txn: &mut TransactionMut<'_>,
    story_id: &str,
) -> OpResult<String> {
    let stories = txn
        .get_map(STORIES)
        .expect("stories root is declared by EditingDoc::new");
    if stories.contains_key(txn, story_id) {
        return Err(OpError::StoryExists(story_id.to_owned()));
    }
    let para_id = doc.next_id();
    let story = stories.insert(txn, story_id, TextPrelim::new(""));
    let pilcrow = story.insert_embed_with_attributes(
        txn,
        0,
        MapPrelim::default(),
        insertion_attrs(None, None),
    );
    write_pilcrow_properties(&pilcrow, txn, &para_id, "Normal", "left");
    Ok(para_id)
}

fn remove_story_exact(txn: &mut TransactionMut<'_>, story_id: &str, deleted: &mut Vec<String>) {
    let stories = txn
        .get_map(STORIES)
        .expect("stories root is declared by EditingDoc::new");
    if stories.remove(txn, story_id).is_some() {
        deleted.push(story_id.to_owned());
    }
}

fn remove_story_tree(txn: &mut TransactionMut<'_>, story_id: &str, deleted: &mut Vec<String>) {
    let stories = txn
        .get_map(STORIES)
        .expect("stories root is declared by EditingDoc::new");
    let prefix = format!("{story_id}:");
    let mut ids: Vec<String> = stories
        .iter(txn)
        .map(|(id, _)| id.to_string())
        .filter(|id| id == story_id || id.starts_with(&prefix))
        .collect();
    // Children first is easier to inspect in receipts and avoids temporarily
    // leaving a descendant whose parent has already vanished.
    ids.sort_by_key(|id| std::cmp::Reverse(id.len()));
    for id in ids {
        if stories.remove(txn, &id).is_some() {
            deleted.push(id);
        }
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
            .all(|width| any_number(Some(width)).is_some())
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

/// Collects structural row revisions at their containing table embed. Multiple
/// rows from one table insertion intentionally collapse to one entry when they
/// share a revision id, matching the sidebar's "Inserted table" grouping.
pub(crate) fn table_row_changes<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<TableRowChange> {
    let mut changes = Vec::new();
    let mut offset = 0u32;
    for diff in story.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        if let Out::YMap(map) = diff.insert
            && map_string(&map, txn, KIND_KEY).as_deref() == Some("table")
            && let Ok(data) = read_table(&map, txn)
        {
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
        offset += len;
    }
    changes
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
) -> OpResult<()> {
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
            continue;
        }
        let mut deleted = Vec::new();
        for row in remove.into_iter().rev() {
            remove_row_at(txn, &mut data, row, &mut deleted)?;
        }
        write_table(txn, &table, &data);
    }
    Ok(())
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
        at: Position,
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
        let table_slot = fresh_table_slot(&used, &at.story, ordinal as usize);
        let base = format!("{}:t{table_slot}", at.story);
        let mut created_story_ids = Vec::with_capacity((rows * columns) as usize);
        let mut new_para_ids = Vec::with_capacity((rows * columns) as usize);
        let revision = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        let mut table_rows = Vec::with_capacity(rows as usize);
        for row in 0..rows as usize {
            let mut cells = Vec::with_capacity(columns as usize);
            for column in 0..columns as usize {
                let story_id = fresh_cell_story(&mut used, &base, row, column);
                let para_id = create_cell_story(self, &mut txn, &story_id)?;
                cells.push(CellData {
                    tc_pr: HashMap::from([
                        ("rowspan".to_owned(), Any::Number(1.0)),
                        ("colspan".to_owned(), Any::Number(1.0)),
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
        let mut tr_pr = data.rows[template_row].tr_pr.clone();
        tr_pr.remove(TR_INS);
        tr_pr.remove(TR_DEL);
        if let Some(revision_id) = revision_id.as_ref() {
            tr_pr.insert(
                TR_INS.to_owned(),
                revision_value(revision_id, &ctx.revision_author()),
            );
        }
        let used = story_ids(&txn);
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
            let story_id = fresh_cell_story(&mut used, &base, row_slot, column);
            let para_id = create_cell_story(self, &mut txn, &story_id)?;
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
            let story_id = fresh_cell_story(&mut used, &base, row, column_slot);
            let para_id = create_cell_story(self, &mut txn, &story_id)?;
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
    /// the survivor before their root entries are removed.
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
        let destination = story_ref(&txn, &target.cell.story).map_err(OpError::from)?;
        let mut deleted = Vec::new();
        let source_ids: HashSet<String> = selected
            .iter()
            .skip(1)
            .map(|anchor| anchor.cell.story.clone())
            .collect();
        for source in selected.iter().skip(1) {
            let source_story = story_ref(&txn, &source.cell.story).map_err(OpError::from)?;
            let chunks = snapshot_story(&source_story, &txn)?;
            append_story(&mut txn, &destination, &chunks);
            // Nested table stories remain reachable through the moved embed;
            // only the source cell-story root itself becomes unreachable.
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
        receipt(locator, Some(&data), Vec::new(), deleted, Vec::new())
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
                let story_id = fresh_cell_story(&mut used, &base, row, column);
                let para_id = create_cell_story(self, &mut txn, &story_id)?;
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
                let source = if row < target_row_end {
                    row
                } else if row < target.row + requested_rows {
                    target_row_end - 1
                } else {
                    row - delta_rows
                };
                RowData {
                    tr_pr: original_rows[source].tr_pr.clone(),
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

    /// Replaces the complete `tcPr.borders` object on every selected cell.
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
        let selected = selected_anchors(&data, range)?;
        let (all, _) = anchors(&data)?;

        for anchor in &selected {
            data.rows[anchor.row].cells[anchor.cell_index]
                .tc_pr
                .insert("borders".to_owned(), to_any_map(borders.clone()));
        }

        // Keep shared edges symmetric, matching the PM border commands. A
        // supplied top/bottom/left/right value is mirrored onto every cell
        // touching that edge, including cells just outside the selection.
        for anchor in &selected {
            for (side, value) in borders {
                let (facing, slots): (&str, Vec<(usize, usize)>) = match side.as_str() {
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
                    let cell = &mut data.rows[neighbor.row].cells[neighbor.cell_index];
                    let mut neighbor_borders = cell
                        .tc_pr
                        .get("borders")
                        .and_then(|value| match value {
                            Any::Map(map) => Some(map.as_ref().clone()),
                            _ => None,
                        })
                        .unwrap_or_default();
                    neighbor_borders.insert(facing.to_owned(), value.clone());
                    cell.tc_pr
                        .insert("borders".to_owned(), to_any_map(neighbor_borders));
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
    fn insert_row_creates_stories_and_still_lowers_to_layout_table() {
        let doc = seed_table();
        let receipt = doc.insert_row(&direct(), &cell(0, 0), true).unwrap();
        assert_eq!((receipt.rows, receipt.columns), (3, 2));
        assert_eq!(receipt.created_story_ids.len(), 2);
        assert!(
            receipt
                .created_story_ids
                .contains(&"body:t0:r2c0".to_owned())
        );
        assert!(
            receipt
                .created_story_ids
                .contains(&"body:t0:r2c1".to_owned())
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
                .contains(&"body:t0:r0c2".to_owned())
        );
        assert!(
            inserted
                .created_story_ids
                .contains(&"body:t0:r1c2".to_owned())
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
            let mut ids: Vec<String> = stories.iter(&txn).map(|(id, _)| id.to_string()).collect();
            ids.sort();
            ids
        }
    }
}
