//! Sheet-catalog targets, versioned cell reads and exact text search for edit batches.

use serde::{Deserialize, Serialize};
use xlsx_model::{CellRange, CellRef, CellValue, MAX_COLS, MAX_ROWS, Sheet, SheetId};
use xlsx_render::display_text;

use super::Workbook;
use super::batch::{
    DocumentVersion, EditFailure, EditFailureCode, EditRefusal, MAX_DIAGNOSTIC_CELLS, failure,
    refusal,
};
use crate::{CellAddress, Result};

/// Most cells one read may return.
const MAX_READ_CELLS: u64 = 100_000;
/// Most UTF-8 bytes of cell text one read or search may return.
const MAX_OUTPUT_TEXT_BYTES: usize = 32 * 1024 * 1024;
const DEFAULT_FIND_LIMIT: u32 = 100;
const MAX_FIND_LIMIT: u32 = 10_000;

/// A zero-based cell coordinate.
#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(deny_unknown_fields)]
pub struct CellPosition {
    pub row: u32,
    pub col: u32,
}

/// An inclusive rectangle: an A1 cell or `A1:B2` range (optionally with `$`), or zero-based
/// corners. Sheet-qualified, union, whole-row/column and defined-name references are refused.
#[derive(Clone, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum RangeAddress {
    A1 {
        a1: String,
    },
    RowCol {
        start: CellPosition,
        end: CellPosition,
    },
}

/// A range on the sheet the current catalog names `sheet_id`. Standalone ids are positional
/// (`sheet:{index}`); collaborative ids are the replica's sheet keys. Neither survives
/// save/reopen as an identity.
#[derive(Clone, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct RangeTarget {
    pub sheet_id: String,
    pub range: RangeAddress,
}

/// One cell in canonical form.
#[derive(Clone, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellTarget {
    pub sheet_id: String,
    pub row: u32,
    pub col: u32,
    pub a1: String,
}

#[derive(Clone, Debug, Default, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ReadRequest {
    /// Empty reads only the sheet catalog.
    #[serde(default)]
    pub ranges: Vec<RangeTarget>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetEntry {
    pub sheet_id: String,
    pub name: String,
    /// False for sheets batches refuse to write: non-worksheets and protected worksheets.
    pub editable: bool,
}

/// One cell: its tagged value (a formula's current result), its formula source without the
/// leading `=`, and the text the engine formats it as.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellRead {
    pub a1: String,
    pub value: CellValue,
    pub formula: Option<String>,
    pub display_text: String,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RangeRead {
    pub target: RangeTarget,
    /// Row-major.
    pub cells: Vec<Vec<CellRead>>,
}

/// Cells the last calculation settled by cycle or left at an evaluation limit.
#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ReadCalculation {
    pub cycle_cells: Vec<CellTarget>,
    pub limited_cells: Vec<CellTarget>,
    /// Whether a list stopped at 10,000 cells.
    pub truncated: bool,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellsRead {
    pub version: DocumentVersion,
    pub sheets: Vec<SheetEntry>,
    pub ranges: Vec<RangeRead>,
    pub calculation: ReadCalculation,
}

pub type ReadOutcome = std::result::Result<CellsRead, EditRefusal>;

/// An exact, case-sensitive substring search over formatted cell text.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FindRequest {
    pub text: String,
    /// Defaults to every sheet.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_ids: Option<Vec<String>>,
    /// Defaults to 100; at most 10,000.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub limit: Option<u32>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FindMatch {
    pub cell: CellTarget,
    /// The cell's whole display text.
    pub text: String,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextFound {
    pub version: DocumentVersion,
    /// One per cell, in sheet then row-major order.
    pub matches: Vec<FindMatch>,
    /// Whether more cells matched than `limit`.
    pub truncated: bool,
}

pub type FindOutcome = std::result::Result<TextFound, EditRefusal>;

/// A target resolved against the current catalog.
#[derive(Clone, Copy, Debug)]
pub(super) struct Resolved {
    pub(super) sheet: SheetId,
    pub(super) range: CellRange,
}

impl Resolved {
    pub(super) fn cells(self) -> u64 {
        u64::from(self.range.end.row - self.range.start.row + 1)
            * u64::from(self.range.end.col - self.range.start.col + 1)
    }

    pub(super) fn rows(self) -> usize {
        (self.range.end.row - self.range.start.row + 1) as usize
    }

    pub(super) fn cols(self) -> usize {
        (self.range.end.col - self.range.start.col + 1) as usize
    }

    pub(super) fn target(self, keys: &[String]) -> RangeTarget {
        RangeTarget {
            sheet_id: keys[self.sheet.0 as usize].clone(),
            range: RangeAddress::A1 {
                a1: self.range.to_a1(),
            },
        }
    }

    /// Row-major cells of the range.
    pub(super) fn iter(self) -> impl Iterator<Item = CellRef> {
        let range = self.range;
        (range.start.row..=range.end.row).flat_map(move |row| {
            (range.start.col..=range.end.col).map(move |col| CellRef::new(row, col))
        })
    }
}

pub(super) fn cell_target(keys: &[String], sheet: SheetId, cell: CellRef) -> CellTarget {
    CellTarget {
        sheet_id: keys[sheet.0 as usize].clone(),
        row: cell.row,
        col: cell.col,
        a1: CellRef::new(cell.row, cell.col).to_a1(),
    }
}

/// At most [`MAX_DIAGNOSTIC_CELLS`] targets, and whether `cells` held more.
fn capped_cell_targets(keys: &[String], cells: &[CellAddress]) -> (Vec<CellTarget>, bool) {
    let targets = cells
        .iter()
        .filter(|address| (address.sheet.0 as usize) < keys.len())
        .take(MAX_DIAGNOSTIC_CELLS)
        .map(|address| cell_target(keys, address.sheet, address.cell))
        .collect();
    (targets, cells.len() > MAX_DIAGNOSTIC_CELLS)
}

fn parse_cell(text: &str) -> Option<CellRef> {
    CellRef::parse_a1(&text.to_ascii_uppercase())
        .ok()
        .map(|cell| CellRef::new(cell.row, cell.col))
}

fn parse_range(address: &RangeAddress) -> std::result::Result<CellRange, String> {
    let (start, end) = match address {
        RangeAddress::A1 { a1 } => {
            let (start, end) = a1.split_once(':').unwrap_or((a1, a1));
            match (parse_cell(start), parse_cell(end)) {
                (Some(start), Some(end)) => (start, end),
                _ => {
                    return Err(format!(
                        "{a1:?} is not a cell or rectangular range; sheet-qualified, union, \
                         whole-row, whole-column and named references are not accepted"
                    ));
                }
            }
        }
        RangeAddress::RowCol { start, end } => {
            for position in [start, end] {
                if position.row >= MAX_ROWS || position.col >= MAX_COLS {
                    return Err(format!(
                        "row {} column {} is outside the sheet",
                        position.row, position.col
                    ));
                }
            }
            (
                CellRef::new(start.row, start.col),
                CellRef::new(end.row, end.col),
            )
        }
    };
    if start.row > end.row || start.col > end.col {
        return Err("a range must run from its top-left to its bottom-right corner".to_owned());
    }
    Ok(CellRange { start, end })
}

impl Workbook {
    /// Sheet ids in sheet order: the collaborative sheet keys, or `sheet:{index}` standalone.
    pub(super) fn sheet_keys(&self) -> Vec<String> {
        match &self.mode {
            super::WorkbookMode::Collaborative { structure } => structure.sheet_keys.clone(),
            super::WorkbookMode::Standalone => (0..self.model.sheets.len())
                .map(|index| format!("sheet:{index}"))
                .collect(),
        }
    }

    pub(super) fn resolve_target(
        &self,
        keys: &[String],
        target: &RangeTarget,
    ) -> std::result::Result<Resolved, EditFailure> {
        let refused = |code, message: String| failure(code, message, Some(target.clone()));
        let Some(index) = keys.iter().position(|key| *key == target.sheet_id) else {
            return Err(refused(
                EditFailureCode::MissingTarget,
                format!("no sheet has the id {:?}", target.sheet_id),
            ));
        };
        if index >= self.model.sheets.len() {
            return Err(refused(
                EditFailureCode::MissingTarget,
                format!("no sheet has the id {:?}", target.sheet_id),
            ));
        }
        let range = parse_range(&target.range)
            .map_err(|message| refused(EditFailureCode::InvalidStep, message))?;
        Ok(Resolved {
            sheet: SheetId(index as u32),
            range,
        })
    }

    /// Why batches refuse to write `sheet`, if they do.
    pub(super) fn sheet_write_refusal(&self, sheet: SheetId) -> Option<(EditFailureCode, String)> {
        let origin = self
            .preserved
            .origins
            .get(sheet.0 as usize)
            .copied()
            .flatten()?;
        let package = self.source_package.as_ref()?;
        let name = self
            .model
            .sheet(sheet)
            .map_or("", |sheet| sheet.name.as_str());
        if !package.source_sheet_is_worksheet(origin) {
            return Some((
                EditFailureCode::Unsupported,
                format!("sheet {name:?} is not a worksheet"),
            ));
        }
        package.source_sheet_is_protected(origin).then(|| {
            (
                EditFailureCode::LockedTarget,
                format!("sheet {name:?} is protected"),
            )
        })
    }

    /// Reads cells with the version they were read at. `Err` is an internal failure.
    pub fn read_cells(&self, request: &ReadRequest) -> Result<ReadOutcome> {
        let version = self.version();
        let keys = self.sheet_keys();
        let mut resolved = Vec::with_capacity(request.ranges.len());
        let mut cells = 0_u64;
        for target in &request.ranges {
            let range = match self.resolve_target(&keys, target) {
                Ok(range) => range,
                Err(failure) => return Ok(Err(refusal(version, failure))),
            };
            cells += range.cells();
            if cells > MAX_READ_CELLS {
                return Ok(Err(refusal(
                    version,
                    failure(
                        EditFailureCode::LimitExceeded,
                        format!("a read returns at most {MAX_READ_CELLS} cells"),
                        Some(target.clone()),
                    ),
                )));
            }
            resolved.push(range);
        }
        let mut text_bytes = 0_usize;
        let mut ranges = Vec::with_capacity(resolved.len());
        for range in resolved {
            let sheet = &self.model.sheets[range.sheet.0 as usize];
            let mut rows = Vec::with_capacity(range.rows());
            for row in range.range.start.row..=range.range.end.row {
                let mut read_row = Vec::with_capacity(range.cols());
                for col in range.range.start.col..=range.range.end.col {
                    let read = self.read_cell(sheet, CellRef::new(row, col));
                    text_bytes += read.display_text.len()
                        + read.formula.as_ref().map_or(0, String::len)
                        + match &read.value {
                            CellValue::Text { value } => value.len(),
                            _ => 0,
                        };
                    if text_bytes > MAX_OUTPUT_TEXT_BYTES {
                        return Ok(Err(refusal(
                            version,
                            failure(
                                EditFailureCode::LimitExceeded,
                                format!(
                                    "a read returns at most {MAX_OUTPUT_TEXT_BYTES} bytes of text"
                                ),
                                None,
                            ),
                        )));
                    }
                    read_row.push(read);
                }
                rows.push(read_row);
            }
            ranges.push(RangeRead {
                target: range.target(&keys),
                cells: rows,
            });
        }
        let sheets = self
            .model
            .sheets
            .iter()
            .enumerate()
            .map(|(index, sheet)| SheetEntry {
                sheet_id: keys[index].clone(),
                name: sheet.name.clone(),
                editable: self.sheet_write_refusal(SheetId(index as u32)).is_none(),
            })
            .collect();
        let (cycle_cells, cycles_cut) =
            capped_cell_targets(&keys, &self.last_calculation.cycle_cells);
        let (limited_cells, limited_cut) =
            capped_cell_targets(&keys, &self.last_calculation.limited_cells);
        let calculation = ReadCalculation {
            cycle_cells,
            limited_cells,
            truncated: cycles_cut || limited_cut,
        };
        Ok(Ok(CellsRead {
            version,
            sheets,
            ranges,
            calculation,
        }))
    }

    fn read_cell(&self, sheet: &Sheet, at: CellRef) -> CellRead {
        match sheet.cell(at) {
            Some(cell) => CellRead {
                a1: at.to_a1(),
                value: cell.value.clone(),
                formula: cell.formula.clone(),
                display_text: display_text(&self.model.styles, self.model.date_system, cell),
            },
            None => CellRead {
                a1: at.to_a1(),
                value: CellValue::Empty,
                formula: None,
                display_text: String::new(),
            },
        }
    }

    /// Finds cells whose display text contains `request.text`, case-sensitively, with the
    /// version they were read at. `Err` is an internal failure.
    pub fn find_text(&self, request: &FindRequest) -> Result<FindOutcome> {
        let version = self.version();
        let refused =
            |code, message: String| Ok(Err(refusal(version.clone(), failure(code, message, None))));
        if request.text.is_empty() {
            return refused(
                EditFailureCode::InvalidStep,
                "search text is empty".to_owned(),
            );
        }
        let limit = request.limit.unwrap_or(DEFAULT_FIND_LIMIT);
        if limit == 0 {
            return refused(
                EditFailureCode::InvalidStep,
                "limit must be positive".to_owned(),
            );
        }
        if limit > MAX_FIND_LIMIT {
            return refused(
                EditFailureCode::LimitExceeded,
                format!("a search returns at most {MAX_FIND_LIMIT} matches"),
            );
        }
        let keys = self.sheet_keys();
        let sheets = match &request.sheet_ids {
            None => (0..self.model.sheets.len()).collect::<Vec<_>>(),
            Some(ids) => {
                let mut sheets = Vec::with_capacity(ids.len());
                for id in ids {
                    match keys.iter().position(|key| key == id) {
                        Some(index) if index < self.model.sheets.len() => sheets.push(index),
                        _ => {
                            return refused(
                                EditFailureCode::MissingTarget,
                                format!("no sheet has the id {id:?}"),
                            );
                        }
                    }
                }
                sheets.sort_unstable();
                sheets.dedup();
                sheets
            }
        };
        let mut matches = Vec::new();
        let mut text_bytes = 0_usize;
        for index in sheets {
            for (at, cell) in self.model.sheets[index].iter_cells() {
                let text = display_text(&self.model.styles, self.model.date_system, cell);
                if !text.contains(request.text.as_str()) {
                    continue;
                }
                if matches.len() == limit as usize {
                    return Ok(Ok(TextFound {
                        version,
                        matches,
                        truncated: true,
                    }));
                }
                text_bytes += text.len();
                if text_bytes > MAX_OUTPUT_TEXT_BYTES {
                    return refused(
                        EditFailureCode::LimitExceeded,
                        format!("a search returns at most {MAX_OUTPUT_TEXT_BYTES} bytes of text"),
                    );
                }
                matches.push(FindMatch {
                    cell: cell_target(&keys, SheetId(index as u32), at),
                    text,
                });
            }
        }
        Ok(Ok(TextFound {
            version,
            matches,
            truncated: false,
        }))
    }
}
