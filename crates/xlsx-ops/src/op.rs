//! wire types for the op log: the serializable `Op` vocabulary, a `CellState`
//! mirror of the model's non-serde `Cell`, provenance tags, and transactions.

use serde::{Deserialize, Serialize};
use xlsx_model::{
    Cell, CellRange, CellRef, CellValue, ColId, FreezePane, Hyperlink, RowId, SheetId,
};

use crate::formatting::{CapturedFormat, NumberFormatMutation, StylePatch};

/// serializable mirror of `xlsx_model::Cell`, which deliberately has no serde.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellState {
    #[serde(default)]
    pub value: CellValue,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<u32>,
}

impl From<&Cell> for CellState {
    fn from(c: &Cell) -> Self {
        Self {
            value: c.value.clone(),
            formula: c.formula.clone(),
            style: c.style,
        }
    }
}

impl From<Cell> for CellState {
    fn from(c: Cell) -> Self {
        Self {
            value: c.value,
            formula: c.formula,
            style: c.style,
        }
    }
}

impl From<CellState> for Cell {
    fn from(s: CellState) -> Self {
        Cell {
            value: s.value,
            formula: s.formula,
            style: s.style,
        }
    }
}

/// the invertible edit vocabulary. width/height are `Option` so "reset to
/// default" is expressible and invertible.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum Op {
    SetCell {
        sheet: SheetId,
        at: CellRef,
        cell: CellState,
    },
    InsertRows {
        sheet: SheetId,
        at: RowId,
        count: u32,
    },
    DeleteRows {
        sheet: SheetId,
        at: RowId,
        count: u32,
    },
    InsertCols {
        sheet: SheetId,
        at: ColId,
        count: u32,
    },
    DeleteCols {
        sheet: SheetId,
        at: ColId,
        count: u32,
    },
    SetColWidth {
        sheet: SheetId,
        col: ColId,
        width: Option<f64>,
    },
    SetRowHeight {
        sheet: SheetId,
        row: RowId,
        height: Option<f64>,
    },
    SetFreezePane {
        sheet: SheetId,
        pane: Option<FreezePane>,
    },
    #[doc(hidden)]
    SetHyperlinks {
        sheet: SheetId,
        hyperlinks: Vec<Hyperlink>,
    },
    MergeCells {
        sheet: SheetId,
        range: CellRange,
    },
    UnmergeCells {
        sheet: SheetId,
        range: CellRange,
    },
    PatchRangeStyle {
        sheet: SheetId,
        range: CellRange,
        patch: StylePatch,
    },
    SetRangeNumberFormat {
        sheet: SheetId,
        range: CellRange,
        format: NumberFormatMutation,
    },
    ApplyRangeFormat {
        sheet: SheetId,
        range: CellRange,
        format: CapturedFormat,
    },
    AddSheet {
        index: usize,
        name: String,
    },
    RemoveSheet {
        index: usize,
    },
    RenameSheet {
        sheet: SheetId,
        name: String,
    },
    #[doc(hidden)]
    RestoreSheet {
        sheet: SheetId,
        name: String,
        formulas: Vec<(SheetId, CellRef, CellState)>,
    },
}

/// who authored a transaction.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum Provenance {
    User,
    Agent { id: String },
    Remote { actor: String },
}

/// an atomic, provenance-tagged batch of ops. `proposed = true` marks a HITL
/// proposal awaiting accept or reject; it is not yet applied.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Transaction {
    pub ops: Vec<Op>,
    pub author: Provenance,
    #[serde(default)]
    pub proposed: bool,
}

impl Transaction {
    pub fn new(ops: Vec<Op>, author: Provenance) -> Self {
        Self {
            ops,
            author,
            proposed: false,
        }
    }

    /// a proposal awaiting review — same payload, not yet committed.
    pub fn proposal(ops: Vec<Op>, author: Provenance) -> Self {
        Self {
            ops,
            author,
            proposed: true,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cell_state_round_trips_through_model_cell() {
        let cell = Cell {
            value: CellValue::Number { value: 3.5 },
            formula: Some("1+2.5".into()),
            style: Some(7),
        };
        let state = CellState::from(&cell);
        assert_eq!(Cell::from(state), cell);
    }
}
