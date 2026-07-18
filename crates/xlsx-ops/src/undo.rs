//! undo/redo over committed transactions. each stack entry is the base-op list
//! that reverses one transaction; undoing replays it and captures a redo inverse.

use xlsx_model::Workbook;

use crate::apply::{OpError, apply_ops};
use crate::op::{Op, Transaction};

#[derive(Debug, Clone, Default)]
pub struct UndoStack {
    undo: Vec<Vec<Op>>,
    redo: Vec<Vec<Op>>,
}

impl UndoStack {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn can_undo(&self) -> bool {
        !self.undo.is_empty()
    }

    pub fn can_redo(&self) -> bool {
        !self.redo.is_empty()
    }

    /// apply a transaction's ops and record its inverse for undo; clears the
    /// redo stack.
    pub fn commit(&mut self, wb: &mut Workbook, tx: &Transaction) -> Result<(), OpError> {
        let inverse = apply_ops(wb, &tx.ops)?;
        self.undo.push(inverse);
        self.redo.clear();
        Ok(())
    }

    /// reverse the most recent transaction, returning the ops applied.
    pub fn undo(&mut self, wb: &mut Workbook) -> Result<Option<Vec<Op>>, OpError> {
        let Some(ops) = self.undo.last().cloned() else {
            return Ok(None);
        };
        let redo_inverse = apply_ops(wb, &ops)?;
        self.undo.pop();
        self.redo.push(redo_inverse);
        Ok(Some(ops))
    }

    /// re-apply the most recently undone transaction.
    pub fn redo(&mut self, wb: &mut Workbook) -> Result<Option<Vec<Op>>, OpError> {
        let Some(ops) = self.redo.last().cloned() else {
            return Ok(None);
        };
        let undo_inverse = apply_ops(wb, &ops)?;
        self.redo.pop();
        self.undo.push(undo_inverse);
        Ok(Some(ops))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::op::{CellState, Provenance};
    use xlsx_model::{Cell, CellProvider, CellRef, CellValue, Sheet, SheetId, Workbook};

    fn r(a1: &str) -> CellRef {
        CellRef::parse_a1(a1).unwrap()
    }

    fn set(a1: &str, v: f64) -> Op {
        Op::SetCell {
            sheet: SheetId(0),
            at: r(a1),
            cell: CellState {
                value: CellValue::Number { value: v },
                ..Default::default()
            },
        }
    }

    #[test]
    fn commit_undo_redo_cycle() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));

        let tx = Transaction::new(vec![set("A1", 1.0), set("B1", 2.0)], Provenance::User);
        let mut stack = UndoStack::new();
        stack.commit(&mut wb, &tx).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A1")),
            CellValue::Number { value: 1.0 }
        );
        assert_eq!(
            wb.value(SheetId(0), r("B1")),
            CellValue::Number { value: 2.0 }
        );

        assert!(stack.can_undo());
        stack.undo(&mut wb).unwrap();
        assert_eq!(wb.value(SheetId(0), r("A1")), CellValue::Empty);
        assert_eq!(wb.value(SheetId(0), r("B1")), CellValue::Empty);

        assert!(stack.can_redo());
        stack.redo(&mut wb).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A1")),
            CellValue::Number { value: 1.0 }
        );
        assert_eq!(
            wb.value(SheetId(0), r("B1")),
            CellValue::Number { value: 2.0 }
        );
    }

    #[test]
    fn undo_preserves_prior_value_not_just_empty() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        wb.sheet_mut(SheetId(0)).unwrap().set_cell(
            r("A1"),
            Cell {
                value: CellValue::Number { value: 5.0 },
                ..Default::default()
            },
        );

        let mut stack = UndoStack::new();
        stack
            .commit(
                &mut wb,
                &Transaction::new(vec![set("A1", 99.0)], Provenance::User),
            )
            .unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A1")),
            CellValue::Number { value: 99.0 }
        );
        stack.undo(&mut wb).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A1")),
            CellValue::Number { value: 5.0 }
        );
    }

    #[test]
    fn empty_stacks_return_none() {
        let mut wb = Workbook::default();
        let mut stack = UndoStack::new();
        assert_eq!(stack.undo(&mut wb).unwrap(), None);
        assert_eq!(stack.redo(&mut wb).unwrap(), None);
    }

    #[test]
    fn failed_undo_keeps_history_and_workbook_state() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        let before = wb.clone();
        let mut stack = UndoStack::new();
        stack.undo.push(vec![
            set("A1", 1.0),
            Op::SetCell {
                sheet: SheetId(9),
                at: r("A1"),
                cell: CellState::default(),
            },
        ]);
        assert!(stack.undo(&mut wb).is_err());
        assert!(stack.can_undo());
        assert_eq!(wb, before);
    }
}
