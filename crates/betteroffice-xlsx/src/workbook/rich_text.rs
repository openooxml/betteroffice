use xlsx_model::{CellRef, CellValue, RichRun, SheetId};

use super::Workbook;

impl Workbook {
    /// Source runs in UTF-8 bytes; unavailable after content writes, including undo.
    /// No-op inputs retain runs. Collaborative and recovered sessions return `None`.
    pub fn cell_text_runs(&self, sheet: SheetId, cell: CellRef) -> Option<&[RichRun]> {
        if self.is_collaborative() {
            return None;
        }
        let invalidated = self.invalidated_text_sources.as_ref()?;
        let (source_sheet, source_cell) = self.text_source_cell(sheet, cell)?;
        if invalidated.contains(&(source_sheet, source_cell.row, source_cell.col)) {
            return None;
        }
        let current = self.model.sheets.get(sheet.0 as usize)?.cell(cell)?;
        if current.formula.is_some() {
            return None;
        }
        let CellValue::Text { value } = &current.value else {
            return None;
        };
        let source = self.source_package.as_ref()?;
        if let Some(&index) = self
            .preserved
            .shared_string_cells
            .get(sheet.0 as usize)?
            .get(&(cell.row, cell.col))
        {
            if self.model.shared_strings.get(index)? != value {
                return None;
            }
            source.materialize().ok()?.shared_string_text_runs(index)
        } else {
            let package = source.materialize().ok()?;
            let original = package
                .source_cell_facts(source_sheet)?
                .rich_inline
                .get(&(source_cell.row, source_cell.col))?;
            if original != value {
                return None;
            }
            package.inline_string_text_runs(source_sheet, source_cell)
        }
    }

    /// Whether a rewrite of `cell` could flatten source runs: `false` only when its text is
    /// known to be plain, in every session kind.
    pub(super) fn cell_text_may_be_rich(&self, sheet: SheetId, cell: CellRef) -> bool {
        let Some(current) = self.model.sheet(sheet).and_then(|cells| cells.cell(cell)) else {
            return false;
        };
        let CellValue::Text { value } = &current.value else {
            return false;
        };
        if current.formula.is_some() {
            return false;
        }
        let Some((source_sheet, source_cell)) = self.text_source_cell(sheet, cell) else {
            return false;
        };
        let rewritten = self.invalidated_text_sources.as_ref().is_some_and(|invalidated| {
            invalidated.contains(&(source_sheet, source_cell.row, source_cell.col))
        });
        if rewritten && !self.is_collaborative() {
            return false;
        }
        let Some(source) = self.source_package.as_ref() else {
            return false;
        };
        let Ok(package) = source.materialize() else {
            return true;
        };
        match self
            .preserved
            .shared_string_cells
            .get(sheet.0 as usize)
            .and_then(|cells| cells.get(&(cell.row, cell.col)))
        {
            Some(&index) => {
                self.model.shared_strings.get(index) == Some(value)
                    && package.shared_string_is_rich(index)
            }
            None => package.source_cell_facts(source_sheet).is_none_or(|facts| {
                facts.rich_inline.get(&(source_cell.row, source_cell.col)) == Some(value)
            }),
        }
    }

    pub(super) fn invalidate_text_source(&mut self, sheet: SheetId, cell: CellRef) {
        let Some((source_sheet, source_cell)) = self.text_source_cell(sheet, cell) else {
            return;
        };
        if let Some(invalidated) = &mut self.invalidated_text_sources {
            invalidated.insert((source_sheet, source_cell.row, source_cell.col));
        }
    }

    fn text_source_cell(&self, sheet: SheetId, cell: CellRef) -> Option<(usize, CellRef)> {
        let index = sheet.0 as usize;
        let source_sheet = (*self.preserved.origins.get(index)?)?;
        let axes = self.preserved.axes.get(index)?.as_ref()?;
        Some((
            source_sheet,
            CellRef::new(axes.rows.source(cell.row)?, axes.cols.source(cell.col)?),
        ))
    }
}
