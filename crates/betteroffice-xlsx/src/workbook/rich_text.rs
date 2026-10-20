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
