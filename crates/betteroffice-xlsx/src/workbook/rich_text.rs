use std::cell::OnceCell;
use std::collections::HashSet;

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
    pub(super) fn cell_text_may_be_rich(
        &self,
        sheet: SheetId,
        cell: CellRef,
        rich_texts: &OnceCell<HashSet<String>>,
    ) -> bool {
        let Some(current) = self.model.sheet(sheet).and_then(|cells| cells.cell(cell)) else {
            return false;
        };
        let CellValue::Text { value } = &current.value else {
            return false;
        };
        if current.formula.is_some() {
            return false;
        }
        let Some(&Some(source_sheet)) = self.preserved.origins.get(sheet.0 as usize) else {
            return false;
        };
        let source_cell = self.text_source_cell(sheet, cell).map(|(_, at)| at);
        let rewritten = source_cell.is_some_and(|at| {
            self.invalidated_text_sources
                .as_ref()
                .is_some_and(|invalidated| invalidated.contains(&(source_sheet, at.row, at.col)))
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
            None => {
                let Some(facts) = package.source_cell_facts(source_sheet) else {
                    return true;
                };
                match source_cell {
                    Some(at) => facts.rich_inline.get(&(at.row, at.col)) == Some(value),
                    None => rich_texts
                        .get_or_init(|| {
                            let shared = &package.original_workbook().shared_strings;
                            (0..package.source_sheet_count())
                                .filter_map(|index| package.source_cell_facts(index))
                                .flat_map(|facts| facts.rich_inline.values().cloned())
                                .chain(
                                    shared
                                        .iter()
                                        .enumerate()
                                        .filter(|(index, _)| package.shared_string_is_rich(*index))
                                        .map(|(_, text)| text.clone()),
                                )
                                .collect()
                        })
                        .contains(value),
                }
            }
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

#[cfg(all(test, feature = "unstable-track-moves"))]
mod tests {
    use super::*;
    use crate::Error;
    use crate::workbook::edit_tests::{options, r, workbook_bytes};
    use xlsx_ops::Op;

    /// Sheet 1 whose first row is `row`, with a bold run in shared string 0.
    /// Model-only regeneration of that sheet writes every run as plain text.
    fn rich_bytes(row: &str) -> Vec<u8> {
        let mut parts = ooxml_opc::unzip_parts(&workbook_bytes()).unwrap();
        let (_, sheet) = parts
            .iter_mut()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap();
        let xml = String::from_utf8(std::mem::take(sheet)).unwrap();
        let start = xml.find("<sheetData>").unwrap() + "<sheetData>".len();
        let end = xml.find("</sheetData>").unwrap();
        *sheet = format!("{}{row}{}", &xml[..start], &xml[end..]).into_bytes();
        let (_, strings) = parts
            .iter_mut()
            .find(|(path, _)| path == "xl/sharedStrings.xml")
            .unwrap();
        *strings = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3"><si><r><rPr><b/></rPr><t>Alpha</t></r><r><t xml:space="preserve"> rich</t></r></si><si><t>Alpha rich</t></si><si><t>Beta</t></si></sst>"#.as_bytes().to_vec();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    /// A1 is shared string 0, D1 an inline italic run: a move of B before D
    /// leaves D1 where it is.
    const ROW_WITH_INLINE_RUN: &str = r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1" t="inlineStr"><is><r><rPr><i/></rPr><t>Italic</t></r></is></c></row>"#;

    /// The same row without inline text: its only runs live in the shared table.
    const ROW_WITH_SHARED_RUN: &str = r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row>"#;

    fn sheet_markup(bytes: &[u8]) -> String {
        let parts = ooxml_opc::unzip_parts(bytes).unwrap();
        let (_, sheet) = parts
            .iter()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap();
        String::from_utf8(sheet.clone()).unwrap()
    }

    #[test]
    fn a_track_move_refuses_rather_than_drop_inline_runs_on_untouched_cells() {
        // B before D and row 2 before row 4: D1 keeps its address either way.
        for op in [
            Op::MoveCols {
                sheet: SheetId(0),
                col: 1,
                count: 1,
                before: 3,
            },
            Op::MoveRows {
                sheet: SheetId(0),
                row: 1,
                count: 1,
                before: 3,
            },
        ] {
            let bytes = rich_bytes(ROW_WITH_INLINE_RUN);
            let mut workbook = Workbook::open(&bytes).unwrap();
            let runs = workbook.cell_text_runs(SheetId(0), r("D1"));
            assert!(runs.is_some_and(|runs| runs[0].style.italic == Some(true)));
            let before = workbook.model().clone();
            let saved_before = workbook.save().unwrap();

            // The save cannot patch a permutation, and regenerating the sheet
            // writes D1 as plain text.
            let result = workbook.apply_ops(vec![op], options());
            assert!(
                matches!(
                    result,
                    Err(Error::Operation(xlsx_ops::OpError::MoveNotSupported { .. }))
                ),
                "{result:?}"
            );
            assert_eq!(workbook.model(), &before);
            assert!(!workbook.can_undo());
            let runs = workbook.cell_text_runs(SheetId(0), r("D1"));
            assert!(runs.is_some_and(|runs| runs[0].style.italic == Some(true)));
            let saved = workbook.save().unwrap();
            assert_eq!(saved, saved_before);
            let markup = sheet_markup(&saved);
            assert!(
                markup.contains("<rPr><i/></rPr>"),
                "saved sheet dropped the italic run: {markup}"
            );
        }
    }

    #[test]
    fn a_noop_column_move_in_a_batch_keeps_source_markup() {
        let bytes = rich_bytes(ROW_WITH_INLINE_RUN);
        let mut workbook = Workbook::open(&bytes).unwrap();
        // The move leaves every track where it was: the source coordinates hold,
        // so the save patches the sheet instead of regenerating it.
        workbook
            .apply_ops(
                vec![
                    Op::SetCell {
                        sheet: SheetId(0),
                        at: r("A1"),
                        cell: xlsx_ops::CellState {
                            value: CellValue::Number { value: 5.0 },
                            ..Default::default()
                        },
                    },
                    Op::MoveCols {
                        sheet: SheetId(0),
                        col: 1,
                        count: 1,
                        before: 1,
                    },
                ],
                options(),
            )
            .unwrap();
        let saved = workbook.save().unwrap();
        let markup = sheet_markup(&saved);
        assert!(
            markup.contains("<rPr><i/></rPr>"),
            "saved sheet dropped the italic run: {markup}"
        );
    }

    #[test]
    fn a_column_move_keeps_shared_string_runs() {
        let mut workbook = Workbook::open(&rich_bytes(ROW_WITH_SHARED_RUN)).unwrap();
        assert!(
            workbook
                .cell_text_runs(SheetId(0), r("A1"))
                .is_some_and(|runs| runs[0].style.bold == Some(true))
        );
        // A before C: the cell moves to B1, whose run comes from the shared table.
        workbook
            .apply_ops(
                vec![Op::MoveCols {
                    sheet: SheetId(0),
                    col: 0,
                    count: 1,
                    before: 2,
                }],
                options(),
            )
            .unwrap();
        // The run stays in the shared table: only the query's source binding is
        // dropped with the sheet's source coordinates.
        assert_eq!(
            workbook.model().sheets[0].cell(r("B1")).unwrap().value,
            CellValue::Text {
                value: "Alpha rich".into()
            }
        );
        let saved = workbook.save().unwrap();
        let sheet = sheet_markup(&saved);
        assert!(
            sheet.contains(r#"<c r="B1" t="s"><v>0</v></c>"#),
            "saved sheet lost the shared-string cell: {sheet}"
        );
        let parts = ooxml_opc::unzip_parts(&saved).unwrap();
        let (_, strings) = parts
            .iter()
            .find(|(path, _)| path == "xl/sharedStrings.xml")
            .unwrap();
        let strings = String::from_utf8(strings.clone()).unwrap();
        assert!(
            strings.contains("<rPr><b/></rPr>"),
            "saved shared strings dropped the run: {strings}"
        );
    }
}
