use std::collections::BTreeSet;
use std::ops::Bound::{Excluded, Unbounded};

use xlsx_model::{CellValue, MAX_COLS, MAX_ROWS, SheetChart, Workbook};

use crate::snapshot::{SnapshotBudget, SnapshotError, SnapshotResult};

use super::super::{
    MAX_CHART_FIELD_BYTES, MAX_CHART_REFS_PER_CHART, MAX_CHARTS_PER_SHEET, MAX_COL_WIDTH,
    MAX_HYPERLINKS_PER_SHEET, MAX_ROW_HEIGHT, is_writable_xml_text, ranges_intersect,
    validate_cell_ref, validate_charts, validate_hyperlinks, validate_range, validate_sheet_name,
};

#[derive(Default)]
pub(super) struct ModelValidation {
    phase: u8,
    sheet: usize,
    index: usize,
    other: usize,
    after: Option<(u32, u32)>,
    dimension: Option<u32>,
    names: BTreeSet<String>,
    charts: BTreeSet<(String, usize)>,
}

fn check(result: crate::Result<()>) -> SnapshotResult<()> {
    result.map_err(|failure| SnapshotError::new(failure.to_string()))
}

fn admit(bytes: usize, budget: SnapshotBudget) -> SnapshotResult<usize> {
    if bytes > budget.max_bytes() {
        Err(SnapshotError::new(
            "snapshot validation exceeds advance byte budget",
        ))
    } else {
        crate::snapshot::step::record(1, bytes);
        Ok(bytes)
    }
}

impl ModelValidation {
    pub(super) fn advance(
        &mut self,
        model: &Workbook,
        package_present: bool,
        budget: SnapshotBudget,
    ) -> SnapshotResult<(bool, usize, usize)> {
        if model.sheets.is_empty() {
            return Err(SnapshotError::new("snapshot workbook has no sheets"));
        }
        loop {
            if self.phase == 0 {
                if let Some(name) = model.defined_names.get(self.index) {
                    let bytes = admit(128, budget)?;
                    if name
                        .local_sheet
                        .is_some_and(|sheet| sheet.0 as usize >= model.sheets.len())
                        || name.formula.len() > xlsx_calc::lexer::MAX_FORMULA_BYTES
                    {
                        return Err(SnapshotError::new("snapshot defined name is invalid"));
                    }
                    self.index += 1;
                    return Ok((false, 1, bytes));
                }
                self.phase = 1;
                self.index = 0;
            }
            if self.phase == 10 {
                if let Some(name) = self.names.first() {
                    let bytes = admit(name.len().saturating_add(24), budget)?;
                    self.names.pop_first();
                    return Ok((false, 1, bytes));
                }
                self.phase = 11;
                return Ok((true, 0, 0));
            }
            if self.phase == 11 {
                return Ok((true, 0, 0));
            }
            let Some(sheet) = model.sheets.get(self.sheet) else {
                self.phase = 10;
                continue;
            };
            match self.phase {
                1 => {
                    let bytes = admit(
                        sheet.name.len().saturating_mul(3).saturating_add(64),
                        budget,
                    )?;
                    check(validate_sheet_name(&sheet.name))?;
                    #[cfg(test)]
                    crate::snapshot::step::allocate(bytes);
                    if !self.names.insert(sheet.name.to_lowercase()) {
                        return Err(SnapshotError::new("snapshot has duplicate sheet names"));
                    }
                    if sheet.freeze_pane.is_some_and(|pane| {
                        pane.rows > MAX_ROWS
                            || pane.cols > MAX_COLS
                            || pane.top_left.row >= MAX_ROWS
                            || pane.top_left.col >= MAX_COLS
                    }) || sheet.hyperlinks.len() > MAX_HYPERLINKS_PER_SHEET
                        || sheet.charts.len() > MAX_CHARTS_PER_SHEET
                    {
                        return Err(SnapshotError::new("snapshot sheet metadata is invalid"));
                    }
                    self.phase = 2;
                    self.after = None;
                    return Ok((false, 1, bytes));
                }
                2 => {
                    let start = self.after.map_or((0, 0), |(row, col)| (row, col + 1));
                    let range = xlsx_model::CellRange {
                        start: xlsx_model::CellRef::new(start.0, start.1),
                        end: xlsx_model::CellRef::new(start.0, u32::MAX),
                    };
                    let later = sheet.cells_in_range(xlsx_model::CellRange {
                        start: xlsx_model::CellRef::new(start.0.saturating_add(1), 0),
                        end: xlsx_model::CellRef::new(u32::MAX, u32::MAX),
                    });
                    if let Some((at, cell)) = sheet.cells_in_range(range).chain(later).next() {
                        let text = match &cell.value {
                            CellValue::Text { value } => value.len(),
                            _ => 0,
                        };
                        let bytes = admit(
                            text.saturating_add(cell.formula.as_ref().map_or(0, String::len))
                                .max(16),
                            budget,
                        )?;
                        check(validate_cell_ref(at))?;
                        if (cell.formula.is_none()
                            && matches!(cell.value, CellValue::Number { value } if !value.is_finite()))
                            || matches!(&cell.value, CellValue::Text { value } if value.chars().count() > xlsx_calc::eval::MAX_CELL_TEXT_CHARS)
                            || cell.formula.as_ref().is_some_and(|formula| {
                                formula.len() > xlsx_calc::lexer::MAX_FORMULA_BYTES
                            })
                            || cell.style.is_some_and(|style| {
                                style as usize >= model.styles.cell_xfs.len().max(1)
                            })
                        {
                            return Err(SnapshotError::new("snapshot cell is invalid"));
                        }
                        self.after = Some((at.row, at.col));
                        return Ok((false, 1, bytes));
                    }
                    self.phase = 3;
                    self.dimension = None;
                }
                3 | 4 => {
                    let values = if self.phase == 3 {
                        &sheet.col_widths
                    } else {
                        &sheet.row_heights
                    };
                    let next = match self.dimension {
                        Some(after) => values.range((Excluded(after), Unbounded)).next(),
                        None => values.first_key_value(),
                    };
                    if let Some((&index, &value)) = next {
                        let bytes = admit(16, budget)?;
                        let (limit, maximum) = if self.phase == 3 {
                            (MAX_COLS, MAX_COL_WIDTH)
                        } else {
                            (MAX_ROWS, MAX_ROW_HEIGHT)
                        };
                        if index >= limit || !value.is_finite() || !(0.0..=maximum).contains(&value)
                        {
                            return Err(SnapshotError::new("snapshot sheet dimension is invalid"));
                        }
                        self.dimension = Some(index);
                        return Ok((false, 1, bytes));
                    }
                    self.phase += 1;
                    self.dimension = None;
                    self.index = 0;
                }
                5 => {
                    if let Some(link) = sheet.hyperlinks.get(self.index) {
                        let bytes = admit(128, budget)?;
                        check(validate_hyperlinks(std::slice::from_ref(link)))?;
                        self.index += 1;
                        return Ok((false, 1, bytes));
                    }
                    self.phase = 6;
                    self.index = 0;
                    self.other = 0;
                }
                6 => {
                    if let Some(range) = sheet.merges.get(self.index) {
                        let bytes = admit(64, budget)?;
                        check(validate_range(*range))?;
                        if let Some(other) = sheet
                            .merges
                            .get(self.other)
                            .filter(|_| self.other < self.index)
                        {
                            if ranges_intersect(*range, *other) {
                                return Err(SnapshotError::new("snapshot merged ranges overlap"));
                            }
                            self.other += 1;
                        } else {
                            self.index += 1;
                            self.other = 0;
                        }
                        return Ok((false, 1, bytes));
                    }
                    self.phase = 7;
                    self.index = 0;
                    self.other = 0;
                }
                7 => {
                    if let Some(chart) = sheet.charts.get(self.index) {
                        if !package_present {
                            admit(1, budget)?;
                            return Err(SnapshotError::new(
                                "snapshot charts require a source package",
                            ));
                        }
                        if self.other == 0 {
                            let bytes = admit(
                                chart
                                    .part
                                    .len()
                                    .saturating_add(chart.drawing.len().saturating_mul(2))
                                    .saturating_add(128),
                                budget,
                            )?;
                            #[cfg(test)]
                            crate::snapshot::step::allocate(bytes);
                            if chart.refs.len() > MAX_CHART_REFS_PER_CHART
                                || !self
                                    .charts
                                    .insert((chart.drawing.clone(), chart.anchor_index))
                            {
                                return Err(SnapshotError::new("snapshot chart is invalid"));
                            }
                            let header = SheetChart {
                                part: chart.part.clone(),
                                drawing: chart.drawing.clone(),
                                anchor_index: chart.anchor_index,
                                anchor: chart.anchor,
                                refs: Vec::new(),
                            };
                            check(validate_charts(std::slice::from_ref(&header)))?;
                            self.other = 1;
                            return Ok((false, 1, bytes));
                        }
                        if let Some(reference) = chart.refs.get(self.other - 1) {
                            let bytes = admit(reference.formula.len().saturating_add(16), budget)?;
                            if reference.formula.len() > MAX_CHART_FIELD_BYTES
                                || !is_writable_xml_text(&reference.formula)
                            {
                                return Err(SnapshotError::new(
                                    "snapshot chart reference is invalid",
                                ));
                            }
                            self.other += 1;
                            return Ok((false, 1, bytes));
                        }
                        self.index += 1;
                        self.other = 0;
                        admit(0, budget)?;
                        return Ok((false, 1, 0));
                    }
                    self.phase = 8;
                }
                8 => {
                    if let Some((drawing, _)) = self.charts.first() {
                        let bytes = admit(drawing.len().saturating_add(32), budget)?;
                        self.charts.pop_first();
                        return Ok((false, 1, bytes));
                    }
                    self.sheet += 1;
                    self.phase = 1;
                    self.index = 0;
                }
                _ => unreachable!(),
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{
        AnchorExtent, AnchorPos, CellRange, CellRef, ChartAnchor, ChartRef, ChartRefKind, Sheet,
    };

    fn chart(index: usize) -> SheetChart {
        SheetChart {
            part: "xl/charts/chart1.xml".to_owned(),
            drawing: "xl/drawings/drawing1.xml".to_owned(),
            anchor_index: index,
            anchor: ChartAnchor::Absolute {
                pos: AnchorPos::default(),
                extent: AnchorExtent { cx: 1, cy: 1 },
            },
            refs: (0..32)
                .map(|row| ChartRef {
                    kind: ChartRefKind::Values,
                    formula: format!("Data!$A${}", row + 1),
                })
                .collect(),
        }
    }

    fn step(
        validation: &mut ModelValidation,
        model: &Workbook,
        budget: SnapshotBudget,
    ) -> SnapshotResult<(bool, usize, usize)> {
        crate::snapshot::step::reset();
        let result = validation.advance(model, true, budget);
        let work = crate::snapshot::step::current();
        assert!(work.records <= budget.max_records(), "{work:?}");
        assert!(work.bytes <= budget.max_bytes(), "{work:?}");
        assert!(work.allocated_bytes <= budget.max_bytes(), "{work:?}");
        if let Ok((_, records, bytes)) = &result {
            assert_eq!(work.records, *records);
            assert_eq!(work.bytes, *bytes);
        }
        result
    }

    #[test]
    fn merge_pairs_validation_visits_and_allocations_are_bounded() {
        let budget = SnapshotBudget::new(1, 16_384).unwrap();
        for overlap in [false, true] {
            let mut sheet = Sheet::new("Data");
            sheet.merges = (0..32)
                .map(|row| CellRange {
                    start: CellRef::new(row * 2, 0),
                    end: CellRef::new(row * 2, 1),
                })
                .collect();
            if overlap {
                sheet.merges[31] = sheet.merges[0];
            }
            let model = Workbook {
                sheets: vec![sheet],
                ..Workbook::default()
            };
            let mut validation = ModelValidation::default();
            let mut pairs = 0;
            let mut refused = false;
            loop {
                let pair = validation.phase == 6 && validation.other < validation.index;
                match step(&mut validation, &model, budget) {
                    Ok((ready, _, _)) => {
                        pairs += usize::from(pair);
                        if ready {
                            break;
                        }
                    }
                    Err(failure) => {
                        assert!(overlap);
                        assert_eq!(failure.to_string(), "snapshot merged ranges overlap");
                        refused = true;
                        break;
                    }
                }
            }
            assert_eq!(refused, overlap);
            if !overlap {
                assert!(pairs >= 32 * 31 / 2);
                assert!(validation.names.is_empty());
            }
        }
    }

    #[test]
    fn chart_headers_references_and_index_retirement_are_bounded() {
        let budget = SnapshotBudget::new(1, 16_384).unwrap();
        for invalid in 0..4 {
            let mut model = Workbook {
                sheets: (0..32)
                    .map(|index| Sheet::new(format!("Sheet{index}")))
                    .collect(),
                ..Workbook::default()
            };
            model.sheets[0].charts = (0..16).map(chart).collect();
            match invalid {
                1 => model.sheets[0].charts[15].part.clear(),
                2 => model.sheets[0].charts[15].refs[31].formula = "\0".to_owned(),
                3 => model.sheets[0].charts[15].drawing = "x".repeat(budget.max_bytes() + 1),
                _ => {}
            }
            let mut validation = ModelValidation::default();
            let mut headers = 0;
            let mut references = 0;
            let mut charts_retired = 0;
            let mut names_retired = 0;
            let mut refused = false;
            loop {
                let phase = validation.phase;
                let charts = validation.charts.len();
                let names = validation.names.len();
                let header = phase == 7 && validation.other == 0 && validation.index < 16;
                let reference = phase == 7 && (1..=32).contains(&validation.other);
                match step(&mut validation, &model, budget) {
                    Ok((ready, _, _)) => {
                        headers += usize::from(header);
                        references += usize::from(reference);
                        charts_retired += charts.saturating_sub(validation.charts.len());
                        names_retired += names.saturating_sub(validation.names.len());
                        if ready {
                            break;
                        }
                    }
                    Err(_) => {
                        refused = true;
                        break;
                    }
                }
            }
            assert_eq!(refused, invalid != 0);
            if invalid == 0 {
                assert!(headers >= 15);
                assert_eq!(references, 16 * 32);
                assert_eq!(charts_retired, 16);
                assert_eq!(names_retired, 32);
                assert!(validation.charts.is_empty());
                assert!(validation.names.is_empty());
            }
        }
    }
}
