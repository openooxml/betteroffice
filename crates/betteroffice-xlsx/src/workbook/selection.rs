use super::*;

#[derive(Debug, Clone, Copy, Deserialize, Serialize)]
#[serde(rename_all = "lowercase")]
#[non_exhaustive]
pub enum SelectionDirection {
    Up,
    Down,
    Left,
    Right,
    Merge,
    #[serde(rename = "usedRange")]
    UsedRange,
    #[serde(rename = "rowEnd")]
    RowEnd,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
pub struct SelectionRegion {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
}

#[derive(Debug, Serialize)]
pub struct SelectionData {
    pub merge: Option<SelectionRegion>,
    pub edge: Option<CellRef>,
    pub region: Option<SelectionRegion>,
    #[serde(rename = "usedRange")]
    pub used_range: Option<SelectionRegion>,
}

fn filled(cell: &xlsx_model::Cell) -> bool {
    cell.formula.is_some() || !matches!(cell.value, CellValue::Empty)
}

struct VisibleAxis {
    spans: Vec<(u32, u32)>,
    counts: Vec<u32>,
}

impl VisibleAxis {
    fn new(spans: impl Iterator<Item = (u32, u32)>) -> Self {
        let spans: Vec<_> = spans.collect();
        let mut counts = vec![0];
        for &(start, end) in &spans {
            counts.push(counts.last().unwrap() + end - start + 1);
        }
        Self { spans, counts }
    }

    fn project(&self, index: u32) -> u32 {
        let before = self.spans.partition_point(|&(_, end)| end < index);
        let partial = self
            .spans
            .get(before)
            .map_or(0, |&(start, _)| index.saturating_sub(start));
        index - self.counts[before] - partial
    }

    fn restore(&self, index: u32) -> u32 {
        let before = self
            .spans
            .partition_point(|&(start, _)| self.project(start) <= index);
        index + self.counts[before]
    }

    fn bounds(&self, start: u32, end: u32) -> Option<(u32, u32)> {
        let first = self.restore(self.project(start));
        let after = self.project(end + 1);
        (first <= end && after > self.project(first)).then(|| (self.project(first), after - 1))
    }
}

fn data_edge(sheet: &Sheet, at: CellRef, direction: SelectionDirection) -> Option<CellRef> {
    let horizontal = matches!(
        direction,
        SelectionDirection::Left | SelectionDirection::Right
    );
    let reverse = matches!(direction, SelectionDirection::Up | SelectionDirection::Left);
    let coordinate = |cell: CellRef| if horizontal { cell.col } else { cell.row };
    let orient = |value: u32| if reverse { u32::MAX - value } else { value };
    if sheet.row_hidden(at.row) || sheet.col_hidden(at.col) {
        return None;
    }
    let axis = if horizontal {
        VisibleAxis::new(sheet.hidden_col_spans(0..=MAX_COLS - 1))
    } else {
        VisibleAxis::new(sheet.hidden_row_spans(0..=MAX_ROWS - 1))
    };
    let mut origin = axis.project(coordinate(at));
    for merge in &sheet.merges {
        if merge.contains(at) {
            if let Some((start, end)) = axis.bounds(coordinate(merge.start), coordinate(merge.end))
            {
                origin = if reverse { start } else { end };
            }
            break;
        }
    }
    let line = if horizontal {
        CellRange::new(CellRef::new(at.row, 0), CellRef::new(at.row, MAX_COLS - 1))
    } else {
        CellRange::new(CellRef::new(0, at.col), CellRef::new(MAX_ROWS - 1, at.col))
    };
    let mut intervals: Vec<(u32, u32)> = sheet
        .cells_in_range(line)
        .filter(|(at, cell)| filled(cell) && !sheet.row_hidden(at.row) && !sheet.col_hidden(at.col))
        .map(|(cell, _)| {
            let value = orient(axis.project(coordinate(cell)));
            (value, value)
        })
        .collect();
    for merge in &sheet.merges {
        if merge.start.row <= line.end.row
            && merge.end.row >= line.start.row
            && merge.start.col <= line.end.col
            && merge.end.col >= line.start.col
            && let Some((start, end)) = axis.bounds(coordinate(merge.start), coordinate(merge.end))
        {
            let start = orient(start);
            let end = orient(end);
            intervals.push((start.min(end), start.max(end)));
        }
    }
    intervals.sort_unstable();
    let mut runs: Vec<(u32, u32)> = Vec::new();
    for (start, end) in intervals {
        if let Some(last) = runs.last_mut()
            && start <= last.1.saturating_add(1)
        {
            last.1 = last.1.max(end);
        } else {
            runs.push((start, end));
        }
    }
    let origin = orient(origin);
    let target = runs.into_iter().find_map(|(start, end)| {
        if end <= origin {
            None
        } else if start <= origin {
            Some(end)
        } else {
            Some(start)
        }
    })?;
    let target = axis.restore(orient(target));
    Some(if horizontal {
        CellRef::new(at.row, target)
    } else {
        CellRef::new(target, at.col)
    })
}

fn merge_anchor(sheet: &Sheet, at: CellRef) -> CellRef {
    sheet
        .merges
        .iter()
        .find(|merge| merge.contains(at))
        .map_or(at, |merge| merge.start)
}

fn row_end(sheet: &Sheet, at: CellRef) -> Option<CellRef> {
    if sheet.row_hidden(at.row) || sheet.col_hidden(at.col) {
        return None;
    }
    let cols = VisibleAxis::new(sheet.hidden_col_spans(0..=MAX_COLS - 1));
    let line = CellRange::new(at, CellRef::new(at.row, MAX_COLS - 1));
    let merged = sheet
        .merges
        .iter()
        .filter(|merge| {
            merge.start.row <= at.row && merge.end.row >= at.row && merge.end.col >= at.col
        })
        .filter_map(|merge| cols.bounds(merge.start.col, merge.end.col))
        .map(|(_, end)| cols.restore(end));
    let last = sheet
        .cells_in_range(line)
        .filter(|(cell, value)| filled(value) && !sheet.col_hidden(cell.col))
        .map(|(cell, _)| cell.col)
        .chain(merged)
        .max()
        .or_else(|| {
            let visible = cols.project(sheet.used_range()?.end.col + 1);
            let col = cols.restore(visible.checked_sub(1)?);
            (col > at.col).then_some(col)
        })?;
    Some(CellRef::new(at.row, last))
}

fn current_region(sheet: &Sheet, at: CellRef) -> Option<SelectionRegion> {
    if sheet.row_hidden(at.row) || sheet.col_hidden(at.col) {
        return None;
    }
    let rows = VisibleAxis::new(sheet.hidden_row_spans(0..=MAX_ROWS - 1));
    let cols = VisibleAxis::new(sheet.hidden_col_spans(0..=MAX_COLS - 1));
    let mut cells: Vec<SelectionRegion> = sheet
        .iter_cells()
        .filter(|(at, cell)| filled(cell) && !sheet.row_hidden(at.row) && !sheet.col_hidden(at.col))
        .map(|(cell, _)| SelectionRegion {
            top: rows.project(cell.row),
            left: cols.project(cell.col),
            bottom: rows.project(cell.row),
            right: cols.project(cell.col),
        })
        .collect();
    cells.extend(sheet.merges.iter().filter_map(|merge| {
        let (top, bottom) = rows.bounds(merge.start.row, merge.end.row)?;
        let (left, right) = cols.bounds(merge.start.col, merge.end.col)?;
        Some(SelectionRegion {
            top,
            left,
            bottom,
            right,
        })
    }));
    let mut boundaries: [Vec<(u32, usize)>; 4] =
        std::array::from_fn(|_| Vec::with_capacity(cells.len()));
    for (index, cell) in cells.iter().enumerate() {
        for (boundary, value) in boundaries.iter_mut().zip([
            cell.top,
            u32::MAX - cell.bottom,
            cell.left,
            u32::MAX - cell.right,
        ]) {
            boundary.push((value, index));
        }
    }
    for boundary in &mut boundaries {
        boundary.sort_unstable();
    }
    let mut positions = [0; 4];
    let mut matches = vec![0; cells.len()];
    let mut region = SelectionRegion {
        top: rows.project(at.row),
        left: cols.project(at.col),
        bottom: rows.project(at.row),
        right: cols.project(at.col),
    };
    let mut found = false;
    loop {
        let previous = region;
        let limits = [
            region.bottom.saturating_add(1),
            u32::MAX - region.top.saturating_sub(1),
            region.right.saturating_add(1),
            u32::MAX - region.left.saturating_sub(1),
        ];
        for ((boundary, position), limit) in boundaries.iter().zip(&mut positions).zip(limits) {
            while let Some(&(value, index)) = boundary.get(*position) {
                if value > limit {
                    break;
                }
                *position += 1;
                matches[index] += 1;
                if matches[index] == 4 {
                    found = true;
                    let cell = cells[index];
                    region.top = region.top.min(cell.top);
                    region.left = region.left.min(cell.left);
                    region.bottom = region.bottom.max(cell.bottom);
                    region.right = region.right.max(cell.right);
                }
            }
        }
        if region == previous {
            return found.then(|| SelectionRegion {
                top: rows.restore(region.top),
                left: cols.restore(region.left),
                bottom: rows.restore(region.bottom),
                right: cols.restore(region.right),
            });
        }
    }
}

impl Workbook {
    pub fn selection_data(
        &self,
        sheet: SheetId,
        at: CellRef,
        direction: Option<SelectionDirection>,
    ) -> Result<SelectionData> {
        validate_cell_ref(at)?;
        let sheet = self.sheet(sheet)?;
        let mut data = match direction {
            Some(SelectionDirection::Merge) => SelectionData {
                merge: None,
                edge: None,
                region: None,
                used_range: None,
            },
            Some(SelectionDirection::UsedRange) => SelectionData {
                merge: None,
                edge: sheet
                    .used_range()
                    .map(|range| merge_anchor(sheet, range.end)),
                region: None,
                used_range: sheet.used_range().map(|range| SelectionRegion {
                    top: range.start.row,
                    left: range.start.col,
                    bottom: range.end.row,
                    right: range.end.col,
                }),
            },
            Some(SelectionDirection::RowEnd) => SelectionData {
                merge: None,
                edge: row_end(sheet, at).map(|edge| merge_anchor(sheet, edge)),
                region: None,
                used_range: None,
            },
            Some(direction) => SelectionData {
                merge: None,
                edge: data_edge(sheet, at, direction).map(|edge| merge_anchor(sheet, edge)),
                region: None,
                used_range: None,
            },
            None => SelectionData {
                merge: None,
                edge: None,
                region: current_region(sheet, at),
                used_range: None,
            },
        };
        data.merge = sheet
            .merges
            .iter()
            .find(|range| range.contains(data.edge.unwrap_or(at)))
            .map(|range| SelectionRegion {
                top: range.start.row,
                left: range.start.col,
                bottom: range.end.row,
                right: range.end.col,
            });
        Ok(data)
    }
}

#[cfg(test)]
mod tests;
