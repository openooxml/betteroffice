use xlsx_model::{Cell, CellRange, CellRef, CellValue, ErrorValue, Workbook};

use super::wire::{ChunkKind, Reader, Writer, frame_records};
use super::{SnapshotBudget, SnapshotError, SnapshotResult};

#[derive(Clone, Copy, Default)]
pub(super) struct CellCursor {
    pub(super) sheet: usize,
    pub(super) after: Option<(u32, u32)>,
}

impl CellCursor {
    pub(super) fn next(
        &mut self,
        model: &Workbook,
        budget: SnapshotBudget,
        ordinal: u64,
    ) -> SnapshotResult<Option<Vec<u8>>> {
        while let Some(sheet) = model.sheets.get(self.sheet) {
            let start = match self.after {
                None => Some((0, 0)),
                Some((row, col)) => col
                    .checked_add(1)
                    .map(|col| (row, col))
                    .or_else(|| row.checked_add(1).map(|row| (row, 0))),
            };
            let Some((row, col)) = start else {
                self.sheet += 1;
                self.after = None;
                continue;
            };
            let range = CellRange {
                start: CellRef::new(row, col),
                end: CellRef::new(row, u32::MAX),
            };
            let later_rows = row.checked_add(1).into_iter().flat_map(|next_row| {
                sheet.cells_in_range(CellRange {
                    start: CellRef::new(next_row, 0),
                    end: CellRef::new(u32::MAX, u32::MAX),
                })
            });
            let cells = sheet.cells_in_range(range).chain(later_rows);
            let base = self.after.unwrap_or((0, 0));
            let mut previous = base;
            let mut records = Writer::new();
            let mut count = 0;
            let mut lengths = Vec::new();
            for (at, cell) in cells {
                let mut record = Writer::new();
                record.var_u32(at.row - previous.0);
                record.var_u32(if at.row == previous.0 {
                    at.col - previous.1
                } else {
                    at.col
                });
                write_cell(&mut record, cell);
                let mut header = Writer::new();
                header.var_usize(self.sheet);
                header.var_u32(base.0);
                header.var_u32(base.1);
                header.var_usize(count + 1);
                lengths.push(record.len() + if count == 0 { header.len() } else { 0 });
                let length = frame_records(ChunkKind::Cells, ordinal, &[], &lengths).len()
                    + header.len()
                    + records.len()
                    + record.len();
                if count != 0 && length > budget.max_bytes().min(1024) {
                    lengths.pop();
                    break;
                }
                records.raw(&record.into_bytes());
                previous = (at.row, at.col);
                self.after = Some(previous);
                count += 1;
                if count == budget.max_records().min(64) {
                    break;
                }
            }
            if count != 0 {
                let mut payload = Writer::new();
                payload.var_usize(self.sheet);
                payload.var_u32(base.0);
                payload.var_u32(base.1);
                payload.var_usize(count);
                payload.raw(&records.into_bytes());
                return Ok(Some(frame_records(
                    ChunkKind::Cells,
                    ordinal,
                    &payload.into_bytes(),
                    &lengths,
                )));
            }
            self.sheet += 1;
            self.after = None;
        }
        Ok(None)
    }
}

pub(super) fn write_cell(w: &mut Writer, cell: &Cell) {
    let Cell {
        value,
        formula,
        style,
    } = cell;
    match value {
        CellValue::Empty => w.u8(0),
        CellValue::Number { value } => {
            w.u8(1);
            w.f64(*value);
        }
        CellValue::Text { value } => {
            w.u8(2);
            w.str(value);
        }
        CellValue::Bool { value } => {
            w.u8(3);
            w.bool(*value);
        }
        CellValue::Error { value } => {
            w.u8(4);
            w.u8(match value {
                ErrorValue::Div0 => 0,
                ErrorValue::NA => 1,
                ErrorValue::Name => 2,
                ErrorValue::Null => 3,
                ErrorValue::Num => 4,
                ErrorValue::Ref => 5,
                ErrorValue::Value => 6,
                ErrorValue::Spill => 7,
                ErrorValue::Calc => 8,
            });
        }
    }
    w.option(formula.as_deref(), |w, formula| w.str(formula));
    w.option(*style, |w, style| w.var_u32(style));
}

pub(super) fn read_cell(r: &mut Reader<'_>) -> SnapshotResult<Cell> {
    let value = match r.u8()? {
        0 => CellValue::Empty,
        1 => CellValue::Number { value: r.f64()? },
        2 => CellValue::Text {
            value: r.str()?.to_owned(),
        },
        3 => CellValue::Bool { value: r.bool()? },
        4 => CellValue::Error {
            value: match r.u8()? {
                0 => ErrorValue::Div0,
                1 => ErrorValue::NA,
                2 => ErrorValue::Name,
                3 => ErrorValue::Null,
                4 => ErrorValue::Num,
                5 => ErrorValue::Ref,
                6 => ErrorValue::Value,
                7 => ErrorValue::Spill,
                8 => ErrorValue::Calc,
                _ => return Err(SnapshotError::new("invalid snapshot cell error")),
            },
        },
        _ => return Err(SnapshotError::new("invalid snapshot cell value")),
    };
    Ok(Cell {
        value,
        formula: r.option(|r| Ok(r.str()?.to_owned()))?,
        style: r.option(Reader::var_u32)?,
    })
}

pub(super) fn write_ref(w: &mut Writer, at: &CellRef) {
    let CellRef {
        row,
        col,
        abs_row,
        abs_col,
    } = at;
    w.var_u32(*row);
    w.var_u32(*col);
    w.bool(*abs_row);
    w.bool(*abs_col);
}

pub(super) fn read_ref(r: &mut Reader<'_>) -> SnapshotResult<CellRef> {
    Ok(CellRef {
        row: r.var_u32()?,
        col: r.var_u32()?,
        abs_row: r.bool()?,
        abs_col: r.bool()?,
    })
}

pub(super) fn write_range(w: &mut Writer, range: &CellRange) {
    let CellRange { start, end } = range;
    write_ref(w, start);
    write_ref(w, end);
}

pub(super) fn read_range(r: &mut Reader<'_>) -> SnapshotResult<CellRange> {
    Ok(CellRange {
        start: read_ref(r)?,
        end: read_ref(r)?,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cells_roundtrip_all_tags_float_bits_and_options() {
        let mut values = vec![
            CellValue::Empty,
            CellValue::Text {
                value: String::new(),
            },
            CellValue::Text {
                value: "é水\0".into(),
            },
            CellValue::Bool { value: false },
            CellValue::Bool { value: true },
        ];
        for bits in [
            0,
            0x8000_0000_0000_0000,
            0x7ff0_0000_0000_0000,
            0xfff0_0000_0000_0000,
            0x7ff8_0000_0000_0042,
            0x7ff0_0000_0000_0001,
            1,
        ] {
            values.push(CellValue::Number {
                value: f64::from_bits(bits),
            });
        }
        for value in [
            ErrorValue::Div0,
            ErrorValue::NA,
            ErrorValue::Name,
            ErrorValue::Null,
            ErrorValue::Num,
            ErrorValue::Ref,
            ErrorValue::Value,
            ErrorValue::Spill,
            ErrorValue::Calc,
        ] {
            values.push(CellValue::Error { value });
        }
        for value in values {
            for formula in [None, Some(String::new()), Some("A1+1".into())] {
                for style in [None, Some(0), Some(u32::MAX)] {
                    let cell = Cell {
                        value: value.clone(),
                        formula: formula.clone(),
                        style,
                    };
                    let mut w = Writer::new();
                    write_cell(&mut w, &cell);
                    let bytes = w.into_bytes();
                    let mut r = Reader::new(&bytes);
                    let decoded = read_cell(&mut r).unwrap();
                    r.finish().unwrap();
                    assert_eq!(decoded.formula, cell.formula);
                    assert_eq!(decoded.style, cell.style);
                    match (&cell.value, &decoded.value) {
                        (CellValue::Number { value: a }, CellValue::Number { value: b }) => {
                            assert_eq!(a.to_bits(), b.to_bits());
                        }
                        (a, b) => assert_eq!(a, b),
                    }
                    let mut w = Writer::new();
                    write_cell(&mut w, &decoded);
                    assert_eq!(w.into_bytes(), bytes);
                }
            }
        }
        assert!(read_cell(&mut Reader::new(&[5])).is_err());
        assert!(read_cell(&mut Reader::new(&[4, 9])).is_err());
    }
}
