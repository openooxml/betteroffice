//! recalc cost of a running total: A-H numbers, I = SUM(A:H) of the row and
//! J = J{r-1} + I{r}. an edit to row 1 re-evaluates every J below it.

use std::time::{Duration, Instant};

use xlsx_calc::graph::DepGraph;
use xlsx_calc::recalc_after;
use xlsx_model::{Cell, CellProvider, CellRef, CellValue, Sheet, SheetId, Workbook};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn running_total(rows: u32) -> Workbook {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    let sheet = wb.sheet_mut(SheetId(0)).unwrap();
    let mut running = 0.0;
    for row in 0..rows {
        let n = row + 1;
        for col in 0..8 {
            let cell = Cell {
                value: number(1.0),
                ..Cell::default()
            };
            sheet.set_cell(CellRef::new(row, col), cell);
        }
        running += 8.0;
        let sum = format!("SUM(A{n}:H{n})");
        let total = if n == 1 {
            "I1".to_string()
        } else {
            format!("J{}+I{n}", n - 1)
        };
        for (col, formula, value) in [(8, sum, 8.0), (9, total, running)] {
            let cell = Cell {
                value: number(value),
                formula: Some(formula),
                style: None,
            };
            sheet.set_cell(CellRef::new(row, col), cell);
        }
    }
    wb
}

fn edit_first_row(rows: u32) -> Duration {
    let mut wb = running_total(rows);
    let mut graph = DepGraph::build(&wb);
    let a1 = CellRef::new(0, 0);
    wb.sheet_mut(SheetId(0)).unwrap().set_cell(
        a1,
        Cell {
            value: number(2.0),
            ..Cell::default()
        },
    );
    let start = Instant::now();
    let result = recalc_after(&mut wb, &mut graph, &[(SheetId(0), a1)], None);
    let elapsed = start.elapsed();
    assert_eq!(result.changed.len(), rows as usize + 1);
    let last = wb.value(SheetId(0), CellRef::new(rows - 1, 9));
    assert_eq!(last, number(8.0 * rows as f64 + 1.0));
    elapsed
}

#[test]
fn a_running_total_over_100k_rows_recalculates_within_budget() {
    let elapsed = edit_first_row(100_000);
    assert!(elapsed < Duration::from_secs(20), "recalc took {elapsed:?}");
}

#[test]
#[ignore = "scale report; run with --release --ignored --nocapture"]
fn running_total_scale_report() {
    let sizes = std::env::var("RECALC_SCALE_ROWS").unwrap_or_else(|_| "10000,30000,100000".into());
    let runs = std::env::var("RECALC_SCALE_RUNS").map_or(3, |runs| runs.parse().unwrap());
    for rows in sizes.split(',').map(|s| s.trim().parse::<u32>().unwrap()) {
        let runs: Vec<u128> = (0..runs)
            .map(|_| edit_first_row(rows).as_millis())
            .collect();
        eprintln!("RECALC rows={rows} ms={runs:?}");
    }
}
