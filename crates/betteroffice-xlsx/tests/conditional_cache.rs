use std::cell::Cell;

use betteroffice_xlsx::{
    CONDITIONAL_FRAME_BUDGET, CONDITIONAL_FRAME_MS, CONDITIONAL_RESULT_CAP, CalculationOptions,
    CellInput, CellRange, CellRef, CellState, CellValue, DifferentialFormat, DisplayList, DrawCmd,
    FrameClock, GridGeometry, Op, RenderOptions, SheetId, Viewport, Workbook,
};

const RED: &str = "#ffc7ce";

/// A workbook of `(name, sheetData rows, conditional formats)` sheets whose
/// rules paint `dxfId` 0, a red fill.
fn workbook(sheets: &[(&str, String, String)]) -> Workbook {
    let entries: String = sheets
        .iter()
        .enumerate()
        .map(|(index, (name, _, _))| {
            format!(
                r#"<sheet name="{name}" sheetId="{}" r:id="rId{}"/>"#,
                index + 1,
                index + 1
            )
        })
        .collect();
    let rels: String = (1..=sheets.len())
        .map(|index| format!(r#"<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>"#))
        .collect();
    let styles = sheets.len() + 1;
    let mut parts: Vec<(String, Vec<u8>)> = vec![
        ("xl/workbook.xml".into(), format!("<workbook><sheets>{entries}</sheets></workbook>").into_bytes()),
        ("xl/_rels/workbook.xml.rels".into(), format!(r#"<Relationships>{rels}<Relationship Id="rId{styles}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#).into_bytes()),
        ("xl/styles.xml".into(), br#"<styleSheet><cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>"#.to_vec()),
    ];
    for (index, (_, rows, formats)) in sheets.iter().enumerate() {
        parts.push((
            format!("xl/worksheets/sheet{}.xml", index + 1),
            format!("<worksheet><sheetData>{rows}</sheetData>{formats}</worksheet>").into_bytes(),
        ));
    }
    Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap()
}

fn sheet(rows: String, formats: String) -> Vec<(&'static str, String, String)> {
    vec![("Sheet1", rows, formats)]
}

/// `count` rows, each holding `cells(row)` for its 1-based row number.
fn rows(count: u32, cells: impl Fn(u32) -> String) -> String {
    (1..=count)
        .map(|row| format!(r#"<row r="{row}">{}</row>"#, cells(row)))
        .collect()
}

fn letter(col: u32) -> char {
    char::from(b'A' + col as u8)
}

/// Numbers in columns A onward, `value(row, col)` each.
fn numbers(cols: u32, value: impl Fn(u32, u32) -> u32) -> impl Fn(u32) -> String {
    move |row| {
        (0..cols)
            .map(|col| {
                format!(
                    r#"<c r="{}{row}"><v>{}</v></c>"#,
                    letter(col),
                    value(row, col)
                )
            })
            .collect()
    }
}

fn rule(sqref: &str, rules: &str) -> String {
    format!(r#"<conditionalFormatting sqref="{sqref}">{rules}</conditionalFormatting>"#)
}

fn greater_than(sqref: &str, value: &str) -> String {
    rule(
        sqref,
        &format!(
            r#"<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>{value}</formula></cfRule>"#
        ),
    )
}

fn expression(sqref: &str, priority: u32, formula: &str) -> String {
    rule(
        sqref,
        &format!(
            r#"<cfRule type="expression" dxfId="0" priority="{priority}"><formula>{formula}</formula></cfRule>"#
        ),
    )
}

fn view(y: f32, height: f32) -> Viewport {
    Viewport {
        x: 0.0,
        y,
        width: 640.0,
        height,
    }
}

/// A screen frame and how many visible cells its conditional formats left
/// undecided.
#[derive(Debug, PartialEq)]
struct Frame {
    list: DisplayList,
    conditional_pending: u32,
}

impl From<(DisplayList, u32)> for Frame {
    fn from((list, conditional_pending): (DisplayList, u32)) -> Self {
        Self {
            list,
            conditional_pending,
        }
    }
}

/// A screen frame or a one-shot render.
trait Painted {
    fn list(&self) -> &DisplayList;
}

impl Painted for DisplayList {
    fn list(&self) -> &DisplayList {
        self
    }
}

impl Painted for Frame {
    fn list(&self) -> &DisplayList {
        &self.list
    }
}

/// Where a display list paints the rules' red fill.
fn painted(list: &impl Painted) -> Vec<(u32, u32)> {
    list.list()
        .commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::FillRect { x, y, color, .. } if &**color == RED => {
                Some((x.to_bits(), y.to_bits()))
            }
            _ => None,
        })
        .collect()
}

fn frame(workbook: &Workbook, viewport: &Viewport) -> Frame {
    workbook
        .display_list_frame_for(SheetId(0), viewport)
        .unwrap()
        .into()
}

/// Frames of `viewport` with no input between them until one decides every
/// cell, and how many that took; fails past `limit` frames.
fn settle(workbook: &Workbook, viewport: &Viewport, limit: usize) -> (Frame, usize) {
    for count in 1..=limit {
        let list = frame(workbook, viewport);
        if list.conditional_pending == 0 {
            return (list, count);
        }
    }
    panic!("conditional formats still pending after {limit} frames");
}

fn edit(workbook: &mut Workbook, sheet: SheetId, edits: &[(u32, u32, &str)]) {
    let edits: Vec<CellInput> = edits
        .iter()
        .map(|&(row, col, input)| CellInput {
            cell: CellRef::new(row, col),
            input: input.to_owned(),
        })
        .collect();
    workbook
        .edit_cells(sheet, &edits, CalculationOptions::default())
        .unwrap();
}

/// A direct edit of one cell of the first sheet.
fn edit_one(workbook: &mut Workbook, row: u32, col: u32, input: &str) {
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(row, col),
            input,
            CalculationOptions::default(),
        )
        .unwrap();
}

#[test]
fn a_second_frame_of_the_same_view_evaluates_no_rule() {
    let workbook = workbook(&sheet(
        rows(40, numbers(8, |row, col| row * 8 + col)),
        greater_than("A1:H40", "100"),
    ));
    let viewport = view(0.0, 400.0);
    let first = frame(&workbook, &viewport);
    let before = workbook.conditional_stats();
    assert_eq!(first.conditional_pending, 0);
    assert!(!painted(&first).is_empty());
    assert!(before.evaluated_cells > 0);
    let second = frame(&workbook, &viewport);
    let after = workbook.conditional_stats();
    assert_eq!(second, first);
    assert_eq!(after.evaluated_cells, before.evaluated_cells);
    assert!(after.hits > before.hits);
}

#[test]
fn scrolling_back_reuses_the_cells_already_decided() {
    let workbook = workbook(&sheet(
        rows(2000, numbers(1, |row, _| row)),
        greater_than("A1:A2000", "1000"),
    ));
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let top = view(0.0, 400.0);
    let first = frame(&workbook, &top);
    frame(&workbook, &view(geometry.row_y(1000), 400.0));
    let scrolled = workbook.conditional_stats();
    let back = frame(&workbook, &top);
    assert_eq!(
        workbook.conditional_stats().evaluated_cells,
        scrolled.evaluated_cells
    );
    assert_eq!(back, first);
}

#[test]
fn editing_a_cell_repaints_the_rules_that_read_it() {
    let values = |a1: u32, c2: u32, e1: u32| {
        rows(10, move |row| {
            let a = if row == 1 { a1 } else { row };
            let c = if row == 2 { c2 } else { 0 };
            let e = if row == 1 { e1 } else { row };
            format!(
                r#"<c r="A{row}"><v>{a}</v></c><c r="C{row}"><v>{c}</v></c><c r="E{row}"><v>{e}</v></c>"#
            )
        })
    };
    let data = |a3: u32| {
        rows(3, move |row| {
            format!(
                r#"<c r="A{row}"><v>{}</v></c>"#,
                if row == 3 { a3 } else { 0 }
            )
        })
    };
    let formats = [
        greater_than("A1:A10", "8"),
        expression("B1:B10", 2, "C1&gt;0"),
        expression("D1:D3", 3, "Data!A1&gt;0"),
        rule(
            "E1:E10",
            r#"<cfRule type="top10" dxfId="0" priority="4" rank="1"/>"#,
        ),
    ]
    .concat();
    let book = |a1, c2, e1, a3| {
        workbook(&[
            ("Sheet1", values(a1, c2, e1), formats.clone()),
            ("Data", data(a3), String::new()),
        ])
    };
    let mut workbook = book(1, 0, 1, 0);
    let viewport = view(0.0, 300.0);
    let before = frame(&workbook, &viewport);
    assert_eq!(painted(&before).len(), 3);
    let built = workbook.conditional_stats().aggregate_builds;
    edit(&mut workbook, SheetId(1), &[(2, 0, "1")]);
    let data_only = frame(&workbook, &viewport);
    assert_eq!(data_only.conditional_pending, 0);
    assert_eq!(painted(&data_only).len(), 4);
    assert_eq!(
        painted(&data_only),
        painted(&book(1, 0, 1, 1).display_list(&viewport).unwrap())
    );
    edit(
        &mut workbook,
        SheetId(0),
        &[(0, 0, "100"), (1, 2, "1"), (0, 4, "50")],
    );
    let after = frame(&workbook, &viewport);
    assert_eq!(after.conditional_pending, 0);
    assert_eq!(painted(&after).len(), 6);
    assert_eq!(
        painted(&after),
        painted(&book(100, 1, 50, 1).display_list(&viewport).unwrap())
    );
    assert_eq!(workbook.conditional_stats().aggregate_builds, built + 2);
    for _ in 0..2 {
        assert!(
            workbook
                .undo(CalculationOptions::default())
                .unwrap()
                .applied
        );
    }
    assert_eq!(painted(&frame(&workbook, &viewport)), painted(&before));
    for _ in 0..2 {
        assert!(
            workbook
                .redo(CalculationOptions::default())
                .unwrap()
                .applied
        );
    }
    assert_eq!(painted(&frame(&workbook, &viewport)), painted(&after));
}

#[test]
fn a_recalculated_formula_cell_repaints_its_rule() {
    let mut workbook = workbook(&sheet(
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="C1"><f>A1*2</f><v>2</v></c></row>"#.into(),
        greater_than("C1", "5"),
    ));
    let viewport = view(0.0, 200.0);
    assert!(painted(&frame(&workbook, &viewport)).is_empty());
    assert!(painted(&frame(&workbook, &viewport)).is_empty());
    edit(&mut workbook, SheetId(0), &[(0, 0, "10")]);
    assert_eq!(painted(&frame(&workbook, &viewport)).len(), 1);
}

#[test]
fn adding_or_removing_a_rule_and_deleting_a_sheet_drop_the_cached_results() {
    let values = || rows(5, numbers(1, |row, _| row));
    let mut workbook = workbook(&[
        ("First", values(), greater_than("A1:A5", "3")),
        (
            "Second",
            values(),
            rule(
                "A1:A5",
                r#"<cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>2</formula></cfRule>"#,
            ),
        ),
    ]);
    let viewport = view(0.0, 200.0);
    let options = CalculationOptions::default();
    assert_eq!(painted(&frame(&workbook, &viewport)).len(), 2);
    workbook
        .apply_ops(
            vec![Op::AddConditionalFormat {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:A5").unwrap(),
                kind: "cellIs".into(),
                operator: Some("lessThan".into()),
                formulas: vec!["2.5".into()],
                format: DifferentialFormat {
                    fill: Some(RED.into()),
                    ..DifferentialFormat::default()
                },
            }],
            options,
        )
        .unwrap();
    assert_eq!(painted(&frame(&workbook, &viewport)).len(), 4);
    let added = workbook
        .conditional_formats(SheetId(0))
        .unwrap()
        .into_iter()
        .find(|rule| rule.formulas == ["2.5"])
        .unwrap()
        .id;
    workbook
        .apply_ops(
            vec![Op::RemoveConditionalFormat {
                sheet: SheetId(0),
                id: added,
            }],
            options,
        )
        .unwrap();
    assert_eq!(painted(&frame(&workbook, &viewport)).len(), 2);
    let second = painted(
        &workbook
            .display_list_frame_for(SheetId(1), &viewport)
            .unwrap()
            .0,
    );
    assert_eq!(second.len(), 1);
    workbook
        .apply_ops(vec![Op::RemoveSheet { index: 0 }], options)
        .unwrap();
    assert_eq!(painted(&frame(&workbook, &viewport)), second);
}

#[test]
fn an_exhausted_budget_leaves_cells_unpainted_and_pending() {
    let book = || {
        workbook(&sheet(
            rows(30, numbers(4, |row, col| row + col)),
            greater_than("A1:D30", "0"),
        ))
    };
    let workbook = book();
    let viewport = view(0.0, 800.0);
    let reference = painted(&book().display_list(&viewport).unwrap());
    assert_eq!(reference.len(), 120);
    workbook.set_conditional_budget(10);
    let first = frame(&workbook, &viewport);
    assert!(first.conditional_pending > 0);
    let partial = painted(&first);
    assert!(partial.len() < reference.len());
    assert!(partial.iter().all(|at| reference.contains(at)));
    assert_eq!(
        partial.len() + first.conditional_pending as usize,
        reference.len()
    );
    let (settled, frames) = settle(&workbook, &viewport, 20);
    assert!(frames > 1);
    assert_eq!(painted(&settled), reference);
}

#[test]
fn a_view_beyond_one_frame_budget_settles_without_input() {
    let book = || {
        workbook(&sheet(
            rows(2000, numbers(10, |row, _| row)),
            expression("A1:J2000", 1, "AND(MOD(A1,3)=0,SUM(A$1:A$2000)&gt;0)"),
        ))
    };
    let workbook = book();
    let viewport = view(0.0, 600.0);
    let first = frame(&workbook, &viewport);
    assert!(first.conditional_pending > 0);
    let (settled, frames) = settle(&workbook, &viewport, 10);
    assert!(frames > 1);
    let reference = book().display_list_for(SheetId(0), &viewport).unwrap();
    assert_eq!(settled.list.commands, reference.commands);
    assert!(!painted(&settled).is_empty());
}

#[test]
fn an_undecided_cell_paints_its_previous_result_under_the_same_rules_or_nothing() {
    let book = |zero: Option<u32>| {
        workbook(&sheet(
            rows(
                30,
                numbers(4, move |row, col| {
                    u32::from(col > 0 || zero != Some(row - 1))
                }),
            ),
            greater_than("A1:D30", "0"),
        ))
    };
    let mut workbook = book(None);
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let viewport = view(0.0, 800.0);
    workbook.set_conditional_budget(60);
    let (first, decided) = counted(&workbook, &viewport);
    assert!(first.conditional_pending > 0);
    let early = cells(&geometry, &first);
    assert_eq!(early.len() as u64, decided);
    let deepest = early
        .iter()
        .copied()
        .filter(|&(_, col)| col == 0)
        .max()
        .unwrap();
    workbook.set_conditional_budget(10);
    edit_one(&mut workbook, deepest.0, 0, "0");
    let reference = cells(
        &geometry,
        &book(Some(deepest.0)).display_list(&viewport).unwrap(),
    );
    assert!(!reference.contains(&deepest));
    let (next, decided) = counted(&workbook, &viewport);
    assert!(next.conditional_pending > 0);
    assert!(decided < early.len() as u64 / 2);
    let shown = cells(&geometry, &next);
    assert!(shown.iter().all(|at| early.contains(at)), "{shown:?}");
    assert!(shown.contains(&deepest), "{shown:?}");
    edit_one(&mut workbook, 0, 1, "2");
    let later = cells(&geometry, &frame(&workbook, &viewport));
    assert!(!later.contains(&deepest), "{later:?}");
    assert!(later.iter().all(|at| reference.contains(at)), "{later:?}");
    assert_eq!(
        cells(&geometry, &settle(&workbook, &viewport, 20).0),
        reference
    );
    workbook
        .apply_ops(
            vec![Op::AddConditionalFormat {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:D30").unwrap(),
                kind: "cellIs".into(),
                operator: Some("lessThan".into()),
                formulas: vec!["0.5".into()],
                format: DifferentialFormat {
                    fill: Some(RED.into()),
                    ..DifferentialFormat::default()
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    paints_only_decided(&workbook, &viewport, "added rule");
}

/// Thirty rows of `row + col` in columns A to D, except the `(row, col,
/// value)` cells of `changed`, on each named sheet under a rule that paints
/// every cell above 0.
fn positive(names: &[&'static str], changed: &[(u32, u32, u32)]) -> Workbook {
    let changed = changed.to_vec();
    let value = move |row: u32, col: u32| {
        changed
            .iter()
            .find(|&&(at, of, _)| at + 1 == row && of == col)
            .map_or(row + col, |&(_, _, value)| value)
    };
    let values = rows(30, numbers(4, value));
    let sheets: Vec<(&str, String, String)> = names
        .iter()
        .map(|&name| (name, values.clone(), greater_than("A1:D30", "0")))
        .collect();
    workbook(&sheets)
}

/// A frame of `viewport` and how many cells it decided.
fn counted(workbook: &Workbook, viewport: &Viewport) -> (Frame, u64) {
    let start = workbook.conditional_stats().evaluated_cells;
    let list = frame(workbook, viewport);
    (list, workbook.conditional_stats().evaluated_cells - start)
}

/// Asserts that a frame of `viewport` the budget stops paints no cell it left
/// undecided.
fn paints_only_decided(workbook: &Workbook, viewport: &Viewport, change: &str) {
    let (list, decided) = counted(workbook, viewport);
    assert!(list.conditional_pending > 0, "{change}");
    let shown = painted(&list).len() as u64;
    assert!(
        shown <= decided,
        "{change}: {shown} painted, {decided} decided"
    );
}

#[test]
fn after_an_edit_that_changes_no_result_undecided_cells_paint_their_previous_format() {
    let mut workbook = positive(&["Sheet1"], &[]);
    let viewport = view(0.0, 800.0);
    let before = painted(&frame(&workbook, &viewport));
    assert_eq!(before.len(), 120);
    workbook.set_conditional_budget(10);
    edit_one(&mut workbook, 29, 3, "7");
    let mut frames = 0;
    loop {
        let list = frame(&workbook, &viewport);
        frames += 1;
        assert_eq!(painted(&list), before, "frame {frames}");
        if list.conditional_pending == 0 {
            break;
        }
        assert!(frames < 20, "still pending after {frames} frames");
    }
    assert!(frames > 1);
}

#[test]
fn after_an_edit_that_changes_a_result_its_cell_paints_the_previous_format_until_decided() {
    let mut workbook = positive(&["Sheet1"], &[]);
    let viewport = view(0.0, 800.0);
    let before = painted(&frame(&workbook, &viewport));
    assert_eq!(before.len(), 120);
    workbook.set_conditional_budget(10);
    edit_one(&mut workbook, 29, 3, "0");
    let first = frame(&workbook, &viewport);
    assert!(first.conditional_pending > 0);
    assert_eq!(painted(&first), before);
    let (settled, frames) = settle(&workbook, &viewport, 20);
    assert!(frames > 1);
    let reference = positive(&["Sheet1"], &[(29, 3, 0)])
        .display_list(&viewport)
        .unwrap();
    assert_eq!(painted(&reference).len(), 119);
    assert_eq!(painted(&settled), painted(&reference));
    let one_shot = workbook.display_list_for(SheetId(0), &viewport).unwrap();
    assert_eq!(settled.list.commands, one_shot.commands);
}

#[test]
fn a_cell_with_no_previous_result_paints_unformatted_until_decided() {
    let book = || {
        workbook(&sheet(
            rows(2000, numbers(1, |row, _| row)),
            greater_than("A1:A2000", "0"),
        ))
    };
    let mut workbook = book();
    workbook.set_conditional_budget(10);
    let top = view(0.0, 400.0);
    let (opened, decided) = counted(&workbook, &top);
    assert!(opened.conditional_pending > 0);
    assert_eq!(painted(&opened).len() as u64, decided);
    settle(&workbook, &top, 20);
    edit_one(&mut workbook, 0, 0, "5");
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let far = view(geometry.row_y(1000), 400.0);
    let (scrolled, decided) = counted(&workbook, &far);
    assert!(scrolled.conditional_pending > 0);
    assert_eq!(painted(&scrolled).len() as u64, decided);
    let (settled, _) = settle(&workbook, &far, 20);
    let reference = book().display_list_for(SheetId(0), &far).unwrap();
    assert_eq!(painted(&settled), painted(&reference));
}

#[test]
fn a_one_shot_render_after_an_edit_shows_no_previous_result() {
    let mut workbook = positive(&["Sheet1"], &[]);
    let viewport = view(0.0, 800.0);
    assert_eq!(painted(&frame(&workbook, &viewport)).len(), 120);
    workbook.set_conditional_budget(1);
    edit_one(&mut workbook, 29, 3, "0");
    assert!(frame(&workbook, &viewport).conditional_pending > 0);
    let fresh = positive(&["Sheet1"], &[(29, 3, 0)]);
    let reference = painted(&fresh.display_list(&viewport).unwrap());
    assert_eq!(reference.len(), 119);
    for list in [
        workbook.display_list(&viewport).unwrap(),
        workbook.display_list_for(SheetId(0), &viewport).unwrap(),
        workbook
            .png_display_list_for(SheetId(0), &viewport)
            .unwrap()
            .0,
    ] {
        assert_eq!(painted(&list), reference);
    }
    let range = |book: &Workbook| {
        let (list, _, _) = book
            .range_png_display_list_for(
                SheetId(0),
                &RenderOptions {
                    range: Some(CellRange::parse_a1("A1:D30").unwrap()),
                    ..RenderOptions::default()
                },
            )
            .unwrap();
        list
    };
    assert_eq!(painted(&range(&workbook)), painted(&range(&fresh)));
}

/// A two-sheet workbook whose 120 visible cells are decided, limited to 10
/// cells a frame from here on.
fn decided_then_limited(viewport: &Viewport) -> Workbook {
    let workbook = positive(&["Sheet1", "Sheet2"], &[]);
    assert_eq!(painted(&frame(&workbook, viewport)).len(), 120);
    workbook.set_conditional_budget(10);
    workbook
}

fn paints_no_previous_result_after(op: Op) {
    let viewport = view(0.0, 800.0);
    let change = format!("{op:?}");
    let mut workbook = decided_then_limited(&viewport);
    if let Err(error) = workbook.apply_ops(vec![op], CalculationOptions::default()) {
        panic!("{change}: {error:?}");
    }
    paints_only_decided(&workbook, &viewport, &change);
}

#[test]
fn an_added_rule_paints_no_previous_result() {
    paints_no_previous_result_after(Op::AddConditionalFormat {
        sheet: SheetId(0),
        range: CellRange::parse_a1("A1:D30").unwrap(),
        kind: "cellIs".into(),
        operator: Some("lessThan".into()),
        formulas: vec!["2.5".into()],
        format: DifferentialFormat {
            fill: Some(RED.into()),
            ..DifferentialFormat::default()
        },
    });
}

#[test]
fn inserted_rows_paint_no_previous_result() {
    paints_no_previous_result_after(Op::InsertRows {
        sheet: SheetId(0),
        at: 0,
        count: 1,
    });
}

#[test]
fn deleted_rows_paint_no_previous_result() {
    paints_no_previous_result_after(Op::DeleteRows {
        sheet: SheetId(0),
        at: 0,
        count: 1,
    });
}

#[test]
fn inserted_columns_paint_no_previous_result() {
    paints_no_previous_result_after(Op::InsertCols {
        sheet: SheetId(0),
        at: 0,
        count: 1,
    });
}

#[test]
fn deleted_columns_paint_no_previous_result() {
    paints_no_previous_result_after(Op::DeleteCols {
        sheet: SheetId(0),
        at: 0,
        count: 1,
    });
}

#[test]
fn a_removed_sheet_paints_no_previous_result() {
    paints_no_previous_result_after(Op::RemoveSheet { index: 0 });
}

#[test]
fn a_moved_sheet_paints_no_previous_result() {
    paints_no_previous_result_after(Op::MoveSheet { from: 0, to: 1 });
}

#[cfg(feature = "unstable-track-moves")]
#[test]
fn moved_rows_paint_no_previous_result() {
    paints_no_previous_result_after(Op::MoveRows {
        sheet: SheetId(0),
        row: 0,
        count: 1,
        before: 5,
    });
}

#[cfg(feature = "unstable-track-moves")]
#[test]
fn moved_columns_paint_no_previous_result() {
    paints_no_previous_result_after(Op::MoveCols {
        sheet: SheetId(0),
        col: 0,
        count: 1,
        before: 3,
    });
}

#[test]
fn a_moved_range_paints_no_previous_result() {
    let viewport = view(0.0, 800.0);
    let mut workbook = positive_authored();
    edit_one(&mut workbook, 0, 5, "7");
    assert_eq!(painted(&frame(&workbook, &viewport)).len(), 120);
    workbook.set_conditional_budget(10);
    workbook
        .move_range(
            SheetId(0),
            CellRange::parse_a1("F1").unwrap(),
            CellRef::new(0, 7),
            CalculationOptions::default(),
        )
        .unwrap();
    paints_only_decided(&workbook, &viewport, "move range");
}

#[test]
fn a_refused_range_move_keeps_every_decided_result() {
    let viewport = view(0.0, 800.0);
    let mut workbook = decided_then_limited(&viewport);
    let refused = workbook.move_range(
        SheetId(0),
        CellRange::parse_a1("A1:A5").unwrap(),
        CellRef::new(0, 5),
        CalculationOptions::default(),
    );
    assert!(refused.is_err());
    let (list, decided) = counted(&workbook, &viewport);
    assert_eq!(decided, 0);
    assert_eq!(list.conditional_pending, 0);
    assert_eq!(painted(&list).len(), 120);
}

#[cfg(not(feature = "unstable-track-moves"))]
#[test]
fn a_refused_operation_keeps_every_decided_result() {
    let viewport = view(0.0, 800.0);
    let mut workbook = decided_then_limited(&viewport);
    let refused = workbook.apply_ops(
        vec![Op::MoveRows {
            sheet: SheetId(0),
            row: 0,
            count: 1,
            before: 5,
        }],
        CalculationOptions::default(),
    );
    assert!(refused.is_err());
    let (list, decided) = counted(&workbook, &viewport);
    assert_eq!(decided, 0);
    assert_eq!(list.conditional_pending, 0);
    assert_eq!(painted(&list).len(), 120);
}

#[test]
fn undo_and_redo_of_a_single_edit_paint_no_previous_result() {
    let mut workbook = positive(&["Sheet1"], &[]);
    let viewport = view(0.0, 800.0);
    let options = CalculationOptions::default();
    let before = painted(&frame(&workbook, &viewport));
    assert_eq!(before.len(), 120);
    workbook.set_conditional_budget(10);
    edit_one(&mut workbook, 29, 3, "0");
    let edited = painted(&settle(&workbook, &viewport, 20).0);
    assert_eq!(edited.len(), 119);
    assert!(workbook.undo(options).unwrap().applied);
    paints_only_decided(&workbook, &viewport, "undo");
    assert_eq!(painted(&settle(&workbook, &viewport, 20).0), before);
    assert!(workbook.redo(options).unwrap().applied);
    paints_only_decided(&workbook, &viewport, "redo");
    assert_eq!(painted(&settle(&workbook, &viewport, 20).0), edited);
}

/// The single-sheet workbook of [`positive`] with its 120 visible cells
/// decided, limited to 10 cells a frame from here on.
/// The sheet of `positive` with its rule added in this session, so the
/// engine lets a cell outside the rule's range move.
fn positive_authored() -> Workbook {
    let mut workbook = workbook(&sheet(
        rows(30, numbers(4, |row, col| row + col)),
        String::new(),
    ));
    workbook
        .apply_ops(
            vec![Op::AddConditionalFormat {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:D30").unwrap(),
                kind: "cellIs".into(),
                operator: Some("greaterThan".into()),
                formulas: vec!["0".into()],
                format: DifferentialFormat {
                    fill: Some(RED.into()),
                    ..DifferentialFormat::default()
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
}

fn decided_one_sheet(viewport: &Viewport) -> Workbook {
    let workbook = positive(&["Sheet1"], &[]);
    assert_eq!(painted(&frame(&workbook, viewport)).len(), 120);
    workbook.set_conditional_budget(10);
    workbook
}

/// Asserts that a frame of `viewport` the budget stops paints all 120 cells
/// decided before the change.
fn paints_previous_results(workbook: &Workbook, viewport: &Viewport, change: &str) {
    let list = frame(workbook, viewport);
    assert!(list.conditional_pending > 0, "{change}");
    assert_eq!(painted(&list).len(), 120, "{change}");
}

#[test]
fn edits_of_several_cells_keep_previous_results_until_decided() {
    let viewport = view(0.0, 800.0);
    let mut workbook = decided_one_sheet(&viewport);
    edit(&mut workbook, SheetId(0), &[(29, 3, "7"), (29, 2, "8")]);
    paints_previous_results(&workbook, &viewport, "edit cells");
}

#[test]
fn a_paste_keeps_previous_results_until_decided() {
    let viewport = view(0.0, 800.0);
    let mut workbook = decided_one_sheet(&viewport);
    let pasted = [(29, 3, "7"), (29, 2, "8")].map(|(row, col, input)| CellInput {
        cell: CellRef::new(row, col),
        input: input.to_owned(),
    });
    workbook
        .edit_pasted_cells(
            SheetId(0),
            &pasted,
            &[false, false],
            CalculationOptions::default(),
        )
        .unwrap();
    paints_previous_results(&workbook, &viewport, "paste");
}

#[test]
fn a_sort_written_as_cell_writes_paints_no_previous_result() {
    let viewport = view(0.0, 800.0);
    let mut workbook = decided_one_sheet(&viewport);
    let write = |row: u32, value: f64| Op::SetCell {
        sheet: SheetId(0),
        at: CellRef::new(row, 0),
        cell: CellState {
            value: CellValue::Number { value },
            ..CellState::default()
        },
    };
    workbook
        .apply_ops(
            vec![write(0, 2.0), write(1, 1.0)],
            CalculationOptions::default(),
        )
        .unwrap();
    paints_only_decided(&workbook, &viewport, "sort");
}

#[test]
fn undoing_a_range_move_paints_no_previous_result() {
    let viewport = view(0.0, 800.0);
    let mut workbook = positive_authored();
    edit_one(&mut workbook, 0, 5, "7");
    workbook
        .move_range(
            SheetId(0),
            CellRange::parse_a1("F1").unwrap(),
            CellRef::new(0, 7),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(painted(&frame(&workbook, &viewport)).len(), 120);
    workbook.set_conditional_budget(10);
    assert!(
        workbook
            .undo(CalculationOptions::default())
            .unwrap()
            .applied
    );
    paints_only_decided(&workbook, &viewport, "undo of a range move");
}

#[test]
fn a_direct_edit_or_a_recalculation_keeps_previous_results_until_decided() {
    let viewport = view(0.0, 800.0);
    for direct in [true, false] {
        let mut workbook = decided_one_sheet(&viewport);
        if direct {
            edit_one(&mut workbook, 29, 3, "7");
        } else {
            workbook.recalculate_all(CalculationOptions::default());
        }
        let list = frame(&workbook, &viewport);
        assert!(list.conditional_pending > 0, "{direct}");
        assert_eq!(painted(&list).len(), 120, "{direct}");
    }
}

#[test]
fn a_rule_reading_a_cell_whose_recalculation_was_refused_paints_nothing() {
    let book = |sum: &str, formats: String| {
        let a1 = format!(r#"<c r="A1"><f>{sum}</f><v>1</v></c>"#);
        let row = format!(r#"<row r="1">{a1}<c r="B1"><v>1</v></c></row>"#);
        workbook(&sheet(row, formats))
    };
    let viewport = view(0.0, 200.0);
    let average = rule(
        "A1:B1",
        r#"<cfRule type="aboveAverage" dxfId="0" priority="1"/>"#,
    );
    for (formats, cells) in [(greater_than("A1", "0"), 1), (average, 2)] {
        let mut limited = book("SUM($B$1:$Z$50000)", formats.clone());
        edit_one(&mut limited, 0, 1, "0");
        assert!(
            painted(&limited.display_list(&viewport).unwrap()).is_empty(),
            "{formats}"
        );
        assert_eq!(
            limited.conditional_stats().withheld_refused,
            cells,
            "{formats}"
        );
        let mut computed = book("SUM($B$1:$Z$2)+1", formats.clone());
        edit_one(&mut computed, 0, 1, "2");
        assert_eq!(
            painted(&computed.display_list(&viewport).unwrap()).len(),
            1,
            "{formats}"
        );
        assert_eq!(
            computed.conditional_stats().withheld_refused,
            0,
            "{formats}"
        );
    }
}

#[test]
fn a_range_move_that_changes_no_cell_keeps_previous_results() {
    let viewport = view(0.0, 800.0);
    let mut workbook = positive_authored();
    assert_eq!(painted(&frame(&workbook, &viewport)).len(), 120);
    workbook.set_conditional_budget(10);
    workbook
        .move_range(
            SheetId(0),
            CellRange::parse_a1("F1").unwrap(),
            CellRef::new(0, 7),
            CalculationOptions::default(),
        )
        .unwrap();
    edit_one(&mut workbook, 29, 3, "7");
    let list = frame(&workbook, &viewport);
    assert!(list.conditional_pending > 0);
    assert_eq!(painted(&list).len(), 120);
}

#[test]
fn only_the_latest_previous_generation_is_painted() {
    let mut workbook = positive(&["Sheet1"], &[]);
    let viewport = view(0.0, 800.0);
    assert_eq!(painted(&frame(&workbook, &viewport)).len(), 120);
    workbook.set_conditional_budget(10);
    edit_one(&mut workbook, 29, 3, "7");
    let (second, kept) = counted(&workbook, &viewport);
    assert!(second.conditional_pending > 0);
    assert_eq!(painted(&second).len(), 120);
    edit_one(&mut workbook, 29, 2, "7");
    let (third, decided) = counted(&workbook, &viewport);
    assert!(third.conditional_pending > 0);
    assert_eq!(painted(&third).len() as u64, decided);
    let stats = workbook.conditional_stats();
    assert!(stats.previous_cells as u64 <= kept, "{stats:?}");
}

#[test]
fn a_frame_whose_budget_a_count_table_spends_leaves_the_rest_pending() {
    let workbook = workbook(&sheet(
        rows(50_000, numbers(1, |row, _| row)),
        expression("B1:B20", 1, "COUNTIF($A$1:$A$50000,A1)&gt;0"),
    ));
    let viewport = view(0.0, 600.0);
    workbook.set_conditional_budget(1);
    let first = frame(&workbook, &viewport);
    assert_eq!(painted(&first).len(), 1);
    assert_eq!(first.conditional_pending, 19);
    assert_eq!(painted(&settle(&workbook, &viewport, 30).0).len(), 20);
}

#[test]
fn an_absolute_aggregate_is_computed_once_per_edit() {
    let book = |formula: &str| {
        workbook(&sheet(
            rows(5000, numbers(1, |row, _| 5001 - row)),
            expression("A1:Z5000", 1, formula),
        ))
    };
    let mut workbook = book("A1&gt;AVERAGE($A$1:$A$5000)");
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let top = view(0.0, 600.0);
    let start = workbook.conditional_stats().aggregate_builds;
    let first = frame(&workbook, &top);
    frame(&workbook, &view(geometry.row_y(2500), 600.0));
    assert_eq!(workbook.conditional_stats().aggregate_builds, start + 1);
    let unfolded = book("A1&gt;AVERAGE(INDIRECT(\"$A$1:$A$5000\"))");
    assert_eq!(
        painted(&first),
        painted(&unfolded.display_list(&top).unwrap())
    );
    assert!(!painted(&first).is_empty());
    edit(&mut workbook, SheetId(0), &[(0, 0, "1")]);
    let edited = frame(&workbook, &top);
    assert_eq!(workbook.conditional_stats().aggregate_builds, start + 2);
    assert_eq!(painted(&edited).len(), painted(&first).len() - 1);
}

/// The renders that decide every visible cell before they return.
fn one_shot() -> [fn(&Workbook, &Viewport) -> DisplayList; 3] {
    [
        |workbook, viewport| workbook.display_list(viewport).unwrap(),
        |workbook, viewport| workbook.display_list_for(SheetId(0), viewport).unwrap(),
        |workbook, viewport| {
            workbook
                .png_display_list_for(SheetId(0), viewport)
                .unwrap()
                .0
        },
    ]
}

#[test]
fn one_shot_renders_decide_every_cell() {
    let book = || {
        workbook(&sheet(
            rows(30, numbers(4, |row, col| row + col)),
            greater_than("A1:D30", "0"),
        ))
    };
    let viewport = view(0.0, 800.0);
    let reference = book().display_list(&viewport).unwrap();
    assert_eq!(painted(&reference).len(), 120);
    for entry in one_shot() {
        let workbook = book();
        workbook.set_conditional_budget(1);
        assert!(frame(&workbook, &viewport).conditional_pending > 0);
        assert_eq!(painted(&entry(&workbook, &viewport)), painted(&reference));
    }
    let fresh = book();
    fresh.set_conditional_budget(1);
    let (range, _, _) = fresh
        .range_png_display_list_for(
            SheetId(0),
            &RenderOptions {
                range: Some(CellRange::parse_a1("A1:D30").unwrap()),
                ..RenderOptions::default()
            },
        )
        .unwrap();
    assert_eq!(painted(&range).len(), 120);
}

#[test]
fn decided_and_previous_cells_stay_within_the_cap_across_a_long_scroll_and_an_edit() {
    assert_eq!(CONDITIONAL_RESULT_CAP, 262_144);
    assert_eq!(CONDITIONAL_FRAME_BUDGET, 200_000);
    let mut workbook = workbook(&sheet(
        r#"<row r="1"><c r="A1"><v>1</v></c></row>"#.into(),
        greater_than("A1:H50000", "0"),
    ));
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let end = geometry.row_y(50_000);
    let height = 10_000.0;
    let at = |y| Viewport {
        x: 0.0,
        y,
        width: 512.0,
        height,
    };
    let mut y = 0.0;
    while y < end {
        let list = frame(&workbook, &at(y));
        assert_eq!(list.conditional_pending, 0);
        assert!(workbook.conditional_stats().cached_cells <= CONDITIONAL_RESULT_CAP);
        y += height;
    }
    let stats = workbook.conditional_stats();
    assert!(
        stats.evaluated_cells > CONDITIONAL_RESULT_CAP as u64,
        "{stats:?}"
    );
    assert!(stats.cached_cells <= CONDITIONAL_RESULT_CAP);
    edit_one(&mut workbook, 0, 0, "2");
    let list = frame(&workbook, &at(0.0));
    assert_eq!(list.conditional_pending, 0);
    let stats = workbook.conditional_stats();
    assert!(stats.previous_cells > 0, "{stats:?}");
    assert!(stats.previous_cells <= CONDITIONAL_RESULT_CAP, "{stats:?}");
    assert!(stats.cached_cells <= CONDITIONAL_RESULT_CAP, "{stats:?}");
}

#[test]
fn a_rule_formula_is_parsed_once_per_epoch_across_its_cells() {
    let book = |formula: &str| {
        workbook(&sheet(
            rows(200, numbers(4, |row, col| row + col)),
            expression("A1:D200", 1, formula),
        ))
    };
    let mut workbook = book("A1+$B$1&gt;50");
    let viewport = view(0.0, 2000.0);
    let start = workbook.conditional_stats();
    let first = frame(&workbook, &viewport);
    let stats = workbook.conditional_stats();
    assert_eq!(first.conditional_pending, 0);
    assert!(
        stats.evaluated_cells - start.evaluated_cells >= 200,
        "{stats:?}"
    );
    assert!(
        stats.formula_parses - start.formula_parses <= 2,
        "{stats:?}"
    );
    let reference = workbook_with_threshold("48");
    assert_eq!(
        painted(&first),
        painted(&reference.display_list(&viewport).unwrap())
    );
    assert!(!painted(&first).is_empty());
    edit(&mut workbook, SheetId(0), &[(150, 3, "0")]);
    let edited = frame(&workbook, &viewport);
    let after = workbook.conditional_stats();
    assert_eq!(edited.conditional_pending, 0);
    assert!(
        after.formula_parses - stats.formula_parses <= 2,
        "{after:?}"
    );
    assert_eq!(painted(&edited), painted(&first));
}

fn whole_reference_book(formula: &str) -> Workbook {
    workbook(&sheet(
        rows(200, numbers(4, |row, col| row + col)),
        expression("A1:D200", 1, formula),
    ))
}

#[test]
fn an_absolute_whole_column_rule_is_parsed_once_per_epoch() {
    let workbook = whole_reference_book("A1*1000&gt;SUM($D:$D)+SUM($1:$1)");
    let viewport = view(0.0, 2000.0);
    let start = workbook.conditional_stats();
    let list = frame(&workbook, &viewport);
    let stats = workbook.conditional_stats();
    assert_eq!(list.conditional_pending, 0);
    assert!(
        stats.evaluated_cells - start.evaluated_cells >= 200,
        "{stats:?}"
    );
    assert!(
        stats.formula_parses - start.formula_parses <= 2,
        "{stats:?}"
    );
    assert_eq!(stats.slow_rules, start.slow_rules, "{stats:?}");
    assert!(!painted(&list).is_empty());
}

#[test]
fn a_relative_whole_row_or_column_rule_parses_each_cell() {
    for formula in [
        "A1*1000&gt;SUM(D:D)",
        "A1*1000&gt;SUM($D:E)",
        "A1*1000&gt;SUM(D:$E)",
        "A1&gt;SUM(1:1)",
    ] {
        let workbook = whole_reference_book(formula);
        let viewport = view(0.0, 2000.0);
        let start = workbook.conditional_stats();
        let (list, _) = settle(&workbook, &viewport, 64);
        let stats = workbook.conditional_stats();
        let cells = stats.evaluated_cells - start.evaluated_cells;
        assert_eq!(list.conditional_pending, 0, "{formula}");
        assert!(cells >= 200, "{formula}: {stats:?}");
        assert!(
            stats.formula_parses - start.formula_parses >= cells,
            "{formula}: {stats:?}"
        );
        assert_eq!(
            stats.slow_rules,
            start.slow_rules + 1,
            "{formula}: {stats:?}"
        );
    }
}

/// The `(row, col)` of each cell a display list from the sheet's top-left
/// corner paints red, sorted.
fn cells(geometry: &GridGeometry, list: &impl Painted) -> Vec<(u32, u32)> {
    let at = |offset: f32, start: fn(&GridGeometry, u32) -> f32| {
        (0..1000)
            .rfind(|&index| start(geometry, index) <= offset + 0.5)
            .unwrap()
    };
    let mut cells: Vec<(u32, u32)> = painted(list)
        .into_iter()
        .map(|(x, y)| {
            (
                at(f32::from_bits(y), GridGeometry::row_y),
                at(f32::from_bits(x), GridGeometry::col_x),
            )
        })
        .collect();
    cells.sort_unstable();
    cells
}

#[test]
fn one_shot_renders_decide_every_cell_of_rules_on_the_general_path() {
    let formats = [
        expression("A1:D200", 1, "A1*1000&lt;SUM($D:E)"),
        expression("A1:D200", 2, "AND(MOD(A1,7)=0,SUM(D:D)=0)"),
        expression("A1:D200", 3, "MOD(SUM(1:1),5)=0"),
    ]
    .concat();
    let book = || {
        workbook(&sheet(
            rows(200, numbers(4, |row, col| row + col)),
            formats.clone(),
        ))
    };
    let expected: Vec<(u32, u32)> = (1..=200u32)
        .flat_map(|row| (0..4u32).map(move |col| (row, col)))
        .filter(|&(row, col)| {
            let value = row + col;
            value <= 20 || (col > 0 && value.is_multiple_of(7)) || row % 5 == 1
        })
        .map(|(row, col)| (row - 1, col))
        .collect();
    assert!(expected.contains(&(19, 0)));
    assert!(expected.contains(&(26, 1)));
    assert!(!expected.contains(&(21, 0)));
    assert!(!expected.contains(&(27, 0)));
    let workbook = book();
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let viewport = view(0.0, geometry.row_y(200));
    let start = workbook.conditional_stats();
    assert!(frame(&workbook, &viewport).conditional_pending > 0);
    assert_eq!(
        workbook.conditional_stats().slow_rules,
        start.slow_rules + 3
    );
    for entry in one_shot() {
        let workbook = book();
        assert!(frame(&workbook, &viewport).conditional_pending > 0);
        assert_eq!(cells(&geometry, &entry(&workbook, &viewport)), expected);
    }
    let (range, _, _) = book()
        .range_png_display_list_for(
            SheetId(0),
            &RenderOptions {
                range: Some(CellRange::parse_a1("A1:D200").unwrap()),
                ..RenderOptions::default()
            },
        )
        .unwrap();
    assert_eq!(cells(&geometry, &range), expected);
    let (settled, frames) = settle(&book(), &viewport, 64);
    assert!(frames > 1);
    assert_eq!(cells(&geometry, &settled), expected);
}

/// The `A1+$B$1>50` workbook of the parse test as a cellIs rule over the
/// same values: `$B$1` holds 2.
fn workbook_with_threshold(value: &str) -> Workbook {
    workbook(&sheet(
        rows(200, numbers(4, |row, col| row + col)),
        greater_than("A1:D200", value),
    ))
}

#[test]
fn a_countif_of_each_cell_over_a_fixed_range_reads_the_range_once() {
    let value = |row: u32| if row <= 20 { row * 1000 } else { row % 50 };
    let book = |formula: &str| {
        workbook(&sheet(
            rows(2000, numbers(1, move |row, _| value(row))),
            expression("A1:A2000", 1, formula),
        ))
    };
    let workbook = book("COUNTIF($A$1:$A$2000,A1)&gt;1");
    let viewport = view(0.0, 600.0);
    let start = workbook.conditional_stats();
    let list = frame(&workbook, &viewport);
    let stats = workbook.conditional_stats();
    let cells = stats.evaluated_cells - start.evaluated_cells;
    assert_eq!(list.conditional_pending, 0);
    let reference = book("COUNTIF(INDIRECT(\"$A$1:$A$2000\"),A1)&gt;1");
    let expected = painted(&reference.display_list(&viewport).unwrap());
    assert_eq!(painted(&list), expected);
    assert!(!expected.is_empty() && (expected.len() as u64) < cells);
    frame(&workbook, &viewport);
    let scanned = workbook.conditional_stats().units;
    assert!(stats.units - scanned <= 2_000 + 8 * cells, "{stats:?}");
}

/// A clock that moves a quarter of a frame each time it is read.
struct Ticks(Cell<f64>);

impl FrameClock for Ticks {
    fn now_ms(&self) -> f64 {
        let now = self.0.get() + CONDITIONAL_FRAME_MS / 4.0;
        self.0.set(now);
        now
    }
}

#[test]
fn a_timed_frame_stops_when_its_clock_runs_out() {
    let book = || {
        workbook(&sheet(
            rows(30, numbers(4, |row, col| row + col)),
            greater_than("A1:D30", "0"),
        ))
    };
    let workbook = book();
    let viewport = view(0.0, 800.0);
    let reference = painted(&book().display_list(&viewport).unwrap());
    assert_eq!(reference.len(), 120);
    assert_eq!(frame(&book(), &viewport).conditional_pending, 0);
    let clock = Ticks(Cell::new(0.0));
    let timed = || -> Frame {
        workbook
            .display_list_frame_timed(SheetId(0), &viewport, &clock)
            .unwrap()
            .into()
    };
    let first = timed();
    assert!(first.conditional_pending > 0);
    let partial = painted(&first);
    assert!(partial.iter().all(|at| reference.contains(at)));
    assert_eq!(
        partial.len() + first.conditional_pending as usize,
        reference.len()
    );
    let mut frames = 1;
    let mut last = first;
    while last.conditional_pending > 0 {
        assert!(frames < 40, "still pending after {frames} timed frames");
        last = timed();
        frames += 1;
    }
    assert!(frames > 1);
    assert_eq!(painted(&last), reference);
}

/// Asserts that a frame the budget stops leaves nothing pending, and that it
/// and a one-shot render of `workbook` paint nothing.
fn settles_unpainted(workbook: &Workbook) {
    let viewport = view(0.0, 800.0);
    workbook.set_conditional_budget(1);
    let stopped = frame(workbook, &viewport);
    assert_eq!(stopped.conditional_pending, 0);
    assert!(painted(&stopped).is_empty());
    assert!(painted(&workbook.display_list(&viewport).unwrap()).is_empty());
}

#[test]
fn cells_whose_first_rule_is_unsupported_are_not_pending() {
    settles_unpainted(&workbook(&sheet(
        rows(30, numbers(4, |row, col| row + col)),
        rule(
            "A1:D30",
            r#"<cfRule type="timePeriod" dxfId="0" priority="1" timePeriod="today"><formula>FLOOR(A1,1)=TODAY()</formula></cfRule><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>0</formula></cfRule>"#,
        ),
    )));
}

#[test]
fn a_sheet_whose_only_rule_reads_random_numbers_has_nothing_pending() {
    settles_unpainted(&workbook(&sheet(
        rows(30, numbers(4, |row, col| row + col)),
        expression("A1:D30", 1, "RAND()&gt;=0"),
    )));
}

#[test]
fn range_statistics_too_large_for_the_units_left_wait_for_a_frame_that_starts_with_them() {
    let workbook = workbook(&sheet(
        rows(100_000, numbers(2, |row, _| row)),
        [
            greater_than("A1:A30", "0"),
            rule(
                "B1:B100000",
                r#"<cfRule type="top10" dxfId="0" priority="2" rank="5"/>"#,
            ),
        ]
        .concat(),
    ));
    let viewport = view(0.0, 400.0);
    workbook.set_conditional_budget(1_000);
    let start = workbook.conditional_stats().aggregate_builds;
    let first = frame(&workbook, &viewport);
    assert!(first.conditional_pending > 0);
    assert_eq!(workbook.conditional_stats().aggregate_builds, start);
    frame(&workbook, &viewport);
    assert_eq!(workbook.conditional_stats().aggregate_builds, start + 1);
}

#[test]
fn the_frame_that_compiles_the_rules_counts_one_unit_per_rule_formula() {
    let workbook = workbook(&sheet(
        rows(10, numbers(1, |row, _| row)),
        [
            expression("A500:A600", 1, "A500&gt;0"),
            expression("B500:B600", 2, "B500&gt;1"),
            expression("C500:C600", 3, "C500&gt;2"),
        ]
        .concat(),
    ));
    let viewport = view(0.0, 300.0);
    frame(&workbook, &viewport);
    let compiled = workbook.conditional_stats().units;
    frame(&workbook, &viewport);
    let scanned = workbook.conditional_stats().units;
    assert_eq!(compiled.checked_sub(scanned), Some(3));
}
