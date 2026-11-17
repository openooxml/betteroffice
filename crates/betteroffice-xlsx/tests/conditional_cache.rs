use betteroffice_xlsx::{
    CONDITIONAL_FRAME_BUDGET, CONDITIONAL_RESULT_CAP, CalculationOptions, CellInput, CellRange,
    CellRef, DifferentialFormat, DisplayList, DrawCmd, GridGeometry, Op, RenderOptions, SheetId,
    Viewport, Workbook,
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

/// Where a display list paints the rules' red fill.
fn painted(list: &DisplayList) -> Vec<(u32, u32)> {
    list.commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::FillRect { x, y, color, .. } if &**color == RED => {
                Some((x.to_bits(), y.to_bits()))
            }
            _ => None,
        })
        .collect()
}

fn frame(workbook: &Workbook, viewport: &Viewport) -> DisplayList {
    workbook
        .display_list_frame_for(SheetId(0), viewport)
        .unwrap()
}

/// Frames of `viewport` with no input between them until one decides every
/// cell, and how many that took; fails past `limit` frames.
fn settle(workbook: &Workbook, viewport: &Viewport, limit: usize) -> (DisplayList, usize) {
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
    edit(
        &mut workbook,
        SheetId(0),
        &[(0, 0, "100"), (1, 2, "1"), (0, 4, "50")],
    );
    edit(&mut workbook, SheetId(1), &[(2, 0, "1")]);
    let after = frame(&workbook, &viewport);
    assert_eq!(after.conditional_pending, 0);
    assert_eq!(painted(&after).len(), 6);
    assert_eq!(
        painted(&after),
        painted(&book(100, 1, 50, 1).display_list(&viewport).unwrap())
    );
    assert_eq!(workbook.conditional_stats().aggregate_builds, built + 1);
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
            .unwrap(),
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
    assert_eq!(settled.commands, reference.commands);
    assert!(!painted(&settled).is_empty());
}

#[test]
fn an_edit_during_a_partial_view_never_paints_the_old_fill() {
    let book = |a: u32| {
        workbook(&sheet(
            rows(30, numbers(4, move |_, col| if col == 0 { a } else { 1 })),
            greater_than("A1:D30", "0"),
        ))
    };
    let mut workbook = book(1);
    let viewport = view(0.0, 800.0);
    workbook.set_conditional_budget(10);
    assert!(frame(&workbook, &viewport).conditional_pending > 0);
    let zeros: Vec<(u32, u32, &str)> = (0..30).map(|row| (row, 0, "0")).collect();
    edit(&mut workbook, SheetId(0), &zeros);
    let reference = painted(&book(0).display_list(&viewport).unwrap());
    let next = frame(&workbook, &viewport);
    assert!(next.conditional_pending > 0);
    assert!(painted(&next).iter().all(|at| reference.contains(at)));
    assert_eq!(painted(&settle(&workbook, &viewport, 20).0), reference);
}

#[test]
fn a_formula_refused_budget_is_pending_not_false() {
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

#[test]
fn one_shot_renders_decide_every_cell() {
    let book = || {
        workbook(&sheet(
            rows(30, numbers(4, |row, col| row + col)),
            greater_than("A1:D30", "0"),
        ))
    };
    let workbook = book();
    let viewport = view(0.0, 800.0);
    let reference = book().display_list(&viewport).unwrap();
    assert_eq!(painted(&reference).len(), 120);
    workbook.set_conditional_budget(1);
    assert!(frame(&workbook, &viewport).conditional_pending > 0);
    for list in [
        workbook.display_list(&viewport).unwrap(),
        workbook.display_list_for(SheetId(0), &viewport).unwrap(),
        workbook
            .png_display_list_for(SheetId(0), &viewport)
            .unwrap()
            .0,
    ] {
        assert_eq!(list.conditional_pending, 0);
        assert_eq!(painted(&list), painted(&reference));
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
    assert_eq!(range.conditional_pending, 0);
    assert_eq!(painted(&range).len(), 120);
}

#[test]
fn decided_cells_stay_within_the_cap_across_a_long_scroll() {
    assert_eq!(CONDITIONAL_RESULT_CAP, 262_144);
    assert_eq!(CONDITIONAL_FRAME_BUDGET, 200_000);
    let workbook = workbook(&sheet(
        r#"<row r="1"><c r="A1"><v>1</v></c></row>"#.into(),
        greater_than("A1:H50000", "0"),
    ));
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let end = geometry.row_y(50_000);
    let height = 10_000.0;
    let mut y = 0.0;
    while y < end {
        let list = workbook
            .display_list_frame_for(
                SheetId(0),
                &Viewport {
                    x: 0.0,
                    y,
                    width: 512.0,
                    height,
                },
            )
            .unwrap();
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
}
