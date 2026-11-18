use std::collections::BTreeMap;

use betteroffice_xlsx::{
    CalculationOptions, CellInput, CellRef, CellValue, DisplayList, DrawCmd, GridGeometry, SheetId,
    Viewport, Workbook,
};

const RED: &str = "#ffc7ce";
const FIRST: f64 = 46_000.25;
const SECOND: f64 = 46_003.75;
const NAMES: &str = r#"<definedNames><definedName name="highlight">Sheet1!$C$1</definedName><definedName name="Stamp">TODAY()</definedName></definedNames>"#;

/// A one-sheet package whose column A holds the days 45999 to 46006 in rows
/// 1 to 8, plus `cells` as `(reference, text or =formula)`, the workbook's
/// `names` and conditional `formats` painting red.
fn package(cells: &[(&str, &str)], names: &str, formats: &str) -> Vec<u8> {
    let mut rows = BTreeMap::<u32, String>::new();
    for row in 1..=8u32 {
        let day = 45_998 + row;
        rows.insert(row, format!(r#"<c r="A{row}"><v>{day}</v></c>"#));
    }
    for (at, value) in cells {
        let row = at
            .trim_start_matches(|c: char| c.is_ascii_alphabetic())
            .parse()
            .unwrap();
        let cell = match value.strip_prefix('=') {
            Some(formula) => format!(r#"<c r="{at}"><f>{formula}</f></c>"#),
            None => format!(r#"<c r="{at}" t="inlineStr"><is><t>{value}</t></is></c>"#),
        };
        rows.entry(row).or_default().push_str(&cell);
    }
    let rows: String = rows
        .iter()
        .map(|(row, cells)| format!(r#"<row r="{row}">{cells}</row>"#))
        .collect();
    let parts: Vec<(String, Vec<u8>)> = vec![
        ("xl/workbook.xml".into(), format!(r#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>{names}</workbook>"#).into_bytes()),
        ("xl/_rels/workbook.xml.rels".into(), br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_vec()),
        ("xl/worksheets/sheet1.xml".into(), format!("<worksheet><sheetData>{rows}</sheetData>{formats}</worksheet>").into_bytes()),
        ("xl/styles.xml".into(), br#"<styleSheet><cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>"#.to_vec()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn cell_is_equal(formula: &str) -> String {
    format!(
        r#"<conditionalFormatting sqref="A1:A8"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>{formula}</formula></cfRule></conditionalFormatting>"#
    )
}

fn expression(formula: &str) -> String {
    format!(
        r#"<conditionalFormatting sqref="A1:A8"><cfRule type="expression" dxfId="0" priority="1"><formula>{formula}</formula></cfRule></conditionalFormatting>"#
    )
}

fn options(now_serial: Option<f64>) -> CalculationOptions {
    CalculationOptions { now_serial }
}

fn open(bytes: &[u8], now_serial: Option<f64>) -> Workbook {
    Workbook::open_recalculated(bytes, options(now_serial)).unwrap()
}

fn viewport() -> Viewport {
    Viewport {
        x: 0.0,
        y: 0.0,
        width: 400.0,
        height: 300.0,
    }
}

/// The rows a display list paints red, in order.
fn red_rows(workbook: &Workbook, list: &DisplayList) -> Vec<u32> {
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let mut rows: Vec<u32> = list
        .commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::FillRect { y, color, .. } if &**color == RED => {
                (0..20).rfind(|&row| geometry.row_y(row) <= y + 0.5)
            }
            _ => None,
        })
        .collect();
    rows.sort_unstable();
    rows
}

/// The rows a screen frame paints red; the frame leaves no cell undecided.
fn frame(workbook: &Workbook) -> Vec<u32> {
    let (list, pending) = workbook
        .display_list_frame_for(SheetId(0), &viewport())
        .unwrap();
    assert_eq!(pending, 0);
    red_rows(workbook, &list)
}

/// The rows a one-shot render paints red.
fn complete(workbook: &Workbook) -> Vec<u32> {
    red_rows(workbook, &workbook.display_list(&viewport()).unwrap())
}

fn value(workbook: &Workbook, row: u32, col: u32) -> CellValue {
    workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(CellRef::new(row, col))
        .map_or(CellValue::Empty, |cell| cell.value.clone())
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

#[test]
fn rules_on_today_paint_the_day_the_cells_were_calculated_with() {
    for formats in [
        cell_is_equal("TODAY()"),
        expression("A1=TODAY()"),
        expression(r#"IF(highlight="Today",A1=TODAY(),A1=highlight)"#),
        expression("A1=Stamp"),
    ] {
        let bytes = package(&[("B1", "=TODAY()"), ("C1", "Today")], NAMES, &formats);
        for (now_serial, row) in [(FIRST, 1), (SECOND, 4)] {
            let workbook = open(&bytes, Some(now_serial));
            assert_eq!(
                value(&workbook, 0, 1),
                number(now_serial.floor()),
                "{formats}"
            );
            assert_eq!(complete(&workbook), [row], "{formats}");
            assert_eq!(frame(&workbook), [row], "{formats}");
        }
    }
}

#[test]
fn a_rule_on_now_paints_where_a_cell_comparing_with_now_is_true() {
    let cells: Vec<(String, String)> = (1..=8)
        .map(|row| (format!("B{row}"), format!("=A{row}&lt;NOW()")))
        .collect();
    let cells: Vec<(&str, &str)> = cells
        .iter()
        .map(|(at, formula)| (at.as_str(), formula.as_str()))
        .collect();
    let bytes = package(&cells, "", &expression("A1&lt;NOW()"));
    for (now_serial, last) in [(FIRST, 1), (SECOND, 4)] {
        let workbook = open(&bytes, Some(now_serial));
        let true_cells: Vec<u32> = (0..8)
            .filter(|&row| value(&workbook, row, 1) == CellValue::Bool { value: true })
            .collect();
        assert_eq!(true_cells, (0..=last).collect::<Vec<u32>>());
        assert_eq!(complete(&workbook), true_cells);
    }
}

#[test]
fn without_a_calculation_date_a_rule_on_the_clock_paints_nothing_and_is_not_pending() {
    for formats in [
        cell_is_equal("TODAY()"),
        expression("A1&lt;NOW()"),
        expression("A1=Stamp"),
    ] {
        let bytes = package(&[("B1", "=TODAY()"), ("C1", "Today")], NAMES, &formats);
        for workbook in [Workbook::open(&bytes).unwrap(), open(&bytes, None)] {
            assert!(frame(&workbook).is_empty(), "{formats}");
            assert!(complete(&workbook).is_empty(), "{formats}");
            assert_eq!(
                workbook.conditional_stats().withheld_volatile,
                8,
                "{formats}"
            );
        }
    }
}

#[test]
fn a_recalculation_with_another_date_repaints_the_next_frame() {
    let bytes = package(&[("B1", "=TODAY()")], "", &cell_is_equal("TODAY()"));
    let mut workbook = open(&bytes, Some(FIRST));
    assert_eq!(frame(&workbook), [1]);
    workbook.recalculate_all(options(Some(FIRST)));
    assert_eq!(frame(&workbook), [1]);
    workbook.recalculate_all(options(Some(SECOND)));
    assert_eq!(value(&workbook, 0, 1), number(46_003.0));
    assert_eq!(frame(&workbook), [4]);
    workbook.recalculate_all(options(None));
    assert!(frame(&workbook).is_empty());
    workbook.recalculate_all(options(Some(FIRST)));
    assert_eq!(frame(&workbook), [1]);
}

#[test]
fn an_edit_with_another_date_repaints_the_next_frame() {
    let bytes = package(&[("B1", "=TODAY()")], "", &cell_is_equal("TODAY()"));
    let mut workbook = open(&bytes, Some(FIRST));
    assert_eq!(frame(&workbook), [1]);
    workbook
        .edit_cell(SheetId(0), CellRef::new(0, 4), "1", options(Some(SECOND)))
        .unwrap();
    assert_eq!(value(&workbook, 0, 1), number(46_003.0));
    assert_eq!(frame(&workbook), [4]);
    workbook
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: CellRef::new(1, 4),
                input: "2".to_owned(),
            }],
            options(Some(FIRST)),
        )
        .unwrap();
    assert_eq!(frame(&workbook), [1]);
}

#[test]
fn a_saved_workbook_opened_again_without_recalculation_has_no_calculation_date() {
    let bytes = package(&[("B1", "=TODAY()")], "", &cell_is_equal("TODAY()"));
    let workbook = open(&bytes, Some(FIRST));
    assert_eq!(frame(&workbook), [1]);
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert!(frame(&reopened).is_empty());
    assert_eq!(reopened.conditional_stats().withheld_volatile, 8);
}

#[test]
fn the_calculation_date_is_no_part_of_the_model_or_the_saved_file() {
    let bytes = package(&[], "", &cell_is_equal("TODAY()"));
    let first = open(&bytes, Some(FIRST));
    let second = open(&bytes, Some(SECOND));
    assert_eq!(frame(&first), [1]);
    assert_eq!(frame(&second), [4]);
    assert!(first.model() == second.model());
    assert_eq!(first.save().unwrap(), second.save().unwrap());
}

#[test]
fn a_peer_hydrated_from_a_calculated_workbook_decides_with_its_date() {
    let bytes = package(&[("B1", "=TODAY()")], "", &cell_is_equal("TODAY()"));
    let worker = open(&bytes, Some(SECOND));
    let peer =
        Workbook::open_with_peer_hydration_json(&bytes, &worker.peer_hydration_json().unwrap())
            .unwrap();
    assert_eq!(value(&peer, 0, 1), number(46_003.0));
    assert_eq!(frame(&peer), [4]);
}

fn snapshot_peer(worker: &Workbook) -> Workbook {
    use betteroffice_xlsx::{SnapshotBudget, WorkbookSnapshotBuilder, WorkbookSnapshotEncoder};

    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(worker, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(worker).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    builder.finish().unwrap().into_parts().0
}

#[test]
fn a_snapshot_peer_keeps_the_date_through_a_refused_edit() {
    let bytes = package(&[("B1", "=TODAY()")], "", &cell_is_equal("TODAY()"));
    let worker = open(&bytes, Some(SECOND));
    let mut peer = snapshot_peer(&worker);
    assert_eq!(frame(&peer), [4]);
    let hydration = worker.peer_hydration_json().unwrap();
    assert_eq!(peer.peer_hydration_json().unwrap(), hydration);
    let stale = serde_json::from_value(serde_json::json!({
        "expectVersion": "stale",
        "steps": [{
            "op": "setCellInputs",
            "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "E1" } },
            "inputs": [["1"]]
        }]
    }))
    .unwrap();
    let refusal = peer.apply_edits(&stale).unwrap().unwrap_err();
    assert_eq!(refusal.failure.code, betteroffice_xlsx::EditFailureCode::StaleVersion);
    assert_eq!(peer.peer_hydration_json().unwrap(), hydration);
    assert_eq!(frame(&peer), [4]);
}

#[test]
fn rules_on_random_numbers_paint_nothing_with_a_calculation_date() {
    let names = r#"<definedNames><definedName name="Draw">RANDARRAY(1)</definedName><definedName name="Later">Draw+1</definedName></definedNames>"#;
    for formula in [
        "RAND()&gt;=0",
        "RANDBETWEEN(1,1)=1",
        "SUM(RANDARRAY(1))&gt;=0",
        "ISNUMBER(Later)",
        "AND(A1=TODAY(),RAND()&gt;=0)",
    ] {
        let workbook = open(&package(&[], names, &expression(formula)), Some(FIRST));
        assert!(frame(&workbook).is_empty(), "{formula}");
        assert_eq!(
            workbook.conditional_stats().withheld_volatile,
            8,
            "{formula}"
        );
    }
}
