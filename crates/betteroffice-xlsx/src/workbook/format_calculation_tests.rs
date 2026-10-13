use super::edit_tests::{options, r, workbook_bytes};
use super::*;
use crate::{EditRequest, XlsxAnchor, XlsxExportOptions, XlsxFormulaResult};
use serde_json::json;
use xlsx_calc::engine::{take_evaluated_formulas, track_formula_evaluations};
use xlsx_model::ErrorValue;
use xlsx_ops::BorderPatch;

/// Cached results the engine would not compute: it gives 5, 10, 6 and 2.
const STALE: &str = concat!(
    r#"<c r="C1"><f>A1+B1</f><v>999</v></c>"#,
    r#"<c r="D1"><f>C1*2</f><v>1234.5</v></c>"#,
    r#"<c r="E1"><f>Sheet1!C1+1</f><v>-7</v></c>"#,
    r#"<c r="F1" t="e"><f>IF(FALSE,1,2)</f><v>#N/A</v></c>"#,
);

const UNCACHED: &str = concat!(
    r#"<c r="C1"><f>A1+B1</f></c><c r="D1"><f>C1*2</f></c>"#,
    r#"<c r="E1"><f>Sheet1!C1+1</f></c><c r="F1"><f>IF(FALSE,1,2)</f></c>"#,
);

const VOLATILE: &str = concat!(
    r#"<c r="G1"><f>INDIRECT("Sheet1!A1")</f><v>2</v></c>"#,
    r#"<c r="H1"><f>G1*3</f><v>6</v></c>"#,
    r#"<c r="I1"><f>INDIRECT("Sheet1!A1")+Sheet1!B1</f><v>5</v></c>"#,
);

type Values = Vec<(SheetId, CellRef, CellValue)>;

fn source(formulas: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&workbook_bytes()).unwrap();
    for (name, bytes) in &mut parts {
        if !name.starts_with("xl/worksheets/sheet") {
            continue;
        }
        let row = if name == "xl/worksheets/sheet1.xml" {
            format!(r#"<row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c>{formulas}</row>"#)
        } else {
            String::new()
        };
        *bytes = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{row}</sheetData></worksheet>"#
        )
        .into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// A package whose three sheets hold `rows` in their first row and whose
/// workbook defines `names`.
fn package(rows: [&str; 3], names: &str) -> Vec<u8> {
    let rows = rows.map(|cells| match cells {
        "" => String::new(),
        cells => format!(r#"<row r="1">{cells}</row>"#),
    });
    sheets_package(rows, names)
}

/// A package whose three sheets hold the `sheet_data` rows.
fn sheets_package(sheet_data: [String; 3], names: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&workbook_bytes()).unwrap();
    for (name, bytes) in &mut parts {
        if name == "xl/workbook.xml" {
            if !names.is_empty() {
                let xml = String::from_utf8(std::mem::take(bytes)).unwrap();
                let defined = format!("</sheets><definedNames>{names}</definedNames>");
                *bytes = xml.replace("</sheets>", &defined).into_bytes();
            }
            continue;
        }
        let Some(index) = name
            .strip_prefix("xl/worksheets/sheet")
            .and_then(|rest| rest.strip_suffix(".xml"))
        else {
            continue;
        };
        let row = &sheet_data[index.parse::<usize>().unwrap() - 1];
        *bytes = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{row}</sheetData></worksheet>"#
        )
        .into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn bold(range: &str) -> Op {
    Op::PatchRangeStyle {
        sheet: SheetId(0),
        range: CellRange::parse_a1(range).unwrap(),
        patch: StylePatch {
            bold: Some(true),
            ..StylePatch::default()
        },
    }
}

fn format_op(kind: &str) -> Op {
    let sheet = SheetId(0);
    let range = CellRange::parse_a1("A1:F1").unwrap();
    let patch = match kind {
        "bold" => return bold("A1:F1"),
        "fill" => StylePatch {
            fill_color: Some("#123456".into()),
            ..StylePatch::default()
        },
        "border" => StylePatch {
            border: Some(BorderPatch {
                preset: Some(BorderPreset::All),
                ..BorderPatch::default()
            }),
            ..StylePatch::default()
        },
        "alignment" => StylePatch {
            horizontal_alignment: Some(HorizontalAlignment::Center),
            ..StylePatch::default()
        },
        "font" => StylePatch {
            font_size: Some(18.0),
            ..StylePatch::default()
        },
        "number" => {
            return Op::SetRangeNumberFormat {
                sheet,
                range,
                format: NumberFormatMutation::Custom {
                    pattern: "0.0000".into(),
                },
            };
        }
        _ => unreachable!("unknown format {kind}"),
    };
    Op::PatchRangeStyle {
        sheet,
        range,
        patch,
    }
}

fn rename() -> Op {
    rename_sheet(0, "Renamed")
}

fn rename_sheet(index: u32, name: &str) -> Op {
    Op::RenameSheet {
        sheet: SheetId(index),
        name: name.into(),
    }
}

fn values(workbook: &Workbook) -> Values {
    workbook
        .model
        .sheets
        .iter()
        .enumerate()
        .flat_map(|(index, sheet)| {
            sheet
                .iter_cells()
                .map(move |(at, cell)| (SheetId(index as u32), at, cell.value.clone()))
        })
        .collect()
}

fn value(workbook: &Workbook, at: &str) -> CellValue {
    workbook.model.sheets[0].cell(r(at)).unwrap().value.clone()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn error(value: ErrorValue) -> CellValue {
    CellValue::Error { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn sorted(addresses: &[CellAddress]) -> Vec<(SheetId, CellRef)> {
    let mut cells: Vec<_> = addresses
        .iter()
        .map(|address| (address.sheet, address.cell))
        .collect();
    cells.sort_by_key(|(sheet, at)| (sheet.0, at.row, at.col));
    cells
}

fn cycle_cells(workbook: &Workbook) -> Vec<(SheetId, CellRef)> {
    sorted(&workbook.last_calculation().cycle_cells)
}

/// Cycle cells, then limited cells.
type Status = [Vec<(SheetId, CellRef)>; 2];

fn status(cycle_cells: &[CellAddress], limited_cells: &[CellAddress]) -> Status {
    [sorted(cycle_cells), sorted(limited_cells)]
}

fn last_status(workbook: &Workbook) -> Status {
    let calculation = workbook.last_calculation();
    status(&calculation.cycle_cells, &calculation.limited_cells)
}

/// The cells a structured export marks as cycle and as limited.
fn exported_status(workbook: &Workbook) -> Status {
    let export = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap();
    let mut marked: [Vec<CellAddress>; 2] = Default::default();
    for (index, sheet) in export.content.sheets.iter().enumerate() {
        for cell in &sheet.cells {
            let slot = match cell.formula_result {
                Some(XlsxFormulaResult::Cycle) => 0,
                Some(XlsxFormulaResult::Limited) => 1,
                _ => continue,
            };
            if let XlsxAnchor::Cell { a1, .. } = &cell.anchor {
                marked[slot].push(CellAddress {
                    sheet: SheetId(index as u32),
                    cell: r(a1),
                });
            }
        }
    }
    status(&marked[0], &marked[1])
}

/// A full recalculation finds nothing the scoped one left behind.
fn assert_fully_calculated(workbook: &mut Workbook) {
    let scoped = values(workbook);
    let status = calculation_status(workbook.last_calculation());
    workbook.recalculate_all(options());
    assert_eq!(values(workbook), scoped);
    assert_eq!(calculation_status(workbook.last_calculation()), status);
}

fn calculation_counters(workbook: &Workbook) -> (bool, u64) {
    (
        workbook.recalculated_since_open,
        workbook.calculations_since_open,
    )
}

fn evaluated() -> Vec<(SheetId, CellRef)> {
    let mut cells = take_evaluated_formulas();
    cells.sort_by_key(|(sheet, at)| (sheet.0, at.row, at.col));
    cells
}

fn cells(addresses: &[&str]) -> Vec<(SheetId, CellRef)> {
    addresses.iter().map(|at| (SheetId(0), r(at))).collect()
}

fn assert_kept(
    workbook: &Workbook,
    result: &MutationResult,
    before: &Values,
    counters: (bool, u64),
) {
    assert!(result.applied);
    assert!(result.changed.is_empty());
    assert!(result.cycle_cells.is_empty());
    assert!(result.limited_cells.is_empty());
    assert_eq!(&values(workbook), before);
    assert_eq!(calculation_counters(workbook), counters);
    assert_eq!(evaluated(), cells(&[]));
}

fn check_format(kind: &str, history: bool) {
    let bytes = source(&format!("{STALE}{VOLATILE}"));
    let mut workbook = Workbook::open(&bytes).unwrap();
    let before = values(&workbook);
    let counters = calculation_counters(&workbook);
    let epoch = workbook.model_epoch;
    track_formula_evaluations();
    let result = workbook
        .apply_ops(vec![format_op(kind)], options())
        .unwrap();
    assert_kept(&workbook, &result, &before, counters);
    assert!(workbook.model_epoch > epoch);
    if !history {
        return;
    }
    for redo in [false, true] {
        let epoch = workbook.model_epoch;
        track_formula_evaluations();
        let result = if redo {
            workbook.redo(options())
        } else {
            workbook.undo(options())
        }
        .unwrap();
        assert_kept(&workbook, &result, &before, counters);
        assert!(workbook.model_epoch > epoch);
    }
}

macro_rules! format_tests {
    ($($edit:ident, $history:ident, $kind:literal;)*) => {
        $(
            #[test]
            fn $edit() {
                check_format($kind, false);
            }

            #[test]
            fn $history() {
                check_format($kind, true);
            }
        )*
    };
}

format_tests! {
    bold_keeps_formula_results, bold_history_keeps_formula_results, "bold";
    fill_keeps_formula_results, fill_history_keeps_formula_results, "fill";
    border_keeps_formula_results, border_history_keeps_formula_results, "border";
    alignment_keeps_formula_results, alignment_history_keeps_formula_results, "alignment";
    font_keeps_formula_results, font_history_keeps_formula_results, "font";
    number_format_keeps_formula_results, number_format_history_keeps_formula_results, "number";
}

/// The `<v>` text of every formula cell in the saved worksheets, by address.
fn formula_caches(bytes: &[u8]) -> BTreeMap<String, String> {
    let mut caches = BTreeMap::new();
    for (name, data) in ooxml_opc::unzip_parts(bytes).unwrap() {
        if !name.starts_with("xl/worksheets/sheet") {
            continue;
        }
        let xml = String::from_utf8(data).unwrap();
        for cell in xml.split("<c ").skip(1) {
            let tag = &cell[..cell.find('>').unwrap()];
            if tag.ends_with('/') {
                continue;
            }
            let cell = cell.split("</c>").next().unwrap();
            if !cell.contains("<f") {
                continue;
            }
            let at = tag.split('"').nth(1).unwrap();
            let cached = cell
                .split("<v>")
                .nth(1)
                .map_or("", |rest| rest.split("</v>").next().unwrap());
            caches.insert(at.to_owned(), cached.to_owned());
        }
    }
    caches
}

fn assert_saved_results_kept(op: Op) {
    let bytes = source(STALE);
    let expected = formula_caches(&bytes);
    assert_eq!(expected.len(), 4);
    let mut workbook = Workbook::open(&bytes).unwrap();
    workbook.apply_ops(vec![op], options()).unwrap();
    assert_eq!(formula_caches(&workbook.save().unwrap()), expected);
}

#[test]
fn bold_saves_formula_results_byte_identical() {
    assert_saved_results_kept(format_op("bold"));
}

#[test]
fn number_format_saves_formula_results_byte_identical() {
    assert_saved_results_kept(format_op("number"));
}

#[test]
fn rename_saves_formula_results_byte_identical() {
    assert_saved_results_kept(rename());
}

#[test]
fn rename_reevaluates_only_volatile_formulas_and_their_dependents() {
    let bytes = source(&format!("{STALE}{VOLATILE}"));
    let mut workbook = Workbook::open(&bytes).unwrap();
    let ordinary = values(&workbook)[..6].to_vec();
    track_formula_evaluations();
    let result = workbook
        .apply_ops(vec![format_op("bold"), rename()], options())
        .unwrap();
    assert_eq!(evaluated(), cells(&["G1", "H1", "I1"]));
    assert_eq!(values(&workbook)[..6], ordinary);
    let changed: Vec<_> = result
        .changed
        .iter()
        .map(|address| (address.sheet, address.cell))
        .collect();
    assert_eq!(changed, cells(&["G1", "H1", "I1"]));
    for at in ["G1", "H1", "I1"] {
        assert_eq!(value(&workbook, at), error(ErrorValue::Ref));
    }
    let rewritten = workbook.model.sheets[0].cell(r("E1")).unwrap();
    assert!(rewritten.formula.as_deref().unwrap().contains("Renamed!"));
    workbook
        .edit_cell(SheetId(0), r("A1"), "7", options())
        .unwrap();
    assert_eq!(value(&workbook, "C1"), number(10.0));
    assert_eq!(value(&workbook, "E1"), number(11.0));
}

/// D1 shows C1's `formula`, which names the renamed sheet.
fn assert_rename_reevaluates_formula_text(formula: &str) {
    let first = format!(
        r#"<c r="A1"><v>2</v></c><c r="C1"><f>{formula}</f><v>2</v></c><c r="D1" t="str"><f>FORMULATEXT(C1)</f><v>={formula}</v></c>"#
    );
    let rows = [
        first.as_str(),
        "",
        r#"<c r="A1"><f>Sheet2!A1+1</f><v>1</v></c>"#,
    ];
    let mut workbook = Workbook::open(&package(rows, "")).unwrap();
    track_formula_evaluations();
    workbook.apply_ops(vec![rename()], options()).unwrap();
    let rewritten = workbook.model.sheets[0].cell(r("C1")).unwrap();
    let rewritten = rewritten.formula.clone().unwrap();
    assert!(rewritten.contains("Renamed!"));
    assert_eq!(value(&workbook, "D1"), text(&format!("={rewritten}")));
    assert_eq!(evaluated(), cells(&["D1"]));
    assert_fully_calculated(&mut workbook);
}

#[test]
fn rename_reevaluates_what_shows_a_rewritten_formula() {
    assert_rename_reevaluates_formula_text("Sheet1!A1");
}

#[test]
fn rename_reevaluates_what_shows_a_rewritten_formula_the_engine_cannot_parse() {
    assert_rename_reevaluates_formula_text("SUM(Sheet1!A1#)");
}

#[test]
fn rename_reevaluates_what_resolves_by_the_new_name() {
    let rows = [
        concat!(
            r#"<c r="C1" t="e"><f>New!Rate</f><v>#REF!</v></c>"#,
            r#"<c r="D1" t="e"><f>ROW(New!B5)</f><v>#VALUE!</v></c>"#,
            r#"<c r="E1" t="e"><f>C1+1</f><v>#REF!</v></c>"#,
        ),
        "",
        r#"<c r="A1"><f>Sheet2!A1+1</f><v>1</v></c>"#,
    ];
    let names = r#"<definedName name="Rate">7</definedName>"#;
    let mut workbook = Workbook::open(&package(rows, names)).unwrap();
    track_formula_evaluations();
    workbook
        .apply_ops(vec![rename_sheet(0, "New")], options())
        .unwrap();
    for (at, expected) in [("C1", 7.0), ("D1", 5.0), ("E1", 8.0)] {
        assert_eq!(value(&workbook, at), number(expected));
    }
    assert_eq!(evaluated(), cells(&["C1", "D1", "E1"]));
    assert_fully_calculated(&mut workbook);
}

#[test]
fn rename_reevaluates_a_missing_name_qualified_by_the_new_sheet_name() {
    let rows = [
        concat!(
            r#"<c r="C1" t="e"><f>New!Missing</f><v>#REF!</v></c>"#,
            r#"<c r="E1" t="e"><f>Alias</f><v>#REF!</v></c>"#,
        ),
        "",
        r#"<c r="A1"><f>Sheet2!A1+1</f><v>1</v></c>"#,
    ];
    let names = r#"<definedName name="Alias">New!Missing</definedName>"#;
    let mut workbook = Workbook::open(&package(rows, names)).unwrap();
    track_formula_evaluations();
    workbook
        .apply_ops(vec![rename_sheet(0, "New")], options())
        .unwrap();
    for at in ["C1", "E1"] {
        assert_eq!(value(&workbook, at), error(ErrorValue::Name));
    }
    assert_eq!(evaluated(), cells(&["C1", "E1"]));
    assert_fully_calculated(&mut workbook);
}

/// Two cycles and two limited cells. The G1 cycle and K1 are volatile, and
/// Sheet3's far corner puts what J1 and K1 sum past the evaluation budget.
fn status_package() -> Vec<u8> {
    let first = concat!(
        r#"<row r="1"><c r="C1"><f>D1+1</f><v>0</v></c><c r="D1"><f>C1+1</f><v>0</v></c>"#,
        r#"<c r="G1"><f>IF(INDIRECT("Sheet3!A1")=1,H1,0)</f><v>0</v></c>"#,
        r#"<c r="H1"><f>G1+1</f><v>0</v></c>"#,
        r#"<c r="J1"><f>SUM(Sheet3!A1:XFD1048576)</f><v>2</v></c>"#,
        r#"<c r="K1"><f>NOW()*0+SUM(Sheet3!A1:XFD1048576)</f><v>2</v></c></row>"#,
    );
    let third = concat!(
        r#"<row r="1"><c r="A1"><v>1</v></c></row>"#,
        r#"<row r="1048576"><c r="XFD1048576"><v>1</v></c></row>"#,
    );
    sheets_package([first.into(), String::new(), third.into()], "")
}

#[test]
fn rename_keeps_the_status_of_formulas_it_does_not_reach() {
    let mut workbook = Workbook::open(&status_package()).unwrap();
    workbook.recalculate_all(options());
    let cycles = cells(&["C1", "D1", "G1", "H1"]);
    let limited = cells(&["J1", "K1"]);
    let expected = [cycles.clone(), limited.clone()];
    assert_eq!(cycle_cells(&workbook), cycles);
    assert_eq!(last_status(&workbook), expected);
    track_formula_evaluations();
    let result = workbook
        .apply_ops(vec![rename_sheet(1, "Other")], options())
        .unwrap();
    assert_eq!(evaluated(), cells(&["G1", "H1", "K1"]));
    assert_eq!(cycle_cells(&workbook), cycles);
    assert_eq!(last_status(&workbook), expected);
    assert_eq!(status(&result.cycle_cells, &result.limited_cells), expected);
    assert_eq!(exported_status(&workbook), expected);
    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(last_status(&restored), expected);
    let read = workbook
        .read_cells(&super::target::ReadRequest::default())
        .unwrap()
        .unwrap();
    let mut shown: Vec<_> = read
        .calculation
        .cycle_cells
        .iter()
        .map(|cell| cell.a1.as_str())
        .collect();
    shown.sort_unstable();
    assert_eq!(shown, ["C1", "D1", "G1", "H1"]);
    let mut shown: Vec<_> = read
        .calculation
        .limited_cells
        .iter()
        .map(|cell| cell.a1.as_str())
        .collect();
    shown.sort_unstable();
    assert_eq!(shown, ["J1", "K1"]);
    track_formula_evaluations();
    workbook
        .apply_ops(vec![rename_sheet(2, "Gone")], options())
        .unwrap();
    assert_eq!(evaluated(), cells(&["G1", "H1", "K1"]));
    assert_eq!(cycle_cells(&workbook), cells(&["C1", "D1"]));
    assert_eq!(last_status(&workbook), [cells(&["C1", "D1"]), limited]);
    assert_fully_calculated(&mut workbook);
}

/// A value edit after a rename reports what it reports after a full
/// recalculation of that rename, the volatile cycle and limit included.
#[test]
fn a_value_edit_after_a_rename_reports_the_status_a_full_rename_leads_to() {
    let bytes = status_package();
    let mut scoped = Workbook::open(&bytes).unwrap();
    let mut full = Workbook::open(&bytes).unwrap();
    for workbook in [&mut scoped, &mut full] {
        workbook.recalculate_all(options());
        workbook
            .apply_ops(vec![rename_sheet(1, "Other")], options())
            .unwrap();
    }
    full.recalculate_all(options());
    let edited = [&mut scoped, &mut full].map(|workbook| {
        let result = workbook
            .edit_cell(SheetId(1), r("A1"), "5", options())
            .unwrap();
        status(&result.cycle_cells, &result.limited_cells)
    });
    assert_eq!(edited[0], edited[1]);
    assert_eq!(last_status(&scoped), last_status(&full));
    assert_eq!(values(&scoped), values(&full));
    let [cycles, limited] = last_status(&scoped);
    for cell in cells(&["G1", "H1"]) {
        assert!(cycles.contains(&cell));
    }
    assert!(limited.contains(&(SheetId(0), r("K1"))));
}

/// Sheet1 renamed to New; C1 reads A1 = 1 through the sheet name, D1 reads C1.
fn renamed_workbook() -> Workbook {
    let first = concat!(
        r#"<c r="A1"><v>1</v></c><c r="C1"><f>Sheet1!A1</f><v>1</v></c>"#,
        r#"<c r="D1"><f>C1+1</f><v>2</v></c>"#,
    );
    let mut workbook = Workbook::open(&package([first, "", ""], "")).unwrap();
    workbook
        .apply_ops(vec![rename_sheet(0, "New")], options())
        .unwrap();
    workbook
}

/// Without an undo step, C1 gets `formula`, which keeps it at 1, and A1 becomes 2.
fn edit_outside_history(workbook: &mut Workbook, formula: &str) {
    let target = |a1: &str| json!({ "sheetId": "sheet:0", "range": { "kind": "a1", "a1": a1 } });
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": workbook.version(),
        "history": "none",
        "steps": [
            { "op": "setFormulas", "target": target("C1"), "formulas": [[formula]] },
            { "op": "setCellInputs", "target": target("A1"), "inputs": [["2"]] },
        ],
    }))
    .unwrap();
    workbook.apply_edits(&request).unwrap().unwrap();
    assert_eq!(value(workbook, "C1"), number(1.0));
}

/// The restored `formula` reads A1 again, so C1 and D1 follow A1 = 2.
fn assert_restored_formula_recalculated(workbook: &mut Workbook, formula: &str) {
    let restored = workbook.model.sheets[0].cell(r("C1")).unwrap();
    assert_eq!(restored.formula.as_deref(), Some(formula));
    assert_eq!(value(workbook, "C1"), number(2.0));
    assert_eq!(value(workbook, "D1"), number(3.0));
    assert_fully_calculated(workbook);
}

#[test]
fn undoing_a_rename_recalculates_a_formula_edited_outside_history() {
    let mut workbook = renamed_workbook();
    edit_outside_history(&mut workbook, "New!A1*0+1");
    workbook.undo(options()).unwrap();
    assert_restored_formula_recalculated(&mut workbook, "Sheet1!A1");
}

#[test]
fn redoing_a_rename_recalculates_a_formula_edited_outside_history() {
    let mut workbook = renamed_workbook();
    workbook.undo(options()).unwrap();
    edit_outside_history(&mut workbook, "Sheet1!A1*0+1");
    workbook.redo(options()).unwrap();
    assert_restored_formula_recalculated(&mut workbook, "New!A1");
}

#[test]
fn format_edit_keeps_a_graph_value_edits_recalculate_through() {
    let mut workbook = Workbook::open(&source(STALE)).unwrap();
    track_formula_evaluations();
    workbook
        .apply_ops(vec![format_op("bold")], options())
        .unwrap();
    assert_eq!(evaluated(), cells(&[]));
    track_formula_evaluations();
    workbook
        .edit_cell(SheetId(0), r("A1"), "7", options())
        .unwrap();
    assert_eq!(evaluated(), cells(&["C1", "D1", "E1"]));
    for (at, expected) in [("C1", 10.0), ("D1", 20.0), ("E1", 11.0)] {
        assert_eq!(value(&workbook, at), number(expected));
    }
    assert_eq!(value(&workbook, "F1"), error(ErrorValue::NA));
}

#[test]
fn format_edit_moving_the_used_extent_recalculates() {
    let blanks = r#"<c r="C1"><f>SUMPRODUCT((A:A="")*1)</f><v>0</v></c>"#;
    let mut workbook = Workbook::open(&source(blanks)).unwrap();
    workbook.apply_ops(vec![bold("A4")], options()).unwrap();
    assert_eq!(value(&workbook, "C1"), number(3.0));
    workbook.undo(options()).unwrap();
    assert_eq!(value(&workbook, "C1"), number(0.0));
}

#[test]
fn collaborative_format_edit_still_recalculates_every_formula() {
    let mut workbook = Workbook::open_collaborative(&source(STALE), 41).unwrap();
    track_formula_evaluations();
    workbook
        .apply_ops(vec![format_op("bold")], options())
        .unwrap();
    assert_eq!(evaluated(), cells(&["C1", "D1", "E1", "F1"]));
}

#[test]
fn value_edit_with_a_format_edit_recalculates_every_formula() {
    let mut workbook = Workbook::open(&source(STALE)).unwrap();
    track_formula_evaluations();
    let ops = vec![
        format_op("bold"),
        Op::SetCell {
            sheet: SheetId(0),
            at: r("A1"),
            cell: CellState {
                value: number(7.0),
                ..CellState::default()
            },
        },
    ];
    workbook.apply_ops(ops, options()).unwrap();
    assert_eq!(evaluated(), cells(&["C1", "D1", "E1", "F1"]));
    assert_eq!(value(&workbook, "F1"), number(2.0));
    assert_eq!(calculation_counters(&workbook), (true, 1));
}

#[test]
fn recalculate_all_saves_every_result_of_an_uncached_workbook() {
    let mut workbook = Workbook::open(&source(UNCACHED)).unwrap();
    track_formula_evaluations();
    workbook.recalculate_all(options());
    assert_eq!(evaluated(), cells(&["C1", "D1", "E1", "F1"]));
    assert_eq!(calculation_counters(&workbook), (true, 1));
    let saved = workbook.save().unwrap();
    let expected = [("C1", "5"), ("D1", "10"), ("E1", "6"), ("F1", "2")]
        .map(|(at, cached)| (at.to_owned(), cached.to_owned()));
    assert_eq!(formula_caches(&saved), BTreeMap::from(expected));
    let reopened = Workbook::open(&saved).unwrap();
    for (at, expected) in [("C1", 5.0), ("D1", 10.0), ("E1", 6.0), ("F1", 2.0)] {
        assert_eq!(value(&reopened, at), number(expected));
    }
}
