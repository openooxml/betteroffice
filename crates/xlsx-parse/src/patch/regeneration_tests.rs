use std::collections::HashMap;

use xlsx_model::{CellValue, SheetId, Workbook};

use super::style_match_tests::{cells, edit, parsed_sheet, save};
use super::{SheetPatch, StyleMatch, scan_sheet_data};
use crate::{
    ParseError, ParsedWorkbook, SaveEdits, SheetAxes,
    serialize_workbook_with_package_and_origins_after_edits_and_active_sheet_with_axes,
};

fn save_error(parsed: &ParsedWorkbook, workbook: &Workbook) -> ParseError {
    let provenance = vec![parsed.package.source_shared_string_cells(0)];
    serialize_workbook_with_package_and_origins_after_edits_and_active_sheet_with_axes(
        workbook,
        &parsed.package,
        &[Some(0)],
        &provenance,
        &[Some(SheetAxes::default())],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
        SheetId(0),
    )
    .unwrap_err()
}

#[test]
fn regenerated_sheet_data_allows_modeled_markup_and_layout_hints() {
    let source = concat!(
        r#"<sheetData xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">"#,
        r#"<row ht="20" customHeight="1" hidden="0" spans="1:5" thickTop="1" thickBot="1" x14ac:dyDescent="0.25">"#,
        r#"<c r="A1" s="2" t="n"><v>1</v></c>"#,
        r#"<c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c>"#,
        r#"<c r="C1"><f t="normal">1</f><v>1</v></c>"#,
        r#"<c r="D1"><f t="array" ref="D1">1</f><v>1</v></c>"#,
        r#"<c r="E1" t="inlineStr"><is><t xml:space="preserve"> plain </t></is></c></row>"#,
        r#"<row r="2" ht="0" customHeight="1" hidden="1"><c r="A2" t="b"><v>1</v></c>"#,
        r#"<c r="B2"><f t="shared" si="0"/><v>2</v></c></row></sheetData>"#,
    );
    assert!(scan_sheet_data(source.as_bytes()).unwrap().is_none());
    let parsed = parsed_sheet(source);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "A1", |cell| {
        cell.value = CellValue::Number { value: 9.0 }
    });

    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));

    assert!(cells(&saved)["A1"].contains("<v>9</v>"));
    assert!(saved.contains(r#"ht="20" customHeight="1""#));
    assert!(saved.contains(r#"ht="0" customHeight="1" hidden="1""#));
    assert!(cells(&saved)["D1"].contains(r#"<f t="array" ref="D1">1</f>"#));
}

#[test]
fn regenerated_sheet_data_refuses_unmodeled_markup() {
    for (row_attributes, cell_markup, expected, location) in [
        (
            r#"s="2" customFormat="1""#,
            "<v>1</v>",
            r#"<row> attribute s="2""#,
            "row 1",
        ),
        (
            r#"outlineLevel="1""#,
            "<v>1</v>",
            r#"<row> attribute outlineLevel="1""#,
            "row 1",
        ),
        (
            r#"customFormat="1""#,
            "<v>1</v>",
            r#"<row> attribute customFormat="1""#,
            "row 1",
        ),
        (
            r#"customFormat="&missing;""#,
            "<v>1</v>",
            r#"<row> attribute customFormat="&missing;""#,
            "row 1",
        ),
        (
            r#"collapsed="1""#,
            "<v>1</v>",
            r#"<row> attribute collapsed="1""#,
            "row 1",
        ),
        (
            r#"ph="1""#,
            "<v>1</v>",
            r#"<row> attribute ph="1""#,
            "row 1",
        ),
        (
            r#"x:flag="1""#,
            "<v>1</v>",
            r#"<row> attribute x:flag="1""#,
            "row 1",
        ),
        (
            "",
            r#"<c r="A1" cm="1"><v>1</v></c>"#,
            r#"<c> attribute cm="1""#,
            "cell A1",
        ),
        (
            "",
            r#"<c r="A1" vm="1"><v>1</v></c>"#,
            r#"<c> attribute vm="1""#,
            "cell A1",
        ),
        (
            "",
            r#"<c r="A1" ph="1"><v>1</v></c>"#,
            r#"<c> attribute ph="1""#,
            "cell A1",
        ),
        (
            "",
            r#"<c r="A1" t="d"><v>1</v></c>"#,
            r#"<c> attribute t="d""#,
            "cell A1",
        ),
        (
            "",
            r#"<v>1</v><extLst><ext uri="value"/></extLst>"#,
            "<extLst>",
            "cell A1",
        ),
        (
            "",
            r#"<f t="dataTable" ref="A1:A2" dt2D="1" r1="B1">1</f><v>1</v>"#,
            r#"<f> attribute t="dataTable""#,
            "cell A1",
        ),
        (
            "",
            r#"<c r="A1" t="inlineStr"><is><r><t>plain</t></r></is></c>"#,
            "<r>",
            "cell A1",
        ),
        (
            "",
            r#"<f ca="1">1</f><v>1</v>"#,
            r#"<f> attribute ca="1""#,
            "cell A1",
        ),
        (
            "",
            r#"<c cm="1"><v>1</v></c>"#,
            r#"<c> attribute cm="1""#,
            "row 1, cell 1",
        ),
        (
            "",
            r#"<c r="A1" t="inlineStr"><is><t>plain</t><rPh sb="0" eb="1"><t>p</t></rPh></is></c>"#,
            "<rPh>",
            "cell A1",
        ),
        (
            "",
            r#"<c r="A1" t="inlineStr"><is><t>plain</t><phoneticPr fontId="0"/></is></c>"#,
            "<phoneticPr>",
            "cell A1",
        ),
        ("", "<v>1</v><v>2</v>", "<v>", "cell A1"),
        ("", "<v>1</v><is><t>plain</t></is>", "<is>", "cell A1"),
        (
            "",
            r#"<c r="A1" t="inlineStr"><is><t>one</t><t>two</t></is></c>"#,
            "<t>",
            "cell A1",
        ),
    ] {
        let cell = if cell_markup.starts_with("<c ") {
            cell_markup.to_owned()
        } else {
            format!(r#"<c r="A1">{cell_markup}</c>"#)
        };
        let source = format!(
            r#"<sheetData><row {row_attributes}>{cell}<c r="B1"><v>2</v></c></row></sheetData>"#
        );
        assert!(scan_sheet_data(source.as_bytes()).unwrap().is_none());
        let parsed = parsed_sheet(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "B1", |cell| {
            cell.value = CellValue::Number { value: 9.0 }
        });

        let ParseError::UnsupportedEdit(message) = save_error(&parsed, &workbook) else {
            panic!("expected UnsupportedEdit for {source}");
        };
        assert!(message.starts_with("sheet Sheet1:"), "{message}");
        assert!(message.contains(expected), "{expected}: {message}");
        assert!(message.ends_with(&format!("at {location}")), "{message}");
    }
}

#[test]
fn regeneration_refusals_report_source_addresses_and_ordinals() {
    for (row, expected) in [
        (
            r#"<row r="7" s="2"><c r="A7"><v>1</v></c></row>"#,
            "at row 7",
        ),
        (r#"<row><c r="D7" cm="1"><v>1</v></c></row>"#, "at cell D7"),
        (
            r#"<row><c><v>1</v></c><c cm="1"><v>2</v></c></row>"#,
            "at row 2, cell 2",
        ),
    ] {
        let source = format!(r#"<sheetData><row><c r="B1"><v>2</v></c></row>{row}</sheetData>"#);
        assert!(scan_sheet_data(source.as_bytes()).unwrap().is_none());
        let parsed = parsed_sheet(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "B1", |cell| {
            cell.value = CellValue::Number { value: 9.0 }
        });

        let ParseError::UnsupportedEdit(message) = save_error(&parsed, &workbook) else {
            panic!("expected UnsupportedEdit for {source}");
        };
        assert!(message.starts_with("sheet Sheet1:"), "{message}");
        assert!(message.ends_with(expected), "{message}");
    }
}

#[test]
fn patch_errors_keep_the_sheet_name_and_original_error() {
    for columns in [false, true] {
        let row = if columns {
            r#"<row r="1">"#
        } else {
            r#"<row r="1" ht="20" hidden="0" customFormat="&missing;">"#
        };
        let cols = if columns {
            r#"<cols><col min="1" max="1" width="10" hidden="0" style="0" customWidth="&missing;"/></cols>"#
        } else {
            ""
        };
        let source = format!(r#"{cols}<sheetData>{row}<c r="A1"><v>1</v></c></row></sheetData>"#);
        let parsed = parsed_sheet(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| {
            cell.value = CellValue::Number { value: 9.0 }
        });
        let axes = SheetAxes::default();
        let styles = StyleMatch::new(&parsed.workbook.styles, &workbook.styles);
        let indices = HashMap::new();
        let retained = parsed.package.source_shared_string_cells(0);
        let patch = SheetPatch {
            sheet: &workbook.sheets[0],
            original: &parsed.workbook.sheets[0],
            axes: &axes,
            workbook: &workbook,
            sst_index: &indices,
            retained: &retained,
            plan: None,
            styles: &styles,
        };
        let template = &parsed.package.sheets[0].template;
        let original_error = if columns {
            patch
                .cols(Some(&template.child("cols").unwrap().bytes))
                .unwrap_err()
        } else {
            patch
                .sheet_data(&template.child("sheetData").unwrap().bytes)
                .unwrap_err()
        };

        assert_eq!(
            save_error(&parsed, &workbook),
            ParseError::UnsupportedEdit(format!("sheet Sheet1: {original_error}")),
        );
    }
}
