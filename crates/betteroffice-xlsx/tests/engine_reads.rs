use betteroffice_xlsx::{
    CalculationOptions, Cell, CellEdit, CellRange, CellRef, CellValue, ConditionalFormatRule,
    ConditionalFormatting, DefinedName, DefinedNamesRequest, DifferentialFormat, Error, Op, Sheet,
    SheetId, SheetInfo, Workbook, WorkbookModel,
};
use serde_json::json;

fn cell(a1: &str) -> CellRef {
    CellRef::parse_a1(a1).unwrap()
}

fn model() -> WorkbookModel {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    model.sheets.push(Sheet::new("Other"));
    model
}

#[test]
fn defined_names_fixture_supports_typed_and_byte_identical_json_reads() {
    let workbook = Workbook::open(include_bytes!(
        "../../../packages/xlsx/test-fixtures/defined-names.xlsx"
    ))
    .unwrap();
    let request = DefinedNamesRequest {
        limit: Some(2),
        ..Default::default()
    };
    let page = workbook.list_defined_names(&request).unwrap();
    assert_eq!(page.version, workbook.version().to_string());
    assert_eq!(page.total, 5);
    assert_eq!(page.next_offset, Some(2));
    assert_eq!(page.items[0].name, "Rows");
    assert_eq!(page.items[0].scope, None);
    assert_eq!(page.items[1].name, "Header");
    assert_eq!(
        workbook.list_defined_names_json(r#"{"limit":2}"#).unwrap(),
        format!(
            r#"{{"items":[{{"hidden":false,"name":"Rows","refersTo":"Data!$A$1:INDEX(Data!$A:$A,COUNTA(Data!$A:$A))","scope":null}},{{"hidden":false,"name":"Header","refersTo":"SUM(Data!$1:$1)","scope":null}}],"nextOffset":2,"total":5,"version":"{}"}}"#,
            workbook.version(),
        ),
    );
    let last = workbook
        .list_defined_names(&DefinedNamesRequest {
            offset: Some(4),
            ..Default::default()
        })
        .unwrap();
    assert_eq!(last.items[0].name, "Spaced");
    assert_eq!(last.next_offset, None);
    let past = workbook
        .list_defined_names(&DefinedNamesRequest {
            offset: Some(usize::MAX),
            ..Default::default()
        })
        .unwrap();
    assert!(past.items.is_empty());
    assert_eq!(past.next_offset, None);
}

#[test]
fn defined_name_scopes_distinguish_missing_null_and_sheet_indices() {
    let mut model = model();
    model.defined_names = vec![
        DefinedName {
            name: "Global".into(),
            formula: "Data!$A$1".into(),
            local_sheet: None,
            hidden: false,
        },
        DefinedName {
            name: "Local".into(),
            formula: "Data!$B$1".into(),
            local_sheet: Some(SheetId(0)),
            hidden: true,
        },
        DefinedName {
            name: "Local".into(),
            formula: "Other!$C$1".into(),
            local_sheet: Some(SheetId(1)),
            hidden: false,
        },
    ];
    let workbook = Workbook::from_model(model).unwrap();
    for (request, scope, names) in [
        ("{}", None, vec!["Global", "Local", "Local"]),
        (r#"{"scope":null}"#, Some(None), vec!["Global"]),
        (r#"{"scope":0}"#, Some(Some(0)), vec!["Local"]),
        (r#"{"scope":1}"#, Some(Some(1)), vec!["Local"]),
    ] {
        let typed: DefinedNamesRequest = serde_json::from_str(request).unwrap();
        assert_eq!(typed.scope, scope);
        let roundtrip: DefinedNamesRequest = serde_json::from_value(json!(typed)).unwrap();
        assert_eq!(roundtrip, typed);
        let result = workbook.list_defined_names(&typed).unwrap();
        assert_eq!(
            result
                .items
                .iter()
                .map(|item| item.name.as_str())
                .collect::<Vec<_>>(),
            names
        );
        assert_eq!(
            serde_json::from_str::<serde_json::Value>(
                &workbook.list_defined_names_json(request).unwrap()
            )
            .unwrap(),
            json!(result),
        );
    }
    let local = workbook
        .list_defined_names(&DefinedNamesRequest {
            scope: Some(Some(0)),
            ..Default::default()
        })
        .unwrap();
    assert!(local.items[0].hidden);
    for (request, message) in [
        (r#"{"limit":0}"#, "limit must be between 1 and 100"),
        (r#"{"limit":101}"#, "limit must be between 1 and 100"),
        (r#"{"scope":2}"#, "scope sheet is out of range"),
    ] {
        assert_eq!(
            workbook
                .list_defined_names_json(request)
                .unwrap_err()
                .to_string(),
            message
        );
    }
    for invalid in [
        r#"{"extra":1}"#,
        r#"{"scope":-1}"#,
        r#"{"offset":-1}"#,
        "null",
        "{",
    ] {
        assert!(workbook.list_defined_names_json(invalid).is_err());
    }
}

#[test]
fn conditional_formats_preserve_rule_order_optional_fields_and_json_bytes() {
    let mut model = model();
    model.sheets[0].conditional_formats = vec![ConditionalFormatting {
        sqref: "A1:A3 C1:C3".into(),
        rules: vec![
            ConditionalFormatRule {
                id: "cf:0".into(),
                rule_type: "cellIs".into(),
                operator: Some("greaterThan".into()),
                formulas: vec!["1".into()],
                priority: 2,
                dxf_id: Some(3),
                format: DifferentialFormat {
                    font_color: Some("#FF0000".into()),
                    bold: Some(false),
                    italic: Some(true),
                    fill: Some("#00FF00".into()),
                },
                raw_xml: None,
            },
            ConditionalFormatRule {
                id: "cf:1".into(),
                rule_type: "dataBar".into(),
                operator: None,
                formulas: Vec::new(),
                priority: 1,
                dxf_id: None,
                format: DifferentialFormat::default(),
                raw_xml: Some("<cfRule/>".into()),
            },
        ],
    }];
    let workbook = Workbook::from_model(model).unwrap();
    let rules = workbook.conditional_formats(SheetId(0)).unwrap();
    assert_eq!(rules.len(), 2);
    assert_eq!(rules[0].range, "A1:A3 C1:C3");
    assert_eq!(rules[0].format.bold, Some(false));
    assert!(rules[0].editable);
    assert!(!rules[1].editable);
    assert_eq!(
        workbook.conditional_formats_json(SheetId(0)).unwrap(),
        r##"[{"dxfId":3,"editable":true,"format":{"bold":false,"fill":"#00FF00","fontColor":"#FF0000","italic":true},"formulas":["1"],"id":"cf:0","operator":"greaterThan","priority":2,"range":"A1:A3 C1:C3","type":"cellIs"},{"editable":false,"format":{},"formulas":[],"id":"cf:1","priority":1,"range":"A1:A3 C1:C3","type":"dataBar"}]"##,
    );
    assert_eq!(workbook.conditional_formats_json(SheetId(1)).unwrap(), "[]");
    assert_eq!(
        workbook
            .conditional_formats_json(SheetId(2))
            .unwrap_err()
            .to_string(),
        "sheet 2 not found"
    );
}

#[test]
fn auto_filter_reads_current_range_after_cached_sheet_info_edits_and_history() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    workbook.sheet_info().unwrap();
    assert_eq!(workbook.auto_filter(SheetId(0)).unwrap(), None);
    for range in [Some("A1:B3"), Some("A1:C3"), None] {
        workbook
            .apply_ops(
                vec![Op::SetAutoFilter {
                    sheet: SheetId(0),
                    range: range.map(|a1| CellRange::parse_a1(a1).unwrap()),
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        let info = workbook.sheet_info().unwrap();
        assert_eq!(
            workbook.auto_filter(info.active_sheet).unwrap().as_deref(),
            range
        );
        assert_eq!(
            workbook.sheet_info_for(SheetId(0)).unwrap(),
            workbook.sheet_info().unwrap()
        );
        assert_eq!(workbook.auto_filter(SheetId(1)).unwrap(), None);
    }
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(
        workbook.auto_filter(SheetId(0)).unwrap().as_deref(),
        Some("A1:C3")
    );
    workbook.set_active_sheet(SheetId(1)).unwrap();
    assert_eq!(
        workbook
            .auto_filter(workbook.sheet_info().unwrap().active_sheet)
            .unwrap(),
        None
    );
    assert_eq!(
        workbook.auto_filter(SheetId(0)).unwrap().as_deref(),
        Some("A1:C3")
    );
    assert_eq!(workbook.active_sheet(), SheetId(1));
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.auto_filter(SheetId(0)).unwrap(), None);
    assert!(matches!(
        workbook.auto_filter(SheetId(2)),
        Err(Error::SheetOutOfRange(SheetId(2)))
    ));
}

#[test]
fn array_formula_reads_include_anchors_members_and_empty_members() {
    let mut model = model();
    model.sheets[0].set_cell(
        cell("B1"),
        Cell {
            formula: Some("SEQUENCE(2,2)".into()),
            value: CellValue::Number { value: 1.0 },
            ..Cell::default()
        },
    );
    model.sheets[0].set_cell(
        cell("B2"),
        Cell {
            value: CellValue::Number { value: 3.0 },
            ..Cell::default()
        },
    );
    model.sheets[0].set_array_formula(cell("B1"), CellRange::parse_a1("B1:C2").unwrap());
    let bytes = ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap();
    let workbook = Workbook::open_for_read(&bytes).unwrap();
    for a1 in ["B1", "C1", "B2", "C2"] {
        assert!(workbook.is_array_formula(SheetId(0), cell(a1)).unwrap());
    }
    assert_eq!(
        workbook.cell(SheetId(0), cell("B1")).unwrap(),
        CellEdit {
            cell: cell("B1"),
            input: "=SEQUENCE(2,2)".into(),
            is_formula: true
        },
    );
    assert_eq!(
        workbook.cell(SheetId(0), cell("A1")).unwrap(),
        CellEdit {
            cell: cell("A1"),
            input: String::new(),
            is_formula: false
        },
    );
    let range: Vec<Vec<CellEdit>> = workbook
        .range_cells(SheetId(0), CellRange::parse_a1("A1:C2").unwrap())
        .unwrap();
    let CellEdit {
        cell: at,
        input,
        is_formula,
    } = &range[0][0];
    assert!(!workbook.is_array_formula(SheetId(0), *at).unwrap());
    assert!(input.is_empty());
    assert!(!*is_formula);
    assert!(
        workbook
            .is_array_formula(SheetId(0), range[1][2].cell)
            .unwrap()
    );
    assert!(!workbook.is_array_formula(SheetId(1), cell("B1")).unwrap());
    assert!(matches!(
        workbook.is_array_formula(SheetId(2), cell("A1")),
        Err(Error::SheetOutOfRange(SheetId(2)))
    ));
    assert!(matches!(
        workbook.is_array_formula(SheetId(0), CellRef::new(1_048_576, 0)),
        Err(Error::CellOutOfRange(_))
    ));
}

#[test]
fn sheet_info_keeps_its_public_literal_and_destructuring_shape() {
    let workbook = Workbook::from_model(model()).unwrap();
    let SheetInfo {
        sheet_ids,
        sheet_names,
        active_sheet,
        content_width,
        content_height,
        frozen_rows,
        frozen_cols,
        initial_scroll_x,
        initial_scroll_y,
    } = workbook.sheet_info().unwrap();
    let info = SheetInfo {
        sheet_ids,
        sheet_names,
        active_sheet,
        content_width,
        content_height,
        frozen_rows,
        frozen_cols,
        initial_scroll_x,
        initial_scroll_y,
    };
    assert_eq!(info, workbook.sheet_info().unwrap());
}

#[test]
fn calculation_status_formats_limited_addresses_and_omits_an_unspent_budget() {
    let mut model = model();
    model.sheets[1].set_cell(
        cell("XFD1048576"),
        Cell {
            value: CellValue::Number { value: 1.0 },
            ..Cell::default()
        },
    );
    model.sheets[0].set_cell(
        cell("A1"),
        Cell {
            formula: Some("SUM(Other!A1:XFD1048576)".into()),
            ..Cell::default()
        },
    );
    let mut workbook = Workbook::from_model(model).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(workbook.calculation_status().limited_cells, vec!["A1"]);
    assert!(!workbook.calculation_status().budget_exhausted);
    assert_eq!(
        serde_json::to_string(&workbook.calculation_status()).unwrap(),
        r#"{"limitedCells":["A1"]}"#
    );
    workbook.set_active_sheet(SheetId(1)).unwrap();
    assert_eq!(workbook.calculation_status().limited_cells, vec!["Data!A1"]);
}
