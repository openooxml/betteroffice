use xlsx_calc::graph::DepGraph;
use xlsx_calc::{EvalContext, Expr, TableBand, evaluate, parse_formula};
use xlsx_model::{
    Cell, CellRange, CellRef, CellValue, ErrorValue, Sheet, SheetId, Table, Workbook,
};

fn number(value: f64) -> Cell {
    Cell {
        value: CellValue::Number { value },
        ..Cell::default()
    }
}

fn label(value: &str) -> Cell {
    Cell {
        value: CellValue::Text {
            value: value.into(),
        },
        ..Cell::default()
    }
}

/// `Sales` over A1:C5: a header row, three data rows, and a totals row.
fn workbook() -> Workbook {
    let mut sheet = Sheet::new("Data");
    for (index, name) in ["Region", "Amount", "Extra\nCost"].iter().enumerate() {
        sheet.set_cell(CellRef::new(0, index as u32), label(name));
    }
    for row in 1..4u32 {
        sheet.set_cell(CellRef::new(row, 0), label("north"));
        sheet.set_cell(CellRef::new(row, 1), number(row as f64 * 10.0));
        sheet.set_cell(CellRef::new(row, 2), number(row as f64));
    }
    sheet.set_cell(CellRef::new(4, 1), number(60.0));
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    workbook.sheets.push(Sheet::new("Report"));
    workbook.tables.push(Table {
        name: "Sales".into(),
        sheet: SheetId(0),
        range: CellRange::parse_a1("A1:C5").unwrap(),
        header_rows: 1,
        totals_rows: 1,
        columns: ["Region", "Amount", "Extra\nCost"]
            .iter()
            .map(|name| name.to_string())
            .collect(),
    });
    workbook
}

fn value(source: &str, at: Option<&str>) -> CellValue {
    let workbook = workbook();
    let mut ctx = EvalContext::new(&workbook, SheetId(0));
    ctx.cell = at.map(|cell| CellRef::parse_a1(cell).unwrap());
    evaluate(&parse_formula(source).unwrap(), &ctx)
}

fn error(source: &str, at: Option<&str>) -> Option<ErrorValue> {
    match value(source, at) {
        CellValue::Error { value } => Some(value),
        _ => None,
    }
}

#[test]
fn resolves_bands_and_columns_as_areas() {
    for (source, expected) in [
        ("SUM(Sales[])", 66.0),
        ("SUM(Sales[Amount])", 60.0),
        ("SUM(Sales[[#Data],[Amount]])", 60.0),
        ("SUM(Sales[[#All],[Amount]])", 120.0),
        ("SUM(Sales[[#Totals],[Amount]])", 60.0),
        ("COUNTA(Sales[[#Headers],[Amount]])", 1.0),
        ("SUM(Sales[[Amount]:[Extra\nCost]])", 66.0),
        ("SUM(Sales[[#All],[Amount]:[Extra\nCost]])", 126.0),
        ("COLUMNS(Sales[])", 3.0),
        ("ROWS(Sales[])", 3.0),
        ("ROWS(Sales[#All])", 5.0),
    ] {
        assert_eq!(
            value(source, None),
            CellValue::Number { value: expected },
            "{source}"
        );
    }
}

#[test]
fn this_row_binds_to_the_formula_cell() {
    assert_eq!(
        value("Sales[[#This Row],[Amount]]", Some("E3")),
        CellValue::Number { value: 20.0 }
    );
    assert_eq!(
        value("Sales[@Amount]", Some("E3")),
        CellValue::Number { value: 20.0 }
    );
    assert_eq!(error("Sales[@Amount]", Some("E9")), Some(ErrorValue::Value));
    assert_eq!(error("Sales[@Amount]", None), Some(ErrorValue::Value));
}

#[test]
fn unknown_tables_columns_and_bands_report_ref() {
    assert_eq!(error("SUM(Missing[Amount])", None), Some(ErrorValue::Ref));
    assert_eq!(error("SUM(Sales[Nope])", None), Some(ErrorValue::Ref));
    assert_eq!(
        error("SUM(Sales[[#Totals],[Nope]])", None),
        Some(ErrorValue::Ref)
    );
    assert_eq!(error("Sales[Amount]", None), Some(ErrorValue::Value));
}

#[test]
fn a_headerless_or_totalless_table_refuses_that_band() {
    let mut workbook = workbook();
    workbook.tables[0].header_rows = 0;
    workbook.tables[0].totals_rows = 0;
    let ctx = EvalContext::new(&workbook, SheetId(0));
    for (source, expected) in [
        ("SUM(Sales[#Headers])", ErrorValue::Ref),
        ("SUM(Sales[#Totals])", ErrorValue::Ref),
    ] {
        assert_eq!(
            evaluate(&parse_formula(source).unwrap(), &ctx),
            CellValue::Error { value: expected },
            "{source}"
        );
    }
    assert_eq!(
        evaluate(&parse_formula("SUM(Sales[])").unwrap(), &ctx),
        CellValue::Number { value: 126.0 }
    );
}

#[test]
fn parses_every_written_form() {
    let Expr::TableRef { table, spec } = parse_formula("Sales[]").unwrap() else {
        panic!("expected a table reference");
    };
    assert_eq!(table, "Sales");
    assert!(spec.bands.is_empty() && spec.first_column.is_none());

    let Expr::TableRef { spec, .. } = parse_formula("Sales[[#This Row],[Amount]]").unwrap() else {
        panic!("expected a table reference");
    };
    assert_eq!(spec.bands, vec![TableBand::ThisRow]);
    assert_eq!(spec.first_column.as_deref(), Some("Amount"));

    let Expr::TableRef { spec, .. } = parse_formula("Sales[[#All],[A]:[B]]").unwrap() else {
        panic!("expected a table reference");
    };
    assert_eq!(spec.bands, vec![TableBand::All]);
    assert_eq!(spec.first_column.as_deref(), Some("A"));
    assert_eq!(spec.last_column.as_deref(), Some("B"));

    let Expr::TableRef { spec, .. } = parse_formula("Sales['[odd'] name]").unwrap() else {
        panic!("expected a table reference");
    };
    assert_eq!(spec.first_column.as_deref(), Some("[odd] name"));
}

#[test]
fn printing_a_structured_reference_re_parses_to_the_same_spec() {
    for source in [
        "Sales[]",
        "Sales[Amount]",
        "Sales[#Headers]",
        "Sales[@Amount]",
        "Sales[[#This Row],[Amount]]",
        "Sales[[#All],[A]:[B]]",
        "SUM(Sales[Amount])+1",
    ] {
        let expression = parse_formula(source).unwrap();
        let printed = expression.to_formula();
        assert_eq!(parse_formula(&printed).unwrap(), expression, "{printed}");
    }
}

#[test]
fn rejects_malformed_structured_references() {
    for source in [
        "Sales[Amount",
        "Sales[[#Nope],[Amount]]",
        "Sales[[A],[B],[C]]",
        "Sales[[A]",
    ] {
        assert!(parse_formula(source).is_err(), "{source}");
    }
}

#[test]
fn the_graph_reads_the_rectangle_a_structured_reference_designates() {
    let mut workbook = workbook();
    workbook.sheets[1].set_cell(
        CellRef::parse_a1("A1").unwrap(),
        Cell {
            formula: Some("SUM(Sales[Amount])".into()),
            ..Cell::default()
        },
    );
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("E3").unwrap(),
        Cell {
            formula: Some("Sales[[#This Row],[Amount]]".into()),
            ..Cell::default()
        },
    );
    let graph = DepGraph::build(&workbook);
    let dependents: Vec<_> = graph
        .dependents_of(SheetId(0), CellRef::parse_a1("B3").unwrap())
        .collect();
    assert!(dependents.contains(&(SheetId(1), CellRef::parse_a1("A1").unwrap())));
    assert!(dependents.contains(&(SheetId(0), CellRef::parse_a1("E3").unwrap())));
    assert!(
        !graph
            .dependents_of(SheetId(0), CellRef::parse_a1("B2").unwrap())
            .any(|node| node == (SheetId(0), CellRef::parse_a1("E3").unwrap()))
    );
}

#[test]
fn index_over_a_table_keeps_a_reference() {
    assert_eq!(
        value("COUNTA(INDEX(Sales[],,2))", None),
        CellValue::Number { value: 3.0 }
    );
    assert_eq!(
        value("SUM(INDEX(Sales[],2,0))", None),
        CellValue::Number { value: 22.0 }
    );
    assert_eq!(
        value("INDEX(Sales[],2,2)", None),
        CellValue::Number { value: 20.0 }
    );
    assert_eq!(error("SUM(INDEX(Sales[],,9))", None), Some(ErrorValue::Ref));
}
