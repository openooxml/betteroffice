use betteroffice_xlsx::{
    CalculationOptions, CapturedFormat, Cell, CellInput, CellRange, CellRef, CellValue, Error,
    NumberFormatMutation, ProposalEditInput, ProposalRequest, Sheet, SheetId, StylePatch, Workbook,
    WorkbookModel,
};
use xlsx_model::{CellFormat, Protection, Xf};

const FORM: SheetId = SheetId(0);
const OPEN: SheetId = SheetId(1);
const STYLED: SheetId = SheetId(2);

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn range(a1: &str) -> CellRange {
    CellRange::parse_a1(a1).unwrap()
}

fn put(sheet: &mut Sheet, address: &str, value: f64, style: Option<u32>) {
    sheet.set_cell(
        cell(address),
        Cell {
            value: CellValue::Number { value },
            formula: None,
            style,
        },
    );
}

/// Form is protected with A1 locked and B1:B2 unlocked, Open is unprotected and Styled is
/// protected but lets users format cells.
fn workbook() -> Workbook {
    fixture(true)
}

/// The sheets of [`workbook`], protected or not.
fn fixture(protected: bool) -> Workbook {
    let mut model = WorkbookModel::default();
    model.styles.cell_xfs.push(Xf::default());
    model.styles.cell_xfs.push(Xf {
        protection: Some(Protection {
            locked: false,
            hidden: false,
        }),
        ..Xf::default()
    });
    let mut form = Sheet::new("Form");
    put(&mut form, "A1", 1.0, None);
    put(&mut form, "B1", 2.0, Some(1));
    put(&mut form, "B2", 4.0, Some(1));
    let mut open = Sheet::new("Open");
    put(&mut open, "A1", 3.0, None);
    put(&mut open, "B1", 5.0, None);
    let mut styled = Sheet::new("Styled");
    put(&mut styled, "A1", 6.0, None);
    model.sheets.extend([form, open, styled]);
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let protections = [
        ("xl/worksheets/sheet1.xml", r#"sheet="1""#),
        ("xl/worksheets/sheet3.xml", r#"sheet="1" formatCells="0""#),
    ];
    for (part, attributes) in protections.into_iter().filter(|_| protected) {
        let sheet = parts.iter_mut().find(|(name, _)| name == part).unwrap();
        let xml = String::from_utf8(sheet.1.clone()).unwrap().replace(
            "</sheetData>",
            &format!("</sheetData><sheetProtection {attributes}/>"),
        );
        assert!(xml.contains("sheetProtection"));
        sheet.1 = xml.into_bytes();
    }
    Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap()
}

fn input(book: &Workbook, sheet: SheetId, address: &str) -> String {
    book.cell(sheet, cell(address)).unwrap().input
}

fn locked(result: betteroffice_xlsx::Result<impl std::fmt::Debug>) -> String {
    match result {
        Err(Error::LockedTarget(message)) => message,
        other => panic!("expected a protection refusal, got {other:?}"),
    }
}

fn options() -> CalculationOptions {
    CalculationOptions::default()
}

fn inputs(edits: &[(&str, &str)]) -> Vec<CellInput> {
    edits
        .iter()
        .map(|(address, input)| CellInput {
            cell: cell(address),
            input: (*input).to_owned(),
        })
        .collect()
}

fn bold() -> StylePatch {
    StylePatch {
        bold: Some(true),
        ..Default::default()
    }
}

fn captured() -> CapturedFormat {
    CapturedFormat {
        rows: 1,
        columns: 1,
        formats: vec![CellFormat::default()],
        kept: Vec::new(),
        source_styles: Vec::new(),
        source_formats: Vec::new(),
    }
}

#[test]
fn edit_cell_refuses_a_locked_cell_on_a_protected_sheet() {
    let mut book = workbook();
    let version = book.version();
    let message = locked(book.edit_cell(FORM, cell("A1"), "changed", options()));
    assert_eq!(message, r#"A1 is locked on protected sheet "Form""#);
    assert_eq!(input(&book, FORM, "A1"), "1");
    assert_eq!(book.version(), version);
    assert!(!book.can_undo());
    let unchanged = locked(book.edit_cell(FORM, cell("A1"), "1", options()));
    assert_eq!(unchanged, message);
    let formula = locked(book.edit_cell(FORM, cell("C3"), "=B1*2", options()));
    assert_eq!(formula, r#"C3 is locked on protected sheet "Form""#);
    assert_eq!(input(&book, FORM, "C3"), "");
}

#[test]
fn edit_cell_writes_unlocked_cells_and_unprotected_sheets() {
    let mut book = workbook();
    assert!(
        book.edit_cell(FORM, cell("B1"), "unlocked", options())
            .unwrap()
            .applied
    );
    assert!(
        book.edit_cell(OPEN, cell("A1"), "unprotected", options())
            .unwrap()
            .applied
    );
    assert_eq!(input(&book, FORM, "B1"), "unlocked");
    assert_eq!(input(&book, OPEN, "A1"), "unprotected");
    assert_eq!(input(&book, FORM, "A1"), "1");
}

#[test]
fn edit_cells_refuses_the_whole_write_when_one_cell_is_locked() {
    let mut book = workbook();
    let version = book.version();
    let message = locked(book.edit_cells(
        FORM,
        &inputs(&[("B1", "kept out"), ("B2", ""), ("A1", "")]),
        options(),
    ));
    assert_eq!(message, r#"A1 is locked on protected sheet "Form""#);
    assert_eq!(input(&book, FORM, "A1"), "1");
    assert_eq!(input(&book, FORM, "B1"), "2");
    assert_eq!(input(&book, FORM, "B2"), "4");
    assert_eq!(book.version(), version);
    assert!(
        book.edit_cells(FORM, &inputs(&[("B1", "7"), ("B2", "8")]), options())
            .unwrap()
            .applied
    );
    assert!(
        book.edit_cells(OPEN, &inputs(&[("A1", ""), ("B1", "9")]), options())
            .unwrap()
            .applied
    );
    assert_eq!(
        [
            input(&book, FORM, "B1"),
            input(&book, FORM, "B2"),
            input(&book, OPEN, "A1"),
            input(&book, OPEN, "B1"),
        ],
        ["7", "8", "", "9"]
    );
}

#[test]
fn edit_cells_reports_an_out_of_range_cell_before_a_lock() {
    let mut book = workbook();
    let error = book
        .edit_cells(
            FORM,
            &[
                CellInput {
                    cell: cell("A1"),
                    input: "1".to_owned(),
                },
                CellInput {
                    cell: CellRef::new(u32::MAX, 0),
                    input: "1".to_owned(),
                },
            ],
            options(),
        )
        .unwrap_err();
    assert!(matches!(error, Error::CellOutOfRange(_)), "{error:?}");
}

#[test]
fn formatting_refuses_a_protected_sheet_unless_it_allows_formatting_cells() {
    let mut book = workbook();
    let version = book.version();
    let sheet = r#"sheet "Form" is protected"#;
    assert_eq!(
        locked(book.patch_range_style(FORM, range("B1"), bold(), options())),
        sheet
    );
    assert_eq!(
        locked(book.set_range_number_format(
            FORM,
            range("B1"),
            NumberFormatMutation::Percent,
            options()
        )),
        sheet
    );
    assert_eq!(
        locked(book.apply_format(FORM, range("A1"), captured(), options())),
        sheet
    );
    assert_eq!(book.version(), version);
    assert!(
        !book
            .selection_formatting(FORM, range("B1"))
            .unwrap()
            .bold
            .unwrap_or(false)
    );
    for target in [STYLED, OPEN] {
        assert!(
            book.patch_range_style(target, range("A1"), bold(), options())
                .unwrap()
                .applied
        );
        assert!(
            book.set_range_number_format(
                target,
                range("A1"),
                NumberFormatMutation::Percent,
                options()
            )
            .unwrap()
            .applied
        );
        book.apply_format(target, range("A1"), captured(), options())
            .unwrap();
    }
    assert_eq!(
        locked(book.edit_cell(STYLED, cell("A1"), "content", options())),
        r#"A1 is locked on protected sheet "Styled""#
    );
}

#[test]
fn move_range_refuses_locked_source_or_destination_cells() {
    let mut book = workbook();
    let version = book.version();
    assert_eq!(
        locked(book.move_range(FORM, range("B1"), cell("A1"), options())),
        r#"A1 is locked on protected sheet "Form""#
    );
    assert_eq!(
        locked(book.move_range(FORM, range("A1"), cell("B2"), options())),
        r#"A1 is locked on protected sheet "Form""#
    );
    assert_eq!(book.version(), version);
    assert_eq!(input(&book, FORM, "A1"), "1");
    assert_eq!(input(&book, FORM, "B1"), "2");
    let mut open = fixture(false);
    assert!(
        open.move_range(FORM, range("A1"), cell("C3"), options())
            .unwrap()
            .applied
    );
    assert_eq!(input(&open, FORM, "C3"), "1");
}

#[test]
fn proposals_refuse_locked_cells_on_a_protected_sheet() {
    let mut book = workbook();
    let request = |sheet, address: &str| ProposalRequest {
        agent_id: "agent".to_owned(),
        note: None,
        edits: vec![ProposalEditInput {
            sheet,
            cell: cell(address),
            input: "proposed".to_owned(),
            number_format: None,
        }],
    };
    assert_eq!(
        locked(book.propose(request(FORM, "A1"), options())),
        r#"A1 is locked on protected sheet "Form""#
    );
    assert!(book.proposals().is_empty());
    for (sheet, address) in [(FORM, "B1"), (OPEN, "A1")] {
        let proposal = book.propose(request(sheet, address), options()).unwrap();
        book.accept_proposal(&proposal.id, false, options())
            .unwrap();
        assert_eq!(input(&book, sheet, address), "proposed");
    }
}

#[test]
fn undo_and_redo_replay_an_unlocked_edit_on_a_protected_sheet() {
    let mut book = workbook();
    book.edit_cell(FORM, cell("B1"), "typed", options())
        .unwrap();
    book.undo(options()).unwrap();
    assert_eq!(input(&book, FORM, "B1"), "2");
    book.redo(options()).unwrap();
    assert_eq!(input(&book, FORM, "B1"), "typed");
}
