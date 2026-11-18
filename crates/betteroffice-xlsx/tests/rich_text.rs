use std::sync::Arc;

use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, CellState, CellValue, DisplayList, DrawCmd, Op,
    PrintMetrics, RenderOptions, Sheet, SheetId, Viewport, Workbook, WorkbookModel,
};
use xlsx_model::styles::{Alignment, Font, Xf};
use xlsx_model::{Color, RichRun, RunUnderline};
use xlsx_render::{TextRun, TextRunBaseline};

fn source() -> Vec<u8> {
    source_with(
        r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="inlineStr"><is><t>plain</t></is></c><c r="D1" t="inlineStr"><is><r><rPr><i/></rPr><t>Inline</t></r></is></c></row>"#,
        r#"<si><r><rPr><b/></rPr><t>Rich</t></r></si>"#,
    )
}

fn source_with(sheet_data: &str, items: &str) -> Vec<u8> {
    let model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..WorkbookModel::default()
    };
    package(&model, sheet_data, "", items)
}

fn package(model: &WorkbookModel, sheet_data: &str, tail: &str, items: &str) -> Vec<u8> {
    let mut parts = xlsx_parse::serialize_workbook(model).unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    let xml = format!("<worksheet><sheetData>{sheet_data}</sheetData>{tail}</worksheet>");
    sheet.1 = xml.into_bytes();
    let strings = format!("<sst>{items}</sst>").into_bytes();
    match parts
        .iter_mut()
        .find(|(path, _)| path == "xl/sharedStrings.xml")
    {
        Some(part) => part.1 = strings,
        None => parts.push(("xl/sharedStrings.xml".into(), strings)),
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn rich_queries_preserve_no_edit_save_and_plain_cells() {
    let bytes = source();
    let workbook = Workbook::open(&bytes).unwrap();
    let rich = workbook
        .cell_text_runs(SheetId(0), CellRef::new(0, 0))
        .unwrap();
    assert_eq!(rich[0].style.bold, Some(true));
    assert!(std::ptr::eq(
        rich,
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 1))
            .unwrap()
    ));
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 2))
            .is_none()
    );
    assert_eq!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 3))
            .unwrap()[0]
            .style
            .italic,
        Some(true)
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(99), CellRef::new(0, 0))
            .is_none()
    );
    assert_eq!(workbook.save().unwrap(), bytes);
}

#[test]
fn content_writes_do_not_restore_runs_after_undo_or_editing_back() {
    for col in [0, 3] {
        let mut workbook = Workbook::open(&source()).unwrap();
        let at = CellRef::new(0, col);
        assert!(workbook.cell_text_runs(SheetId(0), at).is_some());
        workbook
            .edit_cell(SheetId(0), at, "replacement", CalculationOptions::default())
            .unwrap();
        assert!(workbook.cell_text_runs(SheetId(0), at).is_none());
        workbook.undo(CalculationOptions::default()).unwrap();
        assert!(workbook.cell_text_runs(SheetId(0), at).is_none());
        workbook.redo(CalculationOptions::default()).unwrap();
        workbook
            .edit_cell(
                SheetId(0),
                at,
                if col == 0 { "Rich" } else { "Inline" },
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(workbook.cell_text_runs(SheetId(0), at).is_none());
        assert!(
            workbook
                .cell_text_runs(SheetId(0), CellRef::new(0, 1))
                .is_some()
        );
    }
}

#[test]
fn equal_input_noop_keeps_source_runs() {
    let mut workbook = Workbook::open(&source()).unwrap();
    let result = workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "Rich",
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(!result.applied);
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_some()
    );
}

#[test]
fn committed_equal_content_assignment_invalidates_only_its_source_cell() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook
        .apply_ops(
            vec![
                Op::SetCell {
                    sheet: SheetId(0),
                    at: CellRef::new(0, 0),
                    cell: CellState {
                        value: CellValue::Text {
                            value: "Rich".into(),
                        },
                        ..CellState::default()
                    },
                },
                Op::SetCell {
                    sheet: SheetId(0),
                    at: CellRef::new(1, 0),
                    cell: CellState {
                        value: CellValue::Number { value: 2.0 },
                        ..CellState::default()
                    },
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_none()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 1))
            .is_some()
    );
}

#[test]
fn rejected_transaction_does_not_invalidate_source_runs() {
    let mut workbook = Workbook::open(&source()).unwrap();
    let result = workbook.apply_ops(
        vec![
            Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(0, 0),
                cell: CellState {
                    value: CellValue::Text {
                        value: "replacement".into(),
                    },
                    ..CellState::default()
                },
            },
            Op::SetCell {
                sheet: SheetId(99),
                at: CellRef::new(0, 0),
                cell: CellState::default(),
            },
        ],
        CalculationOptions::default(),
    );
    assert!(result.is_err());
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_some()
    );
}

#[test]
fn source_bindings_follow_rows_and_sheets() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 2,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    for col in [0, 3] {
        assert!(
            workbook
                .cell_text_runs(SheetId(0), CellRef::new(0, col))
                .is_none()
        );
        assert!(
            workbook
                .cell_text_runs(SheetId(0), CellRef::new(2, col))
                .is_some()
        );
    }
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(2, 3),
            "changed",
            CalculationOptions::default(),
        )
        .unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_some()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 3))
            .is_none()
    );
}

#[test]
fn collaborative_sessions_return_no_runs() {
    let workbook = Workbook::open_collaborative(&source(), 7).unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_none()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 3))
            .is_none()
    );
}

#[test]
fn style_only_edits_preserve_runs_and_sheet_reordering_keeps_source_identity() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook
        .apply_ops(
            vec![Op::PatchRangeStyle {
                sheet: SheetId(0),
                range: betteroffice_xlsx::CellRange::new(CellRef::new(0, 0), CellRef::new(0, 3)),
                patch: betteroffice_xlsx::StylePatch {
                    bold: Some(false),
                    ..Default::default()
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_some()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 3))
            .is_some()
    );
    workbook
        .apply_ops(
            vec![
                Op::AddSheet {
                    index: 1,
                    name: "Other".into(),
                },
                Op::MoveSheet { from: 0, to: 1 },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_none()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(1), CellRef::new(0, 0))
            .is_some()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(1), CellRef::new(0, 3))
            .is_some()
    );
}

#[test]
fn recovery_without_run_provenance_stays_plain() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "changed",
            CalculationOptions::default(),
        )
        .unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let (restored, _, _, _) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert!(
        restored
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_none()
    );
    assert!(
        restored
            .cell_text_runs(SheetId(0), CellRef::new(0, 3))
            .is_none()
    );
}

const EXCEL_ITEMS: [&str; 4] = [
    r#"<si><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>Bold</t></r><r><rPr><sz val="11"/><color rgb="FFFF0000"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve"> red</t></r></si>"#,
    r#"<si><r><rPr><sz val="10"/><rFont val="NanumGothic"/><family val="3"/><charset val="129"/></rPr><t>합계</t></r><r><rPr><b/><sz val="10"/><rFont val="NanumGothic"/><family val="3"/><charset val="129"/></rPr><t xml:space="preserve"> 금액</t></r><phoneticPr fontId="1" type="noConversion"/></si>"#,
    r#"<si><r><rPr><b/><condense val="0"/><extend val="0"/><outline/><shadow/></rPr><t>Shaped</t></r><r><rPr><i/></rPr><t xml:space="preserve"> text</t></r></si>"#,
    r#"<si><r><rPr><b/></rPr><t>漢</t></r><r><t>字</t></r><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="1"/></si>"#,
];

const EXCEL_ROW: &str = concat!(
    r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c>"#,
    r#"<c r="E1" t="inlineStr"><is><r><rPr><i/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>Inline</t></r><r><rPr><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve"> text</t></r></is></c>"#,
    r#"<c r="F1" t="inlineStr"><is><r><rPr><shadow/><u/></rPr><t>Under</t></r><r><t>lined</t></r></is></c>"#,
    r#"<c r="G1" t="inlineStr"><is><r><rPr><rFont val="MS PGothic"/><charset val="128"/></rPr><t>字</t></r><rPh sb="0" eb="1"><t>じ</t></rPh></is></c></row>"#,
);

fn excel_source() -> Vec<u8> {
    source_with(EXCEL_ROW, &EXCEL_ITEMS.concat())
}

fn text_at(workbook: &Workbook, col: u32) -> &str {
    let cell = workbook.model().sheets[0].cell(CellRef::new(0, col));
    match cell.map(|cell| &cell.value) {
        Some(CellValue::Text { value }) => value.as_str(),
        other => panic!("expected text at column {col}, got {other:?}"),
    }
}

fn cell_runs(workbook: &Workbook, col: u32) -> &[RichRun] {
    workbook
        .cell_text_runs(SheetId(0), CellRef::new(0, col))
        .unwrap_or_else(|| panic!("no runs at column {col}"))
}

fn run_texts(workbook: &Workbook, col: u32) -> Vec<&str> {
    let text = text_at(workbook, col);
    cell_runs(workbook, col)
        .iter()
        .map(|run| &text[run.start as usize..run.end as usize])
        .collect()
}

fn part(package: &[u8], path: &str) -> String {
    let parts = ooxml_opc::unzip_parts(package).unwrap();
    let (_, bytes) = parts.into_iter().find(|(name, _)| name == path).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn excel_written_runs_with_family_and_scheme_are_rich() {
    let workbook = Workbook::open(&excel_source()).unwrap();
    assert_eq!(run_texts(&workbook, 0), ["Bold", " red"]);
    let runs = cell_runs(&workbook, 0);
    assert_eq!(runs[0].style.bold, Some(true));
    assert_eq!(runs[0].style.name.as_deref(), Some("Calibri"));
    assert_eq!(runs[0].style.size_pt, Some(11.0));
    assert_eq!(
        runs[0].style.color,
        Some(Color::Theme { idx: 1, tint: 0.0 })
    );
    assert_eq!(runs[1].style.bold, None);
    assert_eq!(runs[1].style.color, Some(Color::Rgb("#ff0000".into())));
    assert_eq!(run_texts(&workbook, 4), ["Inline", " text"]);
    let runs = cell_runs(&workbook, 4);
    assert_eq!(runs[0].style.italic, Some(true));
    assert_eq!(runs[1].style.italic, None);
    assert_eq!(runs[1].style.name.as_deref(), Some("Calibri"));
}

#[test]
fn korean_runs_with_charset_are_rich() {
    let workbook = Workbook::open(&excel_source()).unwrap();
    assert_eq!(run_texts(&workbook, 1), ["합계", " 금액"]);
    let runs = cell_runs(&workbook, 1);
    assert_eq!(runs[0].style.name.as_deref(), Some("NanumGothic"));
    assert_eq!(runs[0].style.bold, None);
    assert_eq!(runs[1].style.bold, Some(true));
    assert_eq!(runs[1].style.size_pt, Some(10.0));
}

#[test]
fn unknown_run_properties_keep_the_run_and_its_known_properties() {
    let workbook = Workbook::open(&excel_source()).unwrap();
    assert_eq!(run_texts(&workbook, 2), ["Shaped", " text"]);
    let shaped = cell_runs(&workbook, 2);
    assert_eq!(shaped[0].style.bold, Some(true));
    assert_eq!(shaped[1].style.italic, Some(true));
    assert_eq!(run_texts(&workbook, 5), ["Under", "lined"]);
    let underlined = cell_runs(&workbook, 5);
    assert_eq!(underlined[0].style.underline, Some(RunUnderline::Single));
    assert_eq!(underlined[1].style.underline, None);
}

#[test]
fn phonetic_runs_keep_the_base_runs() {
    let workbook = Workbook::open(&excel_source()).unwrap();
    assert_eq!(text_at(&workbook, 3), "漢字かんじ");
    assert_eq!(run_texts(&workbook, 3), ["漢", "字"]);
    assert_eq!(cell_runs(&workbook, 3)[0].style.bold, Some(true));
    assert_eq!(text_at(&workbook, 6), "字じ");
    assert_eq!(run_texts(&workbook, 6), ["字"]);
    let gothic = cell_runs(&workbook, 6);
    assert_eq!(gothic[0].style.name.as_deref(), Some("MS PGothic"));
}

#[test]
fn no_edit_save_keeps_excel_written_items_byte_for_byte() {
    let bytes = excel_source();
    let workbook = Workbook::open(&bytes).unwrap();
    let before: Vec<Vec<RichRun>> = (0..7)
        .map(|col| cell_runs(&workbook, col).to_vec())
        .collect();
    let saved = workbook.save().unwrap();
    let strings = part(&saved, "xl/sharedStrings.xml");
    for item in EXCEL_ITEMS {
        assert!(strings.contains(item), "{item}");
    }
    assert_eq!(strings, part(&bytes, "xl/sharedStrings.xml"));
    assert_eq!(saved, bytes);
    let reopened = Workbook::open(&saved).unwrap();
    for (col, runs) in (0..7).zip(&before) {
        assert_eq!(cell_runs(&reopened, col), runs.as_slice(), "{col}");
    }
}

#[test]
fn editing_another_cell_keeps_the_rich_items_and_their_runs() {
    let mut workbook = Workbook::open(&excel_source()).unwrap();
    let before: Vec<Vec<RichRun>> = (0..4)
        .map(|col| cell_runs(&workbook, col).to_vec())
        .collect();
    let at = CellRef::new(1, 0);
    let receipt = workbook
        .edit_cell(SheetId(0), at, "2", CalculationOptions::default())
        .unwrap();
    assert_eq!(
        receipt,
        betteroffice_xlsx::MutationResult {
            applied: true,
            ..Default::default()
        }
    );
    for (col, runs) in (0..4).zip(&before) {
        assert_eq!(cell_runs(&workbook, col), runs.as_slice(), "{col}");
    }
    let saved = workbook.save().unwrap();
    let strings = part(&saved, "xl/sharedStrings.xml");
    for item in EXCEL_ITEMS {
        assert!(strings.contains(item), "{item}");
    }
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().sheets[0].cell(at).map(|cell| &cell.value),
        Some(&CellValue::Number { value: 2.0 })
    );
    for (col, runs) in (0..4).zip(&before) {
        assert_eq!(cell_runs(&reopened, col), runs.as_slice(), "{col}");
    }
}

#[test]
fn editing_an_excel_written_rich_cell_writes_plain_text() {
    for col in [0, 4] {
        let mut workbook = Workbook::open(&excel_source()).unwrap();
        let at = CellRef::new(0, col);
        workbook
            .edit_cell(SheetId(0), at, "replacement", CalculationOptions::default())
            .unwrap();
        let saved = workbook.save().unwrap();
        let sheet = part(&saved, "xl/worksheets/sheet1.xml");
        let markup = sheet + &part(&saved, "xl/sharedStrings.xml");
        assert!(
            markup.contains(r#"<si><t xml:space="preserve">replacement</t></si>"#)
                || markup.contains(r#"<is><t xml:space="preserve">replacement</t></is>"#),
            "{markup}"
        );
        let reopened = Workbook::open(&saved).unwrap();
        assert_eq!(text_at(&reopened, col), "replacement");
        assert!(reopened.cell_text_runs(SheetId(0), at).is_none());
    }
}

const MIXED_ITEM: &str = concat!(
    r#"<si><r><rPr><b/><i/><sz val="14"/><rFont val="Cambria"/></rPr><t>Ab</t></r>"#,
    r#"<r><rPr><u val="double"/><strike/><color rgb="FFFF0000"/></rPr>"#,
    r#"<t xml:space="preserve"> 中</t></r>"#,
    r#"<r><rPr><vertAlign val="superscript"/><u/></rPr><t>😀</t></r>"#,
    r#"<r><rPr><vertAlign val="subscript"/><u val="none"/></rPr><t>z</t></r></si>"#,
);
const MIXED_TEXT: &str = "Ab 中😀z";
const RIGHT_TO_LEFT_ITEM: &str =
    r#"<si><r><rPr><b/></rPr><t>שלום</t></r><r><t xml:space="preserve"> abc</t></r></si>"#;
const TURNED_ITEM: &str = r#"<si><r><rPr><b/></rPr><t>Up</t></r><r><t>right</t></r></si>"#;

const DISPLAY_ROW: &str = concat!(
    r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c>"#,
    r#"<c r="E1" t="s"><v>2</v></c><c r="G1" t="inlineStr"><is><t>plain</t></is></c>"#,
    r#"<c r="I1" s="1" t="s"><v>3</v></c><c r="K1" s="2" t="s"><v>3</v></c>"#,
    r#"<c r="M1" t="str"><f>"ok"</f><v>ok</v></c>"#,
    r#"<c r="O1" t="inlineStr"><is><r><rPr><b/></rPr><t>Inline</t></r>"#,
    r#"<r><t xml:space="preserve"> text</t></r></is></c></row>"#,
);

const SCREEN: Viewport = Viewport {
    x: 0.0,
    y: 0.0,
    width: 1200.0,
    height: 100.0,
};

fn styled_model() -> WorkbookModel {
    let mut model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..WorkbookModel::default()
    };
    model.styles.fonts.push(Font {
        name: Some("Calibri".into()),
        size_pt: Some(11.0),
        ..Font::default()
    });
    let turned = |degrees| Xf {
        font: Some(0),
        alignment: Some(Alignment {
            text_rotation: Some(degrees),
            ..Alignment::default()
        }),
        ..Xf::default()
    };
    model.styles.cell_xfs.extend([
        Xf {
            font: Some(0),
            ..Xf::default()
        },
        turned(45),
        turned(255),
    ]);
    model
}

fn display_source() -> Vec<u8> {
    let items = [MIXED_ITEM, EXCEL_ITEMS[3], RIGHT_TO_LEFT_ITEM, TURNED_ITEM].concat();
    package(&styled_model(), DISPLAY_ROW, "", &items)
}

fn print_metrics() -> PrintMetrics {
    PrintMetrics {
        dpi: 96.0,
        max_digit_width: 7.0,
        default_row_height_pt: 15.0,
        default_column_width: None,
        font_size_pt: 11.0,
        font_family: "Calibri".into(),
        font_ascent: 10.0,
        font_descent: 3.0,
    }
}

fn text_runs(list: &DisplayList, wanted: &str) -> Vec<Vec<TextRun>> {
    list.commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::Text { text, runs, .. } if &**text == wanted => Some(runs.clone()),
            _ => None,
        })
        .collect()
}

fn mixed_runs() -> Vec<TextRun> {
    vec![
        TextRun {
            start: 0,
            end: 2,
            font_family: Some("Cambria".into()),
            font_size: Some(14.0),
            bold: Some(true),
            italic: Some(true),
            ..TextRun::default()
        },
        TextRun {
            start: 2,
            end: 4,
            underline: Some(true),
            underline_style: Some(RunUnderline::Double),
            strike: Some(true),
            color: Some("#ff0000".into()),
            ..TextRun::default()
        },
        TextRun {
            start: 4,
            end: 6,
            underline: Some(true),
            baseline: Some(TextRunBaseline::Superscript),
            ..TextRun::default()
        },
        TextRun {
            start: 6,
            end: 7,
            underline: Some(false),
            baseline: Some(TextRunBaseline::Subscript),
            ..TextRun::default()
        },
    ]
}

#[test]
fn rich_cells_carry_their_runs_on_the_screen_print_and_png_paths() {
    let workbook = Workbook::open(&display_source()).unwrap();
    let screen = workbook.display_list_for(SheetId(0), &SCREEN).unwrap();
    let (png, _, _) = workbook
        .range_png_display_list_for(SheetId(0), &RenderOptions::default())
        .unwrap();
    let range = CellRange::new(CellRef::new(0, 0), CellRef::new(0, 14));
    let print = workbook
        .print_display_list(SheetId(0), range, &print_metrics(), false)
        .unwrap();
    let inline = vec![
        TextRun {
            start: 0,
            end: 6,
            bold: Some(true),
            ..TextRun::default()
        },
        TextRun {
            start: 6,
            end: 11,
            ..TextRun::default()
        },
    ];
    for (path, list) in [("screen", &screen), ("png", &png), ("print", &print)] {
        assert_eq!(text_runs(list, MIXED_TEXT), [mixed_runs()], "{path}");
        assert_eq!(text_runs(list, "Inline text"), [inline.clone()], "{path}");
        assert_eq!(text_runs(list, "plain"), [Vec::new()], "{path}");
    }
}

#[test]
fn the_phonetic_tail_is_covered_by_no_run() {
    let workbook = Workbook::open(&display_source()).unwrap();
    let screen = workbook.display_list_for(SheetId(0), &SCREEN).unwrap();
    let runs = text_runs(&screen, "漢字かんじ");
    let spans: Vec<Vec<(u32, u32, Option<bool>)>> = runs
        .iter()
        .map(|runs| {
            runs.iter()
                .map(|run| (run.start, run.end, run.bold))
                .collect()
        })
        .collect();
    assert_eq!(spans, [vec![(0, 1, Some(true)), (1, 2, None)]]);
}

#[test]
fn cells_that_do_not_paint_their_stored_text_as_it_is_carry_no_runs() {
    let workbook = Workbook::open(&display_source()).unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 4))
            .is_some()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 8))
            .is_some()
    );
    let screen = workbook.display_list_for(SheetId(0), &SCREEN).unwrap();
    assert_eq!(text_runs(&screen, "שלום abc"), [Vec::new()]);
    assert_eq!(text_runs(&screen, "Upright"), [Vec::new(), Vec::new()]);
    assert_eq!(text_runs(&screen, "ok"), [Vec::new()]);
    assert_eq!(text_runs(&screen, MIXED_TEXT).len(), 1);
}

#[test]
fn a_hyperlink_cell_keeps_the_link_colour_and_underline_over_its_runs() {
    let row = concat!(
        r#"<row r="1"><c r="A1" t="s"><v>0</v></c>"#,
        r#"<c r="C1" t="inlineStr"><is><t>link</t></is></c></row>"#,
    );
    let links = concat!(
        r#"<hyperlinks><hyperlink ref="A1" location="Data!C3"/>"#,
        r#"<hyperlink ref="C1" location="Data!C3"/></hyperlinks>"#,
    );
    let workbook = Workbook::open(&package(&styled_model(), row, links, MIXED_ITEM)).unwrap();
    let screen = workbook.display_list_for(SheetId(0), &SCREEN).unwrap();
    let kept: Vec<TextRun> = mixed_runs()
        .into_iter()
        .map(|run| TextRun {
            underline: None,
            underline_style: None,
            color: None,
            ..run
        })
        .collect();
    assert_eq!(text_runs(&screen, MIXED_TEXT), [kept]);
    let look: Vec<(Arc<str>, bool)> = screen
        .commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::Text {
                color, underline, ..
            } => Some((color.clone(), *underline)),
            _ => None,
        })
        .collect();
    assert_eq!(look.len(), 2);
    assert_eq!(look[0], look[1]);
    assert!(look[0].1);
    assert_ne!(&*look[0].0, "#000000");
}
