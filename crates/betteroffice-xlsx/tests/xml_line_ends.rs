use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, CellState, CellValue, DefinedName, Hyperlink, Op,
    Sheet, SheetId, StylePatch, Workbook, WorkbookModel,
};
use xlsx_model::CellProvider;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let mut xml =
        format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}">"#);
    for (id, kind, target) in entries {
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{DOC_RELS}/{kind}" Target="{target}"/>"#
        ));
    }
    xml + "</Relationships>"
}

/// Table `Sales` names its second column `Price_x000a_[GBP]`; the formulas spell
/// that line break as `eol` and carry the values Excel cached for them. Row 5
/// holds cell text with a CR LF, a lone CR and an escaped CR, a rich shared
/// string whose first run ends in a CR LF, shared strings with `_x000D_` and a
/// literal `_x005F_x000D_`, and formulas holding or returning a CR. A1 links
/// with a two-line tooltip.
fn book(eol: &str) -> Vec<u8> {
    let price = format!("Sales[[#This Row],[Price{eol}'[GBP']]]");
    let sheet = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{main}" xmlns:r="{rels}"><sheetData>"#,
            r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Year</t></is></c><c r="B1" t="inlineStr"><is><t>Price&#10;[GBP]</t></is></c><c r="C1" t="inlineStr"><is><t>Double</t></is></c><c r="E1"><f>SUM(Sales[Price{eol}'[GBP']])</f><v>30</v></c></row>"#,
            r#"<row r="2"><c r="A2"><v>2003</v></c><c r="B2"><v>10</v></c><c r="C2"><f>{price}*2</f><v>20</v></c></row>"#,
            r#"<row r="3"><c r="A3"><v>2004</v></c><c r="B3"><v>20</v></c><c r="C3"><f>{price}*2</f><v>40</v></c></row>"#,
            "<row r=\"5\"><c r=\"A5\" t=\"inlineStr\"><is><t>a\r\nb</t></is></c><c r=\"B5\" t=\"inlineStr\"><is><t>c\rd</t></is></c><c r=\"C5\" t=\"inlineStr\"><is><t>e&#13;&#10;f</t></is></c><c r=\"D5\" t=\"s\"><v>0</v></c>",
            r#"<c r="E5" t="s"><v>1</v></c><c r="F5" t="s"><v>2</v></c><c r="G5" t="s"><v>3</v></c><c r="H5"><f>CODE(MID("p&#13;q",2,1))</f><v>13</v></c><c r="I5" t="str"><f>"m"&amp;CHAR(13)&amp;"n"</f><v>m&#13;n</v></c></row>"#,
            r#"<row r="6"><c r="A6"><f>LEN(A5)</f><v>3</v></c><c r="B6"><f>LEN(B5)</f><v>3</v></c><c r="C6"><f>LEN(C5)</f><v>4</v></c><c r="D6"><f>LEN(D5)</f><v>3</v></c><c r="E6"><f>LEN(E5)</f><v>3</v></c><c r="F6"><f>LEN(F5)</f><v>9</v></c><c r="G6"><f>LEN(G5)</f><v>3</v></c></row>"#,
            r#"<row r="8"><c r="A8" t="s"><v>4</v></c><c r="B8" t="s"><v>5</v></c><c r="D8" t="inlineStr"><is><r><t>_</t></r><r><t>x0041_</t></r></is></c></row>"#,
            r#"</sheetData><hyperlinks><hyperlink ref="A1" location="Data!A2" tooltip="t&#13;&#10;u" display="Year"/></hyperlinks><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts></worksheet>"#
        ),
        main = MAIN,
        rels = DOC_RELS,
        eol = eol,
        price = price,
    );
    let parts = vec![
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "sharedStrings", "sharedStrings.xml"),
            ]),
        ),
        ("xl/worksheets/sheet1.xml", sheet),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationships(&[("rIdTable", "table", "../tables/table1.xml")]),
        ),
        (
            "xl/tables/table1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><table xmlns="{MAIN}" id="1" name="Sales" displayName="Sales" ref="A1:C3"><tableColumns count="3"><tableColumn id="1" name="Year"/><tableColumn id="2" name="Price_x000a_[GBP]"/><tableColumn id="3" name="Double"/></tableColumns></table>"#
            ),
        ),
        (
            "xl/sharedStrings.xml",
            format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"{MAIN}\" count=\"6\" uniqueCount=\"6\"><si><r><t>g\r\n</t></r><r><rPr><b/></rPr><t>h</t></r></si><si><t>i_x000D_j</t></si><si><t>k_x005F_x000D_l</t></si><si><r><t>v_x000D_</t></r><r><rPr><b/></rPr><t>w</t></r></si><si><t>a_x0001_b</t></si><si><t>c_x0007_d</t></si></sst>"
            ),
        ),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn recalculated(bytes: &[u8]) -> Workbook {
    let mut workbook = Workbook::open_for_read(bytes).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    workbook
}

fn values(workbook: &Workbook, cells: &[&str]) -> Vec<CellValue> {
    cells
        .iter()
        .map(|at| workbook.model().value(SheetId(0), cell(at)))
        .collect()
}

fn table_values(workbook: &Workbook) -> Vec<CellValue> {
    values(workbook, &["C2", "C3", "E1"])
}

fn part(bytes: &[u8], name: &str) -> Vec<u8> {
    ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == name)
        .unwrap()
        .1
}

#[test]
fn a_line_feed_in_formula_text_matches_a_multi_line_table_column() {
    let workbook = recalculated(&book("\n"));
    assert_eq!(
        table_values(&workbook),
        vec![number(20.0), number(40.0), number(30.0)]
    );
}

#[test]
fn a_literal_crlf_in_formula_text_reads_as_a_line_feed() {
    let workbook = recalculated(&book("\r\n"));
    assert_eq!(
        table_values(&workbook),
        vec![number(20.0), number(40.0), number(30.0)]
    );
    let formula = workbook.model().sheets[0]
        .cell(cell("C2"))
        .and_then(|c| c.formula.clone())
        .unwrap();
    assert_eq!(formula, "Sales[[#This Row],[Price\n'[GBP']]]*2");
}

#[test]
fn a_lone_cr_in_formula_text_reads_as_a_line_feed() {
    let workbook = recalculated(&book("\r"));
    assert_eq!(
        table_values(&workbook),
        vec![number(20.0), number(40.0), number(30.0)]
    );
}

#[test]
fn editing_a_precedent_recalculates_a_crlf_table_formula() {
    let mut workbook = Workbook::open(&book("\r\n")).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("B2"), "11", CalculationOptions::default())
        .unwrap();
    assert_eq!(
        table_values(&workbook),
        vec![number(22.0), number(40.0), number(31.0)]
    );
}

#[test]
fn a_crlf_table_formula_still_resolves_after_save_and_reopen() {
    let mut workbook = Workbook::open(&book("\r\n")).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("B3"), "21", CalculationOptions::default())
        .unwrap();
    let reopened = recalculated(&workbook.save().unwrap());
    assert_eq!(
        table_values(&reopened),
        vec![number(20.0), number(42.0), number(31.0)]
    );
}

#[test]
fn literal_line_ends_in_cell_text_read_as_line_feeds_and_an_escaped_cr_stays() {
    let workbook = recalculated(&book("\n"));
    assert_eq!(
        values(&workbook, &["A5", "B5", "C5", "D5"]),
        vec![text("a\nb"), text("c\nd"), text("e\r\nf"), text("g\nh")]
    );
    assert_eq!(
        values(&workbook, &["A6", "B6", "C6", "D6"]),
        vec![number(3.0), number(3.0), number(4.0), number(3.0)]
    );
}

#[test]
fn a_rich_shared_string_with_a_crlf_keeps_its_runs() {
    let workbook = Workbook::open(&book("\n")).unwrap();
    let runs = workbook
        .cell_text_runs(SheetId(0), cell("D5"))
        .expect("rich runs survive line-end normalization");
    assert_eq!(runs.len(), 2);
    assert_eq!(runs[1].style.bold, Some(true));
}

#[test]
fn saving_untouched_keeps_every_byte_and_an_edit_keeps_unrelated_text() {
    for eol in ["\n", "\r\n"] {
        let source = book(eol);
        assert_eq!(Workbook::open(&source).unwrap().save().unwrap(), source);
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .edit_cell(SheetId(0), cell("H9"), "7", CalculationOptions::default())
            .unwrap();
        let saved = workbook.save().unwrap();
        assert_eq!(
            part(&saved, "xl/sharedStrings.xml"),
            part(&source, "xl/sharedStrings.xml")
        );
        let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml")).unwrap();
        assert!(sheet.contains("<t>a\r\nb</t>"), "{sheet}");
        assert!(
            sheet.contains(&format!("[Price{eol}'[GBP']]]*2</f>")),
            "{sheet}"
        );
    }
}

const CR_CELLS: [&str; 10] = ["C5", "E5", "F5", "G5", "H5", "I5", "C6", "E6", "F6", "G6"];

fn cr_values() -> Vec<CellValue> {
    vec![
        text("e\r\nf"),
        text("i\rj"),
        text("k_x000D_l"),
        text("v\rw"),
        number(13.0),
        text("m\rn"),
        number(4.0),
        number(3.0),
        number(9.0),
        number(3.0),
    ]
}

fn reopened(saved: &[u8]) -> (Workbook, Workbook) {
    (Workbook::open(saved).unwrap(), recalculated(saved))
}

fn assert_cr_values(saved: &[u8]) {
    let (cached, recalculated) = reopened(saved);
    assert_eq!(values(&cached, &CR_CELLS), cr_values());
    assert_eq!(values(&recalculated, &CR_CELLS), cr_values());
    assert_eq!(
        cached.model().sheets[0].hyperlinks[0].tooltip.as_deref(),
        Some("t\r\nu")
    );
}

fn bold(workbook: &mut Workbook, range: &str) {
    workbook
        .patch_range_style(
            SheetId(0),
            CellRange::parse_a1(range).unwrap(),
            StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
}

#[test]
fn a_cr_reads_from_xstring_and_character_escapes_and_a_literal_escape_stays_text() {
    let workbook = recalculated(&book("\n"));
    assert_eq!(values(&workbook, &CR_CELLS), cr_values());
    let runs = workbook
        .cell_text_runs(SheetId(0), cell("G5"))
        .expect("runs over decoded text");
    assert_eq!(runs.len(), 2);
}

#[test]
fn a_cr_survives_editing_another_cell_saving_and_reopening() {
    let mut workbook = Workbook::open(&book("\n")).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("H9"), "7", CalculationOptions::default())
        .unwrap();
    assert_cr_values(&workbook.save().unwrap());
}

#[test]
fn a_cr_survives_restyling_its_own_cell_saving_and_reopening() {
    let mut workbook = Workbook::open(&book("\n")).unwrap();
    bold(&mut workbook, "C5:I6");
    let saved = workbook.save().unwrap();
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml")).unwrap();
    assert!(!sheet.contains("e\r\nf"), "{sheet}");
    assert_cr_values(&saved);
}

#[test]
fn a_cr_survives_a_hyperlink_rewrite() {
    let mut workbook = Workbook::open(&book("\n")).unwrap();
    workbook
        .set_hyperlink(
            SheetId(0),
            Hyperlink {
                range: CellRange::parse_a1("B1").unwrap(),
                external_target: None,
                location: Some("Data!A3".into()),
                tooltip: Some("x\ry\tz".into()),
                display: None,
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let (cached, _) = reopened(&workbook.save().unwrap());
    let tooltips = cached.model().sheets[0]
        .hyperlinks
        .iter()
        .map(|link| link.tooltip.clone())
        .collect::<Vec<_>>();
    assert_eq!(
        tooltips,
        vec![Some("t\r\nu".to_owned()), Some("x\ry\tz".to_owned())]
    );
}

#[test]
fn typed_text_with_a_cr_or_a_literal_escape_survives_save_restyle_and_reopen() {
    let typed = [
        ("J5", "a\rb"),
        ("K5", "_x000D_"),
        ("L5", "x\r\ny"),
        ("M5", "_X005f_"),
    ];
    let mut workbook = Workbook::open(&book("\n")).unwrap();
    for (at, input) in typed {
        workbook
            .edit_cell(SheetId(0), cell(at), input, CalculationOptions::default())
            .unwrap();
    }
    let saved = workbook.save().unwrap();
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml")).unwrap();
    for written in [
        "<t xml:space=\"preserve\">a_x000D_b</t>",
        "<t xml:space=\"preserve\">_x005F_x000D_</t>",
        "<t xml:space=\"preserve\">x_x000D_\ny</t>",
        "<t xml:space=\"preserve\">_x005F_X005f_</t>",
    ] {
        assert!(sheet.contains(written), "{written} in {sheet}");
    }
    let expected = typed.map(|(_, input)| text(input)).to_vec();
    let cells = typed.map(|(at, _)| at);
    let (mut cached, _) = reopened(&saved);
    assert_eq!(values(&cached, &cells), expected);
    bold(&mut cached, "J5:M5");
    let (restyled, _) = reopened(&cached.save().unwrap());
    assert_eq!(values(&restyled, &cells), expected);
}

#[test]
fn control_characters_from_xstring_escapes_round_trip_through_cached_values() {
    let mut workbook = Workbook::open(&book("\n")).unwrap();
    for (at, input) in [("C8", "=A8"), ("E8", "=B8")] {
        workbook
            .edit_cell(SheetId(0), cell(at), input, CalculationOptions::default())
            .unwrap();
    }
    let expected = vec![
        text("a\u{1}b"),
        text("c\u{7}d"),
        text("a\u{1}b"),
        text("c\u{7}d"),
    ];
    assert_eq!(values(&workbook, &["A8", "B8", "C8", "E8"]), expected);
    let saved = workbook.save().unwrap();
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml")).unwrap();
    assert!(sheet.contains("<v>a_x0001_b</v>"), "{sheet}");
    assert!(sheet.contains("<v>c_x0007_d</v>"), "{sheet}");
    let (cached, recalculated) = reopened(&saved);
    assert_eq!(values(&cached, &["A8", "B8", "C8", "E8"]), expected);
    assert_eq!(values(&recalculated, &["A8", "B8", "C8", "E8"]), expected);
}

#[test]
fn an_escape_split_across_runs_stays_literal_text_through_a_restyle() {
    let mut workbook = Workbook::open(&book("\n")).unwrap();
    assert_eq!(values(&workbook, &["D8"]), vec![text("_x0041_")]);
    assert_eq!(
        workbook
            .cell_text_runs(SheetId(0), cell("D8"))
            .map(<[_]>::len),
        Some(2)
    );
    bold(&mut workbook, "D8");
    let (cached, _) = reopened(&workbook.save().unwrap());
    assert_eq!(values(&cached, &["D8"]), vec![text("_x0041_")]);
}

#[test]
fn a_typed_or_applied_new_formula_writes_its_control_character_with_char() {
    let mut workbook = Workbook::open(&book("\n")).unwrap();
    let formula = "\"a\u{1}b\"";
    workbook
        .edit_cell(
            SheetId(0),
            cell("H9"),
            &format!("={formula}"),
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .apply_ops(
            vec![Op::SetCell {
                sheet: SheetId(0),
                at: cell("I9"),
                cell: CellState {
                    formula: Some(formula.to_owned()),
                    ..CellState::default()
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml")).unwrap();
    assert_eq!(
        sheet
            .matches("<f>(&quot;a&quot;&amp;CHAR(1)&amp;&quot;b&quot;)</f>")
            .count(),
        2,
        "{sheet}"
    );
    let reopened = Workbook::open(&saved).unwrap();
    for at in ["H9", "I9"] {
        assert_eq!(
            reopened.model().sheets[0]
                .cell(cell(at))
                .and_then(|cell| cell.formula.clone()),
            Some("(\"a\"&CHAR(1)&\"b\")".to_owned()),
            "{at}"
        );
        assert_eq!(values(&reopened, &[at]), vec![text("a\u{1}b")], "{at}");
    }
    let error = workbook
        .edit_cell(
            SheetId(0),
            cell("J9"),
            "=\"a\u{0}\"",
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert_eq!(
        error.to_string(),
        "formula contains U+0000, which a workbook file cannot store and CHAR cannot produce"
    );
}

#[test]
fn a_formula_or_name_with_a_control_character_saves_it_as_a_character_reference() {
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        cell("A1"),
        xlsx_model::Cell {
            value: text("ab"),
            formula: Some("\"a\u{1}b\"".into()),
            style: None,
        },
    );
    let mut model = WorkbookModel::default();
    model.sheets.push(sheet);
    model.defined_names.push(DefinedName {
        name: "Carried".into(),
        formula: "\"a\u{1}\"".into(),
        local_sheet: None,
        hidden: false,
    });
    let saved = Workbook::from_model(model).unwrap().save().unwrap();
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml")).unwrap();
    assert!(sheet.contains("<f>&quot;a&#1;b&quot;</f>"), "{sheet}");
    let workbook = String::from_utf8(part(&saved, "xl/workbook.xml")).unwrap();
    assert!(
        workbook.contains(">&quot;a&#1;&quot;</definedName>"),
        "{workbook}"
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().sheets[0]
            .cell(cell("A1"))
            .and_then(|cell| cell.formula.clone()),
        Some("\"a\u{1}b\"".to_owned())
    );
    assert_eq!(reopened.model().defined_names[0].formula, "\"a\u{1}\"");
}
