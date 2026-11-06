use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, CellValue, Hyperlink, SheetId, StylePatch, Workbook,
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
/// string whose first run ends in a CR LF, a shared string with an escaped CR,
/// and formulas holding or returning a CR. A1 links with a two-line tooltip.
fn book(eol: &str) -> Vec<u8> {
    let price = format!("Sales[[#This Row],[Price{eol}'[GBP']]]");
    let sheet = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{main}" xmlns:r="{rels}"><sheetData>"#,
            r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Year</t></is></c><c r="B1" t="inlineStr"><is><t>Price&#10;[GBP]</t></is></c><c r="C1" t="inlineStr"><is><t>Double</t></is></c><c r="E1"><f>SUM(Sales[Price{eol}'[GBP']])</f><v>30</v></c></row>"#,
            r#"<row r="2"><c r="A2"><v>2003</v></c><c r="B2"><v>10</v></c><c r="C2"><f>{price}*2</f><v>20</v></c></row>"#,
            r#"<row r="3"><c r="A3"><v>2004</v></c><c r="B3"><v>20</v></c><c r="C3"><f>{price}*2</f><v>40</v></c></row>"#,
            "<row r=\"5\"><c r=\"A5\" t=\"inlineStr\"><is><t>a\r\nb</t></is></c><c r=\"B5\" t=\"inlineStr\"><is><t>c\rd</t></is></c><c r=\"C5\" t=\"inlineStr\"><is><t>e&#13;&#10;f</t></is></c><c r=\"D5\" t=\"s\"><v>0</v></c>",
            r#"<c r="E5" t="s"><v>1</v></c><c r="H5"><f>CODE(MID("p&#13;q",2,1))</f><v>13</v></c><c r="I5" t="str"><f>"m"&amp;CHAR(13)&amp;"n"</f><v>m&#13;n</v></c></row>"#,
            r#"<row r="6"><c r="A6"><f>LEN(A5)</f><v>3</v></c><c r="B6"><f>LEN(B5)</f><v>3</v></c><c r="C6"><f>LEN(C5)</f><v>4</v></c><c r="D6"><f>LEN(D5)</f><v>3</v></c><c r="E6"><f>LEN(E5)</f><v>3</v></c></row>"#,
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
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"{MAIN}\" count=\"2\" uniqueCount=\"2\"><si><r><t>g\r\n</t></r><r><rPr><b/></rPr><t>h</t></r></si><si><t>i&#13;j</t></si></sst>"
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

const CR_CELLS: [&str; 7] = ["C5", "E5", "H5", "I5", "C6", "E6", "D5"];

fn cr_values() -> Vec<CellValue> {
    vec![
        text("e\r\nf"),
        text("i\rj"),
        number(13.0),
        text("m\rn"),
        number(4.0),
        number(3.0),
        text("g\nh"),
    ]
}

fn assert_cr_values(saved: &[u8]) {
    let cached = Workbook::open(saved).unwrap();
    assert_eq!(values(&cached, &CR_CELLS), cr_values());
    assert_eq!(values(&recalculated(saved), &CR_CELLS), cr_values());
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
fn a_cr_survives_editing_another_cell_saving_and_reopening() {
    let mut workbook = Workbook::open(&book("\n")).unwrap();
    assert_eq!(values(&workbook, &CR_CELLS), cr_values());
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
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    let tooltips = reopened.model().sheets[0]
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
fn typed_text_with_a_cr_survives_save_restyle_and_reopen() {
    let typed = [("J5", "a\rb"), ("K5", "x\r\ny"), ("L5", "_x000D_")];
    let mut workbook = Workbook::open(&book("\n")).unwrap();
    for (at, input) in typed {
        workbook
            .edit_cell(SheetId(0), cell(at), input, CalculationOptions::default())
            .unwrap();
    }
    let saved = workbook.save().unwrap();
    let written = format!(
        "{}{}",
        String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml")).unwrap(),
        String::from_utf8(part(&saved, "xl/sharedStrings.xml")).unwrap()
    );
    assert!(written.contains("a&#13;b"), "{written}");
    assert!(written.contains("x&#13;\ny"), "{written}");
    let expected = typed.map(|(_, input)| text(input)).to_vec();
    let cells = typed.map(|(at, _)| at);
    let mut reopened = Workbook::open(&saved).unwrap();
    assert_eq!(values(&reopened, &cells), expected);
    bold(&mut reopened, "J5:L5");
    let restyled = Workbook::open(&reopened.save().unwrap()).unwrap();
    assert_eq!(values(&restyled, &cells), expected);
}
