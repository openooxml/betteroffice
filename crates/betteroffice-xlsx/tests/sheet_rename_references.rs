//! A sheet rename writes the new name into every formula, defined name and
//! hyperlink location that references the sheet. That text has no escape for
//! a character xml 1.0 cannot carry, so a rename to such a name is refused up
//! front when something references the sheet; an unreferenced sheet takes any
//! name, and undo puts an opened name back without refusal.

use betteroffice_xlsx::{CalculationOptions, Op, SheetId, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// A sheet named `first` as the file spells it, holding 1 in A1, and a sheet
/// `Other` whose cells, hyperlinks and the workbook's names are as given.
fn book(first: &str, other_cells: &str, other_tail: &str, names: &str) -> Vec<u8> {
    let worksheet = |cells: &str, tail: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}"><sheetData>{cells}</sheetData>{tail}</worksheet>"#
        )
    };
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{DOC_RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="{first}" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>{names}</workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            worksheet(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#, ""),
        ),
        ("xl/worksheets/sheet2.xml", worksheet(other_cells, other_tail)),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn rename(workbook: &mut Workbook, name: &str) -> betteroffice_xlsx::Result<()> {
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: name.into(),
            }],
            CalculationOptions::default(),
        )
        .map(|_| ())
}

fn formula(workbook: &Workbook, a1: &str) -> Option<String> {
    workbook.model().sheets[1]
        .cell(xlsx_model::CellRef::parse_a1(a1).unwrap())
        .and_then(|cell| cell.formula.clone())
}

/// Each place that references the sheet: a cell formula, a defined name and a
/// hyperlink location. A rename to a name xml cannot carry is refused with its
/// own reason and changes nothing; the workbook still saves.
#[test]
fn a_rename_to_a_name_xml_cannot_carry_is_refused_when_the_sheet_is_referenced() {
    for (site, source) in [
        (
            "cell formula",
            book(
                "Data",
                r#"<row r="1"><c r="A1"><f>Data!A1</f><v>1</v></c></row>"#,
                "",
                "",
            ),
        ),
        (
            "defined name",
            book(
                "Data",
                "",
                "",
                "<definedNames><definedName name=\"Total\">Data!$A$1</definedName></definedNames>",
            ),
        ),
        (
            "hyperlink location",
            book(
                "Data",
                r#"<row r="1"><c r="B1" t="inlineStr"><is><t>go</t></is></c></row>"#,
                r#"<hyperlinks><hyperlink ref="B1" location="Data!A1"/></hyperlinks>"#,
                "",
            ),
        ),
    ] {
        for name in ["Data\u{1}", "Data\u{fffe}"] {
            let mut workbook = Workbook::open(&source).unwrap();
            let before = workbook.model().clone();
            let error = rename(&mut workbook, name).unwrap_err();
            assert!(
                error
                    .to_string()
                    .contains("sheet name contains a character xml cannot carry"),
                "{site} {name:?}: {error}"
            );
            assert_eq!(workbook.model(), &before, "{site} {name:?}");
            Workbook::open(&workbook.save().unwrap()).unwrap();
        }
    }
}

/// A sheet nothing references takes such a name, written as `ST_Xstring`, and
/// a writable rename of a referenced sheet still rewrites its references.
#[test]
fn an_unreferenced_sheet_takes_any_name_and_a_writable_rename_still_rewrites() {
    let mut workbook = Workbook::open(&book("Data", "", "", "")).unwrap();
    rename(&mut workbook, "Data\u{1}").unwrap();
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(reopened.model().sheets[0].name, "Data\u{1}");

    let mut workbook = Workbook::open(&book(
        "Data",
        r#"<row r="1"><c r="A1"><f>Data!A1</f><v>1</v></c></row>"#,
        "",
        "",
    ))
    .unwrap();
    rename(&mut workbook, "New Data").unwrap();
    assert_eq!(formula(&workbook, "A1").as_deref(), Some("'New Data'!A1"));
}

/// A sheet the file already named with such a character, referenced from a
/// formula, can be renamed to a name xml can carry, and undo puts the opened
/// name and formula back without refusal.
#[test]
fn undo_puts_an_opened_name_back_without_refusal() {
    let source = book(
        "A&#1;B",
        r#"<row r="1"><c r="A1"><f>'A&#1;B'!A1</f><v>1</v></c></row>"#,
        "",
        "",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    let opened = workbook.model().clone();
    assert_eq!(formula(&workbook, "A1").as_deref(), Some("'A\u{1}B'!A1"));
    rename(&mut workbook, "Clean").unwrap();
    assert_eq!(formula(&workbook, "A1").as_deref(), Some("Clean!A1"));
    workbook
        .undo(CalculationOptions::default())
        .unwrap_or_else(|error| panic!("undo refused: {error}"));
    assert_eq!(workbook.model().sheets, opened.sheets);
}
