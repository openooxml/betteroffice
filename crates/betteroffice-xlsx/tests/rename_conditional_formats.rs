//! A sheet rename rewrites the sheet name in conditional-format rule formulas,
//! as Excel does and as it already does for cell formulas. Each formula follows
//! its sheet, so swaps and chains land where Excel puts them, undo restores the
//! text, references to a deleted sheet are kept as cell formulas keep them, and
//! a rename that would write a name xml cannot carry is refused up front.

use betteroffice_xlsx::{CalculationOptions, Op, SheetId, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OTHER: &str = "xl/worksheets/sheet3.xml";

/// Sheets `first` and `second` (part 1 and 2), then `Other` (part 3), whose
/// conditional-format rule holds the formula given.
fn book(first: &str, second: &str, cf: &str) -> Vec<u8> {
    let other = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
            r#"<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>{cf}</formula></cfRule></conditionalFormatting>"#,
            r#"</worksheet>"#
        ),
        MAIN = MAIN,
        DOC_RELS = DOC_RELS,
        cf = cf,
    );
    let plain = |cells: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}"><sheetData>{cells}</sheetData></worksheet>"#
        )
    };
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#.to_owned(),
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
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="{first}" sheetId="1" r:id="rId1"/><sheet name="{second}" sheetId="2" r:id="rId2"/><sheet name="Other" sheetId="3" r:id="rId3"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId4" Type="{DOC_RELS}/styles" Target="styles.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/styles.xml",
            format!(
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
                    r#"<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>"#,
                    r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                    r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>"#
                ),
                MAIN = MAIN
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            plain(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#),
        ),
        (
            "xl/worksheets/sheet2.xml",
            plain(r#"<row r="1"><c r="A1"><v>2</v></c></row>"#),
        ),
        (OTHER, other),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn part(bytes: &[u8], name: &str) -> String {
    String::from_utf8(
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == name)
            .map(|(_, bytes)| bytes)
            .unwrap_or_default(),
    )
    .unwrap()
}

/// The decoded text of every `<tag>` element in `xml`, in document order.
fn texts(xml: &str, tag: &str) -> Vec<String> {
    let open = format!("<{tag}>");
    let close = format!("</{tag}>");
    let mut out = Vec::new();
    let mut rest = xml;
    while let Some(at) = rest.find(&open) {
        let body = &rest[at + open.len()..];
        let end = body.find(&close).unwrap();
        out.push(
            body[..end]
                .replace("&lt;", "<")
                .replace("&gt;", ">")
                .replace("&quot;", "\"")
                .replace("&apos;", "'")
                .replace("&amp;", "&"),
        );
        rest = &body[end..];
    }
    out
}

/// The conditional-format rule formula the saved `Other` sheet holds.
fn saved_formulas(saved: &[u8]) -> Vec<String> {
    texts(&part(saved, OTHER), "formula")
}

fn apply(workbook: &mut Workbook, op: Op) -> betteroffice_xlsx::Result<()> {
    workbook
        .apply_ops(vec![op], CalculationOptions::default())
        .map(|_| ())
}

fn rename(workbook: &mut Workbook, sheet: u32, name: &str) -> betteroffice_xlsx::Result<()> {
    apply(
        workbook,
        Op::RenameSheet {
            sheet: SheetId(sheet),
            name: name.into(),
        },
    )
}

/// A rename rewrites the rule formula naming the sheet, in the file and in
/// what the workbook reports, and it survives a reopen.
#[test]
fn a_rename_rewrites_a_conditional_format_formula() {
    let source = book("Data", "Spare", "Data!$A$1&gt;0");
    let mut workbook = Workbook::open(&source).unwrap();
    rename(&mut workbook, 0, "New Data").unwrap();
    let saved = workbook.save().unwrap();
    assert_eq!(saved_formulas(&saved), ["'New Data'!$A$1>0"]);
    let reported = workbook.conditional_formats(SheetId(2)).unwrap();
    assert_eq!(reported[0].formulas, ["'New Data'!$A$1>0"]);
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        saved_formulas(&reopened.save().unwrap()),
        saved_formulas(&saved)
    );
}

/// Two sheets swap names through a temporary one; the formula follows each
/// sheet, not the name.
#[test]
fn a_swap_through_a_temporary_name_follows_each_sheet() {
    let source = book("A", "B", "A!$A$1&gt;B!$A$1");
    let mut workbook = Workbook::open(&source).unwrap();
    rename(&mut workbook, 0, "Temp").unwrap();
    rename(&mut workbook, 1, "A").unwrap();
    rename(&mut workbook, 0, "B").unwrap();
    assert_eq!(saved_formulas(&workbook.save().unwrap()), ["B!$A$1>A!$A$1"]);
}

/// A chain of renames lands on the last name; undoing both puts the sheet part
/// back exactly as the file wrote it.
#[test]
fn a_chain_of_renames_then_undo_restores_the_text() {
    let source = book("A", "Spare", "A!$A$1&gt;0");
    let mut workbook = Workbook::open(&source).unwrap();
    rename(&mut workbook, 0, "B").unwrap();
    rename(&mut workbook, 0, "C").unwrap();
    assert_eq!(saved_formulas(&workbook.save().unwrap()), ["C!$A$1>0"]);
    for _ in 0..2 {
        workbook.undo(CalculationOptions::default()).unwrap();
    }
    assert_eq!(part(&workbook.save().unwrap(), OTHER), part(&source, OTHER));
}

/// References to a deleted sheet are kept as cell formulas keep them, even
/// when a new sheet then takes the deleted sheet's name.
#[test]
fn references_to_a_deleted_sheet_are_kept() {
    let source = book("Data", "Spare", "Data!$A$1&gt;0");
    let mut workbook = Workbook::open(&source).unwrap();
    apply(&mut workbook, Op::RemoveSheet { index: 0 }).unwrap();
    apply(
        &mut workbook,
        Op::AddSheet {
            index: 0,
            name: "Data".into(),
        },
    )
    .unwrap();
    assert_eq!(saved_formulas(&workbook.save().unwrap()), ["Data!$A$1>0"]);
}

/// A sheet's name inside a string literal, or a defined name spelled like the
/// sheet, is not a reference to it and is left alone.
#[test]
fn a_name_in_a_string_literal_or_a_defined_name_is_left_alone() {
    let source = book(
        "Data",
        "Spare",
        "AND(Data&gt;0,A1&lt;&gt;&quot;Data!A1&quot;)",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    rename(&mut workbook, 0, "Renamed").unwrap();
    assert_eq!(
        saved_formulas(&workbook.save().unwrap()),
        ["AND(Data>0,A1<>\"Data!A1\")"]
    );
}

/// A rename that would write a name xml cannot carry into a rule formula
/// referencing the sheet is refused up front, and nothing changes.
#[test]
fn a_rename_to_a_name_xml_cannot_carry_is_refused_when_a_rule_references_it() {
    let source = book("Data", "Spare", "Data!$A$1&gt;0");
    let mut workbook = Workbook::open(&source).unwrap();
    let before = workbook.model().clone();
    let error = rename(&mut workbook, 0, "Data\u{1}").unwrap_err();
    assert!(
        error
            .to_string()
            .contains("sheet name contains a character xml cannot carry"),
        "{error}"
    );
    assert_eq!(workbook.model(), &before);
    assert_eq!(part(&workbook.save().unwrap(), OTHER), part(&source, OTHER));
}
