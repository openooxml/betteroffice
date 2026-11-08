//! A sheet rename rewrites the sheet name in conditional-format and
//! data-validation formulas, classic and x14 forms alike, as Excel does and as
//! it already does for cell formulas. Names map by sheet identity, so swaps and
//! chains land where Excel puts them, undo restores the text, references to a
//! deleted sheet are kept as cell formulas keep them, and a rename that would
//! write a name xml cannot carry is refused up front.

use betteroffice_xlsx::{CalculationOptions, Op, SheetId, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM: &str = "http://schemas.microsoft.com/office/excel/2006/main";
const OTHER: &str = "xl/worksheets/sheet3.xml";

/// Sheets `first` and `second` (part 1 and 2), then `Other` (part 3), whose
/// conditional formats and data validations hold the formulas given, in
/// classic and x14 extension form.
fn book(first: &str, second: &str, cf: &str, dv: &str) -> Vec<u8> {
    let other = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
            r#"<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>{cf}</formula></cfRule></conditionalFormatting>"#,
            r#"<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="B1"><formula1>{dv}</formula1></dataValidation></dataValidations>"#,
            r#"<extLst><ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}" xmlns:x14="{X14}"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="{XM}"><x14:cfRule type="expression" priority="2" id="{{00000000-0000-0000-0000-000000000001}}"><xm:f>{cf}</xm:f><x14:dxf/></x14:cfRule><xm:sqref>C1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>"#,
            r#"<ext uri="{{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}}" xmlns:x14="{X14}"><x14:dataValidations count="1" xmlns:xm="{XM}"><x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>{dv}</xm:f></x14:formula1><xm:sqref>D1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>"#
        ),
        MAIN = MAIN,
        DOC_RELS = DOC_RELS,
        X14 = X14,
        XM = XM,
        cf = cf,
        dv = dv,
    );
    package(first, second, other)
}

/// Sheets `first` and `second` (part 1 and 2), then `Other` (part 3) written
/// as given.
fn package(first: &str, second: &str, other: String) -> Vec<u8> {
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

/// Every conditional-format and data-validation formula the saved `Other`
/// sheet holds, classic then x14: CF formula, DV formula1, then the x14 `xm:f`
/// texts (CF rule, DV list).
fn saved_formulas(saved: &[u8]) -> Vec<String> {
    let xml = part(saved, OTHER);
    let mut out = texts(&xml, "formula");
    out.extend(texts(&xml, "formula1"));
    out.extend(texts(&xml, "xm:f"));
    out
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

/// A rename rewrites every CF and DV formula naming the sheet, in both forms,
/// in the file and in what the workbook reports, and it survives a reopen.
#[test]
fn a_rename_rewrites_conditional_format_and_validation_formulas() {
    let source = book("Data", "Spare", "Data!$A$1&gt;0", "Data!$A$1:$A$5");
    let mut workbook = Workbook::open(&source).unwrap();
    rename(&mut workbook, 0, "New Data").unwrap();
    let saved = workbook.save().unwrap();
    assert_eq!(
        saved_formulas(&saved),
        [
            "'New Data'!$A$1>0",
            "'New Data'!$A$1:$A$5",
            "'New Data'!$A$1>0",
            "'New Data'!$A$1:$A$5"
        ]
    );
    let reported = workbook.conditional_formats(SheetId(2)).unwrap();
    assert!(
        reported.iter().all(|rule| rule
            .formulas
            .iter()
            .all(|f| !f.contains("Data!") || f.contains("'New Data'!"))),
        "{reported:?}"
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        saved_formulas(&reopened.save().unwrap()),
        saved_formulas(&saved)
    );
}

/// Two sheets swap names through a temporary one; each formula follows its
/// sheet, not the name.
#[test]
fn a_swap_through_a_temporary_name_follows_each_sheet() {
    let source = book("A", "B", "A!$A$1&gt;B!$A$1", "B!$A$1:$A$5");
    let mut workbook = Workbook::open(&source).unwrap();
    rename(&mut workbook, 0, "Temp").unwrap();
    rename(&mut workbook, 1, "A").unwrap();
    rename(&mut workbook, 0, "B").unwrap();
    assert_eq!(
        saved_formulas(&workbook.save().unwrap()),
        [
            "B!$A$1>A!$A$1",
            "A!$A$1:$A$5",
            "B!$A$1>A!$A$1",
            "A!$A$1:$A$5"
        ]
    );
}

/// A chain of renames lands on the last name, quoted as Excel quotes a sheet
/// named like an R1C1 reference; undoing all of them puts the sheet part back
/// exactly as the file wrote it.
#[test]
fn a_chain_of_renames_then_undo_restores_the_text() {
    let source = book("A", "Spare", "A!$A$1&gt;0", "A!$A$1:$A$5");
    let mut workbook = Workbook::open(&source).unwrap();
    rename(&mut workbook, 0, "B").unwrap();
    rename(&mut workbook, 0, "C").unwrap();
    assert_eq!(
        saved_formulas(&workbook.save().unwrap()),
        ["'C'!$A$1>0", "'C'!$A$1:$A$5", "'C'!$A$1>0", "'C'!$A$1:$A$5"]
    );
    for _ in 0..2 {
        workbook.undo(CalculationOptions::default()).unwrap();
    }
    assert_eq!(part(&workbook.save().unwrap(), OTHER), part(&source, OTHER));
}

/// References to a deleted sheet are kept as cell formulas keep them, even
/// when a new sheet then takes the deleted sheet's name.
#[test]
fn references_to_a_deleted_sheet_are_kept() {
    let source = book("Data", "Spare", "Data!$A$1&gt;0", "Data!$A$1:$A$5");
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
    assert_eq!(
        saved_formulas(&workbook.save().unwrap()),
        [
            "Data!$A$1>0",
            "Data!$A$1:$A$5",
            "Data!$A$1>0",
            "Data!$A$1:$A$5"
        ]
    );
}

/// A sheet's name inside a string literal, or a defined name spelled like the
/// sheet, is not a reference to it and is left alone.
#[test]
fn a_name_in_a_string_literal_or_a_defined_name_is_left_alone() {
    let cf = "AND(Data&gt;0,A1&lt;&gt;&quot;Data!A1&quot;)";
    let dv = "&quot;Data!A1,Other&quot;";
    let source = book("Data", "Spare", cf, dv);
    let mut workbook = Workbook::open(&source).unwrap();
    rename(&mut workbook, 0, "Renamed").unwrap();
    assert_eq!(
        saved_formulas(&workbook.save().unwrap()),
        [
            "AND(Data>0,A1<>\"Data!A1\")",
            "\"Data!A1,Other\"",
            "AND(Data>0,A1<>\"Data!A1\")",
            "\"Data!A1,Other\""
        ]
    );
}

/// A rename that would write a name xml cannot carry into a CF or DV formula
/// referencing the sheet is refused up front, and the file is unchanged.
#[test]
fn a_rename_to_a_name_xml_cannot_carry_is_refused_when_a_cf_or_dv_formula_references_it() {
    for (cf, dv) in [("Data!$A$1&gt;0", "1"), ("1", "Data!$A$1:$A$5")] {
        let source = book("Data", "Spare", cf, dv);
        let mut workbook = Workbook::open(&source).unwrap();
        let before = workbook.model().clone();
        let error = rename(&mut workbook, 0, "Data\u{1}").unwrap_err();
        assert!(
            error
                .to_string()
                .contains("sheet name contains a character xml cannot carry"),
            "{cf} / {dv}: {error}"
        );
        assert_eq!(workbook.model(), &before);
        assert_eq!(part(&workbook.save().unwrap(), OTHER), part(&source, OTHER));
    }
}

/// `Other` with a sparkline over `formula` and a data validation whose
/// formula is `dv`, both kept as the file wrote them.
fn sparkline_book(formula: &str, dv: &str) -> Vec<u8> {
    package(
        "Data",
        "Spare",
        format!(
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
                r#"<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="B1"><formula1>{dv}</formula1></dataValidation></dataValidations>"#,
                r#"<extLst><ext uri="{{05C60535-1F16-4fd2-B633-F4F36F0B64E0}}" xmlns:x14="{X14}"><x14:sparklineGroups xmlns:xm="{XM}"><x14:sparklineGroup displayEmptyCellsAs="gap"><x14:colorSeries rgb="FF376092"/><x14:sparklines><x14:sparkline><xm:f>{formula}</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst></worksheet>"#
            ),
            MAIN = MAIN,
            DOC_RELS = DOC_RELS,
            X14 = X14,
            XM = XM,
            formula = formula,
            dv = dv,
        ),
    )
}

/// A sparkline's data range follows the rename like the other extension
/// formulas, and everything else in the part keeps its bytes; undo writes
/// the part back as the file had it.
#[test]
fn a_sparkline_range_follows_the_rename_and_the_rest_of_the_part_is_kept() {
    let source = sparkline_book("Data!A1:C1", "1");
    let mut workbook = Workbook::open(&source).unwrap();
    rename(&mut workbook, 0, "New Data").unwrap();
    assert_eq!(
        part(&workbook.save().unwrap(), OTHER),
        part(&source, OTHER).replace(
            "<xm:f>Data!A1:C1</xm:f>",
            "<xm:f>&apos;New Data&apos;!A1:C1</xm:f>"
        )
    );
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(part(&workbook.save().unwrap(), OTHER), part(&source, OTHER));
}

/// A preserved formula that does not read as text the rename can rewrite
/// refuses the rename up front, naming the kind of formula and its sheet,
/// instead of saving it unrenamed.
#[test]
fn a_preserved_formula_the_rename_cannot_read_refuses_the_rename() {
    let source = sparkline_book("Data!A1:C1", "Data!A1<x/>");
    let mut workbook = Workbook::open(&source).unwrap();
    let before = workbook.model().clone();
    let error = rename(&mut workbook, 0, "New Data").unwrap_err();
    assert_eq!(
        error.to_string(),
        "the data validation formulas on sheet 'Other' cannot follow the sheet rename"
    );
    assert_eq!(workbook.model(), &before);
    assert_eq!(part(&workbook.save().unwrap(), OTHER), part(&source, OTHER));
}

/// Renames are judged by sheet identity: a batch that moves sheets and
/// renames them back to the names their positions held still renamed both
/// sheets, so a fragment that cannot follow refuses it.
#[test]
fn a_batch_of_moves_and_compensating_renames_is_still_a_rename() {
    let source = sparkline_book("Data!A1:C1", "Data!A1<x/>");
    let mut workbook = Workbook::open(&source).unwrap();
    let before = workbook.model().clone();
    let rename = |sheet: u32, name: &str| Op::RenameSheet {
        sheet: SheetId(sheet),
        name: name.into(),
    };
    let error = workbook
        .apply_ops(
            vec![
                Op::MoveSheet { from: 1, to: 0 },
                rename(0, "Temp"),
                rename(1, "Spare"),
                rename(0, "Data"),
            ],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert_eq!(
        error.to_string(),
        "the data validation formulas on sheet 'Other' cannot follow the sheet rename"
    );
    assert_eq!(workbook.model(), &before);
}

/// A sheet that took a deleted sheet's name holds the references that name
/// it, as cell formulas do, but has no source identity for the save to
/// follow, so renaming it while a preserved formula names it is refused.
#[test]
fn renaming_a_sheet_that_took_a_deleted_sheets_name_is_refused_when_a_fragment_names_it() {
    let source = book("Data", "Spare", "1", "Data!$A$1:$A$5");
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
    let before = workbook.model().clone();
    let error = rename(&mut workbook, 0, "Z").unwrap_err();
    assert_eq!(
        error.to_string(),
        "the data validation formulas on sheet 'Other' name a deleted sheet whose name this sheet took, and cannot follow the sheet rename"
    );
    assert_eq!(workbook.model(), &before);
}
