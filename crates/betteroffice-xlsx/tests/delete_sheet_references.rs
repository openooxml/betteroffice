//! Deleting a sheet turns every reference to it into `#REF!`, as Excel writes
//! it in the file: the whole qualified reference, address included, in cell
//! formulas, defined names, conditional formats, data validations and
//! extension formulas. A 3-D span that loses an endpoint narrows onto the
//! sheets left between its endpoints, and becomes `#REF!` when none are
//! left. Hyperlink locations keep their text. Undo puts every text back.

use betteroffice_xlsx::{CalculationOptions, Op, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM: &str = "http://schemas.microsoft.com/office/excel/2006/main";

/// Plain sheets named `sheets`, each holding 1 in A1, then `Other`, whose A1
/// holds `formula` with the cached `value` and whose body ends with
/// `other_tail`; `defined` is the workbook's names element.
fn book(sheets: &[&str], formula: &str, value: &str, other_tail: &str, defined: &str) -> Vec<u8> {
    let other_index = sheets.len() + 1;
    let mut overrides = String::new();
    let mut entries = String::new();
    let mut relationships = String::new();
    let mut parts: Vec<(String, String)> = Vec::new();
    for (at, name) in sheets.iter().chain(std::iter::once(&"Other")).enumerate() {
        let n = at + 1;
        overrides.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        ));
        entries.push_str(&format!(
            r#"<sheet name="{name}" sheetId="{n}" r:id="rId{n}"/>"#
        ));
        relationships.push_str(&format!(
            r#"<Relationship Id="rId{n}" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet{n}.xml"/>"#
        ));
        let body = if n == other_index {
            format!(
                r#"<sheetData><row r="1"><c r="A1"><f>{formula}</f><v>{value}</v></c><c r="B1" t="inlineStr"><is><t>go</t></is></c></row></sheetData>{other_tail}"#
            )
        } else {
            r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#.to_owned()
        };
        parts.push((
            format!("xl/worksheets/sheet{n}.xml"),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}">{body}</worksheet>"#
            ),
        ));
    }
    let styles = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
        r#"<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>"#,
        r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
        r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>"#
    );
    let style_rel = sheets.len() + 2;
    parts.extend([
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>{overrides}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#
            ),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{DOC_RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/workbook.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets>{entries}</sheets>{defined}</workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}">{relationships}<Relationship Id="rId{style_rel}" Type="{DOC_RELS}/styles" Target="styles.xml"/></Relationships>"#
            ),
        ),
        ("xl/styles.xml".to_owned(), styles.to_owned()),
    ]);
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name, xml.into_bytes()))
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

fn remove(workbook: &mut Workbook, index: usize) {
    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index }],
            CalculationOptions::default(),
        )
        .unwrap();
}

/// Every formula naming the sheet, wherever the workbook keeps it.
const TAIL: &str = concat!(
    r#"<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>Data!$A$1&gt;0</formula></cfRule></conditionalFormatting>"#,
    r#"<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="C1"><formula1>Data!$A$1:$A$5</formula1></dataValidation></dataValidations>"#,
    r#"<hyperlinks><hyperlink ref="B1" location="Data!A1"/></hyperlinks>"#,
    r#"<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="X14NS"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="XMNS"><x14:cfRule type="expression" priority="2" id="{00000000-0000-0000-0000-000000000001}"><xm:f>Data!$A$1&gt;0</xm:f><x14:dxf/></x14:cfRule><xm:sqref>D1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>"#,
    r#"<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="X14NS"><x14:dataValidations count="1" xmlns:xm="XMNS"><x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Data!$A$1:$A$5</xm:f></x14:formula1><xm:sqref>E1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>"#
);

fn every_site() -> Vec<u8> {
    book(
        &["Data", "Spare"],
        "SUM(Data!$A$1:$B$2)+Spare!A1",
        "2",
        &TAIL.replace("X14NS", X14).replace("XMNS", XM),
        r#"<definedNames><definedName name="Total">Data!$A$1</definedName><definedName name="Kept">Spare!$A$1</definedName></definedNames>"#,
    )
}

/// Deleting a sheet writes `#REF!` for the whole reference in the cell
/// formula, the defined name, the classic and the extension conditional
/// format and the classic and the extension data validation; what names
/// another sheet, and the hyperlink location, keep their text.
#[test]
fn deleting_a_sheet_writes_ref_wherever_it_was_named() {
    let mut workbook = Workbook::open(&every_site()).unwrap();
    remove(&mut workbook, 0);
    let other = &workbook.model().sheets[1];
    assert_eq!(
        other
            .cell(xlsx_model::CellRef::parse_a1("A1").unwrap())
            .and_then(|cell| cell.formula.clone())
            .as_deref(),
        Some("SUM(#REF!)+Spare!A1")
    );
    let saved = workbook.save().unwrap();
    let sheet = part(&saved, "xl/worksheets/sheet3.xml");
    for expected in [
        "<f>SUM(#REF!)+Spare!A1</f>",
        "<formula>#REF!&gt;0</formula>",
        "<formula1>#REF!</formula1>",
        "<xm:f>#REF!&gt;0</xm:f>",
        "<xm:f>#REF!</xm:f>",
        r#"location="Data!A1""#,
    ] {
        assert!(sheet.contains(expected), "{expected} in {sheet}");
    }
    assert!(!sheet.contains("Data!$A$1"), "{sheet}");
    let workbook_part = part(&saved, "xl/workbook.xml");
    assert!(
        workbook_part.contains(r#"<definedName name="Total">#REF!</definedName>"#),
        "{workbook_part}"
    );
    assert!(
        workbook_part.contains(r#"<definedName name="Kept">Spare!$A$1</definedName>"#),
        "{workbook_part}"
    );
}

/// The text of the first `<tag>…</tag>` element in `xml`.
fn element<'a>(xml: &'a str, tag: &str) -> &'a str {
    let start = xml.find(&format!("<{tag}>")).unwrap();
    let end = xml.find(&format!("</{tag}>")).unwrap() + tag.len() + 3;
    &xml[start..end]
}

/// Undo puts every text the deletion rewrote back as the file wrote it: the
/// sheet part whole, and the workbook part's sheets and defined names. (The
/// workbook part also gains the active-sheet view a deletion records, which
/// undo leaves on main as well; that is not reference text.)
#[test]
fn undo_restores_every_text_byte_for_byte() {
    let source = every_site();
    let mut workbook = Workbook::open(&source).unwrap();
    let opened = workbook.model().clone();
    remove(&mut workbook, 0);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().sheets, opened.sheets);
    assert_eq!(workbook.model().defined_names, opened.defined_names);
    let saved = workbook.save().unwrap();
    let sheet = "xl/worksheets/sheet3.xml";
    assert_eq!(part(&saved, sheet), part(&source, sheet));
    let (now, then) = (
        part(&saved, "xl/workbook.xml"),
        part(&source, "xl/workbook.xml"),
    );
    for tag in ["sheets", "definedNames"] {
        assert_eq!(element(&now, tag), element(&then, tag), "{tag}");
    }
}

fn span_formula(workbook: &Workbook) -> String {
    let other = workbook.model().sheets.last().unwrap();
    other
        .cell(xlsx_model::CellRef::parse_a1("A1").unwrap())
        .and_then(|cell| cell.formula.clone())
        .unwrap()
}

/// A 3-D span narrows onto the sheets left between its endpoints when it
/// loses one, as Microsoft documents for deleting an endpoint sheet, keeps
/// its text when a sheet inside it goes, and becomes `#REF!` when every
/// sheet it covered is gone.
#[test]
fn a_span_narrows_when_it_loses_an_endpoint_and_becomes_ref_when_nothing_is_left() {
    let source = book(&["Jan", "Feb", "Mar"], "SUM(Jan:Mar!A1)", "3", "", "");
    for (removals, expected) in [
        (vec![0], "SUM(Feb:Mar!A1)"),
        (vec![2], "SUM(Jan:Feb!A1)"),
        (vec![1], "SUM(Jan:Mar!A1)"),
        (vec![0, 0], "SUM(Mar!A1)"),
        (vec![0, 0, 0], "SUM(#REF!)"),
    ] {
        let mut workbook = Workbook::open(&source).unwrap();
        for index in &removals {
            remove(&mut workbook, *index);
        }
        assert_eq!(span_formula(&workbook), expected, "{removals:?}");
        let saved = workbook.save().unwrap();
        let other = format!("xl/worksheets/sheet{}.xml", 4);
        assert!(
            part(&saved, &other).contains(&format!("<f>{expected}</f>")),
            "{removals:?}"
        );
    }
}

/// Undo brings back the deleted sheet's own conditional formats along with
/// the rest of the sheet.
#[test]
fn undo_restores_the_deleted_sheets_own_conditional_formats() {
    let mut workbook = Workbook::open(&every_site()).unwrap();
    let opened = workbook.model().clone();
    assert!(!opened.sheets[2].conditional_formats.is_empty());
    remove(&mut workbook, 2);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(
        workbook.model().sheets[2].conditional_formats,
        opened.sheets[2].conditional_formats
    );
}

/// A data validation over a 3-D span narrows when an endpoint sheet is
/// deleted; once sheets have moved or been added since the file was read,
/// the order the span narrows against is no longer known, so that deletion
/// is refused, and so is a later move or added sheet after such a deletion.
#[test]
fn a_preserved_span_narrows_on_delete_but_refuses_after_sheets_moved() {
    let tail = r#"<dataValidations count="1"><dataValidation type="whole" allowBlank="1" sqref="C1"><formula1>SUM(Jan:Mar!A1)</formula1></dataValidation></dataValidations>"#;
    let source = book(&["Jan", "Feb", "Mar"], "1", "1", tail, "");
    let mut workbook = Workbook::open(&source).unwrap();
    remove(&mut workbook, 0);
    assert!(
        part(&workbook.save().unwrap(), "xl/worksheets/sheet4.xml")
            .contains("<formula1>SUM(Feb:Mar!A1)</formula1>")
    );

    for (reorder, jan) in [
        (Op::MoveSheet { from: 2, to: 0 }, 1),
        (
            Op::AddSheet {
                index: 1,
                name: "New".into(),
            },
            0,
        ),
    ] {
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .apply_ops(vec![reorder.clone()], CalculationOptions::default())
            .unwrap();
        let before = workbook.model().clone();
        let error = workbook
            .apply_ops(
                vec![Op::RemoveSheet { index: jan }],
                CalculationOptions::default(),
            )
            .unwrap_err();
        assert_eq!(
            error.to_string(),
            "the data validation formulas on sheet 'Other' cannot follow the sheet removal",
            "{reorder:?}"
        );
        assert_eq!(workbook.model(), &before);
    }

    for reorder in [
        Op::MoveSheet { from: 1, to: 0 },
        Op::AddSheet {
            index: 1,
            name: "New".into(),
        },
    ] {
        let mut workbook = Workbook::open(&source).unwrap();
        remove(&mut workbook, 0);
        let before = workbook.model().clone();
        let error = workbook
            .apply_ops(vec![reorder.clone()], CalculationOptions::default())
            .unwrap_err();
        assert_eq!(
            error.to_string(),
            "the data validation formulas on sheet 'Other' cannot follow the sheet removal",
            "{reorder:?}"
        );
        assert_eq!(workbook.model(), &before);
        let (saved, warnings) = workbook.save_with_warnings().unwrap();
        assert_eq!(warnings, [], "{reorder:?}");
        assert!(
            part(&saved, "xl/worksheets/sheet4.xml")
                .contains("<formula1>SUM(Feb:Mar!A1)</formula1>"),
            "{reorder:?}"
        );
    }
}
