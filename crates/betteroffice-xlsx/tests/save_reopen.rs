//! A workbook that opens saves a package that opens again with the same
//! sheets and values, whatever malformed package it was opened from.

use betteroffice_xlsx::{CalculationOptions, CellRef, SheetId, Workbook};
use xlsx_model::CellValue;

const WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const CALC_CHAIN: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
const CUSTOM_XML: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const STYLES: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const STYLESHEET: &str = r#"<styleSheet><cellXfs count="2"><xf/><xf/></cellXfs></styleSheet>"#;
const STYLED_SHEET: &str = concat!(
    r#"<worksheet><sheetData><row r="1">"#,
    r#"<c r="A1" s="1" t="s"><v>0</v></c></row></sheetData></worksheet>"#
);
const SHEET: &str =
    r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;

fn package(workbook: &str, relationships: &str, parts: &[(&str, &str)]) -> Vec<u8> {
    let mut all = vec![
        ("xl/workbook.xml".to_owned(), workbook.as_bytes().to_vec()),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            format!("<Relationships>{relationships}</Relationships>").into_bytes(),
        ),
    ];
    all.extend(
        parts
            .iter()
            .map(|(path, xml)| ((*path).to_owned(), xml.as_bytes().to_vec())),
    );
    ooxml_opc::rezip_parts(&all).unwrap()
}

fn one_sheet(name: &str) -> String {
    format!(
        r#"<workbook><sheets><sheet name="{name}" sheetId="1" r:id="rId1"/></sheets></workbook>"#
    )
}

type Contents = Vec<(String, Vec<(CellRef, CellValue)>)>;

fn contents(workbook: &Workbook) -> Contents {
    workbook
        .model()
        .sheets
        .iter()
        .map(|sheet| {
            let cells = sheet
                .iter_cells()
                .map(|(at, cell)| (at, cell.value.clone()))
                .collect();
            (sheet.name.clone(), cells)
        })
        .collect()
}

/// Recalculates `workbook` (which counts as a change, so the save rewrites
/// the workbook parts) and requires its save to reopen with the same contents.
fn assert_save_reopens(mut workbook: Workbook) -> (Workbook, Vec<u8>) {
    workbook.recalculate_all(CalculationOptions::default());
    let saved = workbook.save().expect("save");
    let reopened = Workbook::open(&saved)
        .unwrap_or_else(|error| panic!("the saved workbook does not reopen: {error:?}"));
    assert_eq!(contents(&reopened), contents(&workbook));
    (reopened, saved)
}

fn saved_part(saved: &[u8], path: &str) -> String {
    let parts = ooxml_opc::unzip_parts(saved).unwrap();
    let (_, bytes) = parts.iter().find(|(name, _)| name == path).unwrap();
    String::from_utf8(bytes.clone()).unwrap()
}

fn unique_workbook_relationships(saved: &[u8]) -> String {
    let relationships = saved_part(saved, "xl/_rels/workbook.xml.rels");
    let ids = relationships
        .split("Id=\"")
        .skip(1)
        .map(|rest| rest.split('"').next().unwrap())
        .collect::<Vec<_>>();
    let unique = ids.iter().collect::<std::collections::HashSet<_>>();
    assert_eq!(unique.len(), ids.len(), "{relationships}");
    relationships
}

#[test]
fn a_sheet_keeps_its_part_when_its_relationship_id_has_a_target_less_duplicate() {
    let source = package(
        &one_sheet("Data"),
        &format!(
            r#"<Relationship Id="rId1" Type="{WORKSHEET}"/><Relationship Id="rId1" Type="{WORKSHEET}" Target="worksheets/data.xml"/>"#
        ),
        &[("xl/worksheets/data.xml", SHEET)],
    );
    let (reopened, _) = assert_save_reopens(Workbook::open(&source).unwrap());
    assert_eq!(
        contents(&reopened)[0].1,
        vec![(
            CellRef::parse_a1("A1").unwrap(),
            CellValue::Number { value: 1.0 }
        )]
    );
}

#[test]
fn a_sheet_on_the_calc_chain_part_survives_a_save_that_drops_the_calc_chain() {
    for kind in [CALC_CHAIN, "urn:unknown"] {
        let source = package(
            &one_sheet("Chain"),
            &format!(r#"<Relationship Id="rId1" Type="{kind}" Target="calcChain.xml"/>"#),
            &[(
                "xl/calcChain.xml",
                r#"<calcChain><c r="A1" i="1"/></calcChain>"#,
            )],
        );
        assert_save_reopens(Workbook::open(&source).unwrap());
    }
}

#[test]
fn duplicate_workbook_relationship_ids_are_saved_unique() {
    let source = package(
        &one_sheet("Data"),
        &format!(
            r#"<Relationship Id="rId1" Type="{WORKSHEET}" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{CUSTOM_XML}" Target="../customXml/item1.xml"/><Relationship Id="rId2" Type="{CUSTOM_XML}" Target="../customXml/item2.xml"/>"#
        ),
        &[
            ("xl/worksheets/sheet1.xml", SHEET),
            ("customXml/item1.xml", "<one/>"),
            ("customXml/item2.xml", "<two/>"),
        ],
    );
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::parse_a1("B1").unwrap(),
            "2",
            CalculationOptions::default(),
        )
        .unwrap();
    let (_, saved) = assert_save_reopens(workbook);
    let relationships = unique_workbook_relationships(&saved);
    assert!(
        relationships.contains(&format!(
            r#"Id="rId2" Type="{CUSTOM_XML}" Target="../customXml/item2.xml""#
        )),
        "the id keeps the relationship it resolves to: {relationships}"
    );
    assert!(
        relationships.contains("customXml/item1.xml"),
        "{relationships}"
    );
}

#[test]
fn a_workbook_part_sharing_a_relationship_id_keeps_its_part() {
    let sheet =
        format!(r#"<Relationship Id="rId1" Type="{WORKSHEET}" Target="worksheets/sheet1.xml"/>"#);
    let strings =
        format!(r#"<Relationship Id="rId2" Type="{SHARED_STRINGS}" Target="sharedStrings.xml"/>"#);
    for relationships in [
        format!(r#"{sheet}{strings}<Relationship Id="rId2" Type="{STYLES}" Target="styles.xml"/>"#),
        format!(r#"{sheet}<Relationship Id="rId1" Type="{STYLES}"/>{strings}"#),
        format!(r#"{sheet}<Relationship Id="rId1" Type="{STYLES}" Target="styles.xml"/>{strings}"#),
    ] {
        let source = package(
            &one_sheet("Data"),
            &relationships,
            &[
                ("xl/worksheets/sheet1.xml", STYLED_SHEET),
                ("xl/styles.xml", STYLESHEET),
                ("xl/sharedStrings.xml", r#"<sst><si><t>x</t></si></sst>"#),
            ],
        );
        let (_, saved) = assert_save_reopens(Workbook::open(&source).unwrap());
        unique_workbook_relationships(&saved);
    }
}

#[test]
fn only_the_workbook_sheet_list_names_sheets() {
    let source = package(
        r#"<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><extLst><ext><sheets><sheet name="Other"/></sheets></ext></extLst></workbook>"#,
        &format!(r#"<Relationship Id="rId1" Type="{WORKSHEET}" Target="worksheets/sheet1.xml"/>"#),
        &[("xl/worksheets/sheet1.xml", SHEET)],
    );
    let (reopened, _) = assert_save_reopens(Workbook::open(&source).unwrap());
    assert_eq!(reopened.model().sheets.len(), 1);
}

/// The save rewrites the `sheets` child of the workbook root; a package whose
/// sheets are listed anywhere else has no sheets to open.
#[test]
fn a_sheet_list_outside_the_workbook_root_does_not_open() {
    for workbook in [
        r#"<sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets>"#,
        r#"<sheet name="A"/>"#,
    ] {
        let source = package(
            workbook,
            &format!(
                r#"<Relationship Id="rId1" Type="{WORKSHEET}" Target="worksheets/sheet1.xml"/>"#
            ),
            &[("xl/worksheets/sheet1.xml", SHEET)],
        );
        assert!(Workbook::open(&source).is_err(), "{workbook}");
    }
}
