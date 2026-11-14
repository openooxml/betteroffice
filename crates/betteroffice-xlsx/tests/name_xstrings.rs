//! Sheet names (`sheet@name`), defined names (`definedName@name`) and font
//! names (`font/name@val`) are `ST_Xstring`: read decoded and written the way
//! Excel writes them, so every name the API accepts saves as valid XML 1.0 and
//! reads back unchanged. Formula text, which is not `ST_Xstring`, is unchanged.

use betteroffice_xlsx::{CalculationOptions, CellRange, Op, SheetId, StylePatch, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORKBOOK: &str = "xl/workbook.xml";
const STYLES: &str = "xl/styles.xml";
const APP: &str = "docProps/app.xml";

/// Two sheets, `Data` and one named `sheet` as the file spells it, a
/// workbook-level name spelled `defined` referring to `Data!$A$1`, and one
/// font whose name is spelled `font`. `docProps/app.xml` lists both sheets.
fn book(sheet: &str, defined: &str, font: &str) -> Vec<u8> {
    let worksheet = |cells: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}"><sheetData>{cells}</sheetData></worksheet>"#
        )
    };
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{DOC_RELS}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="{DOC_RELS}/extended-properties" Target="docProps/app.xml"/></Relationships>"#
            ),
        ),
        (
            WORKBOOK,
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="{sheet}" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="{defined}">Data!$A$1</definedName></definedNames></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{DOC_RELS}/styles" Target="styles.xml"/></Relationships>"#
            ),
        ),
        (
            STYLES,
            format!(
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{MAIN}"><fonts count="1"><font><sz val="11"/><name val="{font}"/></font></fonts>"#,
                    r#"<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>"#,
                    r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                    r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#
                ),
                MAIN = MAIN,
                font = font
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            worksheet(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#),
        ),
        ("xl/worksheets/sheet2.xml", worksheet("")),
        (
            APP,
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Data</vt:lpstr><vt:lpstr>{sheet}</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#
            ),
        ),
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

/// `xml` holds only characters and character references XML 1.0 allows.
fn assert_xml10(xml: &str) {
    let legal = |code: u32| {
        matches!(code, 0x9 | 0xA | 0xD)
            || (0x20..=0xD7FF).contains(&code)
            || (0xE000..=0xFFFD).contains(&code)
            || (0x10000..=0x10FFFF).contains(&code)
    };
    assert!(
        xml.chars().all(|ch| legal(ch as u32)),
        "raw illegal character in {xml}"
    );
    let mut rest = xml;
    while let Some(at) = rest.find("&#") {
        let reference = &rest[at + 2..];
        let end = reference.find(';').unwrap();
        let digits = &reference[..end];
        let code = match digits.strip_prefix('x') {
            Some(hex) => u32::from_str_radix(hex, 16).unwrap(),
            None => digits.parse().unwrap(),
        };
        assert!(legal(code), "illegal reference &#{digits}; in {xml}");
        rest = &reference[end..];
    }
}

fn sheet_name(workbook: &Workbook, index: usize) -> String {
    workbook.model().sheets[index].name.clone()
}

fn has_font(workbook: &Workbook, name: &str) -> bool {
    workbook
        .model()
        .styles
        .fonts
        .iter()
        .any(|font| font.name.as_deref() == Some(name))
}

/// Excel's `ST_Xstring` escapes in a sheet name, a defined name and a font
/// name read decoded, and a save that leaves them alone keeps their bytes.
#[test]
fn excel_escapes_in_sheet_defined_and_font_names_read_decoded_and_keep_their_bytes() {
    let source = book("A_x0009_B", "_x005F_x0041_", "Cal_x0001_ibri");
    let workbook = Workbook::open(&source).unwrap();
    assert_eq!(sheet_name(&workbook, 1), "A\tB");
    assert_eq!(workbook.model().defined_names[0].name, "_x0041_");
    assert!(
        has_font(&workbook, "Cal\u{1}ibri"),
        "{:?}",
        workbook.model().styles.fonts
    );
    let saved = workbook.save().unwrap();
    assert_eq!(part(&saved, WORKBOOK), part(&source, WORKBOOK));
    assert_eq!(part(&saved, STYLES), part(&source, STYLES));
}

/// Any sheet name the API accepts saves as valid XML 1.0 in the workbook part
/// and in the document properties' sheet titles, and reads back unchanged.
#[test]
fn any_new_sheet_name_saves_as_valid_xml_and_reads_back() {
    for (name, attribute) in [
        ("Q\u{1}4", "Q_x0001_4"),
        ("Tab\tName", "Tab_x0009_Name"),
        ("Line\nBreak", "Line_x000A_Break"),
        ("_x0041_", "_x005F_x0041_"),
        ("a & <b>", "a &amp; &lt;b&gt;"),
    ] {
        let mut workbook = Workbook::open(&book("Second", "Total", "Calibri")).unwrap();
        workbook
            .apply_ops(
                vec![Op::RenameSheet {
                    sheet: SheetId(1),
                    name: name.into(),
                }],
                CalculationOptions::default(),
            )
            .unwrap_or_else(|error| panic!("{name:?}: {error}"));
        let saved = workbook
            .save()
            .unwrap_or_else(|error| panic!("{name:?}: {error}"));
        let xml = part(&saved, WORKBOOK);
        assert!(
            xml.contains(&format!(r#"name="{attribute}""#)),
            "{name:?}: {xml}"
        );
        assert_xml10(&xml);
        assert_xml10(&part(&saved, APP));
        let reopened = Workbook::open(&saved).unwrap();
        assert_eq!(sheet_name(&reopened, 1), name);
        assert_eq!(part(&reopened.save().unwrap(), WORKBOOK), xml, "{name:?}");
    }
}

/// Any font name the API accepts saves as valid XML 1.0 and reads back.
#[test]
fn any_new_font_name_saves_as_valid_xml_and_reads_back() {
    for (name, attribute) in [
        ("Font\u{1}x", "Font_x0001_x"),
        ("Tab\tFont", "Tab_x0009_Font"),
        ("_x0041_", "_x005F_x0041_"),
        ("A&B \"Sans\"", "A&amp;B &quot;Sans&quot;"),
    ] {
        let mut workbook = Workbook::open(&book("Second", "Total", "Calibri")).unwrap();
        workbook
            .patch_range_style(
                SheetId(0),
                CellRange::parse_a1("A1").unwrap(),
                StylePatch {
                    font_family: Some(name.into()),
                    ..StylePatch::default()
                },
                CalculationOptions::default(),
            )
            .unwrap_or_else(|error| panic!("{name:?}: {error}"));
        let saved = workbook
            .save()
            .unwrap_or_else(|error| panic!("{name:?}: {error}"));
        let styles = part(&saved, STYLES);
        assert!(
            styles.contains(&format!(r#"val="{attribute}""#)),
            "{name:?}: {styles}"
        );
        assert_xml10(&styles);
        let reopened = Workbook::open(&saved).unwrap();
        assert!(
            has_font(&reopened, name),
            "{name:?}: {:?}",
            reopened.model().styles.fonts
        );
    }
}

/// An edit that rewrites a defined name, here the formula of a name Excel
/// stored as `_x005F_x0041_`, saves through the patched fragment, keeps the
/// name's escape and reads back.
#[test]
fn an_escaped_defined_name_survives_an_edit_that_rewrites_it() {
    let mut workbook = Workbook::open(&book("Second", "_x005F_x0041_", "Calibri")).unwrap();
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Info".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let xml = part(&saved, WORKBOOK);
    assert!(
        xml.contains(r#"<definedName name="_x005F_x0041_">Info!$A$1</definedName>"#),
        "{xml}"
    );
    assert_xml10(&xml);
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(reopened.model().defined_names[0].name, "_x0041_");
}

/// A pivot cache's `worksheetSource@sheet` is an `ST_Xstring` naming a sheet
/// whose own name Excel wrote escaped. It binds to that sheet, so edits to
/// other sheets go ahead while an edit under the source range is still refused.
#[test]
fn a_pivot_source_names_its_escaped_sheet() {
    let mut model = betteroffice_xlsx::WorkbookModel::default();
    for name in ["A\tB", "Report", "Notes"] {
        model.sheets.push(betteroffice_xlsx::Sheet::new(name));
    }
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    parts.push((
        "xl/pivotCache/pivotCacheDefinition1.xml".to_owned(),
        br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="A1:B4" sheet="A_x0009_B"/></cacheSource></pivotCacheDefinition>"#.to_vec(),
    ));
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    assert!(
        part(&bytes, WORKBOOK).contains(r#"name="A_x0009_B""#),
        "{}",
        part(&bytes, WORKBOOK)
    );
    for op in [
        Op::RenameSheet {
            sheet: SheetId(2),
            name: "Scratch".to_owned(),
        },
        Op::InsertRows {
            sheet: SheetId(2),
            at: 0,
            count: 1,
        },
    ] {
        let mut workbook = Workbook::open(&bytes).unwrap();
        workbook
            .apply_ops(vec![op.clone()], CalculationOptions::default())
            .unwrap_or_else(|error| panic!("{op:?} was refused: {error:?}"));
        workbook.save().unwrap();
    }
    let mut workbook = Workbook::open(&bytes).unwrap();
    assert!(
        workbook
            .apply_ops(
                vec![Op::InsertRows {
                    sheet: SheetId(0),
                    at: 1,
                    count: 1,
                }],
                CalculationOptions::default(),
            )
            .is_err()
    );
}
