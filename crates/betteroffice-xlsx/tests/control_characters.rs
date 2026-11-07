use betteroffice_xlsx::{
    CalculationOptions, CellInput, CellRange, CellRef, CellState, CellValue, Op, SheetId,
    StylePatch, Workbook,
};
use xlsx_model::CellProvider;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let mut xml =
        format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}">"#);
    for (id, kind, target) in entries {
        let mode = if target.starts_with("file:") {
            r#" TargetMode="External""#
        } else {
            ""
        };
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{DOC_RELS}/{kind}" Target="{target}"{mode}/>"#
        ));
    }
    xml + "</Relationships>"
}

/// One workbook that carries `text` (an encoded control character between `a`
/// and `b`) in every kind of part the reader takes text from.
fn book(text: &str) -> Vec<u8> {
    let sheet = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData>"#,
            r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t>{text}</t></is></c>"#,
            r#"<c r="D1" t="str"><f>"{text}"</f><v>x</v></c><c r="E1" t="str"><f>"x"</f><v>{text}</v></c></row>"#,
            r#"<row r="2"><c r="A2"><v>1</v></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>h</t></is></c><c r="B5" t="inlineStr"><is><t>c</t></is></c></row><row r="6"><c r="A6"><v>1</v></c><c r="B6"><v>2</v></c></row></sheetData>"#,
            r#"<conditionalFormatting sqref="A2"><cfRule type="expression" priority="1"><formula>A2="{text}"</formula></cfRule></conditionalFormatting>"#,
            r#"<dataValidations count="1"><dataValidation type="list" sqref="A3"><formula1>"{text}"</formula1></dataValidation></dataValidations>"#,
            r#"<hyperlinks><hyperlink ref="A1" location="Data!A2" tooltip="{text}"/></hyperlinks>"#,
            r#"<headerFooter><oddHeader>&amp;C{text}</oddHeader></headerFooter>"#,
            r#"<tableParts count="1"><tablePart r:id="rIdTable"/></tableParts></worksheet>"#
        ),
        MAIN = MAIN,
        DOC_RELS = DOC_RELS,
        text = text,
    );
    let parts = vec![
        (
            "[Content_Types].xml",
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>"#,
                r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>"#,
                r#"<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>"#
            )
            .to_owned(),
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
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId3"/></externalReferences><definedNames><definedName name="Named">"{text}"</definedName></definedNames></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "sharedStrings", "sharedStrings.xml"),
                ("rId3", "externalLink", "externalLinks/externalLink1.xml"),
            ]),
        ),
        ("xl/worksheets/sheet1.xml", sheet),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationships(&[
                ("rIdTable", "table", "../tables/table1.xml"),
                ("rIdComments", "comments", "../comments1.xml"),
            ]),
        ),
        (
            "xl/sharedStrings.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><sst xmlns="{MAIN}" count="2" uniqueCount="2"><si><t>{text}</t></si><si><r><t>{text}</t></r><r><rPr><b/></rPr><t>!</t></r></si></sst>"#
            ),
        ),
        (
            "xl/tables/table1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><table xmlns="{MAIN}" id="1" name="T" displayName="T" ref="A5:B6"><tableColumns count="2"><tableColumn id="1" name="{text}"/><tableColumn id="2" name="c"/></tableColumns></table>"#
            ),
        ),
        (
            "xl/comments1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><comments xmlns="{MAIN}"><authors><author>x</author></authors><commentList><comment ref="A1" authorId="0"><text><t>{text}</t></text></comment></commentList></comments>"#
            ),
        ),
        (
            "xl/externalLinks/externalLink1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><externalLink xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><externalBook r:id="rId1"><sheetNames><sheetName val="S"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1" t="str"><v>{text}</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#
            ),
        ),
        (
            "xl/externalLinks/_rels/externalLink1.xml.rels",
            relationships(&[("rId1", "externalLinkPath", "file:///other.xlsx")]),
        ),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn part(bytes: &[u8], name: &str) -> Vec<u8> {
    ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == name)
        .map(|(_, bytes)| bytes)
        .unwrap_or_default()
}

fn element(xml: &[u8], open: &str, close: &str) -> Vec<u8> {
    let text = String::from_utf8_lossy(xml);
    match (text.find(open), text.find(close)) {
        (Some(start), Some(end)) if end >= start => {
            text[start..end + close.len()].as_bytes().to_vec()
        }
        _ => Vec::new(),
    }
}

/// Everything the workbook says the control character is, read back from `bytes`.
#[derive(Debug, PartialEq)]
struct Snapshot {
    values: Vec<CellValue>,
    formula: Option<String>,
    names: Vec<String>,
    conditional: Vec<String>,
    tooltips: Vec<Option<String>>,
    columns: Vec<String>,
    validations: Vec<u8>,
    header: Vec<u8>,
    comments: Vec<u8>,
    external: Vec<u8>,
}

fn snapshot(bytes: &[u8]) -> Snapshot {
    let workbook = Workbook::open(bytes).unwrap();
    let model = workbook.model();
    let sheet_xml = part(bytes, "xl/worksheets/sheet1.xml");
    Snapshot {
        values: ["A1", "B1", "C1", "E1"]
            .into_iter()
            .map(|at| model.value(SheetId(0), CellRef::parse_a1(at).unwrap()))
            .collect(),
        formula: model.sheets[0]
            .cell(CellRef::parse_a1("D1").unwrap())
            .and_then(|cell| cell.formula.clone()),
        names: model
            .defined_names
            .iter()
            .map(|name| name.formula.clone())
            .collect(),
        conditional: model.sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|group| group.rules.iter().flat_map(|rule| rule.formulas.clone()))
            .collect(),
        tooltips: model.sheets[0]
            .hyperlinks
            .iter()
            .map(|link| link.tooltip.clone())
            .collect(),
        columns: model
            .tables
            .iter()
            .flat_map(|table| table.columns.clone())
            .collect(),
        validations: element(&sheet_xml, "<dataValidations", "</dataValidations>"),
        header: element(&sheet_xml, "<headerFooter", "</headerFooter>"),
        comments: part(bytes, "xl/comments1.xml"),
        external: part(bytes, "xl/externalLinks/externalLink1.xml"),
    }
}

fn encodings(ch: char) -> Vec<(&'static str, String)> {
    let code = ch as u32;
    vec![
        ("decimal reference", format!("a&#{code};b")),
        ("hex reference", format!("a&#x{code:X};b")),
        ("xstring escape", format!("a_x{code:04X}_b")),
        ("raw", format!("a{ch}b")),
    ]
}

fn controls() -> impl Iterator<Item = char> {
    (1u32..0x20)
        .filter(|code| !matches!(code, 0x9 | 0xa | 0xd))
        .chain([0xfffe, 0xffff])
        .filter_map(char::from_u32)
}

#[test]
fn every_control_character_in_every_part_saves_and_keeps_its_value() {
    let mut failures = Vec::new();
    for ch in controls() {
        for (encoding, text) in encodings(ch) {
            let case = format!("U+{:04X} as {encoding}", ch as u32);
            let source = book(&text);
            let expected = snapshot(&source);
            let plain = match Workbook::open(&source).and_then(|workbook| workbook.save()) {
                Ok(saved) => saved,
                Err(error) => {
                    failures.push(format!("{case}: plain save failed: {error}"));
                    continue;
                }
            };
            if snapshot(&plain) != expected {
                failures.push(format!("{case}: plain save changed a value"));
            }
            let mut workbook = Workbook::open(&plain).unwrap();
            let edited = workbook
                .edit_cell(
                    SheetId(0),
                    CellRef::parse_a1("A2").unwrap(),
                    "2",
                    CalculationOptions::default(),
                )
                .and_then(|_| {
                    workbook.patch_range_style(
                        SheetId(0),
                        CellRange::parse_a1("A1:E1").unwrap(),
                        StylePatch {
                            bold: Some(true),
                            ..StylePatch::default()
                        },
                        CalculationOptions::default(),
                    )
                })
                .and_then(|_| workbook.save());
            match edited {
                Ok(saved) if snapshot(&saved) == expected => {}
                Ok(saved) => failures.push(format!(
                    "{case}: edit and restyle changed a value: {:?} vs {:?}",
                    snapshot(&saved),
                    expected
                )),
                Err(error) => failures.push(format!("{case}: edited save failed: {error}")),
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

const REFUSAL: &str = "which a workbook file cannot store";

fn refused(result: betteroffice_xlsx::Result<betteroffice_xlsx::MutationResult>) -> String {
    let error = result.expect_err("a formula with no storable form was accepted");
    let message = error.to_string();
    assert!(message.contains(REFUSAL), "{message}");
    message
}

fn formula_at(workbook: &Workbook, sheet: usize, a1: &str) -> Option<String> {
    workbook.model().sheets[sheet]
        .cell(CellRef::parse_a1(a1).unwrap())
        .and_then(|cell| cell.formula.clone())
}

fn text_at(workbook: &Workbook, sheet: usize, a1: &str) -> CellValue {
    workbook
        .model()
        .value(SheetId(sheet as u32), CellRef::parse_a1(a1).unwrap())
}

#[test]
fn an_opened_formula_with_a_control_character_stays_editable() {
    let source = book("a&#1;b");
    let mut workbook = Workbook::open(&source).unwrap();
    let d1 = CellRef::parse_a1("D1").unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            d1,
            "=\"a\u{1}b\"",
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::parse_a1("F1").unwrap(),
            "=\"c\u{2}d\"",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        formula_at(&workbook, 0, "F1").as_deref(),
        Some("(\"c\"&CHAR(2)&\"d\")")
    );
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        formula_at(&reopened, 0, "D1"),
        Some("\"a\u{1}b\"".to_owned())
    );
    assert_eq!(
        formula_at(&reopened, 0, "F1").as_deref(),
        Some("(\"c\"&CHAR(2)&\"d\")")
    );
    assert_eq!(
        text_at(&reopened, 0, "F1"),
        CellValue::Text {
            value: "c\u{2}d".into()
        }
    );
    assert_eq!(reopened.model().defined_names[0].formula, "\"a\u{1}b\"");
}

/// D1 holds the workbook's only U+0001, and E1 refers to D1 and carries one too.
fn movable_book() -> Vec<u8> {
    formula_book(r#""a&#1;b""#)
}

/// [`movable_book`] with `d1` as D1's stored formula text.
fn formula_book(d1: &str) -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned(),
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
            relationships(&[("rId1", "worksheet", "worksheets/sheet1.xml")]),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="D1" t="str"><f>{d1}</f><v>x</v></c><c r="E1" t="str"><f>D1&amp;"&#1;"</f><v>x</v></c></row></sheetData></worksheet>"#
            ),
        ),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn an_opened_formula_with_a_control_character_moves_and_shifts() {
    let mut workbook = Workbook::open(&movable_book()).unwrap();
    workbook
        .move_range(
            SheetId(0),
            CellRange::parse_a1("D1").unwrap(),
            CellRef::parse_a1("G1").unwrap(),
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    let formula = |at: &str| {
        reopened.model().sheets[0]
            .cell(CellRef::parse_a1(at).unwrap())
            .and_then(|cell| cell.formula.clone())
    };
    assert_eq!(formula("G2"), Some("\"a\u{1}b\"".to_owned()));
    assert_eq!(formula("E2"), Some("G2&\"\u{1}\"".to_owned()));
}

/// A raw `SetCell` that puts control-character text into a cell that did not
/// hold it is new entry, even when another cell holds the same text: a
/// control character in a text literal is written as `CHAR(n)` and the value
/// stays the same. Text with no storable form is refused and the batch with it.
#[test]
fn apply_ops_writes_new_control_characters_as_char_and_refuses_what_cannot_be_stored() {
    let rewritten = [
        (
            vec![set(
                SheetId(0),
                CellRef::parse_a1("H1").unwrap(),
                Some("\"new\u{1}\"".into()),
            )],
            "H1",
            "(\"new\"&CHAR(1))",
            "new\u{1}",
        ),
        (
            vec![
                set(
                    SheetId(0),
                    CellRef::parse_a1("G1").unwrap(),
                    Some("\"a\u{1}b\"".into()),
                ),
                set(SheetId(0), CellRef::parse_a1("D1").unwrap(), None),
            ],
            "G1",
            "(\"a\"&CHAR(1)&\"b\")",
            "a\u{1}b",
        ),
        (
            vec![
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
                set(
                    SheetId(0),
                    CellRef::parse_a1("D2").unwrap(),
                    Some("\"a\u{1}c\"".into()),
                ),
            ],
            "D2",
            "(\"a\"&CHAR(1)&\"c\")",
            "a\u{1}c",
        ),
    ];
    for (ops, at, formula, value) in rewritten {
        let mut workbook = Workbook::open(&movable_book()).unwrap();
        workbook
            .apply_ops(ops, CalculationOptions::default())
            .unwrap();
        assert_eq!(formula_at(&workbook, 0, at).as_deref(), Some(formula));
        let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
        assert_eq!(formula_at(&reopened, 0, at).as_deref(), Some(formula));
        assert_eq!(
            text_at(&reopened, 0, at),
            CellValue::Text {
                value: value.into()
            }
        );
    }
    let mut workbook = Workbook::open(&movable_book()).unwrap();
    let before = workbook.model().clone();
    for (formula, message) in [
        (
            "\"a\u{0}\"",
            "U+0000, which a workbook file cannot store and CHAR cannot produce",
        ),
        (
            "\"a\u{fffe}\"",
            "U+FFFE, which a workbook file cannot store and CHAR cannot produce",
        ),
        (
            "{\"a\u{1}\",1}",
            "U+0001 inside an array constant, which a workbook file cannot store",
        ),
        (
            "'Data\u{1}'!A1",
            "U+0001 outside a text literal, which a workbook file cannot store",
        ),
    ] {
        let error = workbook
            .apply_ops(
                vec![
                    set(
                        SheetId(0),
                        CellRef::parse_a1("J1").unwrap(),
                        Some("1".into()),
                    ),
                    set(
                        SheetId(0),
                        CellRef::parse_a1("H1").unwrap(),
                        Some(formula.into()),
                    ),
                ],
                CalculationOptions::default(),
            )
            .unwrap_err();
        assert!(error.to_string().contains(message), "{error}");
        assert!(workbook.model() == &before);
    }
    let error = workbook
        .apply_ops(
            vec![set(
                SheetId(0),
                CellRef::parse_a1("H1").unwrap(),
                Some("A1\u{1}".into()),
            )],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(
        error.to_string().contains("U+0001 and cannot be read"),
        "{error}"
    );
    let long = format!("\"{}\"", "\u{1}".repeat(1100));
    let error = workbook
        .apply_ops(
            vec![set(
                SheetId(0),
                CellRef::parse_a1("H1").unwrap(),
                Some(long),
            )],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(
        error.to_string().contains("8192-character limit"),
        "{error}"
    );
    assert!(workbook.model() == &before);
    let saved = workbook.save().unwrap();
    assert_eq!(
        part(&saved, "xl/worksheets/sheet1.xml"),
        part(&movable_book(), "xl/worksheets/sheet1.xml")
    );
}

/// Undo and redo, an agent-style `SetCell` that restyles an opened
/// control-character cell with its formula as it is, a re-commit of a cell
/// shifted by a structural op, and `move_range` carry the opened formula as it
/// is.
#[test]
fn an_opened_control_character_formula_passes_undo_restyle_and_shifted_re_commits() {
    let mut workbook = Workbook::open(&movable_book()).unwrap();
    let d1 = CellRef::parse_a1("D1").unwrap();
    let held = "\"a\u{1}b\"".to_owned();
    workbook
        .patch_range_style(
            SheetId(0),
            CellRange::parse_a1("E1").unwrap(),
            StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let bold = workbook.model().sheets[0]
        .cell(CellRef::parse_a1("E1").unwrap())
        .and_then(|cell| cell.style);
    assert!(bold.is_some());
    workbook
        .apply_ops(
            vec![Op::SetCell {
                sheet: SheetId(0),
                at: d1,
                cell: CellState {
                    formula: Some(held.clone()),
                    style: bold,
                    ..CellState::default()
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(formula_at(&workbook, 0, "D1"), Some(held.clone()));
    assert_eq!(
        workbook.model().sheets[0]
            .cell(d1)
            .and_then(|cell| cell.style),
        bold
    );
    workbook
        .edit_cell(SheetId(0), d1, "", CalculationOptions::default())
        .unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    workbook.redo(CalculationOptions::default()).unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(formula_at(&workbook, 0, "D1"), Some(held.clone()));
    workbook
        .apply_ops(
            vec![
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
                set(
                    SheetId(0),
                    CellRef::parse_a1("D2").unwrap(),
                    Some(held.clone()),
                ),
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(formula_at(&workbook, 0, "D2"), Some(held.clone()));
    workbook
        .move_range(
            SheetId(0),
            CellRange::parse_a1("D2").unwrap(),
            CellRef::parse_a1("G5").unwrap(),
            CalculationOptions::default(),
        )
        .unwrap();
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(formula_at(&reopened, 0, "G5"), Some(held));
}

#[test]
fn a_prefixed_opened_formula_re_commits_and_moves_as_displayed() {
    let stored = "_xlfn.CONCAT(\"a\u{1}b\",\"\")".to_owned();
    let shown = "CONCAT(\"a\u{1}b\",\"\")";
    let set = |at: &str, formula: Option<&str>| Op::SetCell {
        sheet: SheetId(0),
        at: CellRef::parse_a1(at).unwrap(),
        cell: CellState {
            formula: formula.map(str::to_owned),
            ..CellState::default()
        },
    };
    let source = formula_book(r#"_xlfn.CONCAT("a&#1;b","")"#);
    let mut workbook = Workbook::open(&source).unwrap();
    let d1 = CellRef::parse_a1("D1").unwrap();
    workbook
        .apply_ops(vec![set("D1", Some(shown))], CalculationOptions::default())
        .unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            d1,
            &format!("={shown}"),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        workbook.model().sheets[0]
            .cell(d1)
            .and_then(|cell| cell.formula.clone()),
        Some(stored.clone())
    );
    workbook
        .apply_ops(vec![set("H1", Some(shown))], CalculationOptions::default())
        .unwrap();
    let h1 = formula_at(&workbook, 0, "H1").unwrap();
    assert!(
        h1.contains("CONCAT((\"a\"&CHAR(1)&\"b\"),\"\")") && !h1.contains('\u{1}'),
        "{h1}"
    );
    workbook
        .move_range(
            SheetId(0),
            CellRange::parse_a1("D1").unwrap(),
            CellRef::parse_a1("G1").unwrap(),
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml")).unwrap();
    assert!(
        sheet.contains("<f>_xlfn.CONCAT(") && sheet.contains("a&#1;b"),
        "{sheet}"
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().sheets[0]
            .cell(CellRef::parse_a1("G1").unwrap())
            .and_then(|cell| cell.formula.clone()),
        Some(stored)
    );
}

/// Data holds plain and `_xlfn.`-prefixed formulas with U+0001; Other refers to Data.
fn property_book() -> Vec<u8> {
    let sheet = |body: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}"><sheetData>{body}</sheetData></worksheet>"#
        )
    };
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "worksheet", "worksheets/sheet2.xml"),
            ]),
        ),
        (
            "xl/worksheets/sheet1.xml",
            sheet(concat!(
                r#"<row r="1"><c r="A1"><v>1</v></c><c r="D1" t="str"><f>A1&amp;"&#1;"</f><v>x</v></c><c r="E1" t="str"><f>_xlfn.CONCAT(A2,"&#1;")</f><v>x</v></c></row>"#,
                r#"<row r="2"><c r="A2"><v>2</v></c><c r="D2" t="str"><f>"a&#1;b"</f><v>x</v></c></row>"#,
                r#"<row r="3"><c r="A3"><v>3</v></c></row>"#,
            )),
        ),
        (
            "xl/worksheets/sheet2.xml",
            sheet(r#"<row r="1"><c r="F1" t="str"><f>Data!A1&amp;"&#1;"</f><v>x</v></c></row>"#),
        ),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

struct Rng(u64);

impl Rng {
    fn below(&mut self, bound: u64) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0 % bound
    }
}

fn structural_op(rng: &mut Rng, case: usize) -> Op {
    let sheet = SheetId(rng.below(2) as u32);
    match rng.below(6) {
        0 => Op::InsertRows {
            sheet,
            at: rng.below(4) as u32,
            count: 1 + rng.below(2) as u32,
        },
        1 => Op::DeleteRows {
            sheet,
            at: rng.below(6) as u32,
            count: 1,
        },
        2 => Op::InsertCols {
            sheet,
            at: rng.below(4) as u32,
            count: 1,
        },
        3 => Op::DeleteCols {
            sheet,
            at: 1 + rng.below(8) as u32,
            count: 1,
        },
        4 => Op::RenameSheet {
            sheet,
            name: format!("Renamed{case}x{}", rng.below(1000)),
        },
        _ => {
            let from = rng.below(2) as usize;
            Op::MoveSheet { from, to: 1 - from }
        }
    }
}

fn set(sheet: SheetId, at: CellRef, formula: Option<String>) -> Op {
    Op::SetCell {
        sheet,
        at,
        cell: CellState {
            formula,
            ..CellState::default()
        },
    }
}

/// Batches of structural ops with same-cell re-commits of opened
/// control-character formulas, as displayed, then moves through `move_range`
/// and pastes with `copyFrom` provenance, are never refused, save, and keep the
/// formulas as they were. New control-character text, in a fresh cell or a
/// changed one, is written with `CHAR` and never refused, except text with no
/// storable form, which is refused and leaves the workbook as it was.
#[test]
fn opened_control_characters_move_re_commit_and_paste_without_refusal() {
    let source = property_book();
    let mut rng = Rng(0x5eed_c0de_1234_5678);
    let mut failures = Vec::new();
    let mut checked = 0;
    for case in 0..300 {
        let structural = (0..rng.below(4))
            .map(|_| structural_op(&mut rng, case))
            .collect::<Vec<_>>();
        let mut shadow = Workbook::open(&source).unwrap();
        if shadow
            .apply_ops(structural.clone(), CalculationOptions::default())
            .is_err()
        {
            continue;
        }
        let mut held = Vec::new();
        for (index, sheet) in shadow.model().sheets.iter().enumerate() {
            for (at, cell) in sheet.iter_cells() {
                if let Some(formula) = &cell.formula
                    && formula.contains('\u{1}')
                {
                    held.push((SheetId(index as u32), at, formula.clone()));
                }
            }
        }
        if held.is_empty() {
            continue;
        }
        let mut ops = structural.clone();
        for (sheet, at, formula) in &held {
            if rng.below(2) == 0 {
                ops.push(set(*sheet, *at, Some(formula.replace("_xlfn.", ""))));
            }
        }
        checked += 1;
        let mut workbook = Workbook::open(&source).unwrap();
        if let Err(error) = workbook.apply_ops(ops.clone(), CalculationOptions::default()) {
            failures.push(format!("case {case}: refused {ops:?}: {error}"));
            continue;
        }
        let mut probe = Workbook::open(&source).unwrap();
        probe
            .apply_ops(ops.clone(), CalculationOptions::default())
            .unwrap();
        let (sheet, at, formula) = &held[rng.below(held.len() as u64) as usize];
        let fresh = CellRef::new(30, rng.below(4) as u32);
        let written = |probe: &Workbook, at: CellRef| {
            probe.model().sheets[sheet.0 as usize]
                .cell(at)
                .and_then(|cell| cell.formula.clone())
        };
        for (introduced, target) in [
            (
                vec![set(*sheet, fresh, Some(format!("\"n{case}\u{1}\"")))],
                fresh,
            ),
            (
                vec![set(*sheet, *at, Some(format!("{formula}&\"\u{1}\"")))],
                *at,
            ),
            (
                vec![
                    Op::InsertRows {
                        sheet: *sheet,
                        at: 0,
                        count: 1,
                    },
                    set(*sheet, fresh, Some(formula.clone())),
                ],
                fresh,
            ),
        ] {
            if let Err(error) = probe.apply_ops(introduced.clone(), CalculationOptions::default()) {
                failures.push(format!("case {case}: {introduced:?} refused: {error}"));
                continue;
            }
            if !written(&probe, target)
                .is_some_and(|text| text.contains("CHAR(1)") && !text.contains('\u{1}'))
            {
                failures.push(format!(
                    "case {case}: {} holds {:?}",
                    target.to_a1(),
                    written(&probe, target)
                ));
            }
        }
        let before = probe.model().clone();
        match probe.apply_ops(
            vec![set(*sheet, fresh, Some(format!("\"n{case}\u{0}\"")))],
            CalculationOptions::default(),
        ) {
            Err(error) if error.to_string().contains(REFUSAL) => {}
            other => failures.push(format!("case {case}: NUL gave {other:?}")),
        }
        if probe.model() != &before {
            failures.push(format!("case {case}: a refused batch changed the workbook"));
        }
        let mut moves = Vec::new();
        for (round, (sheet, at, _)) in held.iter().enumerate() {
            if rng.below(2) == 0 {
                continue;
            }
            let target = CellRef::new(10 + round as u32, 10 + rng.below(4) as u32);
            if let Err(error) = workbook.move_range(
                *sheet,
                CellRange::new(*at, *at),
                target,
                CalculationOptions::default(),
            ) {
                failures.push(format!(
                    "case {case}: move of {} refused: {error}",
                    at.to_a1()
                ));
                continue;
            }
            moves.push((*sheet, target, *at));
        }
        let mut pasted = Vec::new();
        for (round, (sheet, _, formula)) in held.iter().enumerate() {
            if rng.below(2) == 0 {
                continue;
            }
            let shown = formula.replace("_xlfn.", "");
            let target = CellRef::new(20 + round as u32, 20);
            if let Err(error) = workbook.edit_pasted_cells(
                *sheet,
                &[CellInput {
                    cell: target,
                    input: format!("={shown}"),
                }],
                &[true],
                CalculationOptions::default(),
            ) {
                failures.push(format!("case {case}: paste of {shown:?} refused: {error}"));
            }
            pasted.push((*sheet, target));
        }
        let saved = match workbook.save() {
            Ok(saved) => saved,
            Err(error) => {
                failures.push(format!("case {case}: save failed: {error}"));
                continue;
            }
        };
        let reopened = Workbook::open(&saved).unwrap();
        let formula = |sheet: SheetId, at: CellRef| {
            reopened.model().sheets[sheet.0 as usize]
                .cell(at)
                .and_then(|cell| cell.formula.clone())
        };
        for (sheet, target) in pasted {
            if !formula(sheet, target).is_some_and(|formula| formula.contains('\u{1}')) {
                failures.push(format!("case {case}: paste to {} lost", target.to_a1()));
            }
        }
        for (sheet, target, origin) in moves {
            if !formula(sheet, target).is_some_and(|formula| formula.contains('\u{1}'))
                || formula(sheet, origin).is_some()
            {
                failures.push(format!("case {case}: move to {} lost", target.to_a1()));
            }
        }
    }
    assert!(checked > 100, "only {checked} cases had held formulas");
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

fn batch(workbook: &Workbook, step: serde_json::Value) -> betteroffice_xlsx::EditRequest {
    serde_json::from_value(serde_json::json!({
        "expectVersion": workbook.version(),
        "steps": [step],
    }))
    .unwrap()
}

fn proposal(a1: &str, input: &str) -> betteroffice_xlsx::ProposalRequest {
    betteroffice_xlsx::ProposalRequest {
        agent_id: "agent".into(),
        note: None,
        edits: vec![betteroffice_xlsx::ProposalEditInput {
            sheet: SheetId(0),
            cell: CellRef::parse_a1(a1).unwrap(),
            input: input.into(),
            number_format: None,
        }],
    }
}

/// Every typed entry writes a new formula's control characters in text
/// literals with `CHAR`, refuses text with no storable form and leaves the
/// workbook as it was, and keeps the cell's own stored formula re-committed as
/// displayed.
#[test]
fn typed_entry_writes_new_control_characters_as_char_and_keeps_a_re_commit() {
    let mut workbook = Workbook::open(&movable_book()).unwrap();
    let options = CalculationOptions::default();
    let target =
        |a1: &str| serde_json::json!({ "sheetId": "sheet:0", "range": { "kind": "a1", "a1": a1 } });
    let typed = "=\"n\u{1}\"";
    let written = "(\"n\"&CHAR(1))";
    workbook
        .edit_cell(SheetId(0), CellRef::parse_a1("H1").unwrap(), typed, options)
        .unwrap();
    workbook
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: CellRef::parse_a1("H2").unwrap(),
                input: typed.into(),
            }],
            options,
        )
        .unwrap();
    for step in [
        serde_json::json!({ "op": "setCellInputs", "target": target("H3"), "inputs": [[typed]] }),
        serde_json::json!({ "op": "setFormulas", "target": target("H4"), "formulas": [["\"n\u{1}\""]] }),
    ] {
        workbook
            .apply_edits(&batch(&workbook, step))
            .unwrap()
            .expect("a typed control-character formula was refused");
    }
    let proposed = workbook.propose(proposal("H5", typed), options).unwrap();
    workbook
        .accept_proposal(&proposed.id, false, options)
        .unwrap();
    for at in ["H1", "H2", "H3", "H4", "H5"] {
        assert_eq!(
            formula_at(&workbook, 0, at).as_deref(),
            Some(written),
            "{at}"
        );
        assert_eq!(
            text_at(&workbook, 0, at),
            CellValue::Text {
                value: "n\u{1}".into()
            },
            "{at}"
        );
    }

    let before = workbook.model().clone();
    let nul = "=\"n\u{0}\"";
    refused(workbook.edit_cell(SheetId(0), CellRef::parse_a1("H6").unwrap(), nul, options));
    refused(workbook.edit_cells(
        SheetId(0),
        &[CellInput {
            cell: CellRef::parse_a1("H6").unwrap(),
            input: "={\"n\u{1}\",1}".into(),
        }],
        options,
    ));
    for step in [
        serde_json::json!({ "op": "setCellInputs", "target": target("H6"), "inputs": [[nul]] }),
        serde_json::json!({ "op": "setFormulas", "target": target("H6"), "formulas": [["\"n\u{ffff}\""]] }),
    ] {
        let refusal = workbook
            .apply_edits(&batch(&workbook, step))
            .unwrap()
            .expect_err("a formula with no storable form was applied");
        assert!(
            refusal.failure.message.contains(REFUSAL),
            "{:?}",
            refusal.failure
        );
    }
    let error = workbook.propose(proposal("H6", nul), options).unwrap_err();
    assert!(error.to_string().contains(REFUSAL), "{error}");
    assert!(workbook.model() == &before);

    let shown = "=\"a\u{1}b\"";
    let d1 = CellRef::parse_a1("D1").unwrap();
    workbook.edit_cell(SheetId(0), d1, shown, options).unwrap();
    workbook
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: d1,
                input: shown.into(),
            }],
            options,
        )
        .unwrap();
    for step in [
        serde_json::json!({ "op": "setCellInputs", "target": target("D1"), "inputs": [[shown]] }),
        serde_json::json!({ "op": "setFormulas", "target": target("D1"), "formulas": [["\"a\u{1}b\""]] }),
    ] {
        workbook
            .apply_edits(&batch(&workbook, step))
            .unwrap()
            .expect("a re-commit of the stored formula was refused");
    }
    let proposed = workbook.propose(proposal("D1", shown), options).unwrap();
    workbook
        .accept_proposal(&proposed.id, false, options)
        .unwrap();
    assert_eq!(
        formula_at(&workbook, 0, "D1"),
        Some("\"a\u{1}b\"".to_owned())
    );
}

/// A paste carries its source formula with `copyFrom` provenance, so an opened
/// control-character formula pastes as it is; the same text typed is new entry
/// and is written with `CHAR`.
#[test]
fn a_pasted_copy_of_an_opened_control_character_formula_is_carried_as_it_is() {
    let mut workbook = Workbook::open(&movable_book()).unwrap();
    let input = |a1: &str| CellInput {
        cell: CellRef::parse_a1(a1).unwrap(),
        input: "=\"a\u{1}b\"".into(),
    };
    workbook
        .edit_cells(SheetId(0), &[input("H1")], CalculationOptions::default())
        .unwrap();
    workbook
        .edit_pasted_cells(
            SheetId(0),
            &[input("G1")],
            &[true],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        formula_at(&workbook, 0, "G1"),
        Some("\"a\u{1}b\"".to_owned())
    );
    assert_eq!(
        formula_at(&workbook, 0, "H1").as_deref(),
        Some("(\"a\"&CHAR(1)&\"b\")")
    );
    let saved = workbook.save().unwrap();
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml")).unwrap();
    assert!(sheet.contains(r#"<c r="G1""#), "{sheet}");
    assert_eq!(sheet.matches("a&#1;b").count(), 2, "{sheet}");
    assert!(
        sheet.contains("<f>(&quot;a&quot;&amp;CHAR(1)&amp;&quot;b&quot;)</f>"),
        "{sheet}"
    );
}

#[test]
fn a_hex_reference_is_kept_untouched_and_written_as_decimal_in_a_new_cell() {
    let source = formula_book(r#""a&#x1;b""#);
    let plain = Workbook::open(&source).unwrap().save().unwrap();
    assert_eq!(plain, source);
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .move_range(
            SheetId(0),
            CellRange::parse_a1("D1").unwrap(),
            CellRef::parse_a1("G1").unwrap(),
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml")).unwrap();
    assert!(
        sheet.contains("a&#1;b") && !sheet.contains("a&#x1;b"),
        "{sheet}"
    );
    assert_eq!(
        Workbook::open(&saved).unwrap().model().sheets[0]
            .cell(CellRef::parse_a1("G1").unwrap())
            .and_then(|cell| cell.formula.clone()),
        Some("\"a\u{1}b\"".to_owned())
    );
}

/// Data holds an opened CF rule and Other refers to Data; the workbook holds an
/// opened defined name. `mark` is spliced into all three: `&#1;` for U+0001, or
/// plain text for a control-free twin.
fn names_book(mark: &str) -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="Held">Data!$B$2&amp;"{mark}"</definedName></definedNames></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "worksheet", "worksheets/sheet2.xml"),
            ]),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}"><sheetData><row r="2"><c r="B2"><v>1</v></c><c r="C2"><v>2</v></c></row><row r="3"><c r="B3"><v>3</v></c></row></sheetData><conditionalFormatting sqref="B2:C3"><cfRule type="expression" priority="1"><formula>B2&amp;"{mark}"="1"</formula></cfRule></conditionalFormatting></worksheet>"#
            ),
        ),
        (
            "xl/worksheets/sheet2.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1" t="str"><f>Held&amp;Data!C2</f><v>x</v></c></row></sheetData></worksheet>"#
            ),
        ),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn names_and_rules(workbook: &Workbook) -> (Vec<String>, Vec<String>) {
    let model = workbook.model();
    (
        model
            .defined_names
            .iter()
            .map(|name| name.formula.clone())
            .collect(),
        model
            .sheets
            .iter()
            .flat_map(|sheet| &sheet.conditional_formats)
            .flat_map(|group| group.rules.iter().flat_map(|rule| rule.formulas.clone()))
            .collect(),
    )
}

/// Structural batches that rewrite an opened defined name and an opened CF rule
/// holding U+0001 are refused no more often than on a control-free twin, and
/// what they leave saves and reopens equal.
#[test]
fn opened_names_and_rules_with_control_characters_survive_structural_batches() {
    let held = names_book("&#1;");
    let twin = names_book("x");
    let mut rng = Rng(0x0dd_ba11_cafe_f00d);
    let mut failures = Vec::new();
    let mut applied = 0;
    for case in 0..300 {
        let ops = (0..1 + rng.below(4))
            .map(|_| structural_op(&mut rng, case))
            .collect::<Vec<_>>();
        let twin_ok = Workbook::open(&twin)
            .unwrap()
            .apply_ops(ops.clone(), CalculationOptions::default())
            .is_ok();
        let mut workbook = Workbook::open(&held).unwrap();
        match workbook.apply_ops(ops.clone(), CalculationOptions::default()) {
            Err(error) if twin_ok => {
                failures.push(format!("case {case}: refused {ops:?}: {error}"));
                continue;
            }
            Err(_) => continue,
            Ok(_) => {}
        }
        applied += 1;
        let expected = names_and_rules(&workbook);
        match workbook.save() {
            Ok(saved) => {
                let reopened = names_and_rules(&Workbook::open(&saved).unwrap());
                if reopened != expected {
                    failures.push(format!(
                        "case {case}: {ops:?} reopened {reopened:?}, expected {expected:?}"
                    ));
                }
            }
            Err(error) => failures.push(format!("case {case}: save failed: {error}")),
        }
    }
    assert!(applied > 150, "only {applied} batches applied");
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

/// Data holds `text` in A2 and A4 (shared and inline) and numbers in B; a bar
/// chart plots categories A2:A4 and values B2:B4, and its title, literal series
/// name and category cache carry `text` the way the file spelled it.
fn chart_book(text: &str) -> Vec<u8> {
    chart_book_with(text, "General")
}

/// [`chart_book`] with `format` (already XML-escaped) as the value cache's format code.
fn chart_book_with(text: &str, format: &str) -> Vec<u8> {
    let chart = format!(
        concat!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>"#,
            r#"<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>"#,
            r#"<c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>{text}</c:v></c:tx>"#,
            r#"<c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>{text}</c:v></c:pt><c:pt idx="1"><c:v>mid</c:v></c:pt><c:pt idx="2"><c:v>{text}</c:v></c:pt></c:strCache></c:strRef></c:cat>"#,
            r#"<c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:formatCode>{format}</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>"#,
            r#"</c:ser></c:barChart><c:catAx><c:axId val="1"/></c:catAx></c:plotArea></c:chart></c:chartSpace>"#
        ),
        text = text,
        format = format
    );
    let parts = [
        (
            "[Content_Types].xml".to_owned(),
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels".to_owned(),
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "sharedStrings", "sharedStrings.xml"),
            ]),
        ),
        (
            "xl/worksheets/sheet1.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData><row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>mid</t></is></c><c r="B3"><v>2</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>{text}</t></is></c><c r="B4"><v>3</v></c></row></sheetData><drawing r:id="d1"/></worksheet>"#
            ),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
            relationships(&[("d1", "drawing", "../drawings/drawing1.xml")]),
        ),
        (
            "xl/sharedStrings.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><sst xmlns="{MAIN}" count="1" uniqueCount="1"><si><t>{text}</t></si></sst>"#
            ),
        ),
        (
            "xl/drawings/drawing1.xml".to_owned(),
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><c:chart r:id="c1"/></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>"#.to_owned(),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_owned(),
            relationships(&[("c1", "chart", "../charts/chart1.xml")]),
        ),
        ("xl/charts/chart1.xml".to_owned(), chart),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name, xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The category cache points of a saved chart, with ST_Xstring escapes decoded.
fn category_cache(saved: &[u8]) -> Vec<String> {
    let chart = String::from_utf8(part(saved, "xl/charts/chart1.xml")).unwrap();
    let cache = &chart[chart.find("<c:strCache>").unwrap()..chart.find("</c:strCache>").unwrap()];
    let mut out = Vec::new();
    let mut rest = cache;
    while let Some(start) = rest.find("<c:v>") {
        let tail = &rest[start + 5..];
        let end = tail.find("</c:v>").unwrap();
        let raw = tail[..end]
            .replace("&#13;", "\r")
            .replace("&amp;", "&")
            .replace("&lt;", "<")
            .replace("&gt;", ">");
        out.push(xstring_decoded(&raw));
        rest = &tail[end..];
    }
    out
}

fn xstring_decoded(text: &str) -> String {
    let mut out = String::new();
    let mut index = 0;
    while index < text.len() {
        let rest = &text[index..];
        if rest.len() >= 7
            && rest.starts_with("_x")
            && rest.as_bytes()[6] == b'_'
            && rest[2..6].chars().all(|c| c.is_ascii_hexdigit())
            && let Some(ch) = u32::from_str_radix(&rest[2..6], 16)
                .ok()
                .and_then(char::from_u32)
        {
            out.push(ch);
            index += 7;
            continue;
        }
        let ch = rest.chars().next().unwrap();
        out.push(ch);
        index += ch.len_utf8();
    }
    out
}

/// A chart over cells holding an opened control character survives a row
/// inserted inside its plotted range (which regenerates the cache), a series
/// edit through the chart API, save and reopen; its cache matches the cells.
/// The API still refuses a control character in a NEW title.
#[test]
fn a_chart_over_opened_control_characters_regenerates_and_saves() {
    let mut failures = Vec::new();
    for ch in controls().chain(['\r']) {
        for (encoding, text) in encodings(ch) {
            let case = format!("U+{:04X} as {encoding}", ch as u32);
            let mut workbook = Workbook::open(&chart_book(&text)).unwrap();
            let result = workbook
                .apply_ops(
                    vec![Op::InsertRows {
                        sheet: SheetId(0),
                        at: 2,
                        count: 1,
                    }],
                    CalculationOptions::default(),
                )
                .and_then(|_| {
                    workbook.edit_charts(
                        vec![betteroffice_xlsx::ChartMutation {
                            sheet: SheetId(0),
                            chart: 0,
                            edit: betteroffice_xlsx::ChartEdit::SetSeries {
                                series: 0,
                                name: None,
                                categories: Some("Data!$A$2:$A$5".into()),
                                values: None,
                            },
                        }],
                        CalculationOptions::default(),
                    )
                })
                .and_then(|_| workbook.save());
            let saved = match result {
                Ok(saved) => saved,
                Err(error) => {
                    failures.push(format!("{case}: {error}"));
                    continue;
                }
            };
            let reopened = Workbook::open(&saved).unwrap();
            let cells = ["A2", "A3", "A4", "A5"]
                .into_iter()
                .map(|at| {
                    match reopened
                        .model()
                        .value(SheetId(0), CellRef::parse_a1(at).unwrap())
                    {
                        CellValue::Text { value } => value,
                        _ => String::new(),
                    }
                })
                .collect::<Vec<_>>();
            let cache = category_cache(&saved);
            let expected = cells
                .iter()
                .filter(|value| !value.is_empty())
                .cloned()
                .collect::<Vec<_>>();
            let cached = cache
                .iter()
                .filter(|value| !value.is_empty())
                .cloned()
                .collect::<Vec<_>>();
            if cached != expected {
                failures.push(format!("{case}: cache {cache:?} vs cells {cells:?}"));
            }
            let title = workbook.edit_charts(
                vec![betteroffice_xlsx::ChartMutation {
                    sheet: SheetId(0),
                    chart: 0,
                    edit: betteroffice_xlsx::ChartEdit::SetChartTitle {
                        title: Some("new\u{1}".into()),
                    },
                }],
                CalculationOptions::default(),
            );
            if title.is_ok() {
                failures.push(format!("{case}: a new title with U+0001 was accepted"));
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

/// Random structural batches over an opened chart whose cells, title, literal
/// series name and cache hold a control character are refused no more often
/// than over a control-free twin, and always save.
#[test]
fn opened_chart_control_characters_survive_structural_batches() {
    let held = chart_book("a_x0001_b");
    let twin = chart_book("ab");
    let mut rng = Rng(0xc4a7_0001_5eed_beef);
    let mut failures = Vec::new();
    let mut applied = 0;
    for case in 0..200 {
        let ops = (0..1 + rng.below(3))
            .map(|_| match structural_op(&mut rng, case) {
                Op::MoveSheet { .. } => Op::InsertRows {
                    sheet: SheetId(0),
                    at: 2,
                    count: 1,
                },
                Op::InsertRows { at, count, .. } => Op::InsertRows {
                    sheet: SheetId(0),
                    at,
                    count,
                },
                Op::DeleteRows { at, count, .. } => Op::DeleteRows {
                    sheet: SheetId(0),
                    at,
                    count,
                },
                Op::InsertCols { at, count, .. } => Op::InsertCols {
                    sheet: SheetId(0),
                    at,
                    count,
                },
                Op::DeleteCols { at, count, .. } => Op::DeleteCols {
                    sheet: SheetId(0),
                    at,
                    count,
                },
                Op::RenameSheet { name, .. } => Op::RenameSheet {
                    sheet: SheetId(0),
                    name,
                },
                other => other,
            })
            .collect::<Vec<_>>();
        let twin_ok = Workbook::open(&twin)
            .unwrap()
            .apply_ops(ops.clone(), CalculationOptions::default())
            .is_ok();
        let mut workbook = Workbook::open(&held).unwrap();
        match workbook.apply_ops(ops.clone(), CalculationOptions::default()) {
            Err(error) if twin_ok => {
                failures.push(format!("case {case}: refused {ops:?}: {error}"));
                continue;
            }
            Err(_) => continue,
            Ok(_) => {}
        }
        applied += 1;
        if let Err(error) = workbook.save() {
            failures.push(format!("case {case}: {ops:?} save failed: {error}"));
        }
    }
    assert!(applied > 100, "only {applied} batches applied");
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

fn chart_part(saved: &[u8]) -> String {
    String::from_utf8(part(saved, "xl/charts/chart1.xml")).unwrap()
}

fn insert_row(workbook: &mut Workbook) {
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 2,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
}

/// A cache format code is not cell text: a literal `_xHHHH_` in it survives
/// every regeneration unchanged.
#[test]
fn a_cache_format_code_keeps_a_literal_escape_through_regenerations() {
    let mut workbook = Workbook::open(&chart_book_with("ab", "0&quot;_x0041_&quot;")).unwrap();
    insert_row(&mut workbook);
    let saved = workbook.save().unwrap();
    assert!(
        chart_part(&saved).contains("<c:formatCode>0&quot;_x0041_&quot;</c:formatCode>"),
        "{}",
        chart_part(&saved)
    );
    let mut reopened = Workbook::open(&saved).unwrap();
    insert_row(&mut reopened);
    let again = reopened.save().unwrap();
    assert!(
        chart_part(&again).contains("<c:formatCode>0&quot;_x0041_&quot;</c:formatCode>"),
        "{}",
        chart_part(&again)
    );
}

fn edit_chart(workbook: &mut Workbook, edit: betteroffice_xlsx::ChartEdit) {
    workbook
        .edit_charts(
            vec![betteroffice_xlsx::ChartMutation {
                sheet: SheetId(0),
                chart: 0,
                edit,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
}

fn set_series_name(workbook: &mut Workbook, text: String) {
    edit_chart(
        workbook,
        betteroffice_xlsx::ChartEdit::SetSeries {
            series: 0,
            name: Some(betteroffice_xlsx::ChartSeriesName::Literal { text }),
            categories: None,
            values: None,
        },
    );
}

fn series_name(workbook: &Workbook) -> Option<String> {
    workbook.charts(SheetId(0)).unwrap()[0].series[0]
        .name
        .clone()
}

fn chart_title(workbook: &Workbook) -> Option<String> {
    workbook.charts(SheetId(0)).unwrap()[0].title.clone()
}

/// A literal series name set through the API is written as ST_Xstring, so a
/// literal `_x0041_` stays literal for every reader that decodes `c:v`.
#[test]
fn a_literal_series_name_reads_back_as_set_and_resubmits_unchanged() {
    let mut workbook = Workbook::open(&chart_book("ab")).unwrap();
    set_series_name(&mut workbook, "_x0041_".into());
    assert_eq!(series_name(&workbook).as_deref(), Some("_x0041_"));
    let reported = series_name(&workbook).unwrap();
    set_series_name(&mut workbook, reported);
    assert_eq!(series_name(&workbook).as_deref(), Some("_x0041_"));
    let saved = workbook.save().unwrap();
    assert!(
        chart_part(&saved).contains("<c:tx><c:v>_x005F_x0041_</c:v></c:tx>"),
        "{}",
        chart_part(&saved)
    );
    let mut reopened = Workbook::open(&saved).unwrap();
    assert_eq!(series_name(&reopened).as_deref(), Some("_x0041_"));
    let reported = series_name(&reopened).unwrap();
    set_series_name(&mut reopened, reported);
    let again = Workbook::open(&reopened.save().unwrap()).unwrap();
    assert_eq!(series_name(&again).as_deref(), Some("_x0041_"));
}

#[test]
fn an_opened_literal_series_name_reports_its_escapes_decoded() {
    let mut workbook = Workbook::open(&chart_book("a_x0001_b")).unwrap();
    assert_eq!(series_name(&workbook).as_deref(), Some("a\u{1}b"));
    let saved = workbook.save().unwrap();
    assert!(
        chart_part(&saved).contains("<c:tx><c:v>a_x0001_b</c:v></c:tx>"),
        "{}",
        chart_part(&saved)
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(series_name(&reopened).as_deref(), Some("a\u{1}b"));
    insert_row(&mut workbook);
    let saved = workbook.save().unwrap();
    assert!(
        chart_part(&saved).contains("<c:tx><c:v>a_x0001_b</c:v></c:tx>"),
        "{}",
        chart_part(&saved)
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(series_name(&reopened).as_deref(), Some("a\u{1}b"));
    let reported = series_name(&workbook).unwrap();
    set_series_name(&mut workbook, reported);
    let saved = workbook.save().unwrap();
    assert!(
        chart_part(&saved).contains("<c:tx><c:v>a_x0001_b</c:v></c:tx>"),
        "{}",
        chart_part(&saved)
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(series_name(&reopened).as_deref(), Some("a\u{1}b"));
}

/// A literal series name goes to `c:v`, which carries any character as
/// `_xHHHH_`, so a new one with a control character is accepted; a title goes
/// to `a:t`, which cannot, and is still refused.
#[test]
fn a_new_literal_series_name_takes_a_control_character_and_a_title_does_not() {
    let mut workbook = Workbook::open(&chart_book("ab")).unwrap();
    set_series_name(&mut workbook, "x\u{1}".into());
    assert_eq!(series_name(&workbook).as_deref(), Some("x\u{1}"));
    let saved = workbook.save().unwrap();
    assert!(
        chart_part(&saved).contains("<c:tx><c:v>x_x0001_</c:v></c:tx>"),
        "{}",
        chart_part(&saved)
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(series_name(&reopened).as_deref(), Some("x\u{1}"));
    let error = workbook
        .edit_charts(
            vec![betteroffice_xlsx::ChartMutation {
                sheet: SheetId(0),
                chart: 0,
                edit: betteroffice_xlsx::ChartEdit::SetChartTitle {
                    title: Some("new\u{1}".into()),
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(
        error
            .to_string()
            .contains("title must contain at most 32767 XML text characters"),
        "{error}"
    );
}

#[test]
fn a_rich_text_title_keeps_a_literal_escape() {
    let mut workbook = Workbook::open(&chart_book("_x0041_")).unwrap();
    assert_eq!(chart_title(&workbook).as_deref(), Some("_x0041_"));
    assert_eq!(series_name(&workbook).as_deref(), Some("A"));
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(chart_title(&reopened).as_deref(), Some("_x0041_"));
    let title = chart_title(&workbook);
    edit_chart(
        &mut workbook,
        betteroffice_xlsx::ChartEdit::SetChartTitle { title },
    );
    let saved = workbook.save().unwrap();
    assert!(
        chart_part(&saved).contains("<a:t>_x0041_</a:t>"),
        "{}",
        chart_part(&saved)
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(chart_title(&reopened).as_deref(), Some("_x0041_"));
}

/// A book whose hyperlink target and defined-name comment hold a line break as
/// character references.
fn attribute_book() -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Noted" comment="x&#13;&#10;y">Data!$A$1</definedName></definedNames></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[("rId1", "worksheet", "worksheets/sheet1.xml")]),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><hyperlinks><hyperlink ref="A1" r:id="rIdLink"/></hyperlinks></worksheet>"#
            ),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rIdLink" Type="{DOC_RELS}/hyperlink" Target="https://example.test/a&#13;b" TargetMode="External"/></Relationships>"#
            ),
        ),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Attributes written back from the source (relationship targets, defined-name
/// comments) keep a CR and LF that came as character references.
#[test]
fn source_attributes_keep_line_breaks_when_their_part_is_rewritten() {
    let mut workbook = Workbook::open(&attribute_book()).unwrap();
    workbook
        .set_hyperlink(
            SheetId(0),
            betteroffice_xlsx::Hyperlink {
                range: CellRange::parse_a1("B1").unwrap(),
                external_target: Some("https://example.test/c".into()),
                location: None,
                tooltip: None,
                display: None,
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .apply_ops(
            vec![Op::SetDefinedName {
                name: "Noted".into(),
                refers_to: "Data!$A$1+1".into(),
                scope: None,
                hidden: None,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let names = String::from_utf8(part(&saved, "xl/workbook.xml")).unwrap();
    assert!(names.contains(r#"comment="x&#13;&#10;y""#), "{names}");
    let reopened = Workbook::open(&saved).unwrap();
    let targets = reopened.model().sheets[0]
        .hyperlinks
        .iter()
        .map(|link| link.external_target.clone())
        .collect::<Vec<_>>();
    assert_eq!(
        targets,
        vec![
            Some("https://example.test/a\rb".to_owned()),
            Some("https://example.test/c".to_owned())
        ]
    );
}
