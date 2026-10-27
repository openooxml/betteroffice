use std::collections::BTreeMap;
use xlsx_model::{DefinedName, Sheet, SheetId, Workbook};
use xlsx_parse::{ParseError, PreservedPackage, SaveEdits};

fn source(names: Option<&str>, prefixed: bool) -> (Workbook, PreservedPackage) {
    let mut model = Workbook::default();
    model.sheets.push(Sheet::new("Data"));
    model.sheets.push(Sheet::new("Other"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let workbook = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/workbook.xml")
        .unwrap();
    let mut xml = String::from_utf8(workbook.1.clone()).unwrap();
    if let Some(names) = names {
        xml = xml.replace(
            "</workbook>",
            &format!("<definedNames>{names}</definedNames><calcPr calcId=\"1\"/></workbook>"),
        );
    } else {
        xml = xml.replace("</workbook>", "<calcPr calcId=\"1\"/></workbook>");
    }
    if prefixed {
        xml = xml.replace("xmlns=", "xmlns:s=");
        for name in [
            "workbook",
            "sheets",
            "sheet",
            "definedNames",
            "definedName",
            "calcPr",
        ] {
            xml = xml
                .replace(&format!("<{name}>"), &format!("<s:{name}>"))
                .replace(&format!("<{name} "), &format!("<s:{name} "))
                .replace(&format!("</{name}>"), &format!("</s:{name}>"));
        }
    }
    workbook.1 = xml.into_bytes();
    let parsed = xlsx_parse::parse_workbook_with_package(&parts).unwrap();
    (parsed.workbook, parsed.package)
}

fn save(
    model: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    changed: bool,
) -> Result<BTreeMap<String, Vec<u8>>, ParseError> {
    let strings = origins
        .iter()
        .map(|origin| package.source_shared_string_cells(origin.unwrap()))
        .collect::<Vec<_>>();
    xlsx_parse::serialize_workbook_with_package_and_origins_after_edits(
        model,
        package,
        origins,
        &strings,
        SaveEdits {
            changed,
            moved_references: false,
        },
    )
    .map(|parts| {
        parts
            .into_iter()
            .map(|(path, bytes)| (path, bytes.into_owned()))
            .collect()
    })
}

fn name(name: &str, formula: &str, sheet: Option<SheetId>) -> DefinedName {
    DefinedName {
        name: name.into(),
        formula: formula.into(),
        local_sheet: sheet,
        hidden: false,
    }
}

#[test]
fn writer_inserts_names_into_missing_and_empty_prefixed_containers() {
    for names in [None, Some("")] {
        let (mut model, package) = source(names, true);
        let before = save(&model, &package, &[Some(0), Some(1)], false).unwrap();
        model.defined_names.push(name("Rate", "1", None));
        let saved = save(&model, &package, &[Some(0), Some(1)], true).unwrap();
        let parsed =
            xlsx_parse::parse_workbook(&saved.clone().into_iter().collect::<Vec<_>>()).unwrap();
        assert_eq!(parsed.defined_names, model.defined_names);
        let xml = String::from_utf8(saved["xl/workbook.xml"].clone()).unwrap();
        assert!(xml.find("definedNames").unwrap() < xml.find("calcPr").unwrap());
        for (path, bytes) in before {
            if path != "xl/workbook.xml" {
                assert_eq!(saved[&path], bytes, "{path}");
            }
        }
    }
}

#[test]
fn writer_matches_names_after_insertions_deletions_and_sheet_reordering() {
    let retained =
        r#"<definedName name="Rate" localSheetId="1" comment="same &amp; exact">2</definedName>"#;
    let (mut model, package) = source(
        Some(&format!(
            r#"<definedName name="Rate" localSheetId="0" comment="first">1</definedName>{retained}"#
        )),
        false,
    );
    let original = model.defined_names.clone();
    model.defined_names.insert(0, name("Before", "3", None));
    model.defined_names.push(name("After", "4", None));
    let saved = save(&model, &package, &[Some(0), Some(1)], true).unwrap();
    let xml = String::from_utf8(saved["xl/workbook.xml"].clone()).unwrap();
    assert!(xml.contains(retained));
    assert!(xml.contains("name=\"Before\""));
    assert!(xml.contains("name=\"After\""));
    model.defined_names = original;
    model.sheets.swap(0, 1);
    model.defined_names[0].local_sheet = Some(SheetId(1));
    model.defined_names[0].hidden = true;
    model.defined_names[1].local_sheet = Some(SheetId(0));
    let saved = save(&model, &package, &[Some(1), Some(0)], true).unwrap();
    let xml = String::from_utf8(saved["xl/workbook.xml"].clone()).unwrap();
    assert!(xml.contains("comment=\"first\""));
    assert!(xml.contains("comment=\"same &amp; exact\""));
    let reopened = xlsx_parse::parse_workbook(&saved.into_iter().collect::<Vec<_>>()).unwrap();
    assert_eq!(reopened.defined_names, model.defined_names);
    model.defined_names.remove(0);
    assert_eq!(
        xlsx_parse::parse_workbook(
            &save(&model, &package, &[Some(1), Some(0)], true)
                .unwrap()
                .into_iter()
                .collect::<Vec<_>>()
        )
        .unwrap()
        .defined_names,
        model.defined_names
    );
    model.defined_names.clear();
    assert!(
        !String::from_utf8(
            save(&model, &package, &[Some(1), Some(0)], true).unwrap()["xl/workbook.xml"].clone()
        )
        .unwrap()
        .contains("definedNames")
    );
}

#[test]
fn writer_refuses_source_entry_count_mismatch_instead_of_returning_stale_xml() {
    let (mut model, package) = source(
        Some(r#"<wrapper><definedName name="Rate">1</definedName></wrapper>"#),
        false,
    );
    let before = save(&model, &package, &[Some(0), Some(1)], false).unwrap();
    model.defined_names.push(name("Other", "2", None));
    assert!(matches!(
        save(&model, &package, &[Some(0), Some(1)], true),
        Err(ParseError::UnsupportedEdit(_))
    ));
    model.defined_names.pop();
    assert_eq!(
        save(&model, &package, &[Some(0), Some(1)], false).unwrap(),
        before
    );
}

#[test]
fn writer_preserves_inherited_namespaces_when_entries_redeclare_them() {
    let namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for prefixed in [false, true] {
        let declaration = if prefixed { "xmlns:s" } else { "xmlns" };
        let retained = format!(
            r#"<definedName {declaration}="{namespace}" name="Keep" comment="same &amp; exact">2</definedName>"#
        );
        let entries = format!(r#"<definedName name="Rate">1</definedName>{retained}"#);
        let (mut model, package) = source(Some(&entries), prefixed);
        assert_eq!(model.defined_names.len(), 2);
        let before = save(&model, &package, &[Some(0), Some(1)], false).unwrap();
        model.defined_names[0].formula = "3".into();
        model.defined_names.push(name("Added", "4", None));
        let saved = save(&model, &package, &[Some(0), Some(1)], true).unwrap();
        let xml = String::from_utf8(saved["xl/workbook.xml"].clone()).unwrap();
        let retained = if prefixed {
            retained
                .replace("<definedName ", "<s:definedName ")
                .replace("</definedName>", "</s:definedName>")
        } else {
            retained
        };
        assert!(xml.contains(&retained));
        let reopened =
            xlsx_parse::parse_workbook(&saved.clone().into_iter().collect::<Vec<_>>()).unwrap();
        assert_eq!(reopened.defined_names, model.defined_names);
        for (path, bytes) in before {
            if path != "xl/workbook.xml" {
                assert_eq!(saved[&path], bytes, "{path}");
            }
        }
    }
}

fn junk_names(count: usize) -> String {
    (0..count)
        .map(|index| match index % 3 {
            0 => format!(r#"<definedName name="_junk{index}" hidden="1">#REF!</definedName>"#),
            1 => format!(
                r#"<definedName name="Area{index}" localSheetId="1" comment="a &amp; b">Other!$A${}</definedName>"#,
                index % 1000 + 1
            ),
            _ => format!(r#"<definedName name="Wide{index}">Data!$A$1:$B$2</definedName>"#),
        })
        .collect()
}

fn entries(xml: &str) -> Vec<&str> {
    let start = xml.find("<definedNames>").unwrap() + "<definedNames>".len();
    let end = xml.find("</definedNames>").unwrap();
    xml[start..end].split_inclusive("</definedName>").collect()
}

#[test]
fn a_name_table_past_the_old_count_cap_opens_and_saves_every_entry_verbatim() {
    const COUNT: usize = 66_391;
    let names = junk_names(COUNT);
    let (mut model, package) = source(Some(&names), false);
    assert_eq!(model.defined_names.len(), COUNT);
    assert!(model.defined_names[0].hidden);
    assert_eq!(model.defined_names[1].local_sheet, Some(SheetId(1)));

    let saved = save(&model, &package, &[Some(0), Some(1)], true).unwrap();
    let xml = String::from_utf8(saved["xl/workbook.xml"].clone()).unwrap();
    assert!(xml.contains(&format!("<definedNames>{names}</definedNames>")));

    model.defined_names[1].formula = "Other!$C$3".into();
    let saved = save(&model, &package, &[Some(0), Some(1)], true).unwrap();
    let xml = String::from_utf8(saved["xl/workbook.xml"].clone()).unwrap();
    let written = entries(&xml);
    let source = names.split_inclusive("</definedName>").collect::<Vec<_>>();
    assert_eq!(written.len(), source.len());
    for (index, (written, source)) in written.iter().zip(&source).enumerate() {
        if index != 1 {
            assert_eq!(written, source, "entry {index}");
        }
    }
    assert!(written[1].contains("comment=\"a &amp; b\""));
    assert!(written[1].contains("Other!$C$3"));
    let reopened = xlsx_parse::parse_workbook(&saved.into_iter().collect::<Vec<_>>()).unwrap();
    assert_eq!(reopened.defined_names, model.defined_names);
}

#[test]
fn a_name_table_past_the_byte_bound_is_refused() {
    let formula = "1".repeat(64 * 1024);
    let count = xlsx_parse::MAX_DEFINED_NAME_BYTES / formula.len() + 1;
    let names = (0..count)
        .map(|index| format!(r#"<definedName name="Big{index}">{formula}</definedName>"#))
        .collect::<String>();
    let mut model = Workbook::default();
    model.sheets.push(Sheet::new("Data"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let workbook = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/workbook.xml")
        .unwrap();
    workbook.1 = String::from_utf8(workbook.1.clone())
        .unwrap()
        .replace(
            "</workbook>",
            &format!("<definedNames>{names}</definedNames></workbook>"),
        )
        .into_bytes();
    assert_eq!(
        xlsx_parse::parse_workbook(&parts).unwrap_err(),
        ParseError::TooManyDefinedNames
    );
    assert!(matches!(
        xlsx_parse::parse_workbook_with_package(&parts),
        Err(ParseError::TooManyDefinedNames)
    ));
}
