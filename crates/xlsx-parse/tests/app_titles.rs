use std::collections::BTreeMap;

use xlsx_model::{DefinedName, Sheet, SheetId, Workbook};
use xlsx_parse::{PreservedPackage, SaveEdits};

const APP: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns:x="urn:custom">
<Application>Spreadsheet</Application><x:keep attr="exact &amp; intact"/>
<HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant><vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>
<TitlesOfParts><vt:vector size="5" baseType="lpstr"><vt:lpstr>Instructions</vt:lpstr><vt:lpstr>Version</vt:lpstr><vt:lpstr>Models</vt:lpstr><vt:lpstr>Version!Print_Area</vt:lpstr><vt:lpstr>Version</vt:lpstr></vt:vector></TitlesOfParts>
<Company></Company><x:tail>Keep  spaces</x:tail></Properties>"#;

fn source(app: &str) -> (Workbook, PreservedPackage) {
    let mut model = Workbook::default();
    for name in ["Instructions", "Version", "Models"] {
        model.sheets.push(Sheet::new(name));
    }
    model.defined_names = vec![
        DefinedName {
            name: "_xlnm.Print_Area".into(),
            formula: "Version!$A$1".into(),
            local_sheet: Some(SheetId(1)),
            hidden: false,
        },
        DefinedName {
            name: "Version".into(),
            formula: "1".into(),
            local_sheet: None,
            hidden: false,
        },
    ];
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    parts.push(("docProps/app.xml".into(), app.as_bytes().to_vec()));
    let parsed = xlsx_parse::parse_workbook_with_package(&parts).unwrap();
    (parsed.workbook, parsed.package)
}

fn save(model: &Workbook, package: &PreservedPackage, origins: &[Option<usize>]) -> String {
    let strings = origins
        .iter()
        .map(|origin| {
            origin
                .map(|index| package.source_shared_string_cells(index))
                .unwrap_or_default()
        })
        .collect::<Vec<_>>();
    let parts = xlsx_parse::serialize_workbook_with_package_and_origins_after_edits(
        model,
        package,
        origins,
        &strings,
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap()
    .into_iter()
    .map(|(path, bytes)| (path, bytes.into_owned()))
    .collect::<BTreeMap<_, _>>();
    String::from_utf8(parts["docProps/app.xml"].clone()).unwrap()
}

#[test]
fn rename_patches_only_sheet_titles_and_qualified_names() {
    for (name, qualified, escaped) in [
        ("History", "History", "History"),
        ("Version History", "'Version History'", "Version History"),
        ("Owner's History", "'Owner''s History'", "Owner's History"),
        ("A1", "'A1'", "A1"),
        ("R1C1", "'R1C1'", "R1C1"),
        ("A&B", "'A&amp;B'", "A&amp;B"),
    ] {
        let (mut model, package) = source(APP);
        model.sheets[1].name = name.into();
        let expected = APP
            .replacen(
                "<vt:lpstr>Version</vt:lpstr>",
                &format!("<vt:lpstr>{escaped}</vt:lpstr>"),
                1,
            )
            .replace("Version!Print_Area", &format!("{qualified}!Print_Area"));
        assert_eq!(
            save(&model, &package, &[Some(0), Some(1), Some(2)]),
            expected
        );
    }
}

#[test]
fn ambiguous_or_missing_titles_are_preserved() {
    for app in [
        APP.replace("<vt:i4>3</vt:i4>", "<vt:i4>2</vt:i4>"),
        APP.replace("<vt:lpstr>Version</vt:lpstr>", "<vt:lpstr>Stale</vt:lpstr>"),
        APP.replace("HeadingPairs", "OtherHeadingPairs"),
        APP.replace("TitlesOfParts", "OtherTitlesOfParts"),
    ] {
        let (mut model, package) = source(&app);
        model.sheets[1].name = "History".into();
        assert_eq!(save(&model, &package, &[Some(0), Some(1), Some(2)]), app);
    }
}

#[test]
fn moving_sheets_reorders_only_worksheet_titles() {
    let (mut model, package) = source(APP);
    model.sheets.swap(0, 1);
    model.defined_names[0].local_sheet = Some(SheetId(0));
    let expected = APP.replace(
        "<vt:lpstr>Instructions</vt:lpstr><vt:lpstr>Version</vt:lpstr>",
        "<vt:lpstr>Version</vt:lpstr><vt:lpstr>Instructions</vt:lpstr>",
    );
    assert_eq!(save(&model, &package, &[Some(1), Some(0), Some(2)]), expected);
}

#[test]
fn adding_and_deleting_sheets_currently_preserve_app_titles() {
    let (mut model, package) = source(APP);
    model.sheets.push(Sheet::new("Extra"));
    assert_eq!(
        save(&model, &package, &[Some(0), Some(1), Some(2), None]),
        APP
    );
    model.sheets.pop();
    model.sheets.remove(1);
    model.defined_names.remove(0);
    assert_eq!(save(&model, &package, &[Some(0), Some(2)]), APP);
}

#[test]
fn qualified_titles_before_worksheets_and_quoted_sources_are_patched() {
    let app = APP.replace(
        "<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant>",
        "",
    );
    let app = app.replace(
        "</vt:vector></HeadingPairs>",
        "<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant></vt:vector></HeadingPairs>",
    );
    let app = app.replace(
        "<vt:lpstr>Instructions</vt:lpstr><vt:lpstr>Version</vt:lpstr><vt:lpstr>Models</vt:lpstr><vt:lpstr>Version!Print_Area</vt:lpstr><vt:lpstr>Version</vt:lpstr>",
        "<vt:lpstr>'Version'!Print_Area</vt:lpstr><vt:lpstr>Version</vt:lpstr><vt:lpstr>Instructions</vt:lpstr><vt:lpstr>Version</vt:lpstr><vt:lpstr>Models</vt:lpstr>",
    );
    let (mut model, package) = source(&app);
    model.sheets[1].name = "Version History".into();
    let expected = app
        .replace("'Version'!Print_Area", "'Version History'!Print_Area")
        .replace(
            "<vt:lpstr>Instructions</vt:lpstr><vt:lpstr>Version</vt:lpstr>",
            "<vt:lpstr>Instructions</vt:lpstr><vt:lpstr>Version History</vt:lpstr>",
        );
    assert_eq!(save(&model, &package, &[Some(0), Some(1), Some(2)]), expected);
}
