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
    source_at(app, "docProps/app.xml", "docProps/app.xml")
}

fn source_at(app: &str, path: &str, target: &str) -> (Workbook, PreservedPackage) {
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
    let rels = parts
        .iter_mut()
        .find(|(name, _)| name == "_rels/.rels")
        .unwrap();
    rels.1 = String::from_utf8(rels.1.clone())
        .unwrap()
        .replace(
            "</Relationships>",
            &format!(r#"<Relationship Id="rIdApp" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="{target}"/></Relationships>"#),
        )
        .into_bytes();
    parts.push((path.into(), app.as_bytes().to_vec()));
    let parsed = xlsx_parse::parse_workbook_with_package(&parts).unwrap();
    (parsed.workbook, parsed.package)
}

fn save(model: &Workbook, package: &PreservedPackage, origins: &[Option<usize>]) -> String {
    save_at(model, package, origins, "docProps/app.xml")
}

fn save_at(
    model: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    path: &str,
) -> String {
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
    String::from_utf8(parts[path].clone()).unwrap()
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
    assert_eq!(
        save(&model, &package, &[Some(1), Some(0), Some(2)]),
        expected
    );
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
    assert_eq!(
        save(&model, &package, &[Some(0), Some(1), Some(2)]),
        expected
    );
}

#[test]
fn unreadable_app_properties_are_preserved_and_the_save_succeeds() {
    for app in [
        APP.replace(
            "<Properties ",
            "<!DOCTYPE Properties [<!ENTITY x \"y\">]><Properties ",
        ),
        APP.replace("</Properties>", ""),
    ] {
        let (mut model, package) = source(&app);
        model.sheets[1].name = "History".into();
        assert_eq!(save(&model, &package, &[Some(0), Some(1), Some(2)]), app);
        model.sheets.swap(0, 1);
        model.defined_names[0].local_sheet = Some(SheetId(0));
        assert_eq!(save(&model, &package, &[Some(1), Some(0), Some(2)]), app);
    }
}

#[test]
fn titles_with_comments_or_cdata_are_preserved() {
    for app in [
        APP.replacen(
            "<vt:lpstr>Version</vt:lpstr>",
            "<vt:lpstr>Ver<!-- kept -->sion</vt:lpstr>",
            1,
        ),
        APP.replace(
            "<vt:lpstr>Version!Print_Area</vt:lpstr>",
            "<vt:lpstr>Version!<!-- kept -->Print_Area</vt:lpstr>",
        ),
        APP.replacen(
            "<vt:lpstr>Version</vt:lpstr>",
            "<vt:lpstr><![CDATA[Version]]></vt:lpstr>",
            1,
        ),
    ] {
        let (mut model, package) = source(&app);
        model.sheets[1].name = "History".into();
        assert_eq!(save(&model, &package, &[Some(0), Some(1), Some(2)]), app);
    }
}

#[test]
fn swapping_sheets_and_their_names_patches_qualified_titles_by_origin() {
    let (mut model, package) = source(APP);
    model.sheets.swap(0, 1);
    model.sheets[0].name = "Instructions".into();
    model.sheets[1].name = "Version".into();
    model.defined_names[0].local_sheet = Some(SheetId(0));
    model.defined_names[0].formula = "Instructions!$A$1".into();
    let expected = APP.replace("Version!Print_Area", "Instructions!Print_Area");
    assert_eq!(
        save(&model, &package, &[Some(1), Some(0), Some(2)]),
        expected
    );
}

#[test]
fn app_properties_are_found_through_the_package_relationship() {
    let (mut model, package) = source_at(APP, "meta/extended.xml", "/meta/extended.xml");
    model.sheets[1].name = "History".into();
    let expected = APP
        .replacen(
            "<vt:lpstr>Version</vt:lpstr>",
            "<vt:lpstr>History</vt:lpstr>",
            1,
        )
        .replace("Version!Print_Area", "History!Print_Area");
    assert_eq!(
        save_at(
            &model,
            &package,
            &[Some(0), Some(1), Some(2)],
            "meta/extended.xml"
        ),
        expected
    );
}
