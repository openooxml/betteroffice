use betteroffice_xlsx::{
    CalculationOptions, CellAddress, CellRef, CellValue, SheetId, SnapshotBudget, Workbook,
    WorkbookSnapshotBuilder, WorkbookSnapshotEncoder,
};
use xlsx_model::CellProvider;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DECL: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const STYLES: &str = concat!(
    r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    r#"<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
    r#"<fills count="1"><fill><patternFill patternType="none"/></fill></fills>"#,
    r#"<borders count="1"><border/></borders>"#,
    r#"<cellStyleXfs count="1">"#,
    r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
    r#"<cellXfs count="1">"#,
    r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
    "</styleSheet>",
);
const CONTENT_TYPES: &str = r#"<Types
 xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml"
 ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml"
 ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml"
 ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#;
const SHEET: &str = concat!(
    r#"<row r="1">"#,
    r#"<c r="A1"><f>(A1+10)/2</f><v>10</v></c><c r="B1"><f>A1*2</f><v>20</v></c>"#,
    r#"<c r="C1"><v>5</v></c><c r="D1"><f>C1+1</f><v>6</v></c>"#,
    r#"<c r="E1"><f>F1/2+10</f><v>20</v></c><c r="F1"><f>E1</f><v>20</v></c>"#,
    "</row>",
);
const ITERATE_AND_RECALCULATE: &str = r#"<calcPr iterate="1" fullCalcOnLoad="1"/>"#;
const CACHED: &[(&str, &str)] = &[
    ("A1", "10"),
    ("B1", "20"),
    ("D1", "6"),
    ("E1", "20"),
    ("F1", "20"),
];

fn package(calc_pr: &str) -> Vec<u8> {
    package_with_sheet(calc_pr, SHEET)
}

fn package_with_sheet(calc_pr: &str, sheet: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", CONTENT_TYPES.to_owned()),
        (
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="{PKG}">
<Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#
            ),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>{calc_pr}</workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="{PKG}">
<Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="{RELS}/styles" Target="styles.xml"/>
</Relationships>"#
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(r#"<worksheet xmlns="{MAIN}"><sheetData>{sheet}</sheetData></worksheet>"#),
        ),
        ("xl/styles.xml", STYLES.to_owned()),
    ];
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), [DECL, &xml].concat().into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_values(bytes: &[u8]) -> Vec<(String, Option<String>)> {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let (_, sheet) = parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .unwrap();
    std::str::from_utf8(sheet)
        .unwrap()
        .split("<c ")
        .skip(1)
        .map(|cell| {
            let body = cell.split("</c>").next().unwrap();
            let address = body
                .split_once("r=\"")
                .unwrap()
                .1
                .split('"')
                .next()
                .unwrap();
            let value = body
                .split_once("<v>")
                .and_then(|(_, rest)| rest.split_once("</v>"))
                .map(|(value, _)| value.to_owned());
            (address.to_owned(), value)
        })
        .collect()
}

fn assert_saved_values(bytes: &[u8], expected: &[(&str, &str)], context: &str) {
    let values = saved_values(bytes);
    for &(address, expected) in expected {
        let value = values
            .iter()
            .find(|(name, _)| name == address)
            .and_then(|(_, value)| value.as_deref());
        assert_eq!(value, Some(expected), "{context}: saved {address}");
    }
}

fn open_recalculated(source: &[u8]) -> Workbook {
    Workbook::open_recalculated(source, CalculationOptions::default()).unwrap()
}

fn restore_sessions(workbook: &Workbook) -> [(&'static str, Workbook); 2] {
    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(workbook, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(workbook).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    let snapshot = builder.finish().unwrap().into_parts().0;
    let checkpoint = workbook.recovery_checkpoint(0, None).unwrap();
    let (recovered, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    [("snapshot", snapshot), ("recovery", recovered)]
}

#[test]
fn full_calculation_keeps_iterative_caches() {
    let source = package(ITERATE_AND_RECALCULATE);
    let opened = Workbook::open(&source).unwrap().save().unwrap();
    assert_saved_values(&opened, CACHED, "open");
    assert_eq!(opened, source, "open keeps the original package");

    let workbook = open_recalculated(&source);
    assert_eq!(
        workbook.model().value(SheetId(0), CellRef::new(0, 0)),
        CellValue::Number { value: 10.0 },
        "the recalculated model shows the cached A1 value"
    );
    let saved = workbook.save().unwrap();
    assert_saved_values(&saved, CACHED, "open_recalculated");
    assert_eq!(saved, source);
    let status = workbook.calculation_status();
    let unevaluated: Vec<_> = status
        .unevaluated_cells
        .iter()
        .map(|cell| cell.address.as_str())
        .collect();
    assert_eq!(unevaluated, ["A1", "E1", "F1"]);
}

#[test]
fn iterative_caches_without_a_full_calculation_request_keep_the_original_package() {
    let source = package(r#"<calcPr iterate="true" iterateCount="100" iterateDelta="0.001"/>"#);
    for (mode, workbook) in [
        ("open", Workbook::open(&source).unwrap()),
        ("open_recalculated", open_recalculated(&source)),
    ] {
        assert!(workbook.model().iterate, "{mode} reads iterate=true");
        let saved = workbook.save().unwrap();
        assert_saved_values(&saved, CACHED, mode);
        assert_eq!(saved, source, "{mode} keeps the original package");
    }
}

#[test]
fn editing_outside_an_iterative_cycle_recalculates_the_acyclic_reader() {
    let mut workbook = Workbook::open(&package(ITERATE_AND_RECALCULATE)).unwrap();
    let result = workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 2),
            "7",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        result.changed,
        [CellAddress {
            sheet: SheetId(0),
            cell: CellRef::new(0, 3),
        }]
    );
    assert!(workbook.calculation_status().unevaluated_cells.is_empty());
    assert_saved_values(
        &workbook.save().unwrap(),
        &[
            ("A1", "10"),
            ("B1", "20"),
            ("D1", "8"),
            ("E1", "20"),
            ("F1", "20"),
        ],
        "after editing C1",
    );
}

#[test]
fn editing_an_iterative_cycle_keeps_its_other_members_cached() {
    let mut workbook = Workbook::open(&package(ITERATE_AND_RECALCULATE)).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 4),
            "=F1/2+20",
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    assert_saved_values(
        &saved,
        &[("A1", "10"), ("B1", "20"), ("D1", "6"), ("F1", "20")],
        "after editing E1",
    );
    assert_eq!(
        saved_values(&saved)
            .into_iter()
            .find(|(address, _)| address == "E1"),
        Some(("E1".to_owned(), None)),
        "the edited formula has no value the engine can compute, so it saves none"
    );
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let (_, sheet) = parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .unwrap();
    assert!(
        std::str::from_utf8(sheet)
            .unwrap()
            .contains("<f>F1/2+20</f>"),
        "the edited formula is saved"
    );
    let status = workbook.calculation_status();
    assert_eq!(status.unevaluated_cells.len(), 2);
    for address in ["E1", "F1"] {
        assert!(
            status
                .unevaluated_cells
                .iter()
                .any(|cell| cell.address == address && cell.functions.is_empty()),
            "{address} is reported as unevaluated without a function name"
        );
    }
}

#[test]
fn editing_a_reader_after_cached_open_uses_the_iterative_members_value() {
    let mut workbook = Workbook::open(&package(ITERATE_AND_RECALCULATE)).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 1),
            "=A1*3",
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(workbook.calculation_status().unevaluated_cells.is_empty());
    assert_saved_values(
        &workbook.save().unwrap(),
        &[
            ("A1", "10"),
            ("B1", "30"),
            ("D1", "6"),
            ("E1", "20"),
            ("F1", "20"),
        ],
        "after editing B1",
    );
}

#[test]
fn a_non_iterative_cycle_still_settles_at_zero() {
    let workbook = open_recalculated(&package(r#"<calcPr fullCalcOnLoad="1"/>"#));
    assert_saved_values(
        &workbook.save().unwrap(),
        &[
            ("A1", "0"),
            ("B1", "0"),
            ("D1", "6"),
            ("E1", "0"),
            ("F1", "0"),
        ],
        "ordinary circular references",
    );
    let circular: Vec<_> = workbook
        .last_calculation()
        .cycle_cells
        .iter()
        .map(|address| address.cell.to_a1())
        .collect();
    assert_eq!(
        circular,
        ["A1", "E1", "F1"],
        "only cycle members are circular"
    );
}

#[test]
fn snapshot_and_recovery_restore_iterative_calculation() {
    let workbook = open_recalculated(&package(ITERATE_AND_RECALCULATE));
    for (mode, mut restored) in restore_sessions(&workbook) {
        assert!(restored.model().iterate, "{mode} restores iterate");
        restored
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "=(A1+20)/2",
                CalculationOptions::default(),
            )
            .unwrap();
        let status = restored.calculation_status();
        assert!(
            status
                .unevaluated_cells
                .iter()
                .any(|cell| cell.address == "A1" && cell.functions.is_empty()),
            "{mode} reports the edited cycle member as unevaluated"
        );
        assert!(
            status
                .unevaluated_cells
                .iter()
                .all(|cell| matches!(cell.address.as_str(), "A1" | "E1" | "F1")),
            "{mode} lists only cycle members as unevaluated"
        );
        assert_saved_values(
            &restored.save().unwrap(),
            &[("B1", "0"), ("D1", "6"), ("E1", "20"), ("F1", "20")],
            mode,
        );
    }
}

#[test]
fn a_reader_recalculates_from_the_held_value_in_live_and_restored_sessions() {
    let sheet = concat!(
        r#"<row r="1"><c r="A1"><f>(A1+10)/2</f><v>10</v></c>"#,
        r#"<c r="B1"><f>A1*2+C1</f><v>25</v></c><c r="C1"><v>5</v></c></row>"#,
    );
    let source = package_with_sheet(ITERATE_AND_RECALCULATE, sheet);
    let workbook = open_recalculated(&source);
    let restored = restore_sessions(&workbook);
    for (mode, mut session) in [("live", workbook)].into_iter().chain(restored) {
        assert!(session.model().iterate, "{mode} uses iterative calculation");
        session
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 2),
                "7",
                CalculationOptions::default(),
            )
            .unwrap();
        assert_eq!(
            session.model().value(SheetId(0), CellRef::new(0, 1)),
            CellValue::Number { value: 27.0 },
            "{mode} recalculates B1 from A1 and C1"
        );
        assert!(
            session
                .calculation_status()
                .unevaluated_cells
                .iter()
                .all(|cell| cell.address == "A1"),
            "{mode} reports only the cycle member as unevaluated"
        );
        assert_saved_values(
            &session.save().unwrap(),
            &[("A1", "10"), ("B1", "27"), ("C1", "7")],
            mode,
        );
    }
}

#[test]
fn a_generated_workbook_saves_iterative_calculation() {
    let model = betteroffice_xlsx::WorkbookModel {
        iterate: true,
        sheets: vec![betteroffice_xlsx::Sheet::new("Sheet1")],
        ..Default::default()
    };
    let workbook = Workbook::from_model(model).unwrap();
    let restored = restore_sessions(&workbook);
    for (mode, session) in [("generated", workbook)].into_iter().chain(restored) {
        assert!(session.model().iterate, "{mode} uses iterative calculation");
        let saved = session.save().unwrap();
        let parts = ooxml_opc::unzip_parts(&saved).unwrap();
        let (_, workbook_xml) = parts
            .iter()
            .find(|(name, _)| name == "xl/workbook.xml")
            .unwrap();
        assert!(
            std::str::from_utf8(workbook_xml)
                .unwrap()
                .contains(r#"iterate="1""#),
            "{mode} saves iterative calculation in calcPr"
        );
        assert!(Workbook::open(&saved).unwrap().model().iterate);
    }
}
