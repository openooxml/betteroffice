use betteroffice_xlsx::{CalculationOptions, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

fn book(cells: &[(&str, &str, Option<&str>)], calc_pr: &str) -> Vec<u8> {
    let row: String = cells
        .iter()
        .map(|(address, formula, cache)| {
            let cache = cache
                .map(|value| format!("<v>{value}</v>"))
                .unwrap_or_default();
            format!(r#"<c r="{address}"><f>{formula}</f>{cache}</c>"#)
        })
        .collect();
    package(&format!(r#"<row r="1">{row}</row>"#), calc_pr)
}

fn package(sheet_data: &str, calc_pr: &str) -> Vec<u8> {
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#
                .as_bytes()
                .to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(r#"<Relationships xmlns="{PACKAGE_RELS}"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).into_bytes(),
        ),
        (
            "xl/workbook.xml".to_owned(),
            format!(r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>{calc_pr}</workbook>"#).into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            format!(r#"<Relationships xmlns="{PACKAGE_RELS}"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).into_bytes(),
        ),
        (
            "xl/worksheets/sheet1.xml".to_owned(),
            format!(r#"<worksheet xmlns="{MAIN}"><sheetData>{sheet_data}</sheetData></worksheet>"#).into_bytes(),
        ),
    ];
    let parts: Vec<(String, Vec<u8>)> = parts
        .into_iter()
        .map(|(name, xml)| (name, [DECLARATION.as_bytes(), &xml].concat()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// the `<v>` each cell of the saved first sheet carries, in document order.
fn saved_caches(bytes: &[u8]) -> Vec<Option<String>> {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let sheet = parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .map(|(_, bytes)| String::from_utf8(bytes.clone()).unwrap())
        .unwrap();
    sheet
        .split("<c ")
        .skip(1)
        .map(|cell| {
            let cell = cell.split("</c>").next().unwrap_or(cell);
            cell.split_once("<v>")
                .and_then(|(_, rest)| rest.split_once("</v>"))
                .map(|(value, _)| value.to_owned())
        })
        .collect()
}

fn options() -> CalculationOptions {
    CalculationOptions {
        now_serial: Some(46_327.5),
    }
}

const STALE: &[(&str, &str, Option<&str>)] = &[
    ("A1", "SUM(2,3)", Some("999")),
    ("B1", "REPRO_UNKNOWN_FUNCTION(1)", Some("999")),
    ("C1", "1+1", None),
];

#[test]
fn a_file_asking_for_a_full_calculation_saves_the_values_calculated_at_open() {
    for calc_pr in [
        r#"<calcPr fullCalcOnLoad="1" forceFullCalc="1"/>"#,
        r#"<calcPr fullCalcOnLoad="1"/>"#,
        r#"<calcPr forceFullCalc="true"/>"#,
    ] {
        let source = book(STALE, calc_pr);
        for workbook in [
            Workbook::open_recalculated(&source, options()).unwrap(),
            Workbook::open_collaborative_recalculated(&source, 7, options()).unwrap(),
        ] {
            assert_eq!(
                saved_caches(&workbook.save().unwrap()),
                vec![Some("5".into()), Some("999".into()), Some("2".into())],
                "{calc_pr}"
            );
        }
    }
}

#[test]
fn caches_stay_byte_for_byte_without_a_recalculation_request() {
    for calc_pr in [
        "",
        r#"<calcPr fullCalcOnLoad="0"/>"#,
        r#"<calcPr calcId="191029"/>"#,
    ] {
        let source = book(STALE, calc_pr);
        let workbook = Workbook::open_recalculated(&source, options()).unwrap();
        assert_eq!(workbook.save().unwrap(), source, "{calc_pr}");
    }
    let flagged = book(STALE, r#"<calcPr fullCalcOnLoad="1"/>"#);
    assert_eq!(Workbook::open(&flagged).unwrap().save().unwrap(), flagged);
}

#[test]
fn a_full_calculation_request_without_formulas_keeps_the_file() {
    let formula_free = book(&[], r#"<calcPr fullCalcOnLoad="1"/>"#);
    let workbook = Workbook::open_recalculated(&formula_free, options()).unwrap();
    assert_eq!(workbook.save().unwrap(), formula_free);
}

#[test]
fn an_explicit_recalculation_saves_its_values_and_keeps_unevaluable_caches() {
    let source = book(STALE, "");
    let mut workbook = Workbook::open(&source).unwrap();
    workbook.recalculate_all(options());
    assert_eq!(
        saved_caches(&workbook.save().unwrap()),
        vec![Some("5".into()), Some("999".into()), Some("2".into())]
    );
}

#[test]
fn recalculated_values_survive_a_recovery_checkpoint() {
    let source = book(STALE, r#"<calcPr fullCalcOnLoad="1"/>"#);
    let workbook = Workbook::open_recalculated(&source, options()).unwrap();
    let checkpoint = workbook.recovery_checkpoint(0, None).unwrap();
    let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(
        saved_caches(&restored.save().unwrap()),
        vec![Some("5".into()), Some("999".into()), Some("2".into())]
    );
}

#[test]
fn a_flagged_array_formula_saves_its_calculated_range_and_keeps_every_other_part() {
    let source = package(
        r#"<row r="2"><c r="C2"><f t="array" ref="C2:C3">1</f><v>1</v></c></row><row r="3"><c r="C3"><v>100</v></c></row>"#,
        r#"<calcPr fullCalcOnLoad="1"/>"#,
    );
    let saved = Workbook::open_recalculated(&source, options())
        .unwrap()
        .save()
        .unwrap();
    assert_eq!(
        saved_caches(&saved),
        vec![Some("1".into()), Some("1".into())]
    );
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    assert_eq!(
        before.iter().map(|(name, _)| name).collect::<Vec<_>>(),
        after.iter().map(|(name, _)| name).collect::<Vec<_>>()
    );
    let rewritten: Vec<&str> = before
        .iter()
        .zip(&after)
        .filter(|((_, before), (_, after))| before != after)
        .map(|((name, _), _)| name.as_str())
        .collect();
    assert_eq!(rewritten, vec!["xl/worksheets/sheet1.xml"]);
}
