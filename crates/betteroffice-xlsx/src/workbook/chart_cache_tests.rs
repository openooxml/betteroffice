use super::*;
use crate::workbook::edit_tests::{options, r};

const CHART_PART: &str = "xl/charts/chart1.xml";

const CHART: &str = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:val><c:numRef><c:f>Sheet1!$B$5:$B$7</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1" formatCode="0.00"><c:v>20</c:v></c:pt><c:pt idx="2"><c:formatCode>0.0</c:formatCode><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:ser><c:idx val="1"/><c:val><c:numRef><c:f>Data!$C$2:$C$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:formatCode>0.0</c:formatCode><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

/// Sheet1 anchors a chart whose first series plots its own B5:B7 and whose
/// second plots Data!C2:C3, both through caches whose points carry per-point
/// format codes, content the cache regenerator does not model.
fn charted_workbook() -> Workbook {
    charted_workbook_with_chart(CHART)
}

fn charted_workbook_with_chart(chart: &str) -> Workbook {
    let parts: Vec<(String, Vec<u8>)> = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>North</t></is></c><c r="B5"><v>10</v></c></row><row r="6"><c r="A6" t="inlineStr"><is><t>North</t></is></c><c r="B6"><v>20</v></c></row><row r="7"><c r="A7" t="inlineStr"><is><t>South</t></is></c><c r="B7"><v>30</v></c></row></sheetData><drawing r:id="rIdDrawing"/></worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="C2"><v>5</v></c></row><row r="3"><c r="C3"><v>6</v></c></row></sheetData></worksheet>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
        ),
        (
            "xl/drawings/drawing1.xml",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#,
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
        ),
        (CHART_PART, chart),
    ]
    .into_iter()
    .map(|(path, xml)| (path.to_owned(), xml.as_bytes().to_vec()))
    .collect();
    Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap()
}

fn multi_level_chart() -> String {
    CHART.replace(
        r#"<c:idx val="0"/>"#,
        r#"<c:idx val="0"/><c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$5:$B$7</c:f><c:multiLvlStrCache><c:ptCount val="3"/><c:lvl><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:lvl><c:lvl><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="2"><c:v>South</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>"#,
    )
}

fn apply(workbook: &mut Workbook, op: Op) {
    workbook.apply_ops(vec![op], options()).unwrap();
}

fn saved_chart(workbook: &Workbook) -> String {
    let parts = ooxml_opc::unzip_parts(&workbook.save().unwrap()).unwrap();
    let (_, bytes) = parts
        .into_iter()
        .find(|(path, _)| path == CHART_PART)
        .unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn shifting_a_multi_level_chart_cache_keeps_it_verbatim_and_moves_its_reference() {
    let chart = multi_level_chart();
    let mut workbook = charted_workbook_with_chart(&chart);
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );

    assert_eq!(
        saved_chart(&workbook),
        chart
            .replace("Sheet1!$A$5:$B$7", "Sheet1!$A$6:$B$8")
            .replace("Sheet1!$B$5:$B$7", "Sheet1!$B$6:$B$8")
    );
}

const REFUSAL: &str =
    "xl/charts/chart1.xml references cells this edit would move, and it cannot be rewritten";

const MODELLED_CHART: &str = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:val><c:numRef><c:f>Sheet1!$B$5:$B$7</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

/// a chart whose only obstacle to a rebuilt cache is its multi-level category
/// cache: the values beside it are fully modelled.
fn modelled_multi_level_chart() -> String {
    MODELLED_CHART.replace(
        r#"<c:idx val="0"/>"#,
        r#"<c:idx val="0"/><c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$5:$B$7</c:f><c:multiLvlStrCache><c:ptCount val="3"/><c:lvl><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:lvl><c:lvl><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="2"><c:v>South</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>"#,
    )
}

fn assert_refused_up_front(workbook: &mut Workbook, op: Op, label: &str) {
    let saved = workbook.save().unwrap();
    let version = workbook.version();
    let error = workbook.apply_ops(vec![op], options()).unwrap_err();
    assert_eq!(error.to_string(), REFUSAL, "{label}");
    assert_eq!(workbook.version(), version, "{label}");
    assert!(!workbook.can_undo(), "{label}");
    assert_eq!(workbook.save().unwrap(), saved, "{label}");
}

#[test]
fn reshaping_a_multi_level_chart_cache_is_refused_before_it_applies() {
    for (label, op) in [
        (
            "rows inserted inside",
            Op::InsertRows {
                sheet: SheetId(0),
                at: 5,
                count: 1,
            },
        ),
        (
            "rows deleted inside",
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 5,
                count: 1,
            },
        ),
        (
            "columns inserted inside",
            Op::InsertCols {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
        ),
        (
            "columns deleted inside",
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
        ),
    ] {
        let mut workbook = charted_workbook_with_chart(&modelled_multi_level_chart());
        assert_refused_up_front(&mut workbook, op, label);
    }
}

#[test]
fn reshaping_a_batch_around_a_multi_level_chart_cache_is_refused_whole() {
    let mut workbook = charted_workbook_with_chart(&modelled_multi_level_chart());
    let saved = workbook.save().unwrap();
    let error = workbook
        .apply_ops(
            vec![
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 1,
                    count: 1,
                },
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 7,
                    count: 1,
                },
            ],
            options(),
        )
        .unwrap_err();
    assert_eq!(error.to_string(), REFUSAL);
    assert_eq!(workbook.save().unwrap(), saved);
}

#[test]
fn edits_beside_a_multi_level_chart_cache_still_apply_and_save() {
    let chart = modelled_multi_level_chart();
    for (label, op, from, to) in [
        (
            "rows inserted below",
            Op::InsertRows {
                sheet: SheetId(0),
                at: 7,
                count: 2,
            },
            None,
            None,
        ),
        (
            "rows deleted below",
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 8,
                count: 1,
            },
            None,
            None,
        ),
        (
            "columns inserted after",
            Op::InsertCols {
                sheet: SheetId(0),
                at: 2,
                count: 1,
            },
            None,
            None,
        ),
        (
            "rows deleted above",
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
            Some(["Sheet1!$A$5:$B$7", "Sheet1!$B$5:$B$7"]),
            Some(["Sheet1!$A$3:$B$5", "Sheet1!$B$3:$B$5"]),
        ),
        (
            "columns inserted before",
            Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            Some(["Sheet1!$A$5:$B$7", "Sheet1!$B$5:$B$7"]),
            Some(["Sheet1!$B$5:$C$7", "Sheet1!$C$5:$C$7"]),
        ),
    ] {
        let mut workbook = charted_workbook_with_chart(&chart);
        apply(&mut workbook, op);
        let saved = saved_chart(&workbook);
        let cache = &chart[chart.find("<c:multiLvlStrCache>").unwrap()
            ..chart.find("</c:multiLvlStrCache>").unwrap()];
        assert!(saved.contains(cache), "{label}: {saved}");
        match (from, to) {
            (Some(from), Some(to)) => {
                for (from, to) in from.into_iter().zip(to) {
                    assert!(
                        !saved.contains(&format!("<c:f>{from}</c:f>"))
                            && saved.contains(&format!("<c:f>{to}</c:f>")),
                        "{label}: {saved}"
                    );
                }
            }
            _ => assert_eq!(saved, chart, "{label}"),
        }
    }
}

#[test]
fn changing_either_column_under_a_moved_multi_level_chart_cache_keeps_it_and_saves() {
    for address in ["A7", "B7"] {
        for change_format in [false, true] {
            let mut workbook = charted_workbook_with_chart(&multi_level_chart());
            apply(
                &mut workbook,
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 1,
                    count: 1,
                },
            );
            if change_format {
                workbook
                    .set_range_number_format(
                        SheetId(0),
                        CellRange::new(r(address), r(address)),
                        NumberFormatMutation::Percent,
                        options(),
                    )
                    .unwrap();
            } else {
                workbook
                    .edit_cell(SheetId(0), r(address), "changed", options())
                    .unwrap();
            }

            let chart = multi_level_chart();
            let saved = saved_chart(&workbook);
            assert_eq!(
                saved,
                chart
                    .replace("Sheet1!$A$5:$B$7", "Sheet1!$A$6:$B$8")
                    .replace("Sheet1!$B$5:$B$7", "Sheet1!$B$6:$B$8"),
                "{address}, change_format={change_format}"
            );
            Workbook::open(&workbook.save().unwrap()).unwrap();
        }
    }
}

#[test]
fn edit_beside_an_unmodelled_chart_cache_saves_it_verbatim() {
    let mut workbook = charted_workbook();
    workbook
        .edit_cell(SheetId(0), r("A1"), "2", options())
        .unwrap();

    assert_eq!(saved_chart(&workbook), CHART);
}

#[test]
fn shifting_cells_under_unmodelled_chart_caches_moves_references_and_keeps_caches() {
    for (label, op, from, to) in [
        (
            "rows inserted above",
            Op::InsertRows {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
            "Sheet1!$B$5:$B$7",
            "Sheet1!$B$6:$B$8",
        ),
        (
            "rows deleted above",
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
            "Sheet1!$B$5:$B$7",
            "Sheet1!$B$3:$B$5",
        ),
        (
            "columns inserted before",
            Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            "Sheet1!$B$5:$B$7",
            "Sheet1!$C$5:$C$7",
        ),
        (
            "columns deleted before",
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            "Sheet1!$B$5:$B$7",
            "Sheet1!$A$5:$A$7",
        ),
        (
            "rows inserted above on another sheet",
            Op::InsertRows {
                sheet: SheetId(1),
                at: 0,
                count: 3,
            },
            "Data!$C$2:$C$3",
            "Data!$C$5:$C$6",
        ),
    ] {
        let mut workbook = charted_workbook();
        apply(&mut workbook, op);
        workbook
            .edit_cell(SheetId(0), r("A1"), "2", options())
            .unwrap();

        assert_eq!(saved_chart(&workbook), CHART.replace(from, to), "{label}");
    }
}

#[test]
fn reshaping_an_unmodelled_chart_cache_is_refused_before_it_applies() {
    for (label, op) in [
        (
            "rows inserted inside",
            Op::InsertRows {
                sheet: SheetId(0),
                at: 5,
                count: 1,
            },
        ),
        (
            "rows deleted inside",
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 5,
                count: 1,
            },
        ),
        (
            "rows deleted inside on another sheet",
            Op::DeleteRows {
                sheet: SheetId(1),
                at: 2,
                count: 1,
            },
        ),
    ] {
        let mut workbook = charted_workbook();
        assert_refused_up_front(&mut workbook, op, label);
    }
}

#[test]
fn reshaping_a_modelled_chart_cache_regenerates_it() {
    let mut workbook = charted_workbook_with_chart(MODELLED_CHART);
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 5,
            count: 1,
        },
    );
    let saved = saved_chart(&workbook);
    assert!(
        saved.contains("<c:f>Sheet1!$B$5:$B$8</c:f>") && saved.contains(r#"<c:ptCount val="4"/>"#),
        "{saved}"
    );
}

#[test]
fn edited_chart_previews_read_cleared_cells_the_way_saves_write_them() {
    let chart = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:val><c:numRef><c:f>Sheet1!$B$5:$B$7</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
    let mut workbook = charted_workbook_with_chart(chart);
    workbook
        .edit_charts(
            vec![crate::ChartMutation {
                sheet: SheetId(0),
                chart: 0,
                edit: crate::ChartEdit::SetSeries {
                    series: 0,
                    name: None,
                    categories: None,
                    values: Some("Data!$C$2:$C$3".into()),
                },
            }],
            options(),
        )
        .unwrap();
    let points = |values: &[f64]| {
        values
            .iter()
            .map(|value| value.is_finite().then_some(*value))
            .collect::<Vec<_>>()
    };
    let agree = |workbook: &Workbook, expected: &[Option<f64>]| {
        let chart = workbook.model.sheets[0].charts[0].clone();
        let preview = workbook.resolve_chart_space("Sheet1", &chart).unwrap();
        let saved = saved_chart(workbook);
        let reopened =
            xlsx_parse::chart_space(saved.as_bytes(), &workbook.model.styles.theme).unwrap();
        assert_eq!(points(&preview.series[0].values), expected, "{saved}");
        assert_eq!(points(&reopened.series[0].values), expected, "{saved}");
    };
    agree(&workbook, &[Some(5.0), Some(6.0)]);
    workbook
        .edit_cell(SheetId(1), r("C2"), "", options())
        .unwrap();
    agree(&workbook, &[None, Some(6.0)]);
    workbook
        .edit_cell(SheetId(1), r("C3"), "", options())
        .unwrap();
    agree(&workbook, &[]);
}

#[test]
fn a_control_character_under_a_reshaped_chart_cache_saves_as_an_escape() {
    let chart = MODELLED_CHART.replace(
        r#"<c:idx val="0"/>"#,
        r#"<c:idx val="0"/><c:cat><c:strRef><c:f>Sheet1!$A$5:$A$7</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>North</c:v></c:pt><c:pt idx="2"><c:v>South</c:v></c:pt></c:strCache></c:strRef></c:cat>"#,
    );
    let mut workbook = charted_workbook_with_chart(&chart);
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 5,
            count: 1,
        },
    );
    workbook
        .edit_cell(SheetId(0), r("A5"), "=CHAR(1)", options())
        .unwrap();

    let saved = saved_chart(&workbook);
    assert!(
        saved.contains("<c:f>Sheet1!$A$5:$A$8</c:f>")
            && saved.contains(r#"<c:pt idx="0"><c:v>_x0001_</c:v></c:pt>"#)
            && !saved.contains('\u{1}'),
        "{saved}"
    );
    Workbook::open(&workbook.save().unwrap()).unwrap();
}

#[test]
fn only_batches_that_change_charts_run_the_chart_check() {
    let mut workbook = charted_workbook_with_chart(MODELLED_CHART);
    let before = super::charts::chart_check_count();
    workbook
        .edit_cell(SheetId(0), r("A1"), "9", options())
        .unwrap();
    apply(
        &mut workbook,
        Op::AddSheet {
            index: 2,
            name: "Plain".into(),
        },
    );
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(2),
            at: 0,
            count: 1,
        },
    );
    assert_eq!(super::charts::chart_check_count(), before);
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
    assert_eq!(super::charts::chart_check_count(), before + 1);
}

#[test]
fn finding_changed_chart_parts_is_linear_in_the_charts() {
    let template = charted_workbook().model.sheets[0].charts[0].clone();
    let mut before = xlsx_model::Workbook::default();
    let mut sheet = xlsx_model::Sheet::new("Many");
    sheet.charts = (0..60_000)
        .map(|index| xlsx_model::SheetChart {
            part: format!("xl/charts/chart{index}.xml"),
            anchor_index: index,
            ..template.clone()
        })
        .collect();
    before.sheets.push(sheet);
    let mut after = before.clone();
    after.sheets.push(xlsx_model::Sheet::new("Added"));
    assert!(super::charts::changed_chart_parts(&before, &after).is_empty());
    after.sheets[0].charts[59_999].refs[0].formula = "Sheet1!$B$6:$B$8".into();
    assert_eq!(
        super::charts::changed_chart_parts(&before, &after)
            .into_iter()
            .collect::<Vec<_>>(),
        vec!["xl/charts/chart59999.xml"]
    );
}
