use betteroffice_xlsx::{
    CalculationOptions, ChartEdit, ChartMutation, ChartSeriesName, Error, Op, SheetId,
    SnapshotBudget, Workbook, WorkbookSnapshotBuilder, WorkbookSnapshotEncoder,
};

const CHART: &str = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val='10'/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1" sz="1600"/><a:t>Old </a:t></a:r><a:r><a:rPr i="1"/><a:t>title</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>One</c:v></c:pt><c:pt idx="1"><c:v>Two</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f><c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:ser><c:idx val='1'/><c:order val='1'/><c:tx><c:v>Keep</c:v></c:tx></c:ser></c:barChart><c:catAx><c:axId val='1'/></c:catAx></c:plotArea><c:legend><c:legendPos val='b'/></c:legend></c:chart></c:chartSpace>"#;

fn bytes(chart: &str) -> Vec<u8> {
    let parts = [
        (
            "xl/workbook.xml",
            r#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>New name</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>One</t></is></c><c r="B2"><v>10</v></c><c r="C2"><v>30</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Two</t></is></c><c r="B3"><v>20</v></c><c r="C3"><v>40</v></c></row></sheetData><drawing r:id="d1"/></worksheet>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="d1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
        ),
        (
            "xl/drawings/drawing1.xml",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><c:chart r:id="c1"/></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>"#,
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="c1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
        ),
        ("xl/charts/chart1.xml", chart),
        ("xl/charts/untouched.xml", "<untouched a='b' />"),
    ];
    let mut parts = parts
        .into_iter()
        .map(|(p, x)| (p.to_owned(), x.as_bytes().to_vec()))
        .collect::<Vec<_>>();
    let drawing = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/drawings/drawing1.xml")
        .unwrap();
    let xml = String::from_utf8(drawing.1.clone()).unwrap();
    let start = xml.find("<xdr:twoCellAnchor>").unwrap();
    let end = xml.find("</xdr:twoCellAnchor>").unwrap() + "</xdr:twoCellAnchor>".len();
    let second = xml[start..end].replace("r:id=\"c1\"", "r:id=\"c2\"");
    drawing.1 = xml
        .replace("</xdr:wsDr>", &format!("{second}</xdr:wsDr>"))
        .into_bytes();
    let rels = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/drawings/_rels/drawing1.xml.rels")
        .unwrap();
    rels.1 = String::from_utf8(rels.1.clone()).unwrap().replace("</Relationships>", r#"<Relationship Id="c2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>"#).into_bytes();
    parts.push(("xl/charts/chart2.xml".into(), CHART.as_bytes().to_vec()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn mutation(edit: ChartEdit) -> ChartMutation {
    ChartMutation {
        sheet: SheetId(0),
        chart: 0,
        edit,
    }
}
fn apply(book: &mut Workbook, edit: ChartEdit) {
    book.edit_charts(vec![mutation(edit)], CalculationOptions::default())
        .unwrap();
}
fn chart_xml(book: &Workbook) -> String {
    let parts = ooxml_opc::unzip_parts(&book.save().unwrap()).unwrap();
    String::from_utf8(
        parts
            .into_iter()
            .find(|(p, _)| p == "xl/charts/chart1.xml")
            .unwrap()
            .1,
    )
    .unwrap()
}

#[test]
fn title_replace_add_delete_preserves_properties_and_inverse() {
    let source = bytes(CHART);
    let mut book = Workbook::open(&source).unwrap();
    let charts = book.charts(SheetId(0)).unwrap();
    assert_eq!(charts.len(), 2);
    assert_eq!(charts[0].title.as_deref(), Some("Old title"));
    assert!(charts[0].rich_title);
    assert_eq!(charts[0].anchor_range.as_deref(), Some("D2:J13"));
    apply(
        &mut book,
        ChartEdit::SetChartTitle {
            title: Some("New & title".into()),
        },
    );
    let xml = chart_xml(&book);
    assert_eq!(xml.matches("<a:r>").count(), 1);
    assert!(xml.contains("<a:rPr b=\"1\" sz=\"1600\"/>"));
    assert!(xml.contains("New &amp; title"));
    assert!(xml.contains("<c:overlay val=\"0\"/>"));
    book.undo(CalculationOptions::default()).unwrap();
    assert_eq!(chart_xml(&book), CHART);
    book.redo(CalculationOptions::default()).unwrap();
    assert_eq!(chart_xml(&book), xml);
    apply(&mut book, ChartEdit::SetChartTitle { title: None });
    let deleted = chart_xml(&book);
    assert!(!deleted.contains("<c:title>"));
    assert!(deleted.contains("autoTitleDeleted val=\"1\""));
    apply(
        &mut book,
        ChartEdit::SetChartTitle {
            title: Some("Added".into()),
        },
    );
    let saved = book.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.charts(SheetId(0)).unwrap()[0].title.as_deref(),
        Some("Added")
    );
    assert!(!reopened.charts(SheetId(0)).unwrap()[0].rich_title);
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    for (path, bytes) in before {
        if path != "xl/charts/chart1.xml" {
            assert_eq!(
                after.iter().find(|(p, _)| p == &path).unwrap().1,
                bytes,
                "{path}"
            );
        }
    }
}

#[test]
fn series_names_references_caches_and_atomicity() {
    let mut book = Workbook::open(&bytes(CHART)).unwrap();
    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 0,
            name: Some(ChartSeriesName::Reference {
                reference: "Data!$B$1".into(),
            }),
            categories: None,
            values: Some("Data!$C$2:$C$3".into()),
        },
    );
    let xml = chart_xml(&book);
    assert!(xml.contains("<c:f>Data!$B$1</c:f>"));
    assert!(xml.contains("<c:v>New name</c:v>"));
    assert!(xml.contains("<c:v>30</c:v>"));
    assert!(xml.contains("<c:v>40</c:v>"));
    assert!(
        xml.contains(
            "<c:ser><c:idx val='1'/><c:order val='1'/><c:tx><c:v>Keep</c:v></c:tx></c:ser>"
        )
    );
    book.undo(CalculationOptions::default()).unwrap();
    assert_eq!(chart_xml(&book), CHART);
    book.redo(CalculationOptions::default()).unwrap();
    let info = book.charts(SheetId(0)).unwrap();
    assert_eq!(info[0].series[0].name_ref.as_deref(), Some("Data!$B$1"));
    let before = book.save().unwrap();
    assert!(
        serde_json::from_str::<ChartMutation>(
            r#"{"sheet":0,"chart":0,"action":"set_chart_title"}"#
        )
        .is_err()
    );
    assert!(
        book.edit_charts(
            vec![
                mutation(ChartEdit::SetChartTitle {
                    title: Some("Rollback".into())
                }),
                mutation(ChartEdit::SetSeries {
                    series: 99,
                    name: None,
                    categories: None,
                    values: Some("Data!A1".into())
                })
            ],
            CalculationOptions::default()
        )
        .is_err()
    );
    assert_eq!(book.save().unwrap(), before);
    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 0,
            name: Some(ChartSeriesName::Literal {
                text: "Literal <name>".into(),
            }),
            categories: Some("Data!$A:$A".into()),
            values: None,
        },
    );
    let xml = chart_xml(&book);
    assert!(xml.contains("<c:tx><c:v>Literal &lt;name&gt;</c:v></c:tx>"));
    assert!(xml.contains("<c:cat><c:strRef><c:f>Data!$A:$A</c:f></c:strRef></c:cat>"));
    assert!(!xml.contains("strCache"));
    let reopened = Workbook::open(&book.save().unwrap()).unwrap();
    assert_eq!(
        reopened.charts(SheetId(0)).unwrap()[0].series[0]
            .name
            .as_deref(),
        Some("Literal <name>")
    );
}

#[test]
fn chart_edits_refuse_collaboration() {
    let mut book = Workbook::open_collaborative(&bytes(CHART), 42).unwrap();
    assert!(matches!(
        book.edit_charts(
            vec![mutation(ChartEdit::SetChartTitle { title: None })],
            CalculationOptions::default()
        ),
        Err(Error::CollaborativeStructureOperation)
    ));
}

#[test]
fn cache_updates_after_name_reference_addition_and_cell_edits() {
    let mut book = Workbook::open(&bytes(CHART)).unwrap();
    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 0,
            name: Some(ChartSeriesName::Reference {
                reference: "Data!$B$1".into(),
            }),
            categories: None,
            values: Some("Data!$C$2:$C$3".into()),
        },
    );
    book.edit_cell(
        SheetId(0),
        betteroffice_xlsx::CellRef::parse_a1("C2").unwrap(),
        "75",
        CalculationOptions::default(),
    )
    .unwrap();
    assert!(chart_xml(&book).contains("<c:v>75</c:v>"));
}

#[test]
fn reference_extensions_and_nonstandard_prefixes_survive() {
    let chart = CHART
        .replace("<c:numRef>", "<c:numRef xmlns:keep='urn:keep' flag='yes'>")
        .replace("</c:numRef>", "<keep:meta value='same'/></c:numRef>")
        .replace("xmlns:c=", "xmlns:q=")
        .replace("c:", "q:");
    let mut book = Workbook::open(&bytes(&chart)).unwrap();
    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 0,
            name: None,
            categories: None,
            values: Some("Data!$C$2:$C$3".into()),
        },
    );
    let saved = chart_xml(&book);
    assert!(saved.contains("<keep:meta value='same'/>"));
    assert!(saved.contains("<q:numRef xmlns:keep='urn:keep' flag='yes'>"));
    assert!(saved.contains("<q:v>30</q:v>"));
}

#[test]
fn utf16_chart_edits_preserve_source_encoding() {
    let mut parts = ooxml_opc::unzip_parts(&bytes(CHART)).unwrap();
    let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-16\"?>{CHART}");
    let mut encoded = vec![0xff, 0xfe];
    encoded.extend(xml.encode_utf16().flat_map(u16::to_le_bytes));
    parts
        .iter_mut()
        .find(|(path, _)| path == "xl/charts/chart1.xml")
        .unwrap()
        .1 = encoded;
    let mut book = Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    apply(
        &mut book,
        ChartEdit::SetChartTitle {
            title: Some("Encoded title".into()),
        },
    );
    let saved = book.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    assert!(
        parts
            .iter()
            .find(|(path, _)| path == "xl/charts/chart1.xml")
            .unwrap()
            .1
            .starts_with(&[0xff, 0xfe])
    );
    assert_eq!(
        Workbook::open(&saved).unwrap().charts(SheetId(0)).unwrap()[0]
            .title
            .as_deref(),
        Some("Encoded title")
    );
}

#[test]
fn locally_shadowed_namespaces_survive_title_and_literal_range_edits() {
    let mut chart = CHART.to_owned();
    let start = chart.find("<c:title>").unwrap();
    let end = chart.find("</c:title>").unwrap() + "</c:title>".len();
    chart.replace_range(start..end, r#"<q:title xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:c="urn:foreign"/>"#);
    let start = chart.find("<c:val>").unwrap();
    let end = chart.find("</c:val>").unwrap() + "</c:val>".len();
    chart.replace_range(start..end, r#"<q:val xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:c="urn:foreign"><q:numLit><q:ptCount val="1"/><q:pt idx="0"><q:v>1</q:v></q:pt></q:numLit></q:val>"#);
    let mut book = Workbook::open(&bytes(&chart)).unwrap();
    apply(
        &mut book,
        ChartEdit::SetChartTitle {
            title: Some("Namespaced title".into()),
        },
    );
    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 0,
            name: None,
            categories: None,
            values: Some("Data!$C$2:$C$3".into()),
        },
    );
    let saved = book.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    let info = reopened.charts(SheetId(0)).unwrap();
    assert_eq!(info[0].title.as_deref(), Some("Namespaced title"));
    assert_eq!(
        info[0].series[0].value_ref.as_deref(),
        Some("Data!$C$2:$C$3")
    );
    assert!(chart_xml(&book).contains("<q:v>30</q:v>"));
}

#[test]
fn shared_chart_parts_refuse_edits_without_changing_any_anchor() {
    let mut parts = ooxml_opc::unzip_parts(&bytes(CHART)).unwrap();
    let rels = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/drawings/_rels/drawing1.xml.rels")
        .unwrap();
    rels.1 = String::from_utf8(rels.1.clone())
        .unwrap()
        .replace("../charts/chart2.xml", "../charts/chart1.xml")
        .into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut book = Workbook::open(&source).unwrap();
    let version = book.version();
    let error = book
        .edit_charts(
            vec![mutation(ChartEdit::SetChartTitle {
                title: Some("Shared title".into()),
            })],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(error.to_string().contains("shared"));
    assert_eq!(book.version(), version);
    assert_eq!(book.save().unwrap(), source);
}

fn chartsheet_bytes(protected: bool) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&bytes(CHART)).unwrap();
    let workbook = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/workbook.xml")
        .unwrap();
    workbook.1 = String::from_utf8(workbook.1.clone())
        .unwrap()
        .replace(
            "</sheets>",
            r#"<sheet name="Chart" sheetId="2" r:id="rChart"/></sheets>"#,
        )
        .into_bytes();
    let rels = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/_rels/workbook.xml.rels")
        .unwrap();
    rels.1 = String::from_utf8(rels.1.clone()).unwrap().replace("</Relationships>", r#"<Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>"#).into_bytes();
    let worksheet = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    worksheet.1 = String::from_utf8(worksheet.1.clone())
        .unwrap()
        .replace(r#"<drawing r:id="d1"/>"#, "")
        .into_bytes();
    let rels = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/_rels/sheet1.xml.rels")
        .unwrap();
    rels.0 = "xl/chartsheets/_rels/sheet1.xml.rels".into();
    let protection = if protected {
        r#"<sheetProtection content="1" objects="1"/>"#
    } else {
        ""
    };
    parts.push(("xl/chartsheets/sheet1.xml".into(), format!(r#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{protection}<drawing r:id="d1"/></chartsheet>"#).into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn chartsheet_title_and_series_edits_preserve_sheet_parts() {
    let source = chartsheet_bytes(false);
    let mut book = Workbook::open(&source).unwrap();
    assert!(book.charts(SheetId(0)).unwrap().is_empty());
    assert_eq!(book.charts(SheetId(1)).unwrap().len(), 2);
    book.edit_charts(
        vec![
            ChartMutation {
                sheet: SheetId(1),
                chart: 0,
                edit: ChartEdit::SetChartTitle {
                    title: Some("Chart view".into()),
                },
            },
            ChartMutation {
                sheet: SheetId(1),
                chart: 0,
                edit: ChartEdit::SetSeries {
                    series: 0,
                    name: Some(ChartSeriesName::Literal {
                        text: "Chart series".into(),
                    }),
                    categories: None,
                    values: Some("Data!$C$2:$C$3".into()),
                },
            },
        ],
        CalculationOptions::default(),
    )
    .unwrap();
    let saved = book.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    let info = reopened.charts(SheetId(1)).unwrap();
    assert_eq!(info[0].title.as_deref(), Some("Chart view"));
    assert_eq!(info[0].series[0].name.as_deref(), Some("Chart series"));
    assert_eq!(
        info[0].series[0].value_ref.as_deref(),
        Some("Data!$C$2:$C$3")
    );
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    for path in [
        "xl/worksheets/sheet1.xml",
        "xl/chartsheets/sheet1.xml",
        "xl/drawings/drawing1.xml",
        "xl/charts/chart2.xml",
    ] {
        assert_eq!(
            before.iter().find(|(part, _)| part == path).unwrap().1,
            after.iter().find(|(part, _)| part == path).unwrap().1,
            "{path}"
        );
    }
}

#[test]
fn chart_owner_protection_is_enforced_by_native_edits() {
    let mut worksheet = ooxml_opc::unzip_parts(&bytes(CHART)).unwrap();
    let sheet = worksheet
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    sheet.1 = String::from_utf8(sheet.1.clone())
        .unwrap()
        .replace(
            "</sheetData>",
            r#"</sheetData><sheetProtection sheet="1"/>"#,
        )
        .into_bytes();
    for (source, target) in [
        (chartsheet_bytes(true), SheetId(1)),
        (ooxml_opc::rezip_parts(&worksheet).unwrap(), SheetId(0)),
    ] {
        let mut book = Workbook::open(&source).unwrap();
        let version = book.version();
        let result = book.edit_charts(
            vec![ChartMutation {
                sheet: target,
                chart: 0,
                edit: ChartEdit::SetChartTitle {
                    title: Some("Protected title".into()),
                },
            }],
            CalculationOptions::default(),
        );
        assert!(result.is_err());
        assert_eq!(book.version(), version);
        assert_eq!(book.save().unwrap(), source);
    }
}

#[test]
fn missing_xy_references_use_scatter_and_bubble_slots() {
    for kind in ["scatterChart", "bubbleChart"] {
        let mut chart = CHART
            .replace(
                "<c:barChart><c:barDir val=\"col\"/>",
                &format!("<c:{kind}>"),
            )
            .replace("</c:barChart>", &format!("</c:{kind}>"))
            .replace("<c:val>", "<c:yVal>")
            .replace("</c:val>", "</c:yVal>");
        let start = chart.find("<c:cat>").unwrap();
        let end = chart.find("</c:cat>").unwrap() + "</c:cat>".len();
        chart.replace_range(start..end, "");
        let mut book = Workbook::open(&bytes(&chart)).unwrap();
        book.edit_charts(
            vec![
                mutation(ChartEdit::SetSeries {
                    series: 0,
                    name: None,
                    categories: Some("Data!$C$2:$C$3".into()),
                    values: None,
                }),
                mutation(ChartEdit::SetSeries {
                    series: 1,
                    name: None,
                    categories: None,
                    values: Some("Data!$B$2:$B$3".into()),
                }),
            ],
            CalculationOptions::default(),
        )
        .unwrap();
        let xml = chart_xml(&book);
        assert!(!xml.contains("<c:cat>"), "{kind}");
        assert!(!xml.contains("<c:val>"), "{kind}");
        assert_eq!(xml.matches("<c:xVal><c:numRef>").count(), 1, "{kind}");
        assert_eq!(xml.matches("<c:yVal>").count(), 2, "{kind}");
        let reopened = Workbook::open(&book.save().unwrap()).unwrap();
        let info = reopened.charts(SheetId(0)).unwrap();
        assert_eq!(
            info[0].series[0].category_ref.as_deref(),
            Some("Data!$C$2:$C$3")
        );
        assert_eq!(
            info[0].series[1].value_ref.as_deref(),
            Some("Data!$B$2:$B$3")
        );
        book.undo(CalculationOptions::default()).unwrap();
        assert_eq!(chart_xml(&book), chart);
    }
}

fn edit_parts(source: &[u8], edit: impl FnOnce(&mut Vec<(String, Vec<u8>)>)) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(source).unwrap();
    edit(&mut parts);
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn replace_part(parts: &mut [(String, Vec<u8>)], path: &str, from: &str, to: &str) {
    let part = parts.iter_mut().find(|(name, _)| name == path).unwrap();
    let xml = String::from_utf8(part.1.clone()).unwrap();
    assert!(xml.contains(from), "{path}: {from}");
    part.1 = xml.replace(from, to).into_bytes();
}

/// A second sheet whose name needs quoting, a workbook defined name and a date
/// formatted cell next to the base fixture.
fn reference_bytes(chart: &str) -> Vec<u8> {
    edit_parts(&bytes(chart), |parts| {
        replace_part(
            parts,
            "xl/workbook.xml",
            "</sheets>",
            r#"<sheet name="Q1 Data" sheetId="2" r:id="r2"/></sheets><definedNames><definedName name="Totals">Data!$C$2:$C$3</definedName><definedName name="Local" localSheetId="1">'Q1 Data'!$A$1:$A$2</definedName><definedName name="DataLocal" localSheetId="0">Data!$B$2:$B$3</definedName></definedNames>"#,
        );
        replace_part(
            parts,
            "xl/_rels/workbook.xml.rels",
            "</Relationships>",
            r#"<Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
        );
        replace_part(
            parts,
            "xl/worksheets/sheet1.xml",
            r#"<c r="C2"><v>30</v></c>"#,
            r#"<c r="C2"><v>30</v></c><c r="D2" s="1"><v>45292</v></c>"#,
        );
        parts.push(("xl/worksheets/sheet2.xml".into(), br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>5</v></c></row><row r="2"><c r="A2"><v>7</v></c></row></sheetData></worksheet>"#.to_vec()));
        parts.push(("xl/styles.xml".into(), br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#.to_vec()));
    })
}

fn values(reference: &str) -> ChartEdit {
    ChartEdit::SetSeries {
        series: 0,
        name: None,
        categories: None,
        values: Some(reference.into()),
    }
}

#[test]
fn edited_charts_survive_recovery_checkpoints_and_refuse_peer_snapshots() {
    for history in [false, true] {
        let mut book = Workbook::open(&bytes(CHART)).unwrap();
        apply(
            &mut book,
            ChartEdit::SetChartTitle {
                title: Some("Recovered".into()),
            },
        );
        apply(
            &mut book,
            ChartEdit::SetSeries {
                series: 1,
                name: Some(ChartSeriesName::Reference {
                    reference: "Data!$B$1".into(),
                }),
                categories: None,
                values: Some("Data!$C$2:$C$3".into()),
            },
        );
        let direct = book.save().unwrap();
        let checkpoint = book
            .recovery_checkpoint_with_history(7, None, history)
            .unwrap();
        let (mut restored, through, _, _) =
            Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
        assert_eq!(through, 7);
        assert_eq!(restored.save().unwrap(), direct, "history {history}");
        assert_eq!(
            restored.charts(SheetId(0)).unwrap(),
            book.charts(SheetId(0)).unwrap()
        );
        apply(
            &mut restored,
            ChartEdit::SetChartTitle {
                title: Some("After recovery".into()),
            },
        );
        assert!(chart_xml(&restored).contains("After recovery"));
        assert!(chart_xml(&restored).contains("<c:f>Data!$B$1</c:f>"));
        if history {
            for _ in 0..3 {
                restored.undo(CalculationOptions::default()).unwrap();
            }
            assert_eq!(chart_xml(&restored), CHART);
        }
        let budget = SnapshotBudget::new(8, 65536).unwrap();
        let refused = WorkbookSnapshotEncoder::new(&book, None, budget).and_then(|mut encoder| {
            while encoder.next(&book)?.is_some() {}
            Ok(())
        });
        assert!(refused.is_err());
        assert!(book.peer_hydration().is_err());
    }
}

#[test]
fn chart_references_must_name_one_area_on_an_existing_sheet_or_a_defined_name() {
    let source = reference_bytes(CHART);
    let mut book = Workbook::open(&source).unwrap();
    let version = book.version();
    for reference in [
        "foo bar",
        "1+2",
        "[1]Ext!A1",
        "Q1 Data!$A$1:$A$2",
        "$C$2:$C$3",
        "Missing!$A$1:$A$2",
        "Data!$C$2:$C$3,Data!$B$2",
        "Data!$C$2:$C$3:$D$4",
        "Data!Nope",
        "Unknown",
        "Local",
    ] {
        let error = book
            .edit_charts(
                vec![mutation(values(reference))],
                CalculationOptions::default(),
            )
            .unwrap_err()
            .to_string();
        assert!(error.contains("values"), "{reference}: {error}");
        assert!(error.contains("Sheet1!$B$2:$B$5"), "{reference}: {error}");
        assert_eq!(book.version(), version, "{reference}");
    }
    let error = book
        .edit_charts(
            vec![mutation(values("Q1 Data!$A$1:$A$2"))],
            CalculationOptions::default(),
        )
        .unwrap_err()
        .to_string();
    assert!(error.contains("'Q1 Data'!$A$1:$A$2"), "{error}");
    let error = book
        .edit_charts(
            vec![mutation(ChartEdit::SetSeries {
                series: 0,
                name: Some(ChartSeriesName::Reference {
                    reference: "Missing!$B$1".into(),
                }),
                categories: None,
                values: None,
            })],
            CalculationOptions::default(),
        )
        .unwrap_err()
        .to_string();
    assert!(error.contains("name.ref"), "{error}");
    assert_eq!(book.save().unwrap(), source);

    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 0,
            name: None,
            categories: Some("'Q1 Data'!$A$1:$A$2".into()),
            values: Some("=Totals".into()),
        },
    );
    let xml = chart_xml(&book);
    assert!(xml.contains("<c:f>'Q1 Data'!$A$1:$A$2</c:f>"), "{xml}");
    assert!(
        xml.contains("<c:v>5</c:v>") && xml.contains("<c:v>7</c:v>"),
        "{xml}"
    );
    assert!(xml.contains("<c:f>[0]!Totals</c:f>"), "{xml}");
    let info = book.charts(SheetId(0)).unwrap();
    assert_eq!(
        info[0].series[0].category_ref.as_deref(),
        Some("'Q1 Data'!$A$1:$A$2")
    );
    assert_eq!(info[0].series[0].value_ref.as_deref(), Some("[0]!Totals"));
    for reference in [
        "Data!$C:$C",
        "data!$2:$2",
        "'Q1 Data'!Local",
        "Data!$B$2",
        "[0]!Totals",
    ] {
        apply(&mut book, values(reference));
        assert!(
            chart_xml(&book).contains(&format!("<c:f>{reference}</c:f>")),
            "{reference}"
        );
    }
}

#[test]
fn caches_cover_blank_sparse_and_date_points() {
    let mut book = Workbook::open(&reference_bytes(CHART)).unwrap();
    apply(&mut book, values("Data!$C$1:$C$3"));
    let xml = chart_xml(&book);
    let start = xml.find("<c:val>").unwrap();
    let cache = &xml[start..xml[start..].find("</c:val>").unwrap() + start];
    assert!(
        cache.contains("<c:formatCode>0.00</c:formatCode>"),
        "{cache}"
    );
    assert!(cache.contains("<c:ptCount val=\"3\"/>"), "{cache}");
    assert!(!cache.contains("<c:pt idx=\"0\">"), "{cache}");
    assert!(
        cache.contains("<c:pt idx=\"1\"><c:v>30</c:v></c:pt>"),
        "{cache}"
    );
    assert!(
        cache.contains("<c:pt idx=\"2\"><c:v>40</c:v></c:pt>"),
        "{cache}"
    );
    apply(&mut book, values("Data!$D$2:$D$3"));
    let xml = chart_xml(&book);
    let start = xml.find("<c:val>").unwrap();
    let cache = &xml[start..xml[start..].find("</c:val>").unwrap() + start];
    assert!(cache.contains("<c:f>Data!$D$2:$D$3</c:f>"), "{cache}");
    assert!(
        cache
            .contains("<c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>45292</c:v></c:pt></c:numCache>"),
        "{cache}"
    );
}

#[test]
fn multi_level_categories_are_replaced_by_a_single_reference() {
    let chart = CHART.replace(
        r#"<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>One</c:v></c:pt><c:pt idx="1"><c:v>Two</c:v></c:pt></c:strCache></c:strRef></c:cat>"#,
        r#"<c:cat><c:multiLvlStrRef><c:f>Data!$A$2:$B$3</c:f><c:multiLvlStrCache><c:ptCount val="2"/><c:lvl><c:pt idx="0"><c:v>One</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>"#,
    );
    assert_ne!(chart, CHART);
    let mut book = Workbook::open(&bytes(&chart)).unwrap();
    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 0,
            name: None,
            categories: Some("Data!$A$2:$A$3".into()),
            values: None,
        },
    );
    let xml = chart_xml(&book);
    assert!(!xml.contains("multiLvl"), "{xml}");
    assert!(
        xml.contains("<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache>"),
        "{xml}"
    );
    assert_eq!(
        book.charts(SheetId(0)).unwrap()[0].series[0]
            .category_ref
            .as_deref(),
        Some("Data!$A$2:$A$3")
    );
}

#[test]
fn unreadable_charts_list_with_an_error_and_leave_others_editable() {
    let source = edit_parts(&bytes(CHART), |parts| {
        parts.iter_mut().find(|(path, _)| path == "xl/charts/chart2.xml").unwrap().1 =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:roundedCorners val="0"/></c:chartSpace>"#.to_vec();
    });
    let mut book = Workbook::open(&source).unwrap();
    let charts = book.charts(SheetId(0)).unwrap();
    assert_eq!(charts.len(), 2);
    let listed = serde_json::to_value(&charts).unwrap();
    assert_eq!(listed[0]["editable"], true);
    assert!(listed[0]["error"].is_null());
    assert_eq!(listed[0]["title"], "Old title");
    assert_eq!(listed[1]["editable"], false);
    assert!(
        listed[1]["error"]
            .as_str()
            .unwrap()
            .contains("no chart element"),
        "{listed}"
    );
    assert_eq!(listed[1]["id"], charts[1].id.as_str());
    assert_eq!(listed[1]["series"], serde_json::json!([]));
    let version = book.version();
    let error = book
        .edit_charts(
            vec![ChartMutation {
                sheet: SheetId(0),
                chart: 1,
                edit: ChartEdit::SetChartTitle {
                    title: Some("No".into()),
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(error.to_string().contains("no chart element"), "{error}");
    assert_eq!(book.version(), version);
    apply(
        &mut book,
        ChartEdit::SetChartTitle {
            title: Some("Still editable".into()),
        },
    );
    assert!(chart_xml(&book).contains("Still editable"));
}

#[test]
fn title_markup_matches_excel_output() {
    let chart = CHART.replace(
        r#"<a:t>title</a:t></a:r></a:p>"#,
        r#"<a:t>title</a:t></a:r><a:endParaRPr lang="en-US" b="1"/></a:p>"#,
    );
    assert_ne!(chart, CHART);
    let mut book = Workbook::open(&bytes(&chart)).unwrap();
    apply(
        &mut book,
        ChartEdit::SetChartTitle {
            title: Some("Replaced".into()),
        },
    );
    let xml = chart_xml(&book);
    assert!(!xml.contains("xml:space"), "{xml}");
    assert!(
        xml.contains(r#"<a:t>Replaced</a:t></a:r><a:endParaRPr lang="en-US" b="1"/></a:p>"#),
        "{xml}"
    );
    apply(&mut book, ChartEdit::SetChartTitle { title: None });
    apply(
        &mut book,
        ChartEdit::SetChartTitle {
            title: Some(" Added ".into()),
        },
    );
    let xml = chart_xml(&book);
    assert!(!xml.contains("xml:space"), "{xml}");
    assert!(xml.contains("<a:t> Added </a:t>"), "{xml}");
    assert!(xml.contains(r#"<c:overlay val="0"/></c:title>"#), "{xml}");
    let reopened = Workbook::open(&book.save().unwrap()).unwrap();
    assert_eq!(
        reopened.charts(SheetId(0)).unwrap()[0].title.as_deref(),
        Some(" Added ")
    );
}

#[test]
fn automatic_titles_read_back_as_automatic() {
    let start = CHART.find("<c:title>").unwrap();
    let end = CHART.find("</c:title>").unwrap() + "</c:title>".len();
    let mut chart = CHART.to_owned();
    chart.replace_range(start..end, r#"<c:title><c:overlay val="0"/></c:title>"#);
    let book = Workbook::open(&bytes(&chart)).unwrap();
    let listed = serde_json::to_value(book.charts(SheetId(0)).unwrap()).unwrap();
    assert!(listed[0]["title"].is_null(), "{listed}");
    assert_eq!(listed[0]["autoTitle"], true);
    assert_eq!(listed[1]["autoTitle"], false);
    assert_eq!(listed[1]["title"], "Old title");
}

#[test]
fn missing_values_precede_bar3d_shape() {
    let chart = CHART
        .replace("<c:barChart>", "<c:bar3DChart>")
        .replace("</c:barChart>", "</c:bar3DChart>")
        .replace(
            "<c:tx><c:v>Keep</c:v></c:tx></c:ser>",
            r#"<c:tx><c:v>Keep</c:v></c:tx><c:shape val="box"/></c:ser>"#,
        );
    let mut book = Workbook::open(&bytes(&chart)).unwrap();
    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 1,
            name: None,
            categories: Some("Data!$A$2:$A$3".into()),
            values: Some("Data!$C$2:$C$3".into()),
        },
    );
    let xml = chart_xml(&book);
    assert!(
        xml.contains(r#"</c:cat><c:val><c:numRef><c:f>Data!$C$2:$C$3</c:f>"#),
        "{xml}"
    );
    assert!(
        xml.contains(r#"</c:val><c:shape val="box"/></c:ser>"#),
        "{xml}"
    );
}

#[test]
fn workbook_and_sheet_names_are_written_as_excel_qualifies_them() {
    let mut book = Workbook::open(&reference_bytes(CHART)).unwrap();
    let version = book.version();
    for reference in [
        "[0]!Local",
        "[0]!Unknown",
        "[1]!Totals",
        "[0]!$A$1",
        "[0]!Data!$A$1",
    ] {
        let error = book
            .edit_charts(
                vec![mutation(values(reference))],
                CalculationOptions::default(),
            )
            .unwrap_err()
            .to_string();
        assert!(error.contains("Sheet1!$B$2:$B$5"), "{reference}: {error}");
        assert_eq!(book.version(), version, "{reference}");
    }
    for (input, written) in [
        ("Totals", "[0]!Totals"),
        ("=[0]!Totals", "[0]!Totals"),
        ("DataLocal", "Data!DataLocal"),
        ("Data!DataLocal", "Data!DataLocal"),
    ] {
        apply(&mut book, values(input));
        let xml = chart_xml(&book);
        assert!(
            xml.contains(&format!("<c:f>{written}</c:f>")),
            "{input}: {xml}"
        );
    }
}

#[test]
fn values_and_name_references_are_one_dimensional() {
    let mut book = Workbook::open(&reference_bytes(CHART)).unwrap();
    let version = book.version();
    for edit in [
        values("Data!$B$2:$C$3"),
        values("Data!$B:$C"),
        values("Data!$2:$3"),
        ChartEdit::SetSeries {
            series: 0,
            name: Some(ChartSeriesName::Reference {
                reference: "Data!$B$1:$C$2".into(),
            }),
            categories: None,
            values: None,
        },
    ] {
        let error = book
            .edit_charts(vec![mutation(edit.clone())], CalculationOptions::default())
            .unwrap_err()
            .to_string();
        assert!(error.contains("one row or one column"), "{edit:?}: {error}");
        assert_eq!(book.version(), version);
    }
    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 0,
            name: Some(ChartSeriesName::Reference {
                reference: "Data!$B$1:$C$1".into(),
            }),
            categories: Some("Data!$A$2:$B$3".into()),
            values: Some("Data!$C$2:$C$3".into()),
        },
    );
    let xml = chart_xml(&book);
    assert!(xml.contains("<c:f>Data!$A$2:$B$3</c:f>"), "{xml}");
    assert!(xml.contains("<c:f>Data!$B$1:$C$1</c:f>"), "{xml}");
}

#[test]
fn chartsheets_cannot_be_series_sources() {
    let mut book = Workbook::open(&chartsheet_bytes(false)).unwrap();
    let version = book.version();
    let on_chartsheet = |edit| ChartMutation {
        sheet: SheetId(1),
        chart: 0,
        edit,
    };
    let error = book
        .edit_charts(
            vec![on_chartsheet(values("Chart!$A$1:$A$2"))],
            CalculationOptions::default(),
        )
        .unwrap_err()
        .to_string();
    assert!(error.contains("not a worksheet"), "{error}");
    assert!(error.contains("values"), "{error}");
    let error = book
        .edit_charts(
            vec![on_chartsheet(values("$C$2:$C$3"))],
            CalculationOptions::default(),
        )
        .unwrap_err()
        .to_string();
    assert!(error.contains("as in Data!$C$2:$C$3"), "{error}");
    assert_eq!(book.version(), version);
    book.edit_charts(
        vec![on_chartsheet(values("Data!$C$2:$C$3"))],
        CalculationOptions::default(),
    )
    .unwrap();
}

#[test]
fn missing_bubble_references_precede_bubble3d() {
    let mut chart = CHART
        .replace("<c:barChart><c:barDir val=\"col\"/>", "<c:bubbleChart>")
        .replace("</c:barChart>", "</c:bubbleChart>")
        .replace("<c:val>", "<c:yVal>")
        .replace("</c:val>", "</c:yVal>")
        .replace(
            "<c:tx><c:v>Keep</c:v></c:tx></c:ser>",
            r#"<c:tx><c:v>Keep</c:v></c:tx><c:bubble3D val="0"/></c:ser>"#,
        );
    let start = chart.find("<c:cat>").unwrap();
    let end = chart.find("</c:cat>").unwrap() + "</c:cat>".len();
    chart.replace_range(start..end, "");
    let mut book = Workbook::open(&bytes(&chart)).unwrap();
    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 1,
            name: None,
            categories: Some("Data!$B$2:$B$3".into()),
            values: Some("Data!$C$2:$C$3".into()),
        },
    );
    let xml = chart_xml(&book);
    assert!(xml.contains("</c:xVal><c:yVal>"), "{xml}");
    assert!(
        xml.contains(r#"</c:yVal><c:bubble3D val="0"/></c:ser>"#),
        "{xml}"
    );
}

fn edit_cell(book: &mut Workbook, at: &str, value: &str) {
    book.edit_cell(
        SheetId(0),
        betteroffice_xlsx::CellRef::parse_a1(at).unwrap(),
        value,
        CalculationOptions::default(),
    )
    .unwrap();
}

fn value_cache(xml: &str, series: usize) -> String {
    let start = xml.match_indices("<c:val>").nth(series).unwrap().0;
    xml[start..start + xml[start..].find("</c:val>").unwrap()].to_owned()
}

#[test]
fn self_closing_series_and_chart_elements_take_their_new_children() {
    let chart = CHART.replace(
        "<c:ser><c:idx val='1'/><c:order val='1'/><c:tx><c:v>Keep</c:v></c:tx></c:ser>",
        "<c:ser/>",
    );
    assert_ne!(chart, CHART);
    let mut book = Workbook::open(&bytes(&chart)).unwrap();
    apply(
        &mut book,
        ChartEdit::SetSeries {
            series: 1,
            name: Some(ChartSeriesName::Literal { text: "Rev".into() }),
            categories: Some("Data!$A$2:$A$3".into()),
            values: Some("Data!$C$2:$C$3".into()),
        },
    );
    let xml = chart_xml(&book);
    assert!(!xml.contains("<c:ser/>"), "{xml}");
    assert!(
        xml.contains("<c:ser><c:tx><c:v>Rev</c:v></c:tx><c:cat>"),
        "{xml}"
    );
    assert!(xml.contains("</c:val></c:ser></c:barChart>"), "{xml}");
    let reopened = Workbook::open(&book.save().unwrap()).unwrap();
    let info = reopened.charts(SheetId(0)).unwrap();
    assert_eq!(info[0].series[1].name.as_deref(), Some("Rev"));
    assert_eq!(
        info[0].series[1].value_ref.as_deref(),
        Some("Data!$C$2:$C$3")
    );

    let empty = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart/></c:chartSpace>"#;
    let mut book = Workbook::open(&bytes(empty)).unwrap();
    apply(&mut book, ChartEdit::SetChartTitle { title: None });
    assert!(
        chart_xml(&book).contains(r#"<c:chart><c:autoTitleDeleted val="1"/></c:chart>"#),
        "{}",
        chart_xml(&book)
    );
    let mut book = Workbook::open(&bytes(empty)).unwrap();
    apply(
        &mut book,
        ChartEdit::SetChartTitle {
            title: Some("Added".into()),
        },
    );
    let xml = chart_xml(&book);
    assert!(xml.contains("<c:chart><c:title>"), "{xml}");
    assert!(
        xml.contains(r#"<c:autoTitleDeleted val="0"/></c:chart>"#),
        "{xml}"
    );
}

#[test]
fn edited_chart_caches_follow_later_cell_edits_on_save() {
    let mut book = Workbook::open(&bytes(CHART)).unwrap();
    edit_cell(&mut book, "C2", "75");
    apply(&mut book, values("Data!$C$2:$C$3"));
    assert!(value_cache(&chart_xml(&book), 0).contains("<c:v>75</c:v>"));
    edit_cell(&mut book, "C2", "30");
    let cache = value_cache(&chart_xml(&book), 0);
    assert!(
        cache.contains("<c:pt idx=\"0\"><c:v>30</c:v></c:pt>"),
        "{cache}"
    );
    assert!(!cache.contains("75"), "{cache}");

    apply(&mut book, values("Data!$C$1:$C$3"));
    edit_cell(&mut book, "C3", "55");
    let cache = value_cache(&chart_xml(&book), 0);
    assert!(cache.contains("<c:ptCount val=\"3\"/>"), "{cache}");
    assert!(!cache.contains("<c:pt idx=\"0\">"), "{cache}");
    assert!(
        cache.contains("<c:pt idx=\"2\"><c:v>55</c:v></c:pt>"),
        "{cache}"
    );
}

const FIXTURE: &[u8] = include_bytes!("../../../packages/xlsx/test-fixtures/charts.xlsx");

#[test]
fn moved_references_keep_edited_and_multi_level_charts_readable() {
    let source = edit_parts(FIXTURE, |parts| {
        let part = parts
            .iter_mut()
            .find(|(name, _)| name == "xl/charts/chart1.xml")
            .unwrap();
        let mut xml = String::from_utf8(part.1.clone()).unwrap();
        let start = xml.find("<c:cat>").unwrap();
        let end = xml.find("</c:cat>").unwrap() + "</c:cat>".len();
        xml.replace_range(
            start..end,
            r#"<c:cat><c:multiLvlStrRef><c:f>Charts!$A$2:$B$5</c:f><c:multiLvlStrCache><c:ptCount val="4"/><c:lvl><c:pt idx="0"><c:v>One</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>"#,
        );
        part.1 = xml.into_bytes();
    });
    let mut book = Workbook::open(&source).unwrap();
    let chart = book
        .charts(SheetId(0))
        .unwrap()
        .iter()
        .position(|chart| chart.part == "xl/charts/chart1.xml")
        .unwrap();
    book.apply_ops(
        vec![Op::RenameSheet {
            sheet: SheetId(0),
            name: "Sales".into(),
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    let info = book.charts(SheetId(0)).unwrap();
    assert!(info[chart].editable, "{:?}", info[chart].error);
    assert_eq!(
        info[chart].series[0].category_ref.as_deref(),
        Some("Sales!$A$2:$B$5")
    );
    let title = |text: &str| ChartMutation {
        sheet: SheetId(0),
        chart,
        edit: ChartEdit::SetChartTitle {
            title: Some(text.into()),
        },
    };
    book.edit_charts(vec![title("Edited")], CalculationOptions::default())
        .unwrap();
    book.apply_ops(
        vec![Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    let info = book.charts(SheetId(0)).unwrap();
    assert!(info[chart].editable, "{:?}", info[chart].error);
    assert_eq!(info[chart].title.as_deref(), Some("Edited"));
    assert_eq!(
        info[chart].series[0].category_ref.as_deref(),
        Some("Sales!$A$3:$B$6")
    );
    book.edit_charts(vec![title("Edited again")], CalculationOptions::default())
        .unwrap();
    let xml = chart_xml(&book);
    assert!(
        xml.contains("Edited again") && xml.contains("<c:f>Sales!$A$3:$B$6</c:f>"),
        "{xml}"
    );
}

#[test]
fn snapshots_of_unedited_charted_workbooks_hydrate_and_take_chart_edits() {
    let book = Workbook::open(FIXTURE).unwrap();
    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(&book, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(&book).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    let (mut restored, _) = builder.finish().unwrap().into_parts();
    assert_eq!(
        restored.charts(SheetId(0)).unwrap(),
        book.charts(SheetId(0)).unwrap()
    );
    restored
        .edit_charts(
            vec![ChartMutation {
                sheet: SheetId(0),
                chart: 0,
                edit: ChartEdit::SetChartTitle {
                    title: Some("Hydrated".into()),
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let reopened = Workbook::open(&restored.save().unwrap()).unwrap();
    assert_eq!(
        reopened.charts(SheetId(0)).unwrap()[0].title.as_deref(),
        Some("Hydrated")
    );
}

#[test]
fn quoted_sheet_names_may_hold_commas_parentheses_and_apostrophes() {
    let names = ["Sales (2026)", "Sales, West", "O'Brien"];
    let source = edit_parts(&bytes(CHART), |parts| {
        let (mut sheets, mut rels) = (String::new(), String::new());
        for (index, name) in names.iter().enumerate() {
            let id = index + 2;
            sheets.push_str(&format!(
                r#"<sheet name="{name}" sheetId="{id}" r:id="r{id}"/>"#
            ));
            rels.push_str(&format!(
                r#"<Relationship Id="r{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{id}.xml"/>"#
            ));
            parts.push((
                format!("xl/worksheets/sheet{id}.xml"),
                format!(
                    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2"><v>{}</v></c></row><row r="3"><c r="B3"><v>{}</v></c></row></sheetData></worksheet>"#,
                    id * 100 + 1,
                    id * 100 + 2
                )
                .into_bytes(),
            ));
        }
        replace_part(
            parts,
            "xl/workbook.xml",
            "</sheets>",
            &format!("{sheets}</sheets>"),
        );
        replace_part(
            parts,
            "xl/_rels/workbook.xml.rels",
            "</Relationships>",
            &format!("{rels}</Relationships>"),
        );
    });
    let mut book = Workbook::open(&source).unwrap();
    for (index, name) in names.iter().enumerate() {
        let id = index + 2;
        let reference = format!("'{}'!$B$2:$B$3", name.replace('\'', "''"));
        apply(&mut book, values(&reference));
        let xml = chart_xml(&book);
        assert!(xml.contains(&format!("<c:f>{reference}</c:f>")), "{xml}");
        let cache = value_cache(&xml, 0);
        for value in [id * 100 + 1, id * 100 + 2] {
            assert!(cache.contains(&format!("<c:v>{value}</c:v>")), "{cache}");
        }
        let info = Workbook::open(&book.save().unwrap())
            .unwrap()
            .charts(SheetId(0))
            .unwrap();
        assert_eq!(
            info[0].series[0].value_ref.as_deref(),
            Some(reference.as_str())
        );
    }
    let version = book.version();
    for reference in [
        "'Sales, West'!$B$2,'Sales, West'!$B$3",
        "'Sales (2026)'!$B$2:$B$3,Data!$B$2",
        "SUM('Sales (2026)'!$B$2:$B$3)",
        "('Sales, West'!$B$2:$B$3)",
        "'O''Brien'!$B$2:$B$3 'O''Brien'!$B$2",
    ] {
        let error = book
            .edit_charts(
                vec![mutation(values(reference))],
                CalculationOptions::default(),
            )
            .unwrap_err()
            .to_string();
        assert!(error.contains("values"), "{reference}: {error}");
        assert_eq!(book.version(), version, "{reference}");
    }
}
