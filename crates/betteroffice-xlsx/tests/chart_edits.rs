use betteroffice_xlsx::{
    CalculationOptions, ChartEdit, ChartMutation, ChartSeriesName, Error, SheetId, Workbook,
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
    let mut parts = parts.into_iter().map(|(p, x)| (p.to_owned(), x.as_bytes().to_vec())).collect::<Vec<_>>();
    let drawing = parts.iter_mut().find(|(path, _)| path == "xl/drawings/drawing1.xml").unwrap();
    let xml = String::from_utf8(drawing.1.clone()).unwrap();
    let start = xml.find("<xdr:twoCellAnchor>").unwrap();
    let end = xml.find("</xdr:twoCellAnchor>").unwrap() + "</xdr:twoCellAnchor>".len();
    let second = xml[start..end].replace("r:id=\"c1\"", "r:id=\"c2\"");
    drawing.1 = xml.replace("</xdr:wsDr>", &format!("{second}</xdr:wsDr>")).into_bytes();
    let rels = parts.iter_mut().find(|(path, _)| path == "xl/drawings/_rels/drawing1.xml.rels").unwrap();
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
    assert!(serde_json::from_str::<ChartMutation>(r#"{"sheet":0,"chart":0,"action":"set_chart_title"}"#).is_err());
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
            categories: Some("Missing!$A$1:$A$2".into()),
            values: None,
        },
    );
    let xml = chart_xml(&book);
    assert!(xml.contains("<c:tx><c:v>Literal &lt;name&gt;</c:v></c:tx>"));
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
    apply(&mut book, ChartEdit::SetChartTitle { title: Some("Namespaced title".into()) });
    apply(&mut book, ChartEdit::SetSeries { series: 0, name: None, categories: None, values: Some("Data!$C$2:$C$3".into()) });
    let saved = book.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    let info = reopened.charts(SheetId(0)).unwrap();
    assert_eq!(info[0].title.as_deref(), Some("Namespaced title"));
    assert_eq!(info[0].series[0].value_ref.as_deref(), Some("Data!$C$2:$C$3"));
    assert!(chart_xml(&book).contains("<q:v>30</q:v>"));
}

#[test]
fn shared_chart_parts_refuse_edits_without_changing_any_anchor() {
    let mut parts = ooxml_opc::unzip_parts(&bytes(CHART)).unwrap();
    let rels = parts.iter_mut().find(|(path, _)| path == "xl/drawings/_rels/drawing1.xml.rels").unwrap();
    rels.1 = String::from_utf8(rels.1.clone()).unwrap().replace("../charts/chart2.xml", "../charts/chart1.xml").into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut book = Workbook::open(&source).unwrap();
    let version = book.version();
    let error = book.edit_charts(vec![mutation(ChartEdit::SetChartTitle { title: Some("Shared title".into()) })], CalculationOptions::default()).unwrap_err();
    assert!(error.to_string().contains("shared"));
    assert_eq!(book.version(), version);
    assert_eq!(book.save().unwrap(), source);
}

fn chartsheet_bytes(protected: bool) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&bytes(CHART)).unwrap();
    let workbook = parts.iter_mut().find(|(path, _)| path == "xl/workbook.xml").unwrap();
    workbook.1 = String::from_utf8(workbook.1.clone()).unwrap().replace("</sheets>", r#"<sheet name="Chart" sheetId="2" r:id="rChart"/></sheets>"#).into_bytes();
    let rels = parts.iter_mut().find(|(path, _)| path == "xl/_rels/workbook.xml.rels").unwrap();
    rels.1 = String::from_utf8(rels.1.clone()).unwrap().replace("</Relationships>", r#"<Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>"#).into_bytes();
    let worksheet = parts.iter_mut().find(|(path, _)| path == "xl/worksheets/sheet1.xml").unwrap();
    worksheet.1 = String::from_utf8(worksheet.1.clone()).unwrap().replace(r#"<drawing r:id="d1"/>"#, "").into_bytes();
    let rels = parts.iter_mut().find(|(path, _)| path == "xl/worksheets/_rels/sheet1.xml.rels").unwrap();
    rels.0 = "xl/chartsheets/_rels/sheet1.xml.rels".into();
    let protection = if protected { r#"<sheetProtection content="1" objects="1"/>"# } else { "" };
    parts.push(("xl/chartsheets/sheet1.xml".into(), format!(r#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{protection}<drawing r:id="d1"/></chartsheet>"#).into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn chartsheet_title_and_series_edits_preserve_sheet_parts() {
    let source = chartsheet_bytes(false);
    let mut book = Workbook::open(&source).unwrap();
    assert!(book.charts(SheetId(0)).unwrap().is_empty());
    assert_eq!(book.charts(SheetId(1)).unwrap().len(), 2);
    book.edit_charts(vec![
        ChartMutation { sheet: SheetId(1), chart: 0, edit: ChartEdit::SetChartTitle { title: Some("Chart view".into()) } },
        ChartMutation { sheet: SheetId(1), chart: 0, edit: ChartEdit::SetSeries { series: 0, name: Some(ChartSeriesName::Literal { text: "Chart series".into() }), categories: None, values: Some("Data!$C$2:$C$3".into()) } },
    ], CalculationOptions::default()).unwrap();
    let saved = book.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    let info = reopened.charts(SheetId(1)).unwrap();
    assert_eq!(info[0].title.as_deref(), Some("Chart view"));
    assert_eq!(info[0].series[0].name.as_deref(), Some("Chart series"));
    assert_eq!(info[0].series[0].value_ref.as_deref(), Some("Data!$C$2:$C$3"));
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    for path in ["xl/worksheets/sheet1.xml", "xl/chartsheets/sheet1.xml", "xl/drawings/drawing1.xml", "xl/charts/chart2.xml"] {
        assert_eq!(before.iter().find(|(part, _)| part == path).unwrap().1, after.iter().find(|(part, _)| part == path).unwrap().1, "{path}");
    }
}

#[test]
fn chart_owner_protection_is_enforced_by_native_edits() {
    let mut worksheet = ooxml_opc::unzip_parts(&bytes(CHART)).unwrap();
    let sheet = worksheet.iter_mut().find(|(path, _)| path == "xl/worksheets/sheet1.xml").unwrap();
    sheet.1 = String::from_utf8(sheet.1.clone()).unwrap().replace("</sheetData>", r#"</sheetData><sheetProtection sheet="1"/>"#).into_bytes();
    for (source, target) in [(chartsheet_bytes(true), SheetId(1)), (ooxml_opc::rezip_parts(&worksheet).unwrap(), SheetId(0))] {
        let mut book = Workbook::open(&source).unwrap();
        let version = book.version();
        let result = book.edit_charts(vec![ChartMutation { sheet: target, chart: 0, edit: ChartEdit::SetChartTitle { title: Some("Protected title".into()) } }], CalculationOptions::default());
        assert!(result.is_err());
        assert_eq!(book.version(), version);
        assert_eq!(book.save().unwrap(), source);
    }
}

#[test]
fn missing_xy_references_use_scatter_and_bubble_slots() {
    for kind in ["scatterChart", "bubbleChart"] {
        let mut chart = CHART.replace("<c:barChart><c:barDir val=\"col\"/>", &format!("<c:{kind}>"))
            .replace("</c:barChart>", &format!("</c:{kind}>"))
            .replace("<c:val>", "<c:yVal>").replace("</c:val>", "</c:yVal>");
        let start = chart.find("<c:cat>").unwrap();
        let end = chart.find("</c:cat>").unwrap() + "</c:cat>".len();
        chart.replace_range(start..end, "");
        let mut book = Workbook::open(&bytes(&chart)).unwrap();
        book.edit_charts(vec![
            mutation(ChartEdit::SetSeries { series: 0, name: None, categories: Some("Data!$C$2:$C$3".into()), values: None }),
            mutation(ChartEdit::SetSeries { series: 1, name: None, categories: None, values: Some("Data!$B$2:$B$3".into()) }),
        ], CalculationOptions::default()).unwrap();
        let xml = chart_xml(&book);
        assert!(!xml.contains("<c:cat>"), "{kind}");
        assert!(!xml.contains("<c:val>"), "{kind}");
        assert_eq!(xml.matches("<c:xVal><c:numRef>").count(), 1, "{kind}");
        assert_eq!(xml.matches("<c:yVal>").count(), 2, "{kind}");
        let reopened = Workbook::open(&book.save().unwrap()).unwrap();
        let info = reopened.charts(SheetId(0)).unwrap();
        assert_eq!(info[0].series[0].category_ref.as_deref(), Some("Data!$C$2:$C$3"));
        assert_eq!(info[0].series[1].value_ref.as_deref(), Some("Data!$B$2:$B$3"));
        book.undo(CalculationOptions::default()).unwrap();
        assert_eq!(chart_xml(&book), chart);
    }
}
