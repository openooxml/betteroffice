#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use serde_json::Value;

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const DRAWING: &str = concat!(
    r#"<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">"#,
    r#"<wp:inline><wp:extent cx="2743200" cy="1828800"/><wp:docPr id="1" name="Chart 1"/>"#,
    r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">"#,
    r#"<c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#
);

/// A column chart over 0.123 and 0.8, cached without a format, labelled by
/// `labels`.
fn chart_part(labels: &str) -> String {
    format!(
        concat!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>"#,
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/>"#,
            r#"<c:order val="0"/><c:tx><c:v>Share</c:v></c:tx>{labels}<c:cat><c:strRef><c:strCache>"#,
            r#"<c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt>"#,
            r#"</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="2"/>"#,
            r#"<c:pt idx="0"><c:v>0.123</c:v></c:pt><c:pt idx="1"><c:v>0.8</c:v></c:pt>"#,
            r#"</c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>"#,
            r#"<c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>"#,
            r#"<c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>"#,
            r#"</c:plotArea></c:chart></c:chartSpace>"#
        ),
        labels = labels
    )
}

fn charted_docx(labels: &str) -> Vec<u8> {
    let body = fixture::p("10000001", &format!("<w:r>{DRAWING}</w:r>"));
    let mut parts = ooxml_opc::unzip_parts(&fixture::with_body(&body)).unwrap();
    for (name, bytes) in &mut parts {
        let (end, entry) = match name.as_str() {
            "word/_rels/document.xml.rels" => (
                "</Relationships>",
                format!(
                    r#"<Relationship Id="rIdChart" Type="{REL}/chart" Target="charts/chart1.xml"/>"#
                ),
            ),
            "[Content_Types].xml" => (
                "</Types>",
                r#"<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#
                    .to_owned(),
            ),
            _ => continue,
        };
        let xml = String::from_utf8(std::mem::take(bytes)).unwrap();
        *bytes = xml.replacen(end, &format!("{entry}{end}"), 1).into_bytes();
    }
    parts.push((
        "word/charts/chart1.xml".to_owned(),
        chart_part(labels).into_bytes(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn page_texts(page: &Value) -> Vec<String> {
    page["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|primitive| primitive["text"].as_str().map(str::to_owned))
        .collect()
}

/// The texts the resident frame and the JSON display build paint on the first page.
fn painted_texts(labels: &str) -> (Vec<String>, Vec<String>) {
    let (engine, request) = fixture::laid_out(&charted_docx(labels), 9702);
    engine
        .build_display_list_frame("{}", 0)
        .expect("the resident display list builds");
    let resident = engine
        .with_display_list(|list| serde_json::to_value(list).unwrap())
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    let json: Value = serde_json::from_str(
        &engine
            .build_display_list_json(&layout)
            .expect("the JSON display list builds"),
    )
    .unwrap();
    (
        page_texts(&resident["pages"][0]),
        page_texts(&json["pages"][0]),
    )
}

#[test]
fn a_linked_label_without_a_cached_format_keeps_its_own_code() {
    let (resident, json) = painted_texts(
        r#"<c:dLbls><c:numFmt formatCode="0.00%" sourceLinked="1"/><c:showVal val="1"/></c:dLbls>"#,
    );
    for texts in [&resident, &json] {
        assert!(texts.contains(&"12.30%".to_owned()), "{texts:?}");
        assert!(texts.contains(&"80.00%".to_owned()), "{texts:?}");
    }
}
