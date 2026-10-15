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

/// A column chart over categories A, B and C whose value cache holds `points`
/// under a point count of three.
fn chart_part(points: &str) -> String {
    format!(
        concat!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>"#,
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/>"#,
            r#"<c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:spPr><a:solidFill>"#,
            r#"<a:srgbClr val="4472C4"/></a:solidFill></c:spPr><c:cat><c:strRef><c:strCache>"#,
            r#"<c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt>"#,
            r#"<c:pt idx="2"><c:v>C</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef>"#,
            r#"<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>{points}"#,
            r#"</c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>"#,
            r#"<c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>"#,
            r#"<c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>"#,
            r#"</c:plotArea><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>"#
        ),
        points = points
    )
}

fn charted_docx(points: &str) -> Vec<u8> {
    let body = format!(
        "{}{}",
        fixture::p("10000001", &fixture::r("Quarterly sales")),
        fixture::p("10000002", &format!("<w:r>{DRAWING}</w:r>"))
    );
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
        chart_part(points).into_bytes(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The chart rects in the series colour: its bars and its legend keys.
fn series_rects(page: &Value) -> usize {
    page["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|primitive| {
            primitive["kind"] == "rect"
                && primitive["chart"].is_object()
                && primitive["fill"]
                    .as_str()
                    .is_some_and(|fill| fill.eq_ignore_ascii_case("#4472c4"))
        })
        .count()
}

/// The series rects the resident frame and the JSON display build paint on the first page.
fn painted_rects(points: &str) -> (usize, usize) {
    let (engine, request) = fixture::laid_out(&charted_docx(points), 9701);
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
        series_rects(&resident["pages"][0]),
        series_rects(&json["pages"][0]),
    )
}

#[test]
fn a_chart_with_a_blank_cached_point_renders_with_a_gap() {
    let point =
        |index: usize, value: &str| format!(r#"<c:pt idx="{index}"><c:v>{value}</c:v></c:pt>"#);
    let (resident, json) = painted_rects(&format!(
        "{}{}{}",
        point(0, "2"),
        point(1, "1"),
        point(2, "3")
    ));
    assert!(resident >= 3);
    assert_eq!(json, resident);
    assert_eq!(
        painted_rects(&format!("{}{}", point(0, "2"), point(2, "3"))),
        (resident - 1, resident - 1)
    );
}
