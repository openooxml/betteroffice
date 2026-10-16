use serde_json::{Value, json};

use crate::EngineSession;

const FONT: &[u8] = include_bytes!("../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#;
const DRAWING: &str = r#"<w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="7" name="Chart" descr="Quarterly sales"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId9"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#;
const CHART: &str = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>North</c:v></c:tx><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

fn document(boxed: bool, hidden_chart: bool) -> Vec<u8> {
    let formatting = if hidden_chart {
        "<w:rPr><w:vanish/></w:rPr>"
    } else {
        ""
    };
    let body = if boxed {
        format!(
            r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>457200</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>228600</wp:posOffset></wp:positionV><wp:extent cx="2743200" cy="1371600"/><wp:wrapNone/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Before</w:t></w:r><w:r>{formatting}<mc:AlternateContent><mc:Choice Requires="c">{DRAWING}</mc:Choice><mc:Fallback><w:t>Fallback</w:t></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t>After</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:t>Fallback</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>"#
        )
    } else {
        format!("<w:p><w:r>{formatting}{DRAWING}</w:r></w:p>")
    };
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>"#.to_owned()),
        ("word/charts/chart1.xml", CHART.to_owned()),
        ("word/document.xml", format!(r#"<w:document {NAMESPACES}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#)),
    ];
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, value)| (name.to_owned(), value.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn primitives(boxed: bool, hidden_chart: bool, show_hidden_text: bool) -> Vec<Value> {
    let bytes = document(boxed, hidden_chart);
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75302);
    crate::seed::seed_with_layout_tokens(engine.doc(), &bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body",
                "renderEnv": {"showHiddenText": show_hidden_text},
                "regions": {"sections": [{"properties": package.document.final_section_properties}]},
                "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
            })
            .to_string(),
        )
        .unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    assert_eq!(display["pages"].as_array().unwrap().len(), 1);
    display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .clone()
}

#[test]
fn alternate_content_text_box_chart_paints_inside_the_box() {
    let boxed = primitives(true, false, false);
    let body = primitives(false, false, false);
    let charts: Vec<_> = boxed
        .iter()
        .filter(|primitive| !primitive["chart"].is_null())
        .collect();
    let body_charts: Vec<_> = body
        .iter()
        .filter(|primitive| !primitive["chart"].is_null())
        .collect();
    assert!(!charts.is_empty());
    assert_eq!(charts.len(), body_charts.len());
    assert!(boxed.iter().all(|primitive| primitive["kind"] != "image"));
    assert!(charts.iter().any(|primitive| primitive["text"] == "Sales"));
    let canvas = charts[0];
    assert_eq!(canvas["kind"], "rect");
    assert_eq!(canvas["w"], 192);
    assert_eq!(canvas["h"], 96);
    assert_eq!(canvas["chart"], body_charts[0]["chart"]);
    assert_eq!(canvas["ariaDescription"], "Quarterly sales");
    let shape = boxed
        .iter()
        .find(|primitive| primitive["kind"] == "shape")
        .unwrap();
    for (position, size) in [("x", "w"), ("y", "h")] {
        let start = canvas[position].as_f64().unwrap();
        let end = start + canvas[size].as_f64().unwrap();
        let box_start = shape[position].as_f64().unwrap();
        let box_end = box_start + shape[size].as_f64().unwrap();
        assert!(start >= box_start && end <= box_end);
    }
    assert!(canvas["x"].as_f64().unwrap() > shape["x"].as_f64().unwrap());
    let text: String = boxed
        .iter()
        .filter(|primitive| primitive["chart"].is_null())
        .filter_map(|primitive| primitive["text"].as_str())
        .collect();
    assert_eq!(text, "BeforeAfter");
}

#[test]
fn hidden_text_box_chart_obeys_show_hidden_text() {
    for boxed in [false, true] {
        let hidden = primitives(boxed, true, false);
        assert!(hidden.iter().all(|primitive| primitive["chart"].is_null()));
        assert!(hidden.iter().all(|primitive| primitive["kind"] != "image"));
        if boxed {
            let text: String = hidden
                .iter()
                .filter_map(|primitive| primitive["text"].as_str())
                .collect();
            assert_eq!(text, "BeforeAfter");
        }
        let shown = primitives(boxed, true, true);
        assert!(shown.iter().any(|primitive| primitive["text"] == "Sales"));
        let visible = primitives(boxed, false, false);
        let charts = |primitives: Vec<Value>| {
            primitives
                .into_iter()
                .filter(|primitive| !primitive["chart"].is_null())
                .collect::<Vec<_>>()
        };
        assert_eq!(charts(shown), charts(visible));
    }
}

#[test]
fn text_box_chart_carrier_uses_constrained_chart_dimensions() {
    use docx_layout::types::Run;

    use crate::bridge::RenderEnv;

    for (width, height, expected) in [
        (1, 4762, (1.0 / 914400.0 * 96.0, 4762.0 / 914400.0 * 96.0)),
        (-914400, -457200, (320.0, 220.0)),
        (0, 0, (320.0, 220.0)),
        (1828800, 914400, (192.0, 96.0)),
    ] {
        for page_height in [None, Some(80.0)] {
            let run = json!({"type": "run", "content": [{"type": "chart", "chart": {
                "chartType": "column", "size": {"width": width, "height": height},
                "series": [{"name": "Sales", "categories": ["Q1"], "values": [2.0]}]
            }}]});
            let runs = super::shape_document_runs(
                &run,
                &RenderEnv {
                    page_content_height: page_height,
                    ..RenderEnv::default()
                },
            );
            let Run::Image(image) = &runs[0] else {
                panic!("missing chart carrier")
            };
            let chart = image.inline_chart.as_ref().unwrap();
            assert_eq!((image.width, image.height), (chart.width, chart.height));
            let expected = page_height
                .filter(|height| *height < expected.1)
                .map_or(expected, |height| {
                    ((expected.0 * (height / expected.1)).round(), height)
                });
            assert_eq!((chart.width, chart.height), expected);
        }
    }
}
