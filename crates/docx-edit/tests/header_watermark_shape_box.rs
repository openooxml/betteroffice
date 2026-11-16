use docx_edit::layout_request::{measurement, region_layout_request};
use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const HEADER_START: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">"#,
    r#"<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Watermarks"/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:r><w:pict>"#,
    r#"<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800"><v:path textpathok="t"/><v:textpath on="t" fitshape="t"/></v:shapetype>"#,
);

const HEADER_END: &str = r#"</w:pict></w:r></w:p></w:sdtContent></w:sdt></w:hdr>"#;

fn watermarked_document([width, height, family, size, text]: [&str; 5]) -> Vec<u8> {
    let shape = format!(
        r##"<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:{width};height:{height};rotation:315;z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;{family}&quot;;font-size:{size}" string="{text}"/></v:shape>"##
    );
    let header = format!("{HEADER_START}{shape}{HEADER_END}");
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_owned()),
        ("_rels/.rels".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#.to_owned()),
        ("word/_rels/document.xml.rels".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_owned()),
        ("word/header1.xml".to_owned(), header),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name, value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn painted_watermark(bytes: &[u8], font: u32) -> Value {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let package = serde_json::to_value(package).unwrap();
    let mut request = region_layout_request(&package, 24.0);
    request["measurement"] = measurement(&package);
    request["measurement"]["fontChains"] = json!({"arial|0|0": [font], "calibri|0|0": [font]});
    let engine = EngineSession::new(73103);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    let page = &display["pages"][0];
    assert_eq!(page["watermarkPrimitiveCount"], json!(1));
    page["primitives"][0].clone()
}

fn em_px(run: &Value) -> f64 {
    run["font"]
        .as_str()
        .unwrap()
        .split_whitespace()
        .find_map(|part| part.strip_suffix("px")?.parse().ok())
        .unwrap()
}

fn rotated_em_box_points(run: &Value) -> [f64; 4] {
    let em = em_px(run);
    let x = run["x"].as_f64().unwrap();
    let width = run["width"].as_f64().unwrap();
    let baseline = run["baselineY"].as_f64().unwrap();
    let (pivot_x, pivot_y) = (x + width / 2.0, baseline - 0.2 * em);
    let angle = run["rotationDeg"].as_f64().unwrap().to_radians();
    let (sin, cos) = angle.sin_cos();
    let top = baseline - 0.8 * em;
    let mut bounds = [f64::MAX, f64::MAX, f64::MIN, f64::MIN];
    for (corner_x, corner_y) in [
        (x, top),
        (x + width, top),
        (x, top + em),
        (x + width, top + em),
    ] {
        let (dx, dy) = (corner_x - pivot_x, corner_y - pivot_y);
        let turned_x = pivot_x + dx * cos - dy * sin;
        let turned_y = pivot_y + dx * sin + dy * cos;
        bounds = [
            bounds[0].min(turned_x),
            bounds[1].min(turned_y),
            bounds[2].max(turned_x),
            bounds[3].max(turned_y),
        ];
    }
    bounds.map(|value| value * 0.75)
}

#[test]
fn text_watermark_fills_its_shape_box_centred_on_the_margin_box() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let cases = [
        (
            ["592.1pt", "47.35pt", "Arial", "1pt", "FOR REVIEW"],
            [592.1, 47.35],
            [79.92, 169.92, 532.08, 622.08],
        ),
        (
            ["672pt", "58.5pt", "Calibri", "48pt", "Sample copy"],
            [672.0, 58.5],
            [47.73, 137.73, 564.27, 654.27],
        ),
        (
            ["412.4pt", "247.45pt", "Calibri", "1pt", "DRAFT"],
            [412.4, 247.45],
            [72.71, 162.71, 539.29, 629.29],
        ),
    ];
    for (shape, [width_pt, height_pt], expected) in cases {
        let [_, _, family, _, text] = shape;
        let run = painted_watermark(&watermarked_document(shape), font);
        assert_eq!(run["kind"], "text", "{text}");
        assert_eq!(run["text"], text);
        assert_eq!(run["rotationDeg"].as_f64(), Some(-45.0), "{text}");
        let bounds = rotated_em_box_points(&run);
        for (got, want) in bounds.iter().zip(expected) {
            assert!(
                (got - want).abs() < 0.05,
                "{text}: box {bounds:?}, expected {expected:?}"
            );
        }
        let painted_width = run["width"].as_f64().unwrap() * 0.75;
        assert!((painted_width - width_pt).abs() < 0.01, "{text}: {run}");
        let em_pt = em_px(&run) * 0.75;
        assert!((em_pt - height_pt).abs() < 0.01, "{text}: {run}");
        let face = run["font"].as_str().unwrap();
        assert!(face.contains(family), "{text}: {run}");
        let stretch = run["horizontalScale"].as_f64();
        assert!(stretch.is_some_and(|scale| scale > 0.0), "{text}: {run}");
    }
}
