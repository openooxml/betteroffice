use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

#[path = "support/quality_fixture.rs"]
mod quality_fixture;
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[test]
fn anchored_header_preserves_body_space_and_paints_at_page_coordinates() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(73002);
    seed_from_docx(engine.doc(), &quality_fixture::document(false)).unwrap();
    let request = json!({
        "bodyStory": "body",
        "options": {},
        "regions": {"sections": [{
            "sectionId": "main",
            "pageSize": {"w": 1122.6666666666667, "h": 793.3333333333334},
            "margins": {"top": 113.33333333333333, "right": 113.33333333333333,
                "bottom": 113.33333333333333, "left": 113.33333333333333,
                "header": 56.666666666666664, "footer": 56.666666666666664},
            "headerFooterRefs": {"headerFirst": "rIdHeader"}
        }]},
        "measurement": {
            "fontChains": {"arial|0|0": [font], "calibri|0|0": [font]},
            "defaults": {"fontSize": 12, "fontFamily": "Arial"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    });
    let output = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let value: Value = serde_json::from_str(&output).unwrap();
    assert_eq!(value["layout"]["pages"].as_array().unwrap().len(), 1);
    assert!(
        (value["options"]["margins"]["top"].as_f64().unwrap() - 113.33333333333333).abs() < 0.001
    );
    let variant = &value["headersFooters"]["variants"][0];
    assert!(variant["flowHeight"].as_f64().unwrap() < 30.0);
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    let primitives = display["pages"][0]["header"]["primitives"]
        .as_array()
        .unwrap();
    let shape = primitives
        .iter()
        .find(|primitive| primitive["kind"] == "shape")
        .expect("header shape must paint");
    assert_eq!(shape["x"], 1010);
    assert!((shape["y"].as_f64().unwrap() - 56.667).abs() < 0.001);
    assert_eq!(shape["w"], 50);
}

fn vml_header_document() -> Vec<u8> {
    let header = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word">"#,
        r#"<w:p><w:r><w:pict><v:shape id="Text Box 1" o:spid="_x0000_s1026" style="position:absolute;margin-left:36pt;margin-top:18pt;width:144pt;height:72pt;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;v-text-anchor:top" filled="f" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>Header label</w:t></w:r></w:p></w:txbxContent></v:textbox><w10:wrap type="none"/></v:shape></w:pict></w:r></w:p></w:hdr>"#,
    );
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_owned()),
        ("_rels/.rels".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body text</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="16840" w:h="11900" w:orient="landscape"/><w:pgMar w:top="1700" w:right="1700" w:bottom="1700" w:left="1700" w:header="850" w:footer="850"/></w:sectPr></w:body></w:document>"#.to_owned()),
        ("word/_rels/document.xml.rels".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_owned()),
        ("word/header1.xml".to_owned(), header.to_owned()),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name, value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

#[test]
fn vml_text_box_in_header_paints_text_at_box_position() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(73003);
    seed_from_docx(engine.doc(), &vml_header_document()).unwrap();
    let request = json!({
        "bodyStory": "body",
        "options": {},
        "regions": {"sections": [{
            "sectionId": "main",
            "pageSize": {"w": 1122.6666666666667, "h": 793.3333333333334},
            "margins": {"top": 113.33333333333333, "right": 113.33333333333333,
                "bottom": 113.33333333333333, "left": 113.33333333333333,
                "header": 56.666666666666664, "footer": 56.666666666666664},
            "headerFooterRefs": {"headerFirst": "rIdHeader"}
        }]},
        "measurement": {
            "fontChains": {"arial|0|0": [font], "calibri|0|0": [font]},
            "defaults": {"fontSize": 12, "fontFamily": "Arial"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    });
    let output = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    let primitives = display["pages"][0]["header"]["primitives"]
        .as_array()
        .unwrap();
    let shape = primitives
        .iter()
        .find(|primitive| primitive["kind"] == "shape")
        .expect("header VML text box must paint a shape");
    for (key, expected) in [("x", 48.0), ("y", 24.0), ("w", 192.0), ("h", 96.0)] {
        let actual = shape[key].as_f64().unwrap();
        assert!(
            (actual - expected).abs() < 1.0,
            "shape {key}: {actual} != {expected}"
        );
    }
    let painted: String = primitives
        .iter()
        .filter(|primitive| matches!(primitive["kind"].as_str(), Some("text" | "glyphRun")))
        .filter_map(|primitive| primitive["text"].as_str())
        .collect();
    assert_eq!(painted.matches("Header label").count(), 1, "{painted}");
}
