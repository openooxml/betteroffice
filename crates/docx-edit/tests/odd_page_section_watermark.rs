use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn odd_page_section_document() -> Vec<u8> {
    let header = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">"#,
        r#"<w:p><w:r><w:t>Header</w:t></w:r></w:p>"#,
        r#"<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Watermarks"/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:r><w:pict>"#,
        r##"<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;width:412.4pt;height:206.2pt;rotation:315" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Liberation Sans&quot;;font-size:1pt" string="DRAFT"/></v:shape>"##,
        r#"</w:pict></w:r></w:p></w:sdtContent></w:sdt></w:hdr>"#,
    );
    let section = |start: &str| {
        format!(
            r#"<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>{start}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>"#
        )
    };
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:pPr>{}</w:pPr><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p>{}</w:body></w:document>"#,
        section(""),
        section(r#"<w:type w:val="oddPage"/>"#)
    );
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_owned()),
        ("_rels/.rels".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml".to_owned(), document),
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
fn odd_page_section_start_keeps_two_watermarked_pages() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(73102);
    seed_from_docx(engine.doc(), &odd_page_section_document()).unwrap();
    let watermark = json!({
        "kind": "text",
        "text": "DRAFT",
        "font": "Liberation Sans",
        "color": "silver",
        "semitransparent": true,
        "layout": "diagonal"
    });
    let section = |id: &str| {
        json!({
            "sectionId": id,
            "pageSize": {"w": 816, "h": 1056},
            "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96, "header": 48, "footer": 48},
            "titlePg": false,
            "headerFooterRefs": {"headerDefault": "rIdHeader"}
        })
    };
    let request = json!({
        "bodyStory": "body",
        "options": {},
        "regions": {
            "sections": [section("first"), section("final")],
            "watermark": watermark.clone(),
            "headerWatermarks": {"rIdHeader": watermark}
        },
        "measurement": {
            "fontChains": {"liberation sans|0|0": [font], "calibri|0|0": [font]},
            "defaults": {"fontSize": 12, "fontFamily": "Liberation Sans"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    });
    let output = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    let pages = display["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 2);
    for page in pages {
        assert!(!page["header"].is_null());
        assert_eq!(page["header"]["rId"], "rIdHeader");
        assert_eq!(page["parityFiller"], Value::Null);
        assert_eq!(page["watermarkPrimitiveCount"], json!(1));
    }
}
