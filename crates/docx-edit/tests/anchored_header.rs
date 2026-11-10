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

const CARLITO: &[u8] = include_bytes!("../../../packages/fonts/assets/Carlito-Regular.ttf");

fn band_document(band: &str, ask: bool) -> Vec<u8> {
    let (root, kind) = if band == "header" {
        ("hdr", "header")
    } else {
        ("ftr", "footer")
    };
    let typography = if ask {
        r#"<w14:ligatures w14:val="standard"/>"#
    } else {
        ""
    };
    let band_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>{typography}</w:rPr><w:t xml:space="preserve">office fifty</w:t></w:r></w:p></w:{root}>"#
    );
    let parts = vec![
        ("[Content_Types].xml".to_owned(), format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/{kind}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/></Types>"#)),
        ("_rels/.rels".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml".to_owned(), format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body text</w:t></w:r></w:p><w:sectPr><w:{kind}Reference w:type="default" r:id="rIdBand"/><w:pgSz w:w="16840" w:h="11900" w:orient="landscape"/><w:pgMar w:top="1700" w:right="1700" w:bottom="1700" w:left="1700" w:header="850" w:footer="850"/></w:sectPr></w:body></w:document>"#)),
        ("word/_rels/document.xml.rels".to_owned(), format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdBand" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{kind}1.xml"/></Relationships>"#)),
        (format!("word/{kind}1.xml"), band_xml),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name, value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn band_glyphs(band: &str, ask: bool) -> BandPaint {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(CARLITO).unwrap();
    let engine = EngineSession::new(73004);
    seed_from_docx(engine.doc(), &band_document(band, ask)).unwrap();
    let mut refs = serde_json::Map::new();
    for slot in ["Default", "First"] {
        refs.insert(format!("{band}{slot}"), json!("rIdBand"));
    }
    let request = json!({
        "bodyStory": "body",
        "options": {},
        "regions": {"sections": [{
            "sectionId": "main",
            "pageSize": {"w": 1122.6666666666667, "h": 793.3333333333334},
            "margins": {"top": 113.33333333333333, "right": 113.33333333333333,
                "bottom": 113.33333333333333, "left": 113.33333333333333,
                "header": 56.666666666666664, "footer": 56.666666666666664},
            "headerFooterRefs": refs
        }]},
        "measurement": {
            "fontChains": {"arial|0|0": [font], "calibri|0|0": [font]},
            "defaults": {"fontSize": 12, "fontFamily": "Calibri"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    });
    let output = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let mut display_input: Value = serde_json::from_str(&output).unwrap();
    display_input["fontChains"] = request["measurement"]["fontChains"].clone();
    let display_json = engine
        .build_display_list_json(&display_input.to_string())
        .unwrap();
    let display: Value = serde_json::from_str(&display_json).unwrap();
    let page = &display["pages"][0];
    let primitives = page[band]["primitives"].as_array().unwrap_or_else(|| {
        let keys: Vec<_> = page
            .as_object()
            .map(|object| object.keys().cloned().collect())
            .unwrap_or_default();
        panic!("no {band} band; page keys {keys:?}; request {request}")
    });
    let painted: Vec<&Value> = primitives
        .iter()
        .filter(|primitive| matches!(primitive["kind"].as_str(), Some("text" | "glyphRun")))
        .collect();
    let kinds: Vec<_> = primitives
        .iter()
        .filter_map(|primitive| primitive["kind"].as_str())
        .collect();
    let text: String = painted
        .iter()
        .filter_map(|primitive| primitive["text"].as_str())
        .collect();
    assert_eq!(
        text.split_whitespace().collect::<String>(),
        "officefifty",
        "{band} ask={ask} primitive kinds {kinds:?}"
    );
    let ligated = painted
        .iter()
        .filter(|primitive| {
            primitive["text"]
                .as_str()
                .is_some_and(|t| !t.trim().is_empty())
        })
        .map(|primitive| {
            primitive["fontFeatures"]
                .as_array()
                .is_some_and(|features| features.iter().any(|feature| feature == "liga"))
        })
        .collect();
    let glyphs = painted
        .iter()
        .all(|primitive| primitive["kind"] == "glyphRun")
        .then(|| {
            painted
                .iter()
                .map(|run| run["glyphs"].as_array().map_or(0, Vec::len))
                .sum()
        });
    BandPaint {
        text,
        ligated,
        glyphs,
    }
}

struct BandPaint {
    text: String,
    ligated: Vec<bool>,
    glyphs: Option<usize>,
}

#[test]
fn header_and_footer_runs_ligate_only_when_asked() {
    for band in ["header", "footer"] {
        let plain = band_glyphs(band, false);
        let asked = band_glyphs(band, true);
        assert!(
            plain.ligated.iter().all(|on| !on),
            "{band}: a run that does not ask paints without liga: {:?}",
            plain.ligated
        );
        assert!(
            !asked.ligated.is_empty() && asked.ligated.iter().all(|on| *on),
            "{band}: an asking run paints with liga: {:?}",
            asked.ligated
        );
        let plain_glyphs = plain
            .glyphs
            .unwrap_or_else(|| panic!("{band}: a band without a request paints glyph runs"));
        let asked_glyphs = asked
            .glyphs
            .unwrap_or_else(|| panic!("{band}: an asking band paints glyph runs"));
        assert_eq!(plain_glyphs, plain.text.chars().count(), "{band}");
        assert!(
            asked_glyphs < plain_glyphs,
            "{band}: {asked_glyphs} vs {plain_glyphs}"
        );
    }
}
