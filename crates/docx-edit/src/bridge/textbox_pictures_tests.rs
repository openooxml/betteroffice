use serde_json::{Value, json};

use crate::EngineSession;

const FONT: &[u8] = include_bytes!("../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#;
const PICTURE: &str = r#"<w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="7" name="Picture 7" descr="Figure description"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId9"/><a:srcRect t="10000" r="20000" b="30000" l="40000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#;

fn document(content: &str, boxed: bool) -> Vec<u8> {
    let paragraph = format!("<w:p><w:r>{content}</w:r></w:p>");
    let body = if boxed {
        format!(
            r#"<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2743200" cy="1371600"/><wp:wrapNone/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>{paragraph}</w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>"#
        )
    } else {
        paragraph
    };
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/picture.png"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document {NAMESPACES}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#)),
    ];
    let mut parts: Vec<_> = parts
        .into_iter()
        .map(|(name, value)| (name.to_owned(), value.into_bytes()))
        .collect();
    parts.push((
        "word/media/picture.png".to_owned(),
        vec![
            137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1,
            8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 11, 73, 68, 65, 84, 120, 156, 99, 96, 0, 2,
            0, 0, 5, 0, 1, 169, 245, 232, 39, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
        ],
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn primitives(content: &str, boxed: bool, media_tokens: bool) -> Vec<Value> {
    let bytes = document(content, boxed);
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75301);
    crate::seed::seed_with_layout_tokens(engine.doc(), &bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "renderEnv": {"mediaTokens": media_tokens},
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

fn images(primitives: &[Value]) -> Vec<&Value> {
    primitives
        .iter()
        .filter(|primitive| primitive["kind"] == "image")
        .collect()
}

fn text(primitives: &[Value]) -> String {
    primitives
        .iter()
        .filter(|primitive| matches!(primitive["kind"].as_str(), Some("text" | "glyphRun")))
        .filter_map(|primitive| primitive["text"].as_str())
        .collect()
}

#[test]
fn inline_picture_in_text_box_paints_with_body_image_properties() {
    for media_tokens in [true, false] {
        let boxed = primitives(PICTURE, true, media_tokens);
        let body = primitives(PICTURE, false, media_tokens);
        let boxed_images = images(&boxed);
        let body_images = images(&body);
        assert_eq!(boxed_images.len(), 1);
        assert_eq!(body_images.len(), 1);
        let image = boxed_images[0];
        for key in ["relId", "w", "h", "crop", "altText"] {
            assert_eq!(image[key], body_images[0][key], "{key}");
        }
        if media_tokens {
            assert_eq!(image["relId"], "media:0");
        } else {
            assert!(
                image["relId"]
                    .as_str()
                    .unwrap()
                    .starts_with("data:image/png;base64,")
            );
        }
        assert_eq!(image["w"], 192);
        assert_eq!(image["h"], 96);
        assert_eq!(image["altText"], "Figure description");
        assert_eq!(
            image["crop"],
            json!({"top": 0.1, "right": 0.2, "bottom": 0.3, "left": 0.4})
        );
        let shape = boxed
            .iter()
            .find(|primitive| primitive["kind"] == "shape")
            .unwrap();
        for (position, size) in [("x", "w"), ("y", "h")] {
            let start = image[position].as_f64().unwrap();
            let end = start + image[size].as_f64().unwrap();
            let box_start = shape[position].as_f64().unwrap();
            let box_end = box_start + shape[size].as_f64().unwrap();
            assert!(start >= box_start && end <= box_end);
        }
    }
}

#[test]
fn text_around_a_text_box_picture_still_paints() {
    let display = primitives(
        &format!("<w:t>Before</w:t>{PICTURE}<w:t>After</w:t>"),
        true,
        true,
    );
    assert_eq!(images(&display).len(), 1);
    assert_eq!(text(&display), "BeforeAfter");
}

#[test]
fn text_only_text_box_still_paints() {
    let display = primitives("<w:t>Plain text</w:t>", true, true);
    assert!(images(&display).is_empty());
    assert_eq!(text(&display), "Plain text");
}
