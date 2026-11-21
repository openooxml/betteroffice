use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas""#;

fn fixture(body: &str) -> Vec<u8> {
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned().into_bytes()),
        ("_rels/.rels".to_owned(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned().into_bytes()),
        ("word/document.xml".to_owned(), format!("<w:document {NS}><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>{body}<w:sectPr/></w:body></w:document>").into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn package(source: &[u8]) -> Value {
    let parsed = parse_docx_s9_wire(source, S9ParseOptions::default()).unwrap();
    let mut package: Value = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    package
}

fn find_mut<'a>(value: &'a mut Value, matches: &dyn Fn(&Value) -> bool) -> Option<&'a mut Value> {
    if matches(value) {
        return Some(value);
    }
    match value {
        Value::Object(map) => map.values_mut().find_map(|child| find_mut(child, matches)),
        Value::Array(items) => items.iter_mut().find_map(|child| find_mut(child, matches)),
        _ => None,
    }
}

fn save(source: &[u8], package: &Value) -> String {
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    let saved = write_docx_s13(request, source).unwrap();
    String::from_utf8(
        ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .into_iter()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap()
            .1,
    )
    .unwrap()
}

fn group_with_bookmarked_box() -> String {
    r#"<w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="30" name="Composite"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></wpg:grpSpPr><wps:wsp><wps:cNvPr id="31" name="Box"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Inner</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp><wps:wsp><wps:cNvPr id="32" name="Mark"/><wps:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="ellipse"/></wps:spPr><wps:bodyPr/></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>"#.to_owned()
}

fn word_shaped_group(drawing: &str) -> String {
    let fallback = r#"<w:pict><v:group id="Composite" style="width:144pt;height:72pt" coordsize="1828800,914400"><v:rect style="width:72pt;height:72pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect><v:oval style="left:72pt;width:72pt;height:72pt"/></v:group></w:pict>"#;
    format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wpg">{drawing}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
    )
}

fn assert_group(xml: &str, word_shaped: bool) {
    assert!(xml.contains("<wpg:wgp>"), "{xml}");
    assert_eq!(xml.matches("<wps:wsp>").count(), 2, "{xml}");
    assert!(xml.contains(r#"<wps:cNvPr id="31" name="Box""#), "{xml}");
    assert!(xml.contains(r#"<wps:cNvPr id="32" name="Mark""#), "{xml}");
    if word_shaped {
        assert!(xml.contains("<mc:Fallback>"), "{xml}");
    }
}

fn assert_bookmark(xml: &str) {
    assert!(xml.contains("_GoBack"), "{xml}");
    assert!(xml.contains("<w:bookmarkStart "), "{xml}");
    assert!(xml.contains("<w:bookmarkEnd "), "{xml}");
}

fn check_group(word_shaped: bool) {
    for bookmarked in [true, false] {
        let mut drawing = group_with_bookmarked_box();
        if !bookmarked {
            drawing = drawing.replace(
                r#"<w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>"#,
                "",
            );
        }
        if word_shaped {
            drawing = word_shaped_group(&drawing);
        }
        let source = fixture(&format!(
            r#"<w:p><w:r><w:t>Host</w:t></w:r><w:r>{drawing}</w:r></w:p>"#
        ));
        let mut model = package(&source);
        let xml = save(&source, &model);
        assert_group(&xml, word_shaped);
        assert!(xml.contains(">Inner<"), "{xml}");
        if bookmarked {
            assert_bookmark(&xml);
        } else {
            assert!(!xml.contains("_GoBack"), "{xml}");
        }

        let host = &mut model["document"]["content"][1]["content"];
        let Some(text) = find_mut(host, &|value| value["text"] == "Inner") else {
            assert!(word_shaped, "text box text in the model");
            continue;
        };
        *text = json!({"type":"text","text":"Changed"});
        let xml = save(&source, &model);
        assert_group(&xml, word_shaped);
        assert!(xml.contains(">Changed<"), "{xml}");
        let drawing = xml.split("<mc:Fallback>").next().expect("document xml");
        assert!(!drawing.contains(">Inner<"), "{xml}");
        if bookmarked {
            assert_bookmark(&xml);
        }
    }
}

#[test]
fn a_bookmarked_text_box_in_a_group_saves() {
    check_group(false);
}

#[test]
fn a_bookmarked_text_box_in_a_word_shaped_group_saves() {
    check_group(true);
}
