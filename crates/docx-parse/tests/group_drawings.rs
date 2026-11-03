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

fn group(kind: &str) -> String {
    let (uri, open, close) = match kind {
        "group" => (
            "wordprocessingGroup",
            r#"<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></wpg:grpSpPr>"#,
            "</wpg:wgp>",
        ),
        _ => (
            "wordprocessingCanvas",
            r#"<wpc:wpc><wpc:bg/><wpc:whole/>"#,
            "</wpc:wpc>",
        ),
    };
    format!(
        r#"<w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="30" name="Composite"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/{uri}">{open}<wps:wsp><wps:cNvPr id="31" name="Box"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp><wps:wsp><wps:cNvPr id="32" name="Mark"/><wps:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="ellipse"/></wps:spPr><wps:bodyPr/></wps:wsp>{close}</a:graphicData></a:graphic></wp:inline></w:drawing>"#
    )
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

#[test]
fn groups_and_canvases_in_a_rewritten_paragraph_are_kept() {
    for kind in ["group", "canvas"] {
        let uri = if kind == "group" {
            "wordprocessingGroup"
        } else {
            "wordprocessingCanvas"
        };
        let fallback = r#"<w:pict><v:rect style="width:144pt;height:72pt"/></w:pict>"#;
        let body = format!(
            r#"<w:p><w:bookmarkStart w:id="0" w:name="Lead"/><w:r><w:t>Host</w:t></w:r><w:bookmarkEnd w:id="0"/><w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wpg">{}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent></w:r></w:p>"#,
            group(kind)
        );
        let source = fixture(&body);
        for change in ["none", "host", "text", "frame"] {
            let mut model = package(&source);
            let host = &mut model["document"]["content"][1]["content"];
            match change {
                "host" => host
                    .as_array_mut()
                    .unwrap()
                    .push(json!({"type":"run","content":[{"type":"text","text":"Added"}]})),
                "text" => {
                    *find_mut(host, &|value| value["text"] == "Inner").unwrap() =
                        json!({"type":"text","text":"Changed"});
                    host.as_array_mut()
                        .unwrap()
                        .push(json!({"type":"run","content":[{"type":"text","text":"Added"}]}));
                }
                "frame" => {
                    let shape = find_mut(host, &|value| {
                        value["children"].is_array() && value["size"].is_object()
                    })
                    .unwrap();
                    shape["size"]["width"] = json!(2743200.0);
                }
                _ => {}
            }
            let xml = save(&source, &model);
            assert_eq!(
                xml.matches(&format!(
                    r#"uri="http://schemas.microsoft.com/office/word/2010/{uri}""#
                ))
                .count(),
                1,
                "{kind} {change}: {xml}"
            );
            assert_eq!(
                xml.matches("<wps:wsp>").count(),
                2,
                "{kind} {change}: {xml}"
            );
            let text = if change == "text" {
                ">Changed<"
            } else {
                ">Inner<"
            };
            assert_eq!(xml.matches(text).count(), 1, "{kind} {change}: {xml}");
            if change == "frame" {
                assert!(xml.contains(r#"cx="2743200""#), "{kind} {change}: {xml}");
            }
        }
    }
}

#[test]
fn an_untouched_group_in_a_rewritten_paragraph_is_written_byte_for_byte() {
    for kind in ["group", "canvas"] {
        let drawing = group(kind);
        let source = fixture(&format!(
            r#"<w:p><w:r><w:t>Host</w:t></w:r><w:r>{drawing}</w:r></w:p>"#
        ));
        let mut model = package(&source);
        model["document"]["content"][1]["content"]
            .as_array_mut()
            .unwrap()
            .push(json!({"type":"run","content":[{"type":"text","text":"Added"}]}));
        let xml = save(&source, &model);
        assert!(xml.contains(">Added<"), "{kind}: {xml}");
        assert!(xml.contains(&drawing), "{kind}: {xml}");
    }
}

#[test]
fn a_group_whose_child_changed_is_written_from_the_model() {
    for kind in ["group", "canvas"] {
        let drawing = group(kind);
        let source = fixture(&format!(
            r#"<w:p><w:r><w:t>Host</w:t></w:r><w:r>{drawing}</w:r></w:p>"#
        ));
        let mut model = package(&source);
        let host = &mut model["document"]["content"][1]["content"];
        let mark = find_mut(host, &|value| value["name"] == "Mark").unwrap();
        mark["fill"] = json!({"type":"solid","color":{"rgb":"FF0000"}});
        let xml = save(&source, &model);
        assert!(!xml.contains(&drawing), "{kind}: {xml}");
        assert!(xml.contains(">Host<"), "{kind}: {xml}");
    }
}

#[test]
fn a_text_edit_in_a_group_never_lands_in_another_text_box() {
    let text_box = |id: u32, text: &str| {
        format!(
            r#"<wps:wsp><wps:cNvPr id="{id}" name="Box {id}"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>"#
        )
    };
    let plain = r#"<wps:wsp><wps:cNvPr id="40" name="Plain"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:bodyPr/></wps:wsp>"#;
    let nested = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wps">{}</mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
        text_box(42, "Second")
    );
    let drawing = format!(
        r#"<w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="39" name="Group"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></wpg:grpSpPr>{plain}{}{nested}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        text_box(41, "First")
    );
    let source = fixture(&format!(
        r#"<w:p><w:r><w:t>Host</w:t></w:r><w:r>{drawing}</w:r></w:p>"#
    ));
    let mut model = package(&source);
    let host = &mut model["document"]["content"][1]["content"];
    *find_mut(host, &|value| value["text"] == "First").unwrap() =
        json!({"type":"text","text":"Changed"});
    host.as_array_mut()
        .unwrap()
        .push(json!({"type":"run","content":[{"type":"text","text":"Added"}]}));
    let xml = save(&source, &model);
    assert!(!xml.contains(">First<"), "{xml}");
    assert!(xml.matches(">Second<").count() <= 1, "{xml}");
}

fn unbound_prefixes(xml: &str) -> Vec<String> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut scopes: Vec<Vec<String>> = Vec::new();
    let mut unbound = Vec::new();
    loop {
        let (start, empty) = match reader.read_event().unwrap() {
            Event::Start(start) => (start, false),
            Event::Empty(start) => (start, true),
            Event::End(_) => {
                scopes.pop();
                continue;
            }
            Event::Eof => break,
            _ => continue,
        };
        let mut declared = Vec::new();
        let mut used = Vec::new();
        let name = String::from_utf8(start.name().as_ref().to_vec()).unwrap();
        used.extend(name.split_once(':').map(|(prefix, _)| prefix.to_owned()));
        for attribute in start.attributes() {
            let key = String::from_utf8(attribute.unwrap().key.as_ref().to_vec()).unwrap();
            if let Some(prefix) = key.strip_prefix("xmlns:") {
                declared.push(prefix.to_owned());
            } else if let Some((prefix, _)) = key.split_once(':') {
                used.push(prefix.to_owned());
            }
        }
        for prefix in used {
            if prefix != "xml"
                && !declared.contains(&prefix)
                && !scopes.iter().any(|scope| scope.contains(&prefix))
            {
                unbound.push(prefix);
            }
        }
        if !empty {
            scopes.push(declared);
        }
    }
    unbound
}

fn rewrite_host(source: &[u8]) -> String {
    let mut model = package(source);
    model["document"]["content"][1]["content"]
        .as_array_mut()
        .unwrap()
        .push(json!({"type":"run","content":[{"type":"text","text":"Added"}]}));
    save(source, &model)
}

#[test]
fn a_replayed_group_keeps_prefixes_declared_on_its_wrapper() {
    let a = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;
    let drawing = group("group");
    let source = fixture(&format!(
        r#"<w:p><w:r><w:t>Host</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="wpg" {a}>{drawing}</mc:Choice><mc:Fallback/></mc:AlternateContent></w:r></w:p>"#
    ));
    let source = {
        let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
        let (_, bytes) = parts
            .iter_mut()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap();
        *bytes = String::from_utf8(bytes.clone())
            .unwrap()
            .replacen(&format!(" {a}"), "", 1)
            .into_bytes();
        ooxml_opc::rezip_parts(&parts).unwrap()
    };
    let xml = rewrite_host(&source);
    assert!(
        xml.contains(r#"uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup""#),
        "{xml}"
    );
    assert_eq!(unbound_prefixes(&xml), Vec::<String>::new(), "{xml}");
}

#[test]
fn a_group_with_single_quoted_attributes_is_kept() {
    let drawing = group("group").replace('"', "'");
    let source = fixture(&format!(
        r#"<w:p><w:r><w:t>Host</w:t></w:r><w:r>{drawing}</w:r></w:p>"#
    ));
    let xml = rewrite_host(&source);
    assert!(xml.contains(&drawing), "{xml}");
}

#[test]
fn an_edit_in_a_table_only_text_box_is_never_reverted() {
    let drawing = group("group").replace(
        "<w:txbxContent><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:txbxContent>",
        r#"<w:txbxContent><w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:txbxContent>"#,
    );
    let source = fixture(&format!(
        r#"<w:p><w:r><w:t>Host</w:t></w:r><w:r>{drawing}</w:r></w:p>"#
    ));
    let mut model = package(&source);
    let host = &mut model["document"]["content"][1]["content"];
    *find_mut(host, &|value| value["text"] == "Inner").unwrap() =
        json!({"type":"text","text":"Changed"});
    host.as_array_mut()
        .unwrap()
        .push(json!({"type":"run","content":[{"type":"text","text":"Added"}]}));
    let xml = save(&source, &model);
    assert!(!xml.contains(">Inner<"), "{xml}");
}
