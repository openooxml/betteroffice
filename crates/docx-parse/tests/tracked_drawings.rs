use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml""#;

fn drawing(text: &str) -> String {
    format!(
        r#"<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="20" name="Text Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvPr id="20"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    )
}

fn alternate(text: &str) -> String {
    format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wps">{}</mc:Choice><mc:Fallback><w:pict><v:rect style="width:72pt;height:72pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent>"#,
        drawing(text)
    )
}

const DEL: &str = r#"<w:del w:id="3" w:author="Editor" w:date="2018-07-05T14:06:00Z">"#;

const INS: &str = r#"<w:ins w:id="2" w:author="Editor" w:date="2018-07-05T14:06:00Z">"#;

fn fixture(body: &str) -> Vec<u8> {
    fixture_with(NS, body)
}

fn fixture_with(namespaces: &str, body: &str) -> Vec<u8> {
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned().into_bytes()),
        ("_rels/.rels".to_owned(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned().into_bytes()),
        ("word/document.xml".to_owned(), format!("<w:document {namespaces}><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>{body}<w:sectPr/></w:body></w:document>").into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved(source: &[u8], edit: bool) -> String {
    let parsed = parse_docx_s9_wire(source, S9ParseOptions::default()).unwrap();
    let mut package: Value = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    if edit {
        package["document"]["content"][0]["content"][0]["content"][0]["text"] = json!("Edited");
    }
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

fn drawings(xml: &str) -> (usize, usize) {
    (
        xml.matches("<w:drawing>").count() + xml.matches("<w:drawing ").count(),
        xml.matches("<w:pict>").count() + xml.matches("<w:pict ").count(),
    )
}

#[test]
fn a_text_box_inside_a_tracked_insertion_survives_a_save() {
    for (name, run) in [
        ("alternate content", alternate("Boxed")),
        ("drawing", drawing("Boxed")),
    ] {
        let body = format!(
            r#"<w:p><w:r><w:t>Host</w:t></w:r>{INS}<w:r>{run}</w:r></w:ins><w:hyperlink w:anchor="Target"><w:r>{run}</w:r></w:hyperlink>{DEL}<w:r>{run}</w:r></w:del></w:p>"#
        );
        let source = fixture(&body);
        let source_xml = String::from_utf8(
            ooxml_opc::unzip_parts(&source)
                .unwrap()
                .into_iter()
                .find(|(path, _)| path == "word/document.xml")
                .unwrap()
                .1,
        )
        .unwrap();
        for edit in [false, true] {
            let xml = saved(&source, edit);
            assert_eq!(
                drawings(&xml),
                drawings(&source_xml),
                "{name} {edit}: {xml}"
            );
            assert_eq!(
                xml.matches(">Boxed<").count(),
                source_xml.matches(">Boxed<").count(),
                "{name} {edit}: {xml}"
            );
            let ins = xml
                .find("<w:ins ")
                .unwrap_or_else(|| panic!("{name} {edit}: {xml}"));
            let boxed = xml.find(">Boxed<").unwrap();
            let closed = xml[ins..].find("</w:ins>").unwrap() + ins;
            assert!(ins < boxed && boxed < closed, "{name} {edit}: {xml}");
            assert!(!xml.contains("<w:delText>Boxed"), "{name} {edit}: {xml}");
            assert_eq!(
                xml.matches("<w:t>Boxed</w:t>").count(),
                source_xml.matches("<w:t>Boxed</w:t>").count(),
                "{name} {edit}: {xml}"
            );
        }
    }
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
            let bound = prefix == "xml"
                || declared.contains(&prefix)
                || scopes.iter().any(|scope| scope.contains(&prefix));
            if !bound {
                unbound.push(prefix);
            }
        }
        if !empty {
            scopes.push(declared);
        }
    }
    unbound
}

#[test]
fn two_text_boxes_in_one_tracked_alternate_content_survive_a_save() {
    let pair = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wps">{}{}</mc:Choice><mc:Fallback><w:pict><v:rect style="width:72pt;height:72pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Pair</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent>"#,
        drawing("First"),
        drawing("Second")
    );
    let body = format!(
        r#"<w:p><w:r><w:t>Host</w:t></w:r>{INS}<w:r>{pair}</w:r></w:ins><w:hyperlink w:anchor="Target"><w:r>{pair}</w:r></w:hyperlink></w:p>"#
    );
    let source = fixture(&body);
    for edit in [false, true] {
        let xml = saved(&source, edit);
        for text in [">First<", ">Second<", ">Pair<"] {
            assert_eq!(xml.matches(text).count(), 2, "{text} {edit}: {xml}");
        }
        assert_eq!(drawings(&xml), (4, 2), "{edit}: {xml}");
    }
}

#[test]
fn a_tracked_text_box_keeps_namespaces_declared_on_its_run() {
    let a = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;
    let namespaces = NS.replace(&format!(" {a}"), "");
    for run in [alternate("Boxed"), drawing("Boxed")] {
        let body = format!(
            r#"<w:p><w:r><w:t>Host</w:t></w:r>{INS}<w:r {a}>{run}</w:r></w:ins><w:hyperlink w:anchor="Target"><w:r {a}>{run}</w:r></w:hyperlink></w:p>"#
        );
        let source = fixture_with(&namespaces, &body);
        for edit in [false, true] {
            let xml = saved(&source, edit);
            assert_eq!(
                xml.matches(">Boxed<").count(),
                body.matches(">Boxed<").count(),
                "{edit}: {xml}"
            );
            assert_eq!(
                unbound_prefixes(&xml),
                Vec::<String>::new(),
                "{edit}: {xml}"
            );
        }
    }
}
