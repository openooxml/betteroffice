use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{
    HexIdAllocator, InlineNode, ParagraphContent, ParseBudget, ParseLimits, RunContent,
    S9ParseOptions, XmlDocument, XmlElement, XmlNode, parse_docx_s9_wire, parse_paragraph,
    parse_xml,
};
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

fn fixture(body: &str, namespaces: &str) -> Vec<u8> {
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned().into_bytes()),
        ("_rels/.rels".to_owned(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned().into_bytes()),
        ("word/document.xml".to_owned(), format!("<w:document {namespaces}><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>{body}<w:sectPr/></w:body></w:document>").into_bytes()),
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

fn saved(source: &[u8], edit: bool) -> String {
    let mut package = package(source);
    if edit {
        package["document"]["content"][0]["content"][0]["content"][0]["text"] = json!("Edited");
    }
    save(source, &package)
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

#[test]
fn a_text_box_after_bookmarks_and_other_markers_stays_on_its_run() {
    let markers = r#"<w:bookmarkStart w:id="0" w:name="One"/><w:bookmarkStart w:id="1" w:name="Two"/><w:bookmarkEnd w:id="0"/><w:commentRangeStart w:id="5"/>"#;
    for (name, run) in [
        ("alternate content", alternate("Boxed")),
        ("drawing", drawing("Boxed")),
    ] {
        let body = format!(
            r#"<w:p>{markers}<w:r><w:t>Lead</w:t></w:r><w:bookmarkEnd w:id="1"/><w:r><w:br/></w:r><w:r>{run}</w:r><w:r><w:t>Tail</w:t></w:r></w:p>"#
        );
        let source = fixture(&body, NS);
        for edit in [false, true] {
            let xml = saved(&source, edit);
            assert_eq!(
                xml.matches("<w:drawing>").count(),
                1,
                "{name} {edit}: {xml}"
            );
            assert!(xml.contains(">Boxed<"), "{name} {edit}: {xml}");
            let boxed = xml.find(">Boxed<").unwrap();
            assert!(
                xml.find("<w:br/>").unwrap() < boxed && boxed < xml.find(">Tail<").unwrap(),
                "{name} {edit}: {xml}"
            );
        }
    }
}

fn read_xml(xml: &str) -> XmlDocument {
    parse_xml(
        xml.as_bytes(),
        "word/document.xml",
        &mut ParseBudget::new(&ParseLimits::default()),
    )
    .unwrap()
}

#[test]
fn public_paragraph_parser_keeps_bookmarks_added_to_a_text_box() {
    let start = "<w:bookmarkStart w:id='7' w:name='Bo&#120;'/>";
    let end = "<w:bookmarkEnd w:id='7'/>";
    let outer_paragraph = |before, after| {
        format!(
            r#"<w:p {NS}><w:r><w:pict><v:shape style="width:72pt;height:72pt"><v:textbox><w:txbxContent><w:p>{before}<w:r><w:t>Boxed</w:t></w:r>{after}</w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"#
        )
    };
    let mut edited = read_xml(&outer_paragraph("", "")).roots.pop().unwrap();
    let parse = |element: &XmlElement| {
        parse_paragraph(
            element,
            None,
            None,
            None,
            None,
            None,
            "word/document.xml",
            &mut ParseBudget::new(&ParseLimits::default()),
            &mut HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap(),
            None,
            false,
            0,
        )
        .unwrap()
    };
    let original = parse(&edited);
    let mut boxed = &mut edited;
    for name in [
        "w:r",
        "w:pict",
        "v:shape",
        "v:textbox",
        "w:txbxContent",
        "w:p",
    ] {
        boxed = boxed
            .children
            .iter_mut()
            .find_map(|child| match child {
                XmlNode::Element(element) if element.name == name => Some(element),
                _ => None,
            })
            .unwrap();
    }
    let mut boundaries = read_xml(&format!("<w:p {NS}>{start}{end}</w:p>"))
        .roots
        .pop()
        .unwrap()
        .children
        .into_iter();
    boxed.children.insert(0, boundaries.next().unwrap());
    boxed.children.push(boundaries.next().unwrap());

    let authored = read_xml(&outer_paragraph(start, end));
    let parsed = parse(&edited);
    assert_eq!(parsed, parse(authored.root().unwrap()));
    assert_ne!(parsed, original);
    let [ParagraphContent::Inline(InlineNode::Run(run))] = parsed.content.as_slice() else {
        panic!("expected a run containing the text box");
    };
    let [RunContent::OpaqueDrawing { kind, xml, .. }] = run.content.as_slice() else {
        panic!("expected an opaque VML text box");
    };
    assert_eq!(kind, "pict");
    for fragment in ["<v:textbox>", start, ">Boxed<", end] {
        assert!(xml.contains(fragment), "{xml}");
    }
}

#[test]
fn bookmarked_drawing_keeps_its_inherited_default_namespace_on_save() {
    use quick_xml::NsReader;
    use quick_xml::events::Event;
    use quick_xml::name::{Namespace, ResolveResult};

    let extension_namespace = "urn:box-extension";
    let boxed = drawing("Boxed")
        .replace("<w:p>", r#"<w:p><w:bookmarkStart w:id="7" w:name="Box"/>"#)
        .replace("</w:p>", r#"<w:bookmarkEnd w:id="7"/></w:p>"#)
        .replace(
            "</wps:spPr>",
            r#"<a:extLst><a:ext uri="urn:box-extension"><boxExtension/></a:ext></a:extLst></wps:spPr>"#,
        );
    let source = fixture(
        &format!("<w:p><w:r>{boxed}</w:r></w:p>"),
        &format!(r#"{NS} xmlns="{extension_namespace}""#),
    );
    let model = package(&source);
    let drawing_xml = model
        .pointer("/document/content/1/content/0/content/0/shape/drawingXml")
        .and_then(Value::as_str)
        .unwrap();
    assert!(drawing_xml.contains(&format!(r#"xmlns="{extension_namespace}""#)));

    let xml = save(&source, &model);
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut extensions = 0;
    loop {
        match reader.read_resolved_event().unwrap() {
            (namespace, Event::Start(element) | Event::Empty(element))
                if element.name().as_ref() == b"boxExtension" =>
            {
                assert_eq!(
                    namespace,
                    ResolveResult::Bound(Namespace(extension_namespace.as_bytes())),
                    "{xml}"
                );
                extensions += 1;
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }
    assert_eq!(extensions, 1, "{xml}");
}
