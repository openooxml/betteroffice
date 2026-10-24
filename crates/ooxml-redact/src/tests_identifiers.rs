use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};

use crate::mask::TextMasker;
use crate::styles::StyleMap;
use crate::{Format, RedactionOptions, RedactionReport, redact_with_options, xml};

const W: &str = r##"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""##;
const KEPT_BOOKMARKS: &[&str] = &["_Toc123456", "_GoBack", "_Ref99887", "_Hlk4455"];

fn docx(document_body: &str, extra: &[(&str, String)]) -> Vec<u8> {
    let mut overrides = String::new();
    let mut relationships = String::new();
    for (index, (path, _)) in extra.iter().enumerate() {
        let (content_type, relationship) = if path.contains("header") {
            ("wordprocessingml.header", "header")
        } else if path.contains("footnotes") {
            ("wordprocessingml.footnotes", "footnotes")
        } else {
            ("theme", "theme")
        };
        overrides.push_str(&format!(r##"<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-officedocument.{content_type}+xml"/>"##));
        relationships.push_str(&format!(r##"<Relationship Id="rIdX{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{relationship}" Target="{}"/>"##, path.trim_start_matches("word/")));
    }
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), format!(r##"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{overrides}</Types>"##)),
        ("_rels/.rels".to_owned(), r##"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"##.to_owned()),
        ("word/_rels/document.xml.rels".to_owned(), format!(r##"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"##)),
        ("word/document.xml".to_owned(), format!(r##"<w:document {W}><w:body>{document_body}<w:sectPr/></w:body></w:document>"##)),
    ];
    parts.extend(
        extra
            .iter()
            .map(|(path, xml)| ((*path).to_owned(), xml.clone())),
    );
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(path, xml)| (path, xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn redacted_docx(source: &[u8], random_characters: bool) -> Vec<(String, String)> {
    let output = redact_with_options(
        source,
        Format::Docx,
        &RedactionOptions {
            random_characters,
            ..RedactionOptions::default()
        },
    )
    .unwrap();
    parse_docx_s9_wire(&output, S9ParseOptions::default()).unwrap();
    ooxml_opc::unzip_parts(&output)
        .unwrap()
        .into_iter()
        .filter(|(path, _)| path.ends_with(".xml"))
        .map(|(path, bytes)| (path, String::from_utf8(bytes).unwrap()))
        .collect()
}

fn redact_parts(format: Format, parts: &[(&str, &str)], random_characters: bool) -> Vec<String> {
    let mut masker = TextMasker::new(&RedactionOptions {
        random_characters,
        ..RedactionOptions::default()
    });
    parts
        .iter()
        .map(|(path, source)| {
            let output = xml::redact_xml_with_styles(
                format,
                path,
                source.as_bytes(),
                &mut RedactionReport::default(),
                &StyleMap::default(),
                &mut masker,
            )
            .unwrap();
            String::from_utf8(output).unwrap()
        })
        .collect()
}

fn attributes(xml: &str, element: &str, attribute: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let decoder = reader.decoder();
    let mut values = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(start) | Event::Empty(start)
                if start.name().as_ref() == element.as_bytes() =>
            {
                for item in start.attributes() {
                    let item = item.unwrap();
                    if item.key.as_ref() == attribute.as_bytes() {
                        values.push(
                            item.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                                .unwrap()
                                .into_owned(),
                        );
                    }
                }
            }
            Event::Eof => return values,
            _ => {}
        }
    }
}

fn instructions(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let decoder = reader.decoder();
    let mut open: Vec<Option<String>> = Vec::new();
    let mut found = attributes(xml, "w:fldSimple", "w:instr");
    let mut in_instruction = false;
    loop {
        match reader.read_event().unwrap() {
            Event::Start(start) | Event::Empty(start) if start.name().as_ref() == b"w:fldChar" => {
                let kind = start
                    .try_get_attribute("w:fldCharType")
                    .unwrap()
                    .unwrap()
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .unwrap()
                    .into_owned();
                match kind.as_str() {
                    "begin" => open.push(Some(String::new())),
                    _ => {
                        if let Some(text) = open.last_mut().and_then(Option::take) {
                            found.push(text);
                        }
                        if kind == "end" {
                            open.pop();
                        }
                    }
                }
            }
            Event::Start(start) if start.name().as_ref() == b"w:instrText" => in_instruction = true,
            Event::End(end) if end.name().as_ref() == b"w:instrText" => in_instruction = false,
            Event::Text(text) if in_instruction => {
                if let Some(Some(field)) = open.last_mut() {
                    field.push_str(&text.decode().unwrap());
                }
            }
            Event::GeneralRef(reference) if in_instruction => {
                if let Some(Some(field)) = open.last_mut() {
                    let resolved = match reference.decode().unwrap().as_ref() {
                        "quot" => "\"",
                        "amp" => "&",
                        "apos" => "'",
                        "lt" => "<",
                        "gt" => ">",
                        other => panic!("unexpected entity {other}"),
                    };
                    field.push_str(resolved);
                }
            }
            Event::Eof => return found,
            _ => {}
        }
    }
}

fn bookmark_target(instruction: &str) -> Option<String> {
    let tokens: Vec<&str> = instruction.split_whitespace().collect();
    let keyword = tokens.first()?.to_ascii_uppercase();
    let target = match keyword.as_str() {
        "REF" | "PAGEREF" | "NOTEREF" | "SET" | "ASK" | "GOTOBUTTON" => tokens.get(1)?,
        "HYPERLINK" => tokens.get(tokens.iter().position(|token| *token == "\\l")? + 1)?,
        _ => return None,
    };
    Some(target.trim_matches('"').to_owned())
}

fn assert_attached_ref_switch(switch: &str) {
    let source = docx(
        &format!(
            r#"<w:p><w:bookmarkStart w:id="1" w:name="Secret"/><w:bookmarkEnd w:id="1"/>
<w:fldSimple w:instr=" REF Secret{switch} "><w:r><w:t>Result</w:t></w:r></w:fldSimple>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> REF Se</w:instrText></w:r>
<w:r><w:instrText xml:space="preserve">cret{switch} </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#
        ),
        &[],
    );
    for random in [false, true] {
        let parts = redacted_docx(&source, random);
        let document = &parts
            .iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1;
        let names = attributes(document, "w:bookmarkStart", "w:name");
        let fields = instructions(document);
        assert_eq!(fields.len(), 2);
        for field in fields {
            let parsed = docx_parse::parse_field_instruction(&field);
            let target = parsed.argument.unwrap();
            assert!(
                names.contains(&target),
                "REF target {target} not in {names:?}"
            );
            assert_eq!(parsed.switches.len(), 1, "lost switch in {field}");
            assert_eq!(format!("\\{}", parsed.switches[0].switch_name), switch);
        }
    }
}

#[test]
fn attached_ref_h_switch_preserves_bookmark_targets() {
    assert_attached_ref_switch(r"\h");
}

#[test]
fn attached_ref_p_switch_preserves_bookmark_targets() {
    assert_attached_ref_switch(r"\p");
}

#[test]
fn attached_hyperlink_l_switch_preserves_bookmark_targets() {
    let source = docx(
        r#"<w:p><w:bookmarkStart w:id="1" w:name="Secret"/><w:bookmarkEnd w:id="1"/>
<w:fldSimple w:instr=" HYPERLINK \l&quot;Secret&quot; "><w:r><w:t>Result</w:t></w:r></w:fldSimple>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> HYPERLINK \l"Secret" </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        &[],
    );
    for random in [false, true] {
        let parts = redacted_docx(&source, random);
        let document = &parts
            .iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1;
        let names = attributes(document, "w:bookmarkStart", "w:name");
        let fields = instructions(document);
        assert_eq!(fields.len(), 2);
        for field in fields {
            let parsed = docx_parse::parse_field_instruction(&field);
            assert_eq!(parsed.switches.len(), 1, "lost hyperlink switch in {field}");
            assert_eq!(parsed.switches[0].switch_name, "l");
            let target = parsed.switches[0].value.as_ref().unwrap();
            assert!(
                names.contains(target),
                "HYPERLINK target {target} not in {names:?}"
            );
        }
    }
}

#[test]
fn nested_set_ref_preserves_instruction_slots_and_bookmark_targets() {
    let source = docx(
        r#"<w:p><w:bookmarkStart w:id="1" w:name="Outer"/><w:bookmarkEnd w:id="1"/>
<w:bookmarkStart w:id="2" w:name="Other"/><w:bookmarkEnd w:id="2"/>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> SET Ou</w:instrText></w:r>
<w:r><w:instrText xml:space="preserve">ter "prefix </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> REF O</w:instrText></w:r>
<w:r><w:instrText xml:space="preserve">ther </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>Result</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
<w:r><w:instrText xml:space="preserve"> suffix" </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>Result</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
<w:fldSimple w:instr=" REF Outer "><w:r><w:t>Result</w:t></w:r></w:fldSimple></w:p>"#,
        &[],
    );
    for random in [false, true] {
        let parts = redacted_docx(&source, random);
        let document = &parts
            .iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1;
        let mut reader = Reader::from_str(document);
        let mut slots = Vec::new();
        loop {
            match reader.read_event().unwrap() {
                Event::Start(start) if start.name().as_ref() == b"w:instrText" => {
                    let text = reader.read_text(start.name()).unwrap();
                    let text = text.decode().unwrap();
                    slots.push(quick_xml::escape::unescape(&text).unwrap().into_owned());
                }
                Event::Eof => break,
                _ => {}
            }
        }
        assert_eq!(slots.len(), 5);
        assert_eq!(
            slots[0].matches('"').count(),
            0,
            "SET value moved before its original slot: {slots:?}"
        );
        assert_eq!(
            slots[1].matches('"').count(),
            1,
            "opening quote moved across slots: {slots:?}"
        );
        assert_eq!(
            slots[4].matches('"').count(),
            1,
            "closing quote moved across the nested REF: {slots:?}"
        );
        assert!(
            slots[1].starts_with(" \""),
            "prefix moved across the nested REF: {slots:?}"
        );
        assert!(slots[1].ends_with(' '));
        assert!(slots[4].starts_with(' '));
        assert!(slots[4].ends_with("\" "));
        assert_eq!(slots[1].chars().count(), " \"prefix ".chars().count());
        assert_eq!(slots[4].chars().count(), " suffix\" ".chars().count());
        let names = attributes(document, "w:bookmarkStart", "w:name");
        let fields = instructions(document);
        assert_eq!(fields.len(), 3);
        for field in fields {
            let target = bookmark_target(&field).unwrap();
            assert!(
                names.contains(&target),
                "REF/SET target {target} not in {names:?}"
            );
        }
    }
}

#[test]
fn entity_encoded_field_markers_preserve_split_ref_targets() {
    let source = docx(
        r##"<w:p><w:bookmarkStart w:id="1" w:name="Secret"/><w:bookmarkEnd w:id="1"/>
<w:r><w:fldChar w:fldCharType="beg&#105;n"/></w:r>
<w:r><w:instrText xml:space="preserve"> REF Se</w:instrText></w:r>
<w:r><w:instrText xml:space="preserve">cret </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="sep&#97;rate"/></w:r>
<w:r><w:t>Result</w:t></w:r><w:r><w:fldChar w:fldCharType="&#101;nd"/></w:r></w:p>"##,
        &[],
    );
    for random in [false, true] {
        let parts = redacted_docx(&source, random);
        let document = &parts
            .iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1;
        let names = attributes(document, "w:bookmarkStart", "w:name");
        let fields = instructions(document);
        assert_eq!(fields.len(), 1);
        let target = bookmark_target(&fields[0]).unwrap();
        assert!(
            names.contains(&target),
            "REF target {target} not in {names:?}"
        );
    }
}

#[test]
fn entity_encoded_identifier_definitions_and_references_resolve() {
    let source = format!(
        r##"<w:document {W}><w:body><w:p>
<w:bookmarkStart w:id="1" w:name="Secr&#101;t"/><w:bookmarkEnd w:id="1"/>
<w:hyperlink w:anchor="S&#101;cret"/><w:fldSimple w:instr="REF Secre&#116;"/>
<v:shapetype id="Ty&#112;e"/><v:shape id="Sha&#112;e" o:spid="&#83;hape" type="&#35;Type"/>
<o:OLEObject ShapeID="Shap&#101;"/><o:r idref="&#35;Shape"/>
<w:control w:shapeid="&#83;hape"/><v:textbox style="mso-next-textbox:&#35;Shape"/>
<a:theme name="Offic&#101; Theme"><a:clrScheme name="&#79;ffice"/><a:fontScheme name="Custo&#109;"/><a:fmtScheme name="Custom"/></a:theme>
</w:p></w:body></w:document>"##
    );
    for random in [false, true] {
        let output = redact_parts(Format::Docx, &[("word/document.xml", &source)], random);
        let document = &output[0];
        let names = attributes(document, "w:bookmarkStart", "w:name");
        assert_eq!(attributes(document, "w:hyperlink", "w:anchor"), names);
        assert_eq!(
            bookmark_target(&instructions(document)[0]),
            Some(names[0].clone())
        );
        assert_eq!(
            attributes(document, "v:shape", "id"),
            attributes(document, "v:shape", "o:spid")
        );
        assert_vml_references_resolve(document);
        assert_eq!(attributes(document, "a:theme", "name"), ["Office Theme"]);
        assert_eq!(attributes(document, "a:clrScheme", "name"), ["Office"]);
        assert_eq!(attributes(document, "a:fontScheme", "name"), ["Fonts"]);
        assert_eq!(attributes(document, "a:fmtScheme", "name"), ["Effects"]);
    }
}

#[test]
fn undecodable_identifier_attributes_return_typed_xml_errors() {
    for node in [
        r#"<w:fldChar w:fldCharType="beg&missing;n"/>"#,
        r#"<w:bookmarkStart w:name="Se&missing;cret"/>"#,
        r#"<w:hyperlink w:anchor="Se&missing;cret"/>"#,
        r#"<w:fldSimple w:instr="REF Se&missing;cret"/>"#,
        r#"<v:shape id="Sh&missing;ape"/>"#,
        r#"<v:shape o:spid="Sh&missing;ape"/>"#,
        r##"<v:shape type="#Ty&missing;pe"/>"##,
        r#"<o:OLEObject ShapeID="Sh&missing;ape"/>"#,
        r#"<o:r idref="Sh&missing;ape"/>"#,
        r#"<v:textbox style="mso-next-textbox:&missing;"/>"#,
        r#"<a:theme name="Off&missing;ice"/>"#,
    ] {
        let source = format!(r#"<w:document {W}>{node}</w:document>"#);
        let error = xml::redact_xml_with_styles(
            Format::Docx,
            "word/document.xml",
            source.as_bytes(),
            &mut RedactionReport::default(),
            &StyleMap::default(),
            &mut TextMasker::new(&RedactionOptions::default()),
        )
        .unwrap_err();
        assert!(
            matches!(error, crate::RedactError::Xml { part, .. } if part == "word/document.xml")
        );
    }
}

#[test]
fn entity_encoded_namespace_bindings_still_resolve_identifier_references() {
    let source = format!(
        r##"<w:document {W} xmlns:b="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&#110;"><w:body><w:p>
<w:bookmarkStart w:id="1" w:name="Secret"/><w:bookmarkEnd w:id="1"/>
<w:hyperlink b:anchor="Secret"/>
</w:p></w:body></w:document>"##
    );
    for random in [false, true] {
        let output = redact_parts(Format::Docx, &[("word/document.xml", &source)], random);
        let names = attributes(&output[0], "w:bookmarkStart", "w:name");
        assert!(!names.contains(&"Secret".to_owned()), "{names:?}");
        assert_eq!(attributes(&output[0], "w:hyperlink", "b:anchor"), names);
    }
}

#[test]
fn ignorable_field_marker_attributes_do_not_override_word_markers() {
    let source = format!(
        r##"<w:document {W} xmlns:x="urn:ignored" mc:Ignorable="x"><w:body><w:p>
<w:bookmarkStart w:id="1" w:name="Secret"/><w:bookmarkEnd w:id="1"/>
<w:r><w:fldChar x:fldCharType="end" w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> REF Se</w:instrText></w:r>
<w:r><w:instrText xml:space="preserve">cret </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p></w:body></w:document>"##
    );
    for random in [false, true] {
        let output = redact_parts(Format::Docx, &[("word/document.xml", &source)], random);
        let names = attributes(&output[0], "w:bookmarkStart", "w:name");
        let fields = instructions(&output[0]);
        assert_eq!(fields.len(), 1);
        let target = bookmark_target(&fields[0]).unwrap();
        assert!(
            names.contains(&target),
            "REF target {target} not in {names:?}"
        );
    }
}

#[test]
fn custom_xml_vml_attributes_keep_custom_xml_masking() {
    let source = r##"<v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="font-family:SecretCustomer"/>"##;
    for format in [Format::Docx, Format::Xlsx, Format::Pptx] {
        for random in [false, true] {
            let output = redact_parts(format, &[("customxml/item1.xml", source)], random);
            assert!(!output[0].contains("SecretCustomer"), "{}", output[0]);
        }
    }
}

fn assert_no_prose(parts: &[(String, String)]) {
    for (path, xml) in parts {
        assert!(
            !xml.contains("Zebra") && !xml.contains("ZEBRA"),
            "{path} keeps a user identifier: {xml}"
        );
    }
}

fn bookmark_fixture() -> Vec<u8> {
    let body = [
        r##"<w:p><w:bookmarkStart w:id="1" w:name="ZebraPlan"/><w:r><w:t>Plan</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>"##,
        r##"<w:p><w:bookmarkStart w:id="2" w:name="_TocZebraHeading"/><w:bookmarkStart w:id="3" w:name="_Toc123456"/><w:bookmarkStart w:id="4" w:name="_GoBack"/><w:bookmarkStart w:id="5" w:name="_Ref99887"/><w:bookmarkStart w:id="6" w:name="_Hlk4455"/><w:r><w:t>Heading</w:t></w:r><w:bookmarkEnd w:id="2"/><w:bookmarkEnd w:id="3"/><w:bookmarkEnd w:id="4"/><w:bookmarkEnd w:id="5"/><w:bookmarkEnd w:id="6"/></w:p>"##,
        r##"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF Zeb</w:instrText></w:r><w:r><w:instrText xml:space="preserve">raPlan \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Plan</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"##,
        r##"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _TocZebraHeading \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"##,
        r##"<w:p><w:fldSimple w:instr=" NOTEREF ZEBRAPLAN \h "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>"##,
        r##"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK \l &quot;_Toc123456&quot; \o &quot;Zebra tip&quot; </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Go</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"##,
        r##"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK \l "_TocZebraHeading" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"##,
        r##"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF ZebraPlan </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Plan</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:instrText xml:space="preserve"> = "x" "Zebra yes" "no" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>no</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"##,
        r##"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SET ZebraVar "Zebra value" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ASK ZebraAsk "Zebra prompt" \d "Zebra default" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"##,
        r##"<w:p><w:bookmarkStart w:id="7" w:name="ZebraField"/><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="ZebraField"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>value</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="7"/></w:p>"##,
        r##"<w:p><w:hyperlink w:anchor="zebraplan"><w:r><w:t>back</w:t></w:r></w:hyperlink><w:hyperlink w:anchor="_GoBack"><w:r><w:t>last</w:t></w:r></w:hyperlink></w:p>"##,
    ]
    .concat();
    let header = format!(
        r##"<w:hdr {W}><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref99887 \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF ZebraPlan \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>"##
    );
    let footnotes = format!(
        r##"<w:footnotes {W}><w:footnote w:id="1"><w:p><w:hyperlink w:anchor="_TocZebraHeading"><w:r><w:t>see</w:t></w:r></w:hyperlink></w:p></w:footnote></w:footnotes>"##
    );
    docx(
        &body,
        &[
            ("word/header1.xml", header),
            ("word/footnotes.xml", footnotes),
        ],
    )
}

#[test]
fn bookmark_names_become_neutral_and_every_reference_still_resolves() {
    let source = bookmark_fixture();
    for random_characters in [false, true] {
        let parts = redacted_docx(&source, random_characters);
        assert_no_prose(&parts);
        let document = &parts
            .iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1;
        let names = attributes(document, "w:bookmarkStart", "w:name");
        assert_eq!(names.len(), 7);
        for kept in KEPT_BOOKMARKS {
            assert!(
                names.iter().any(|name| name == kept),
                "{kept} renamed: {names:?}"
            );
        }
        assert!(names.iter().filter(|name| name.starts_with('_')).count() == 5);
        let resolves = |target: &str| names.iter().any(|name| name.eq_ignore_ascii_case(target));
        let mut references = 0;
        for (path, xml) in &parts {
            for anchor in attributes(xml, "w:hyperlink", "w:anchor") {
                assert!(
                    resolves(&anchor),
                    "{path}: dangling anchor {anchor} in {names:?}"
                );
                references += 1;
            }
            for instruction in instructions(xml) {
                let Some(target) = bookmark_target(&instruction) else {
                    continue;
                };
                if instruction.trim_start().starts_with("SET")
                    || instruction.trim_start().starts_with("ASK")
                {
                    continue;
                }
                assert!(
                    resolves(&target),
                    "{path}: dangling `{instruction}` in {names:?}"
                );
                references += 1;
            }
        }
        assert_eq!(references, 11);
        for name in attributes(document, "w:name", "w:val") {
            assert!(resolves(&name), "form field {name} lost its bookmark");
        }
        let all: String = parts.iter().map(|(_, xml)| xml.as_str()).collect();
        assert!(all.contains("\\h"));
        assert!(all.contains("_Ref99887"));
    }
}

fn vml_fixture() -> Vec<u8> {
    let body = [
        r##"<w:p><w:r><w:pict>"##,
        r##"<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>"##,
        r##"<v:shapetype id="ZebraType" coordsize="21600,21600" o:spt="1" path="m,l,21600r21600,l21600,xe"/>"##,
        r##"<v:shape id="Zebra Box" o:spid="_x0000_s1026" type="#_x0000_t202" alt="Zebra alt text" title="Zebra title" style="position:absolute;width:100pt;height:50pt"><v:textbox style="mso-next-textbox:#ZebraLinked;mso-fit-shape-to-text:t"><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>"##,
        r##"<v:shape id="ZebraLinked" o:spid="ZebraSpid" type="#ZebraType" style="width:10pt;height:10pt"><v:imagedata o:title="Zebra image"/></v:shape>"##,
        r##"<v:line id="ZebraConnector" o:spid="_x0000_s1028" from="0,0" to="10pt,10pt" o:connectortype="straight"/>"##,
        r##"<o:rules v:ext="edit"><o:r id="V:Rule1" type="connector" idref="#ZebraConnector"><o:proxy start="" idref="#Zebra Box" connectloc="0"/><o:proxy end="" idref="#ZebraLinked" connectloc="0"/></o:r></o:rules>"##,
        r##"</w:pict></w:r></w:p>"##,
        r##"<w:p><w:r><w:object w:dxaOrig="200" w:dyaOrig="200"><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" path="m@4@5l@4@11@9@11@9@5xe"/><v:shape id="ZebraOle" o:spid="_x0000_i1025" type="#_x0000_t75" style="width:10pt;height:10pt" o:ole=""><v:imagedata o:title=""/></v:shape><o:OLEObject Type="Embed" ProgID="Package" ShapeID="ZebraOle" DrawAspect="Content" ObjectID="_1234567890"/></w:object></w:r></w:p>"##,
        r##"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="12700" cy="12700"/><wp:docPr id="5" name="Zebra Drawing" descr="Zebra description" title="Zebra caption"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Zebra pic" descr="Zebra pic description"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill/><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="Zebra Fallback" o:spid="_x0000_s1030" type="#_x0000_t202" alt="Zebra fallback alt" style="width:1pt;height:1pt"/></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>"##,
    ]
    .concat();
    docx(&body, &[])
}

fn assert_vml_references_resolve(xml: &str) {
    let mut ids = Vec::new();
    for element in ["v:shape", "v:line", "v:rect", "v:group"] {
        ids.extend(attributes(xml, element, "id"));
        ids.extend(attributes(xml, element, "o:spid"));
    }
    let shapetypes = attributes(xml, "v:shapetype", "id");
    let mut checked = 0;
    for element in ["v:shape", "v:line"] {
        for reference in attributes(xml, element, "type") {
            let target = reference.strip_prefix('#').unwrap();
            assert!(
                shapetypes.iter().any(|id| id == target),
                "type {reference} in {shapetypes:?}"
            );
            checked += 1;
        }
    }
    for (element, attribute) in [
        ("o:OLEObject", "ShapeID"),
        ("o:r", "idref"),
        ("o:proxy", "idref"),
        ("w:control", "w:shapeid"),
        ("p:oleObj", "spid"),
    ] {
        for reference in attributes(xml, element, attribute) {
            let target = reference.trim_start_matches('#');
            assert!(
                ids.iter().any(|id| id == target),
                "{element} {reference} in {ids:?}"
            );
            checked += 1;
        }
    }
    for style in attributes(xml, "v:textbox", "style") {
        if let Some(target) = style
            .split(';')
            .find_map(|declaration| declaration.strip_prefix("mso-next-textbox:#"))
        {
            assert!(
                ids.iter().any(|id| id == target),
                "textbox link {target} in {ids:?}"
            );
            checked += 1;
        }
    }
    assert!(checked > 0);
}

#[test]
fn vml_shape_ids_become_neutral_and_references_still_resolve() {
    let source = vml_fixture();
    for random_characters in [false, true] {
        let parts = redacted_docx(&source, random_characters);
        assert_no_prose(&parts);
        let document = &parts
            .iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1;
        for kept in [
            "id=\"_x0000_t202\"",
            "id=\"_x0000_t75\"",
            "o:spid=\"_x0000_s1026\"",
            "o:spid=\"_x0000_i1025\"",
            "ObjectID=\"_1234567890\"",
            "id=\"V:Rule1\"",
        ] {
            assert!(document.contains(kept), "{kept} lost: {document}");
        }
        assert_eq!(attributes(document, "v:shape", "type").len(), 4);
        assert_vml_references_resolve(document);
    }
}

#[test]
fn vml_identifiers_in_spreadsheet_and_presentation_parts_resolve() {
    let vml = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"/><v:shape id="Zebra Note" type="#_x0000_t202" alt="Zebra alt" style="visibility:hidden"><x:ClientData ObjectType="Note"><x:Row>0</x:Row></x:ClientData></v:shape><v:shape id="_x0000_s1025" type="#_x0000_t202"/></xml>"##;
    for random_characters in [false, true] {
        let xlsx = redact_parts(
            Format::Xlsx,
            &[("xl/drawings/vmldrawing1.vml", vml)],
            random_characters,
        );
        assert!(!xlsx[0].contains("Zebra"), "{}", xlsx[0]);
        assert!(xlsx[0].contains("id=\"_x0000_s1025\""));
        assert_vml_references_resolve(&xlsx[0]);

        let slide = r##"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:graphicFrame><p:oleObj spid="Zebra Note" name="Object" r:id="rId2" progId="Package"/></p:graphicFrame><p:controls><p:control spid="_x0000_s1025" name="Control" r:id="rId3"/></p:controls></p:spTree></p:cSld></p:sld>"##;
        let pptx = redact_parts(
            Format::Pptx,
            &[
                ("ppt/drawings/vmldrawing1.vml", vml),
                ("ppt/slides/slide1.xml", slide),
            ],
            random_characters,
        );
        assert!(!pptx.concat().contains("Zebra"));
        assert!(pptx[1].contains("spid=\"_x0000_s1025\""));
        assert_vml_references_resolve(&pptx.concat());

        let control = format!(
            r##"<w:document {W}><w:body><w:p><w:r><w:object><v:shape id="Zebra Control" o:spid="_x0000_i1026" type="#_x0000_t202"/><w:control r:id="rId4" w:name="Control" w:shapeid="Zebra Control"/></w:object></w:r></w:p></w:body></w:document>"##
        );
        let docx = redact_parts(
            Format::Docx,
            &[("word/document.xml", &control)],
            random_characters,
        );
        assert!(!docx[0].contains("Zebra"), "{}", docx[0]);
        let ids = attributes(&docx[0], "v:shape", "id");
        assert_eq!(attributes(&docx[0], "w:control", "w:shapeid"), ids);
    }
}

fn theme(names: [&str; 5]) -> String {
    let [theme, colors, fonts, effects, family] = names;
    let color = |name: &str| format!(r##"<a:{name}><a:srgbClr val="000000"/></a:{name}>"##);
    let colors_xml: String = [
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ]
    .iter()
    .map(|name| color(name))
    .collect();
    format!(
        r##"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="{theme}"><a:themeElements><a:clrScheme name="{colors}">{colors_xml}</a:clrScheme><a:fontScheme name="{fonts}"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="{effects}"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:custClrLst><a:custClr name="Zebra Blue"><a:srgbClr val="0000FF"/></a:custClr></a:custClrLst><a:extLst><a:ext uri="{{05A4C25C-085E-4340-85A3-A5531E510DB2}}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="{family}" id="{{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}}" vid="{{4A3C46E8-61CC-4603-A589-7422A47A8E4A}}"/></a:ext></a:extLst></a:theme>"##
    )
}

#[test]
fn custom_theme_names_become_neutral_and_built_in_names_are_kept() {
    let custom = theme([
        "Zebra Corporate",
        "Zebra Colors",
        "Zebra Fonts",
        "Zebra Effects",
        "Zebra Family",
    ]);
    let builtin = theme([
        "Office Theme",
        "Office",
        "Office 2013 - 2022",
        "Office",
        "Office Theme",
    ]);
    for random_characters in [false, true] {
        for (format, path) in [
            (Format::Docx, "word/theme/theme1.xml"),
            (Format::Xlsx, "xl/theme/theme1.xml"),
            (Format::Pptx, "ppt/theme/theme1.xml"),
        ] {
            let output = redact_parts(
                format,
                &[(path, &custom), (path, &builtin)],
                random_characters,
            );
            assert!(!output[0].contains("Zebra"), "{format}: {}", output[0]);
            for (element, attribute) in [
                ("a:theme", "name"),
                ("a:clrScheme", "name"),
                ("a:fontScheme", "name"),
                ("a:fmtScheme", "name"),
                ("thm15:themeFamily", "name"),
            ] {
                let value = attributes(&output[0], element, attribute);
                assert!(
                    value.len() == 1 && value[0].chars().all(|c| c.is_ascii_alphanumeric()),
                    "{format} {element}: {value:?}"
                );
            }
            for kept in [
                "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}",
                "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}",
            ] {
                assert!(output[0].contains(kept));
            }
            for (element, expected) in [
                ("a:theme", "Office Theme"),
                ("a:clrScheme", "Office"),
                ("a:fontScheme", "Office 2013 - 2022"),
                ("a:fmtScheme", "Office"),
                ("thm15:themeFamily", "Office Theme"),
            ] {
                assert_eq!(
                    attributes(&output[1], element, "name"),
                    [expected],
                    "{format}"
                );
            }
            assert!(!output[1].contains("Zebra"));
        }
    }
    let package = docx("<w:p/>", &[("word/theme/theme1.xml", custom.clone())]);
    assert_no_prose(&redacted_docx(&package, true));
}
