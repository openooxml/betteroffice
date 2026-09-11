use std::io::Write;
use std::process::{Command, Stdio};

use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use quick_xml::{Reader, events::Event};
use serde_json::json;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn package(body: &str, attributes: &str, header: Option<&str>) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="{MC}" xmlns:bofx="urn:foreign" xmlns:bofy="urn:other" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" {attributes}><w:body>{body}<w:sectPr><w:headerReference w:type="default" r:id="header"/></w:sectPr></w:body></w:document>"#
    );
    let mut parts = vec![
        ("word/document.xml".to_owned(), document.into_bytes()),
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="header" Type="{R}/header" Target="header1.xml"/><Relationship Id="footer" Type="{R}/footer" Target="footer1.xml"/></Relationships>"#).into_bytes()),
    ];
    if let Some(header) = header {
        for (root, part) in [("hdr", "header1"), ("ftr", "footer1")] {
            parts.push((format!("word/{part}.xml"), format!(r#"<w:{root} xmlns:w="{W}" xmlns:mc="{MC}" xmlns:bofx="urn:foreign" xmlns:bofy="urn:other" {attributes}>{header}</w:{root}>"#).into_bytes()));
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn roundtrip(bytes: &[u8]) -> Vec<u8> {
    let parsed = parse_docx_s9_wire(bytes, S9ParseOptions::default()).unwrap();
    let package = parsed.document.package;
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {
            "content": package.document.content,
            "finalSectionProperties": package.document.final_section_properties,
            "customRootBindings": package.document.custom_root_bindings,
            "comments": package.document.comments
        },
        "headerEntries": package.header_entries.unwrap_or_default(),
        "footerEntries": package.footer_entries.unwrap_or_default(),
        "footnotes": package.footnotes.unwrap_or_default(),
        "endnotes": package.endnotes.unwrap_or_default(),
        "footnoteSeparators": package.footnote_separators.unwrap_or_default(),
        "endnoteSeparators": package.endnote_separators.unwrap_or_default(),
        "relationshipEntries": package.relationship_entries,
        "options": {"updateModifiedDate":false}
    }))
    .unwrap();
    write_docx_s13(request, bytes).unwrap()
}

fn strict_xml(bytes: &[u8]) {
    let mut reader = Reader::from_reader(bytes);
    loop {
        match reader.read_event().unwrap() {
            Event::Start(start) | Event::Empty(start) => {
                for attribute in start.attributes() {
                    let attribute = attribute.unwrap();
                    assert!(!attribute.value.contains(&b'<'));
                    #[allow(deprecated)]
                    attribute
                        .decode_and_unescape_value(reader.decoder())
                        .unwrap();
                }
            }
            Event::Text(text) => {
                quick_xml::escape::unescape(&text.decode().unwrap()).unwrap();
            }
            Event::GeneralRef(reference) => {
                assert!(
                    reference.is_char_ref()
                        || quick_xml::escape::resolve_predefined_entity(
                            &reference.decode().unwrap()
                        )
                        .is_some()
                );
            }
            Event::Eof => break,
            _ => {}
        }
    }
}

fn python_check(bytes: &[u8], assertions: &str) {
    let script = format!(
        "import sys,io,zipfile,xml.etree.ElementTree as E\nz=zipfile.ZipFile(io.BytesIO(sys.stdin.buffer.read()))\n{assertions}"
    );
    let mut child = Command::new("python3")
        .args(["-c", &script])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    child.stdin.take().unwrap().write_all(bytes).unwrap();
    let output = child.wait_with_output().unwrap();
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
}

#[test]
fn foreign_fragment_values_roundtrip_in_inline_block_and_cell_xml() {
    let fragment = r#"<bofx:marker bofx:q="a&amp;&lt;&gt;&quot;b">a&amp;&lt;&gt;"b</bofx:marker>"#;
    for body in [
        format!("<w:p>{fragment}</w:p>"),
        fragment.to_owned(),
        format!(
            "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc>{fragment}<w:p/></w:tc></w:tr></w:tbl>"
        ),
    ] {
        let saved = roundtrip(&package(&body, "", None));
        for (name, xml) in ooxml_opc::unzip_parts(&saved).unwrap() {
            if name == "word/document.xml" {
                strict_xml(&xml);
            }
        }
        python_check(
            &saved,
            "root=E.fromstring(z.read('word/document.xml'))\nnode=next(root.iter('{urn:foreign}marker'))\nassert node.attrib['{urn:foreign}q']=='a&<>\\\"b',node.attrib\nassert node.text=='a&<>\\\"b',node.text",
        );
    }
}

#[test]
fn omml_angle_bracket_delimiters_roundtrip_as_valid_xml() {
    let saved = roundtrip(&package(
        r#"<w:p><m:oMath><m:d><m:dPr><m:begChr m:val="&lt;"/><m:endChr m:val="&gt;"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d></m:oMath></w:p>"#,
        "",
        None,
    ));
    python_check(
        &saved,
        "root=E.fromstring(z.read('word/document.xml'))\nm='{http://schemas.openxmlformats.org/officeDocument/2006/math}'\nassert next(root.iter(m+'begChr')).attrib[m+'val']=='<'\nassert next(root.iter(m+'endChr')).attrib[m+'val']=='>'",
    );
}

#[test]
fn authored_custom_ignorable_prefixes_survive_all_story_roots() {
    let saved = roundtrip(&package(
        r#"<w:p><bofx:marker/></w:p>"#,
        r#"mc:Ignorable="w14 bofx""#,
        Some(r#"<w:p><bofx:marker/></w:p>"#),
    ));
    python_check(
        &saved,
        "for part in ['document.xml','header1.xml','footer1.xml']:\n root=E.fromstring(z.read('word/'+part))\n prefixes=root.attrib['{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable'].split()\n assert 'bofx' in prefixes,(part,prefixes)\n assert 'bofy' not in prefixes,(part,prefixes)\n assert len(prefixes)==len(set(prefixes))",
    );
}

#[test]
fn comprehensive_custom_ignorable_prefix_survives_save() {
    let saved = roundtrip(include_bytes!(
        "../../betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx"
    ));
    python_check(
        &saved,
        "root=E.fromstring(z.read('word/document.xml'))\nassert 'bofx' in root.attrib['{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable'].split()\nassert next(root.iter('{urn:betteroffice-fixture-x}marker'),None) is not None",
    );
}

#[test]
fn non_modelled_root_attributes_survive_document_header_and_footer() {
    let saved = roundtrip(&package(
        "<w:p/>",
        r#"w:conformance="strict" bofy:flag="a&amp;&lt;&gt;&quot;b""#,
        Some("<w:p/>"),
    ));
    python_check(
        &saved,
        "for part in ['document.xml','header1.xml','footer1.xml']:\n root=E.fromstring(z.read('word/'+part))\n assert root.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}conformance']=='strict'\n assert root.attrib['{urn:other}flag']=='a&<>\\\"b'",
    );
}

#[test]
fn raw_note_blocks_keep_inherited_namespace_bindings() {
    let mut parts = ooxml_opc::unzip_parts(&package("<w:p/>", "", None)).unwrap();
    for (plural, singular) in [("footnotes", "footnote"), ("endnotes", "endnote")] {
        parts.push((format!("word/{plural}.xml"), format!(r#"<w:{plural} xmlns:w="{W}" xmlns:bofx="urn:foreign"><w:{singular} w:id="1"><bofx:block/><w:p/></w:{singular}></w:{plural}>"#).into_bytes()));
    }
    let saved = roundtrip(&ooxml_opc::rezip_parts(&parts).unwrap());
    python_check(
        &saved,
        "for part in ['footnotes.xml','endnotes.xml']:\n root=E.fromstring(z.read('word/'+part))\n assert next(root.iter('{urn:foreign}block'),None) is not None",
    );
}
