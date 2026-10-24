use std::collections::{BTreeMap, HashSet};

use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use ooxml_redact::{Format, RedactionOptions, redact, redact_with_options, redact_with_report};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};

const FIXTURE: &[u8] = include_bytes!("fixtures/redaction-integrity.docx");

#[test]
fn random_characters_keep_style_and_relationship_mappings_consistent() {
    let default = parts(&redact(FIXTURE, Format::Docx).unwrap());
    let random = parts(
        &redact_with_options(
            FIXTURE,
            Format::Docx,
            &RedactionOptions {
                random_characters: true,
                ..Default::default()
            },
        )
        .unwrap(),
    );
    for (path, bytes) in &default {
        if path.ends_with(".rels") {
            assert_eq!(&random[path], bytes, "{path}");
        }
    }
    for path in ["word/styles.xml", "word/document.xml"] {
        for tag in [
            "style",
            "basedOn",
            "next",
            "link",
            "pStyle",
            "rStyle",
            "tblStyle",
            "numStyleLink",
            "styleLink",
        ] {
            assert_eq!(
                elements(&default[path], tag),
                elements(&random[path], tag),
                "{path}: {tag}"
            );
        }
    }
    parse_docx_s9_wire(
        &redact_with_options(
            FIXTURE,
            Format::Docx,
            &RedactionOptions {
                random_characters: true,
                ..Default::default()
            },
        )
        .unwrap(),
        S9ParseOptions::default(),
    )
    .unwrap();
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn elements(bytes: &[u8], tag: &str) -> Vec<BTreeMap<String, String>> {
    let mut reader = Reader::from_reader(bytes);
    let mut values = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(start) | Event::Empty(start)
                if tag == "*" || start.local_name().as_ref() == tag.as_bytes() =>
            {
                values.push(
                    start
                        .attributes()
                        .map(|attribute| {
                            let attribute = attribute.unwrap();
                            (
                                String::from_utf8_lossy(attribute.key.local_name().as_ref())
                                    .into_owned(),
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .unwrap()
                                    .into_owned(),
                            )
                        })
                        .collect(),
                );
            }
            Event::Eof => break,
            _ => {}
        }
    }
    values
}

#[test]
fn custom_xml_parts_relationships_and_overrides_are_removed() {
    let (bytes, report) = redact_with_report(FIXTURE, Format::Docx).unwrap();
    let source = parts(FIXTURE);
    let output = parts(&bytes);
    assert_eq!(
        report.custom_xml_parts,
        source
            .keys()
            .filter(|path| path.starts_with("customXml/"))
            .count()
    );
    assert!(!output.keys().any(|path| path.starts_with("customXml/")));
    for (path, bytes) in &output {
        if path.ends_with(".rels") || path == "[Content_Types].xml" {
            assert!(
                !String::from_utf8_lossy(bytes).contains("customXml/"),
                "{path}"
            );
        }
    }
    parse_docx_s9_wire(&bytes, S9ParseOptions::default()).unwrap();
}

#[test]
fn custom_styles_have_unique_names_and_all_references_resolve() {
    let bytes = redact(FIXTURE, Format::Docx).unwrap();
    parse_docx_s9_wire(&bytes, S9ParseOptions::default()).unwrap();
    let output = parts(&bytes);
    let styles = elements(&output["word/styles.xml"], "style");
    let names: Vec<_> = elements(&output["word/styles.xml"], "name")
        .into_iter()
        .map(|attributes| attributes["val"].to_lowercase())
        .collect();
    assert_eq!(names.len(), names.iter().collect::<HashSet<_>>().len());
    assert!(names.contains(&"normal".to_owned()));
    assert!(names.contains(&"heading 1".to_owned()));
    let ids: HashSet<_> = styles
        .iter()
        .map(|attributes| attributes["styleId"].as_str())
        .collect();
    assert!(ids.contains("Normal"));
    assert!(ids.contains("Heading1"));
    let custom: Vec<_> = styles
        .iter()
        .filter(|attributes| {
            attributes
                .get("customStyle")
                .is_some_and(|value| value == "1")
        })
        .collect();
    assert_eq!(custom.len(), 2);
    assert!(
        custom
            .iter()
            .all(|attributes| attributes["styleId"].starts_with("RedactedStyle"))
    );
    for reference in elements(&output["word/document.xml"], "pStyle") {
        assert!(ids.contains(reference["val"].as_str()));
    }
    for reference in [
        "basedOn",
        "next",
        "link",
        "pStyle",
        "rStyle",
        "tblStyle",
        "numStyleLink",
        "styleLink",
    ] {
        for attributes in elements(&output["word/styles.xml"], reference) {
            assert!(ids.contains(attributes["val"].as_str()));
        }
    }
}

#[test]
fn embedded_gfxdata_is_removed_without_removing_the_shape() {
    let source = parts(FIXTURE);
    let bytes = redact(FIXTURE, Format::Docx).unwrap();
    let output = parts(&bytes);
    let original_shapes = elements(&source["word/document.xml"], "shape");
    assert!(
        original_shapes
            .iter()
            .any(|shape| shape["gfxdata"].starts_with("UEsDB"))
    );
    let redacted_shapes = elements(&output["word/document.xml"], "shape");
    assert_eq!(original_shapes.len(), redacted_shapes.len());
    for (original, redacted) in original_shapes.iter().zip(&redacted_shapes) {
        assert!(!redacted.contains_key("gfxdata"));
        assert_eq!(original["style"], redacted["style"]);
        assert!(redacted["id"].starts_with("Shape"));
    }
    assert_ne!(
        source["word/media/image1.png"],
        output["word/media/image1.png"]
    );
    for (path, data) in &output {
        if path.ends_with(".xml") || path.ends_with(".rels") {
            let text = String::from_utf8_lossy(data);
            assert!(!text.contains("SECRET_"), "unmasked sentinel in {path}");
            assert!(!text.contains("gfxdata"), "embedded payload in {path}");
        }
    }
}

fn updated_fixture(updates: &[(&str, String)]) -> Vec<u8> {
    let mut source = parts(FIXTURE);
    for (path, xml) in updates {
        source.insert((*path).to_owned(), xml.as_bytes().to_vec());
    }
    ooxml_opc::rezip_parts(&source.into_iter().collect::<Vec<_>>()).unwrap()
}

#[test]
fn unmarked_styles_and_builtin_aliases_do_not_leak_metadata() {
    for marker in [
        "",
        r#"w:customStyle="0""#,
        r#"w:customStyle="false""#,
        r#"w:default="1""#,
    ] {
        let styles = format!(
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:aliases w:val="SECRET_BUILTIN_ALIAS"/></w:style>
<w:style w:styleId="SECRET_UNMARKED_ID" {marker}><w:name w:val="SECRET_UNMARKED_NAME"/><w:aliases w:val="SECRET_UNMARKED_ALIAS"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:styleId="SECRET_BUILTIN_ID"><w:name w:val="heading 1"/></w:style></w:styles>"#
        );
        let document = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="SECRET_UNMARKED_ID"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#;
        let source = updated_fixture(&[
            ("word/styles.xml", styles),
            ("word/document.xml", document.to_owned()),
        ]);
        let output = parts(&redact(&source, Format::Docx).unwrap());
        assert!(!String::from_utf8_lossy(&output["word/styles.xml"]).contains("SECRET_"));
        assert!(!String::from_utf8_lossy(&output["word/document.xml"]).contains("SECRET_"));
        let definitions = elements(&output["word/styles.xml"], "style");
        let reference = &elements(&output["word/document.xml"], "pStyle")[0]["val"];
        assert_eq!(reference, &definitions[1]["styleId"]);
        assert_eq!(definitions[0]["styleId"], "Normal");
        assert_eq!(
            elements(&output["word/styles.xml"], "name")[2]["val"],
            "heading 1"
        );
        if marker.contains("default") {
            assert_eq!(definitions[1]["default"], "1");
        }
    }
}

#[test]
fn glossary_references_use_their_own_style_definitions() {
    let glossary_styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:customStyle="1" w:styleId="BBBB2222"><w:name w:val="Glossary B"/></w:style><w:style w:customStyle="1" w:styleId="AAAA1111"><w:name w:val="Glossary A"/></w:style></w:styles>"#;
    let glossary = r#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartBody><w:p><w:pPr><w:pStyle w:val="AAAA1111"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#;
    let original = parts(FIXTURE);
    let relationships = String::from_utf8(original["word/_rels/document.xml.rels"].clone()).unwrap().replace("</Relationships>", r#"<Relationship Id="rIdGlossary" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/></Relationships>"#);
    let types = String::from_utf8(original["[Content_Types].xml"].clone()).unwrap().replace("</Types>", r#"<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/><Override PartName="/word/glossary/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#);
    let glossary_relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let source = updated_fixture(&[
        ("[Content_Types].xml", types),
        ("word/_rels/document.xml.rels", relationships),
        (
            "word/glossary/_rels/document.xml.rels",
            glossary_relationships.to_owned(),
        ),
        ("word/glossary/styles.xml", glossary_styles.to_owned()),
        ("word/glossary/document.xml", glossary.to_owned()),
    ]);
    let output = parts(&redact(&source, Format::Docx).unwrap());
    let glossary_definitions = elements(&output["word/glossary/styles.xml"], "style");
    let glossary_reference = &elements(&output["word/glossary/document.xml"], "pStyle")[0]["val"];
    let main_reference = &elements(&output["word/document.xml"], "pStyle")[0]["val"];
    assert_eq!(glossary_reference, &glossary_definitions[1]["styleId"]);
    assert_ne!(glossary_reference, main_reference);
}
