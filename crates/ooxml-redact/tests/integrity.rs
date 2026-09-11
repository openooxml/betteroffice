use std::collections::{BTreeMap, HashSet};

use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use ooxml_redact::{Format, redact};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};

const FIXTURE: &[u8] = include_bytes!("fixtures/redaction-integrity.docx");

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn elements(bytes: &[u8], tag: &str) -> Vec<BTreeMap<String, String>> {
    let mut reader = Reader::from_reader(bytes);
    let mut values = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(start) | Event::Empty(start)
                if start.local_name().as_ref() == tag.as_bytes() =>
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
fn custom_xml_relationships_keep_their_types_and_existing_targets() {
    let source = parts(FIXTURE);
    let output = parts(&redact(FIXTURE, Format::Docx).unwrap());
    for index in 1..=4 {
        let path = format!("customXml/_rels/item{index}.xml.rels");
        let original = elements(&source[&path], "Relationship");
        let redacted = elements(&output[&path], "Relationship");
        assert_eq!(original, redacted);
        for relationship in redacted {
            assert!(output.contains_key(&format!("customXml/{}", relationship["Target"])));
        }
    }
}

#[test]
fn custom_xml_external_relationships_are_still_redacted() {
    let mut source = parts(FIXTURE);
    let path = "customXml/_rels/item1.xml.rels";
    let xml = String::from_utf8(source[path].clone()).unwrap().replace(
        "</Relationships>",
        r#"<Relationship Id="private" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://secret.example/private" TargetMode="External"/></Relationships>"#,
    );
    source.insert(path.into(), xml.into_bytes());
    let input = ooxml_opc::rezip_parts(&source.into_iter().collect::<Vec<_>>()).unwrap();
    let output = parts(&redact(&input, Format::Docx).unwrap());
    let relationships = elements(&output[path], "Relationship");
    let external = relationships
        .iter()
        .find(|rel| rel["Id"] == "private")
        .unwrap();
    assert_eq!(external["Target"], "https://example.com");
    assert_eq!(external["TargetMode"], "External");
    assert!(external["Type"].ends_with("/hyperlink"));
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
        assert_eq!(original["id"], redacted["id"]);
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

#[test]
fn schema_structure_survives_without_literal_secrets() {
    let source = parts(FIXTURE);
    let output = parts(&redact(FIXTURE, Format::Docx).unwrap());
    let path = "customXml/item2.xml";
    for tag in [
        "schema",
        "element",
        "complexType",
        "simpleType",
        "restriction",
        "maxLength",
    ] {
        let original = elements(&source[path], tag);
        let redacted = elements(&output[path], tag);
        assert_eq!(original.len(), redacted.len(), "{tag}");
        for (before, after) in original.iter().zip(&redacted) {
            for key in [
                "name",
                "ref",
                "type",
                "base",
                "minOccurs",
                "maxOccurs",
                "nillable",
                "targetNamespace",
                "value",
            ] {
                if let Some(value) = before.get(key) {
                    assert_eq!(after.get(key), Some(value), "{tag}@{key}");
                }
            }
            assert!(!after.contains_key("default"));
            assert!(!after.contains_key("fixed") || tag == "maxLength");
        }
    }
    for tag in ["annotation", "enumeration", "pattern"] {
        assert!(!elements(&source[path], tag).is_empty());
        assert!(elements(&output[path], tag).is_empty());
    }
    assert!(!String::from_utf8_lossy(&output[path]).contains("SECRET_"));
}
