use std::collections::{BTreeMap, HashSet};

use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use ooxml_redact::{Format, redact, redact_with_report};
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

#[test]
fn element_derivation_constraints_survive_schema_redaction() {
    for (block, final_value) in [
        ("#all", "#all"),
        ("extension", "restriction"),
        (
            "extension restriction substitution",
            "extension restriction",
        ),
    ] {
        let schema = format!(
            r#"<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:element name="Record" type="xs:string" block="{block}" final="{final_value}"/></xs:schema>"#
        );
        let source = updated_fixture(&[("customXml/item2.xml", schema)]);
        let output = parts(&redact(&source, Format::Docx).unwrap());
        let element = &elements(&output["customXml/item2.xml"], "element")[0];
        assert_eq!(element["block"], block);
        assert_eq!(element["final"], final_value);
    }
}

#[test]
fn gfxdata_removal_uses_the_resolved_office_namespace() {
    let xml = r#"<root xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:q="urn:private" xmlns:gfxdata="urn:kept"><record o:gfxdata="SECRET_PAYLOAD" q:gfxdata="SECRET_FOREIGN" gfxdata="SECRET_UNQUALIFIED"/><nested xmlns:o="urn:private"><record o:gfxdata="SECRET_REBOUND"/></nested></root>"#;
    let source = updated_fixture(&[("customXml/item4.xml", xml.to_owned())]);
    let output = parts(&redact(&source, Format::Docx).unwrap());
    let text = String::from_utf8(output["customXml/item4.xml"].clone()).unwrap();
    assert!(!text.contains("SECRET_"));
    assert!(text.contains(r#"xmlns:gfxdata="urn:kept""#));
    let records = elements(&output["customXml/item4.xml"], "record");
    assert_eq!(records.len(), 2);
    assert!(text.contains(r#"q:gfxdata="xxxxxxxxxxxxxx""#));
    assert!(text.contains(r#" gfxdata="xxxxxxxxxxxxxxxxxx""#));
    assert!(text.contains(r#"o:gfxdata="xxxxxxxxxxxxxx""#));
    assert_eq!(text.matches("o:gfxdata=").count(), 1);
    assert!(!text.contains("SECRET_PAYLOAD"));
}

#[test]
fn removed_schema_subtrees_count_comments_and_processing_instructions() {
    let schema = r#"<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><!--outer--><?outer hidden?><xs:annotation><!--annotation--><?annotation hidden?><xs:documentation><!--nested--><?nested hidden?><![CDATA[SECRET_CDATA]]></xs:documentation></xs:annotation><xs:element name="Record" type="xs:string"/></xs:schema>"#;
    let source = updated_fixture(&[("customXml/item2.xml", schema.to_owned())]);
    let (output, report) = redact_with_report(&source, Format::Docx).unwrap();
    assert_eq!(report.xml_comments, 6);
    let output = parts(&output);
    let schema = String::from_utf8_lossy(&output["customXml/item2.xml"]);
    assert!(!schema.contains("hidden"));
    assert!(!schema.contains("SECRET_CDATA"));
    assert_eq!(elements(schema.as_bytes(), "element").len(), 1);
}
